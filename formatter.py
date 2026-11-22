"""Mise en forme automatique d'un mémoire Word (.docx) – texte arabe (RTL) et latin (LTR).

Pipeline : nettoyage -> détection (page de garde, préliminaires, titres) -> styles
-> sommaire réel (champ TOC pré-rempli) -> sections + numérotation des pages.
"""
import copy
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import unicodedata

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt

# --------------------------------------------------------------------------- règles

DEFAULT_RULES = {
    "page": {"width_cm": 21.0, "height_cm": 29.7,
             "margin_top_cm": 2.5, "margin_bottom_cm": 2.5,
             "margin_left_cm": 3.0, "margin_right_cm": 2.0},
    "body": {"font_latin": "Times New Roman", "size_latin": 12,
             "font_arabic": "Traditional Arabic", "size_arabic": 14,
             "line_spacing": 1.5, "space_after_pt": 6,
             "first_line_indent_cm": 1.0, "align": "justify"},
    "headings": {
        "1": {"size_latin": 16, "size_arabic": 18, "bold": True, "align": "center",
              "space_before_pt": 24, "space_after_pt": 18, "caps": False, "page_break_before": True},
        "2": {"size_latin": 14, "size_arabic": 16, "bold": True, "align": "start",
              "space_before_pt": 18, "space_after_pt": 8, "caps": False},
        "3": {"size_latin": 13, "size_arabic": 15, "bold": True, "align": "start",
              "space_before_pt": 12, "space_after_pt": 6, "caps": False},
        "4": {"size_latin": 12, "size_arabic": 14, "bold": True, "align": "start",
              "space_before_pt": 12, "space_after_pt": 6, "caps": False},
    },
    "front_titles": {"size_latin": 16, "size_arabic": 18, "bold": True, "align": "center",
                     "page_break_before": True, "in_toc": False},
    "caption": {"size_latin": 11, "size_arabic": 12, "bold": False, "italic": True, "align": "center"},
    "table": {"size_latin": 11, "size_arabic": 12},
    "toc": {"enabled": True, "levels": 3, "title_latin": "Sommaire", "title_arabic": "فهرس المحتويات"},
    "page_numbers": {"enabled": True, "position": "center",
                     "front_format": "lowerRoman",   # lowerRoman | arabicAbjad | decimal | none
                     "body_format": "decimal"},
    "cover": {"restyle": False},
    "language": {"latin": "fr-FR", "arabic": "ar-DZ"},
}

HEADING_MAX_LEVEL = 4


FLAT_KEYS = {  # clés « à plat » produites par l'extraction IA du cahier des charges -> chemin imbriqué
    "font_latin": ("body", "font_latin"), "size_latin": ("body", "size_latin"),
    "font_arabic": ("body", "font_arabic"), "size_arabic": ("body", "size_arabic"),
    "line_spacing": ("body", "line_spacing"), "first_line_indent_cm": ("body", "first_line_indent_cm"),
    "space_after_pt": ("body", "space_after_pt"), "text_align": ("body", "align"),
    "margin_top_cm": ("page", "margin_top_cm"), "margin_bottom_cm": ("page", "margin_bottom_cm"),
    "margin_left_cm": ("page", "margin_left_cm"), "margin_right_cm": ("page", "margin_right_cm"),
    "toc_levels": ("toc", "levels"), "toc_title_latin": ("toc", "title_latin"),
    "toc_title_arabic": ("toc", "title_arabic"),
    "page_number_position": ("page_numbers", "position"), "front_numbering": ("page_numbers", "front_format"),
    "body_numbering": ("page_numbers", "body_format"),
}


def normalize_rules(raw):
    """Accepte des règles imbriquées OU à plat ; ignore les valeurs vides/nulles."""
    out = {}
    ZERO_OK = {"first_line_indent_cm", "space_after_pt"}
    for k, v in (raw or {}).items():
        if v is None or v == "" or v == []:
            continue
        # un nombre à 0 venant de l'extraction IA signifie « non précisé » (sauf retrait / espacement)
        if isinstance(v, (int, float)) and not isinstance(v, bool) and v <= 0 and k not in ZERO_OK \
                and (k in FLAT_KEYS or k.startswith("h")):
            continue
        if k in FLAT_KEYS:
            sec, key = FLAT_KEYS[k]
            out.setdefault(sec, {})[key] = v
        elif k in ("h1_size", "h2_size", "h3_size", "h4_size"):
            n = k[1]
            hs = out.setdefault("headings", {}).setdefault(n, {})
            hs["size_latin"] = v
            hs["size_arabic"] = v
        elif isinstance(v, dict):
            out[k] = {kk: vv for kk, vv in v.items() if vv not in (None, "", [])}
        else:
            out[k] = v
    return out


def _coerce(base, override):
    """Convertit les nombres reçus en texte (« 12 », « 2,5 ») selon le type par défaut."""
    for k, v in list(override.items()):
        if isinstance(v, dict):
            _coerce(base.get(k, {}) if isinstance(base.get(k), dict) else {}, v)
        elif isinstance(base.get(k), (int, float)) and not isinstance(base.get(k), bool) and isinstance(v, str):
            try:
                override[k] = float(v.replace(",", ".").strip())
            except ValueError:
                del override[k]
    return override


def deep_merge(base, override):
    out = copy.deepcopy(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


# --------------------------------------------------------------------------- texte

AR_CHAR = re.compile(r"[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]")
AR_LETTER = re.compile(r"[ء-يٱ-ۓݐ-ݿﭐ-﷿ﹰ-﻿]")
LAT_LETTER = re.compile(r"[A-Za-zÀ-ɏ]")


def norm(s):
    """Minuscules, sans accents ni diacritiques arabes, alef/ya/ta marbouta unifiés."""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("ـ", "").replace("ة", "ه").replace("ى", "ي")
    return re.sub(r"\s+", " ", s).strip().lower()


def rx(p):
    """Compile un motif déjà écrit en minuscules, en appliquant la même normalisation que `norm`
    (sans passer en minuscules : \\S ne doit pas devenir \\s)."""
    p = unicodedata.normalize("NFKD", p)
    p = "".join(c for c in p if unicodedata.category(c) != "Mn")
    p = p.replace("\u0640", "").replace("ة", "ه").replace("ى", "ي")
    return re.compile(p)


def char_class(c):
    if AR_LETTER.match(c):
        return "ar"
    if LAT_LETTER.match(c):
        return "lat"
    return "neu"


def dominant_dir(text, default="ltr"):
    a = len(AR_LETTER.findall(text))
    l = len(LAT_LETTER.findall(text))
    if a == 0 and l == 0:
        return default
    return "rtl" if a > l else "ltr"


FRONT_PATTERNS = {
    "dedication": rx(r"^((ال)?اهداء|dedicaces?)\s*[:\-]?$"),
    "ack": rx(r"^(remerciements?|شكر وتقدير|شكر و تقدير|شكر وعرفان|شكر و عرفان|كلمه شكر|تشكرات|شكر|الشكر|avant[- ]propos)\s*[:\-]?$"),
    "abstract": rx(r"^(resume|abstract|ملخص( الدراسه| البحث| المذكره)?|الملخص|مستخلص|resume en francais|ملخص بالعربيه)\s*[:\-]?$"),
    "lists": rx(r"^(liste des (tableaux|figures|abreviations|sigles|acronymes|symboles|annexes|illustrations)|"
                r"table des (figures|tableaux|illustrations)|glossaire|"
                r"قايمه (الجداول|الاشكال|المختصرات|الرموز|الصور|الملاحق|الاختصارات|الاختصارات والرموز)|فهرس (الجداول|الاشكال|الصور|الملاحق))\s*[:\-]?$"),
    "toc": rx(r"^(sommaire|table des matieres|فهرس المحتويات|جدول المحتويات|المحتويات|الفهرس|فهرس)\s*[:\-]?$"),
}
MAIN_PATTERN = rx(r"^(introduction( generale)?|conclusion( generale)?|bibliographie|references( bibliographiques)?|"
                  r"webographie|annexes?|perspectives|مقدمه( عامه)?|المقدمه( العامه)?|خاتمه( عامه)?|الخاتمه( العامه)?|"
                  r"قايمه (المراجع|المصادر والمراجع|المصادر)|المراجع|المصادر والمراجع|المصادر|الملاحق|ملحق( \d+)?)\s*[:\-]?$")
CHAPTER_PATTERN = rx(r"^(chapitre|chapter|partie|الفصل|الباب|الجزء)\s+\S+")
CHAPTER_LABEL_ONLY = rx(r"^(chapitre|chapter|partie|الفصل|الباب|الجزء)\s+\S+(\s+\S+)?\s*[:\-]?$")
SECTION_AR = [(rx(r"^المبحث\s"), 2), (rx(r"^المطلب\s"), 3), (rx(r"^الفرع\s"), 4)]
SUBCHAPTER_AR = rx(r"^(تمهيد|خلاصه الفصل|ملخص الفصل|خاتمه الفصل)\s*[:\-\.]?$")
ANNEX_PATTERN = rx(r"^(الملاحق|ملحق( \d+)?|annexes?)\s*[:\-]?$")
BIBLIO_PART1 = rx(r"^قايمه المصادر\s*$")
BIBLIO_PART2 = rx(r"^والمراجع\s*$")
TOC_TABLE_HDR = rx(r"(العنوان|الموضوع|الصفحه|titre|page)")
ORDINAL_AR = rx(r"^(اولا|ثانيا|ثالثا|رابعا|خامسا|سادسا|سابعا|ثامنا|تاسعا|عاشرا)\s*[:\-–]")
CAPTION_PATTERN = rx(r"^(tableau|figure|fig\.|graphique|schema|photo|جدول|الجدول|شكل|الشكل|صوره|الصوره|رسم بياني|مخطط)\s*[\d٠-٩]")
TOC_ENTRY = re.compile(r"(\.{3,}|…+|\t|\s{3,})\s*[\divxlcIVXLC٠-٩]+\s*$")
NUM4 = re.compile(r"^\d+\.\d+\.\d+\.\d+\.?\s*\S")
NUM3 = re.compile(r"^\d+\.\d+\.\d+\.?\s*\S")
NUM2 = re.compile(r"^\d+\.\d+\.?\s*\S")
NUM1 = re.compile(r"^\d+\s*[\.\-\)]\s*\S")
ROMAN = re.compile(r"^[IVX]{1,5}\s*[\.\-\)]\s+\S")
LETTER = re.compile(r"^[A-Z]\s*[\.\)]\s+\S")


# --------------------------------------------------------------------------- XML helpers

RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
             "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
             "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
SECT_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
              "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
              "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"]


def _local(el):
    return el.tag.split("}")[1]


def ordered_set(parent, el, order):
    """Insère `el` dans `parent` à la bonne place du schéma OOXML (remplace l'élément existant de même nom)."""
    name = _local(el)
    for old in parent.findall(el.tag):
        parent.remove(old)
    idx = order.index(name) if name in order else len(order)
    for child in parent:
        cname = _local(child)
        if cname in order and order.index(cname) > idx:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def w_el(tag, **attrs):
    el = OxmlElement("w:" + tag)
    for k, v in attrs.items():
        el.set(qn("w:" + k), str(v))
    return el


def get_rpr(r_el):
    rpr = r_el.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        r_el.insert(0, rpr)
    return rpr


def get_ppr(p_el):
    ppr = p_el.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p_el.insert(0, ppr)
    return ppr


def fonts_el(latin, arabic):
    return w_el("rFonts", ascii=latin, hAnsi=latin, cs=arabic, eastAsia=latin)


def set_rpr_fonts(rpr, latin, arabic, size_latin, size_arabic, bold=None, italic=None):
    ordered_set(rpr, fonts_el(latin, arabic), RPR_ORDER)
    ordered_set(rpr, w_el("sz", val=int(size_latin * 2)), RPR_ORDER)
    ordered_set(rpr, w_el("szCs", val=int(size_arabic * 2)), RPR_ORDER)
    if bold is not None:
        for tag in ("b", "bCs"):
            if bold:
                ordered_set(rpr, w_el(tag), RPR_ORDER)
            else:
                for old in rpr.findall(qn("w:" + tag)):
                    rpr.remove(old)
    if italic is not None:
        for tag in ("i", "iCs"):
            if italic:
                ordered_set(rpr, w_el(tag), RPR_ORDER)
            else:
                for old in rpr.findall(qn("w:" + tag)):
                    rpr.remove(old)


def strip_rpr(rpr, tags):
    for t in tags:
        for old in rpr.findall(qn("w:" + t)):
            rpr.remove(old)


def set_bidi(p_el, rtl):
    ppr = get_ppr(p_el)
    for old in ppr.findall(qn("w:bidi")):
        ppr.remove(old)
    if rtl:
        ordered_set(ppr, w_el("bidi"), PPR_ORDER)


def set_jc(p_el, value):
    """left/right se comprennent 'début/fin' de ligne dans un paragraphe bidi (comportement de Word)."""
    ppr = get_ppr(p_el)
    ordered_set(ppr, w_el("jc", val=value), PPR_ORDER)


ALIGN_MAP = {"start": "left", "end": "right", "center": "center", "justify": "both", "both": "both"}


def p_text(p_el):
    return "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip()


def has_page_break(p_el):
    for br in p_el.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    ppr = p_el.find(qn("w:pPr"))
    return ppr is not None and ppr.find(qn("w:pageBreakBefore")) is not None


def has_drawing(p_el):
    return any(True for _ in p_el.iter(qn("w:drawing"))) or any(True for _ in p_el.iter(qn("w:pict")))


def style_name(doc, p_el):
    ppr = p_el.find(qn("w:pPr"))
    if ppr is None:
        return ""
    ps = ppr.find(qn("w:pStyle"))
    if ps is None:
        return ""
    sid = ps.get(qn("w:val"))
    for st in doc.styles:
        if st.style_id == sid:
            return st.name or ""
    return sid or ""


def all_bold(p_el):
    runs = [r for r in p_el.iter(qn("w:r")) if "".join(t.text or "" for t in r.iter(qn("w:t"))).strip()]
    if not runs:
        return False
    for r in runs:
        rpr = r.find(qn("w:rPr"))
        if rpr is None or rpr.find(qn("w:b")) is None:
            return False
        b = rpr.find(qn("w:b"))
        if b.get(qn("w:val")) in ("0", "false"):
            return False
    return True


# --------------------------------------------------------------------------- découpage des runs par écriture


def chunk_text(text, para_rtl):
    """Découpe en [(classe, texte)] avec classe ar|lat|neu->hérite de la direction voisine."""
    out = []
    prev = "ar" if para_rtl else "lat"
    for c in text:
        cl = char_class(c)
        if cl == "neu":
            cl = prev
        else:
            prev = cl
        if out and out[-1][0] == cl:
            out[-1][1].append(c)
        else:
            out.append((cl, [c]))
    return [(cl, "".join(chars)) for cl, chars in out]


def split_runs(p_el, para_rtl):
    """Sépare les runs « un seul w:t » mixtes arabe/latin ; renvoie la liste de (run, classe)."""
    result = []
    for r in list(p_el.iter(qn("w:r"))):
        texts = r.findall(qn("w:t"))
        content = [c for c in r if _local(c) not in ("rPr",)]
        if not texts:
            result.append((r, "ar" if para_rtl else "lat"))
            continue
        full = "".join(t.text or "" for t in texts)
        if len(texts) == 1 and len(content) == 1:
            chunks = chunk_text(full, para_rtl)
            if len(chunks) > 1:
                parent = r.getparent()
                pos = parent.index(r)
                parent.remove(r)
                for i, (cl, txt) in enumerate(chunks):
                    nr = copy.deepcopy(r)
                    nt = nr.find(qn("w:t"))
                    nt.text = txt
                    nt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    parent.insert(pos + i, nr)
                    result.append((nr, cl))
                continue
            result.append((r, chunks[0][0] if chunks else "lat"))
        else:
            chunks = chunk_text(full, para_rtl)
            dom = "ar" if sum(len(t) for c, t in chunks if c == "ar") > sum(len(t) for c, t in chunks if c == "lat") else "lat"
            result.append((r, dom))
    return result


# --------------------------------------------------------------------------- noyau


class Formatter:
    def __init__(self, data, rules, page_map=None, markers=False):
        self.doc = Document(io.BytesIO(data))
        self.rules = deep_merge(DEFAULT_RULES, rules)
        self.page_map = page_map or {}
        self.markers = markers
        self.report = {"warnings": [], "outline": [], "counts": {}}
        self.headings = []   # [{"id","level","text","kind"}]

    # ----- utilitaires règles
    def body_font(self):
        b = self.rules["body"]
        return b["font_latin"], b["font_arabic"]

    # ----- 1. nettoyage
    def clean(self):
        body = self.doc.element.body
        removed = 0
        # sommaires intégrés (contrôles de contenu)
        for sdt in list(body.findall(qn("w:sdt"))):
            gal = sdt.find(".//" + qn("w:docPartGallery"))
            if gal is not None and "table of contents" in (gal.get(qn("w:val")) or "").lower():
                body.remove(sdt)
                removed += 1
        # paragraphes de sommaire existants (styles TOC, champs TOC)
        in_field = False
        for p in list(body.findall(qn("w:p"))):
            sname = style_name(self.doc, p).lower()
            instr = " ".join(i.text or "" for i in p.iter(qn("w:instrText")))
            if re.match(r"^\s*toc\b", instr.lower()):
                in_field = True
            if sname.startswith("toc") or sname.startswith("table des mati") or in_field:
                body.remove(p)
                removed += 1
                if any(fc.get(qn("w:fldCharType")) == "end" for fc in p.iter(qn("w:fldChar"))):
                    in_field = False
        # mémoriser les sauts de section avant de les supprimer (utile pour détecter la page de garde)
        self.sect_break_els = set()
        for sp in body.findall(".//" + qn("w:p") + "/" + qn("w:pPr") + "/" + qn("w:sectPr")):
            p_el = sp.getparent().getparent()
            if p_el is not None:
                self.sect_break_els.add(p_el)
            sp.getparent().remove(sp)
        # nettoyer les anciens en-têtes/pieds hérités (ex. « ملخص ») sur la section finale
        final_sect = body.find(qn("w:sectPr"))
        if final_sect is not None:
            for child in list(final_sect):
                if _local(child) in ("headerReference", "footerReference", "titlePg", "pgNumType"):
                    final_sect.remove(child)
        self.report["counts"]["old_toc_elements_removed"] = removed

    def _is_toc_table(self, tbl_el):
        rows = tbl_el.findall(qn("w:tr"))
        if len(rows) < 3:
            return False
        first_cells = ["".join(t.text or "" for t in c.iter(qn("w:t"))).strip() for c in rows[0].findall(qn("w:tc"))]
        hdr_txt = " ".join(first_cells)
        if len(first_cells) <= 3 and TOC_TABLE_HDR.search(norm(hdr_txt)):
            return True
        num_hits = 0
        for row in rows[1:min(len(rows), 15)]:
            cells = row.findall(qn("w:tc"))
            if len(cells) >= 2:
                last_txt = "".join(t.text or "" for t in cells[-1].iter(qn("w:t"))).strip()
                if re.match(r"^[\divxlcIVXLC٠-٩]+$", last_txt):
                    num_hits += 1
        return num_hits >= 3

    def remove_manual_toc(self, items):
        """Supprime un sommaire tapé à la main (lignes avec numéros de page OU tableau de sommaire)."""
        i = 0
        while i < len(items):
            it = items[i]
            if it["tag"] == "p" and it["text"] and len(it["text"]) < 40 and FRONT_PATTERNS["toc"].match(norm(it["text"])):
                # Cas 1 : sommaire sous forme de tableau (ex. الفهرس suivi d'un tableau العنوان / الصفحة)
                j = i + 1
                while j < len(items) and items[j]["tag"] == "p" and (
                    not items[j]["text"] or (len(items[j]["text"]) < 40 and FRONT_PATTERNS["toc"].match(norm(items[j]["text"])))
                ):
                    j += 1
                if j < len(items) and items[j]["tag"] == "tbl" and self._is_toc_table(items[j]["el"]):
                    start = i
                    # supprimer aussi un éventuel cadre décoratif vide juste avant الفهرس
                    prev = start - 1
                    while prev >= 0 and items[prev]["tag"] == "p" and not items[prev]["text"] and not has_drawing(items[prev]["el"]):
                        prev -= 1
                    if prev >= 0 and items[prev]["tag"] == "p" and not items[prev]["text"] and has_drawing(items[prev]["el"]):
                        start = prev
                    for k in range(start, j + 1):
                        el = items[k]["el"]
                        if el.getparent() is not None:
                            el.getparent().remove(el)
                    del items[start:j + 1]
                    self.report["counts"]["manual_toc_removed"] = self.report["counts"].get("manual_toc_removed", 0) + 1
                    i = start
                    continue
                # Cas 2 : sommaire tapé en paragraphes avec points de suite
                j = i + 1
                n_entries = 0
                while j < len(items):
                    t = items[j]
                    if t["tag"] == "p" and (not t["text"] or TOC_ENTRY.search(t["raw"])):
                        n_entries += 1 if t["text"] else 0
                        j += 1
                    else:
                        break
                if n_entries >= 2:
                    for k in range(i, j):
                        el = items[k]["el"]
                        if el.getparent() is not None:
                            el.getparent().remove(el)
                    del items[i:j]
                    self.report["counts"]["manual_toc_removed"] = n_entries
                    continue
            i += 1

    # ----- 2. analyse
    def build_items(self):
        items = []
        for el in self.doc.element.body.iterchildren():
            name = _local(el)
            if name == "p":
                raw = "".join(t.text or "" for t in el.iter(qn("w:t")))
                items.append({"tag": "p", "el": el, "raw": raw, "text": raw.strip()})
            elif name == "tbl":
                items.append({"tag": "tbl", "el": el, "raw": "", "text": ""})
        return items

    def _find_body_start_raw(self, items, cover_end):
        """Repère l'index du premier vrai titre de corps (مقدمة / Introduction / الفصل / Chapitre)."""
        start = (cover_end + 1) if cover_end is not None else 0
        for i in range(start, len(items)):
            it = items[i]
            if it["tag"] != "p" or not it["text"] or len(it["text"]) > 150:
                continue
            n = norm(it["text"])
            if MAIN_PATTERN.match(n) or CHAPTER_PATTERN.match(n):
                return i
        return None

    def classify(self, items, cover_end=None):
        has_chapters = any(it["tag"] == "p" and it["text"] and len(it["text"]) < 150 and
                           CHAPTER_PATTERN.match(norm(it["text"])) for it in items)
        has_ar_hierarchy = any(it["tag"] == "p" and it["text"] and len(it["text"]) < 200 and
                               any(pat.match(norm(it["text"])) for pat, _ in SECTION_AR) for it in items)
        has_fara = any(it["tag"] == "p" and it["text"] and len(it["text"]) < 200 and
                       SECTION_AR[2][0].match(norm(it["text"])) for it in items)
        self.has_chapters = has_chapters
        self.has_ar_hierarchy = has_ar_hierarchy
        self.has_fara = has_fara
        raw_body_start = self._find_body_start_raw(items, cover_end)
        base = 2 if has_chapters else 1
        for idx, it in enumerate(items):
            it["kind"], it["level"], it["uncertain"] = "body", 0, False
            if it["tag"] != "p":
                continue
            if cover_end is not None and idx <= cover_end:
                it["kind"] = "cover"
                continue
            t, n = it["text"], norm(it["text"])
            if not t:
                it["kind"] = "empty"
                continue
            if has_drawing(it["el"]) and not any(p.match(n) for p in FRONT_PATTERNS.values()):
                it["kind"] = "figure"
                continue
            # Préliminaires (شكر، إهداء، ملخص، قائمة المختصرات...)
            matched = False
            if len(t) <= 80:
                for cat, pat in FRONT_PATTERNS.items():
                    if pat.match(n):
                        it["kind"], it["cat"] = "front", cat
                        matched = True
                        break
            if matched:
                continue
            # Ne jamais classer en heading un paragraphe situé avant le début du corps (مقدمة / الفصل)
            if raw_body_start is not None and idx < raw_body_start:
                continue
            # Motifs structurels explicites (vérifiés AVANT la limite de 60 car. ou le point final / numPr)
            if len(t) <= 200:
                if CHAPTER_PATTERN.match(n):
                    it["kind"], it["level"], it["chapter"] = "heading", 1, True
                    continue
                if MAIN_PATTERN.match(n):
                    it["kind"], it["level"] = "heading", 1
                    continue
                if SUBCHAPTER_AR.match(n):
                    it["kind"], it["level"] = "heading", 2
                    continue
                sec_matched = False
                for pat, lvl in SECTION_AR:
                    if pat.match(n):
                        it["kind"], it["level"] = "heading", lvl
                        sec_matched = True
                        break
                if sec_matched:
                    continue
                if CAPTION_PATTERN.match(n):
                    it["kind"] = "caption"
                    continue
            sname = style_name(self.doc, it["el"]).lower()
            m = re.match(r"^(heading|titre)\s*(\d)$", sname)
            if m and len(t) <= 200:
                it["kind"], it["level"] = "heading", min(int(m.group(2)), HEADING_MAX_LEVEL)
                continue
            ppr = it["el"].find(qn("w:pPr"))
            has_num = ppr is not None and ppr.find(qn("w:numPr")) is not None
            short = len(t) <= 150 and not t.endswith((".", ";", ",")) or (len(t) <= 60)
            if not short or has_num:
                continue
            words = len(t.split())
            if ORDINAL_AR.match(n) and words <= 14:
                # Si la hiérarchie arabe (المبحث/المطلب/الفرع) est utilisée, أولاً/ثانياً est un sous-titre de corps (hors sommaire)
                if not has_ar_hierarchy:
                    it["kind"], it["level"] = "heading", base
            elif NUM4.match(t) and words <= 16:
                it["kind"], it["level"] = "heading", base + 3
            elif NUM3.match(t) and words <= 16:
                it["kind"], it["level"] = "heading", base + 2
            elif NUM2.match(t) and words <= 16:
                it["kind"], it["level"] = "heading", base + 1
            elif (ROMAN.match(t) or NUM1.match(t)) and words <= 14 and not t.endswith((".", ";", ",")):
                it["kind"], it["level"] = "heading", base
            elif LETTER.match(t) and words <= 14 and not t.endswith((".", ";", ",")):
                it["kind"], it["level"] = "heading", base + 1
            elif (not has_chapters and not has_ar_hierarchy) and all_bold(it["el"]) and words <= 12 and not t.endswith((".", ":", ";", ",")):
                it["kind"], it["level"], it["uncertain"] = "heading", base + 1, True
            if it["kind"] == "heading":
                it["level"] = max(1, min(it["level"], HEADING_MAX_LEVEL))
        return items

    def merge_chapter_titles(self, items):
        """Fusionne « الفصل الأول / Chapitre 1 » avec les 1 à 3 lignes de titre qui suivent, et قائمة المصادر + والمراجع."""
        i = 0
        while i < len(items):
            it = items[i]
            if it.get("chapter") and CHAPTER_LABEL_ONLY.match(norm(it["text"])) and len(it["text"]) < 40:
                merged_any = False
                for step in range(3):
                    j = i + 1
                    while j < len(items) and items[j]["kind"] == "empty":
                        j += 1
                    if j >= len(items):
                        break
                    nxt = items[j]
                    nt = nxt["text"]
                    nn = norm(nt)
                    if nxt["tag"] != "p" or nxt["kind"] not in ("body", "heading") or not nt or len(nt) > 120:
                        break
                    if nt.endswith(".") or SUBCHAPTER_AR.match(nn) or CHAPTER_PATTERN.match(nn) or MAIN_PATTERN.match(nn) or any(p.match(nn) for p, _ in SECTION_AR):
                        break
                    sep = w_el("r")
                    t = OxmlElement("w:t")
                    t.text = " : " if not merged_any else " "
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    sep.append(t)
                    it["el"].append(sep)
                    for r in list(nxt["el"].findall(qn("w:r"))):
                        it["el"].append(r)
                    it["text"] = p_text(it["el"])
                    it["raw"] = it["text"]
                    if nxt["el"].getparent() is not None:
                        nxt["el"].getparent().remove(nxt["el"])
                    del items[i + 1:j + 1]
                    merged_any = True
            elif it["tag"] == "p" and it["text"] and BIBLIO_PART1.match(norm(it["text"])):
                j = i + 1
                while j < len(items) and items[j]["kind"] == "empty":
                    j += 1
                if j < len(items) and items[j]["tag"] == "p" and items[j]["text"] and BIBLIO_PART2.match(norm(items[j]["text"])):
                    nxt = items[j]
                    sep = w_el("r")
                    t = OxmlElement("w:t")
                    t.text = " "
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    sep.append(t)
                    it["el"].append(sep)
                    for r in list(nxt["el"].findall(qn("w:r"))):
                        it["el"].append(r)
                    it["text"] = p_text(it["el"])
                    it["raw"] = it["text"]
                    it["kind"], it["level"] = "heading", 1
                    if nxt["el"].getparent() is not None:
                        nxt["el"].getparent().remove(nxt["el"])
                    del items[i + 1:j + 1]
            i += 1

    def strict_structural(self, text):
        n = norm(text)
        if len(text) > 80:
            return False
        return (any(p.match(n) for p in FRONT_PATTERNS.values()) or MAIN_PATTERN.match(n) is not None
                or CHAPTER_PATTERN.match(n) is not None)

    def find_cover_end(self, items):
        """Renvoie l'index (inclus) du dernier élément du bloc couverture (incluant double couverture / التصريح الشرفي avant les préliminaires)."""
        first_struct = next((i for i, it in enumerate(items)
                             if it["tag"] == "p" and it["text"] and self.strict_structural(it["text"])), len(items))

        def content_count(upto):
            return sum(1 for x in items[:upto + 1] if x["tag"] == "tbl" or x["text"] or has_drawing(x["el"]))

        # Si un vrai titre préliminaire ou de corps apparaît dans les 120 premiers éléments, tout ce qui le précède est la couverture
        if 0 < first_struct <= 120 and content_count(first_struct - 1) >= 3:
            end = first_struct - 1
            # ne pas absorber un cadre décoratif vide appartenant à la page du premier titre préliminaire
            while end > 0 and items[end]["tag"] == "p" and not items[end]["text"]:
                if has_drawing(items[end]["el"]):
                    end -= 1
                    break
                end -= 1
            return end
        # Sinon, repli sur le premier saut de page ou de section
        sect_breaks = getattr(self, "sect_break_els", set())
        for i in range(min(first_struct + 1, len(items), 80)):
            it = items[i]
            if it["tag"] == "p" and (has_page_break(it["el"]) or it["el"] in sect_breaks):
                return i if content_count(i) >= 2 else None
        return None

    # ----- 3. styles
    def apply_document_styles(self):
        r = self.rules
        doc = self.doc
        lat, ara = self.body_font()
        b = r["body"]
        pg = r["page"]
        # docDefaults
        dd = doc.styles.element.find(qn("w:docDefaults"))
        if dd is not None:
            rpr = dd.find(".//" + qn("w:rPrDefault") + "/" + qn("w:rPr"))
            if rpr is not None:
                set_rpr_fonts(rpr, lat, ara, b["size_latin"], b["size_arabic"])
                ordered_set(rpr, w_el("lang", val=r["language"]["latin"], bidi=r["language"]["arabic"]), RPR_ORDER)
        def ensure_style(name, base=None):
            try:
                return doc.styles[name]
            except KeyError:
                s = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
                if base is not None:
                    s.base_style = base
                return s

        normal = ensure_style("Normal")
        npr = normal.element.get_or_add_rPr()
        set_rpr_fonts(npr, lat, ara, b["size_latin"], b["size_arabic"], bold=False, italic=False)
        pf = normal.paragraph_format
        pf.line_spacing = b["line_spacing"]
        pf.space_after = Pt(b["space_after_pt"])
        pf.space_before = Pt(0)
        # titres
        for lvl in range(1, HEADING_MAX_LEVEL + 1):
            h = r["headings"][str(lvl)]
            st = ensure_style("Heading %d" % lvl, normal)
            ordered_set(st.element.get_or_add_pPr(), w_el("outlineLvl", val=lvl - 1), PPR_ORDER)
            rpr = st.element.get_or_add_rPr()
            for t in ("color", "caps", "smallCaps", "u", "i", "iCs", "spacing"):
                strip_rpr(rpr, [t])
            set_rpr_fonts(rpr, lat, ara, h["size_latin"], h["size_arabic"], bold=h["bold"], italic=False)
            if h.get("caps"):
                ordered_set(rpr, w_el("caps"), RPR_ORDER)
            p = st.paragraph_format
            p.space_before, p.space_after = Pt(h["space_before_pt"]), Pt(h["space_after_pt"])
            p.line_spacing = 1.15
            p.keep_with_next = True
            p.keep_together = True
            p.page_break_before = False if lvl == 1 else bool(h.get("page_break_before"))
            p.alignment = getattr(WD_ALIGN_PARAGRAPH, "CENTER" if h["align"] == "center" else "LEFT")
        # titres préliminaires (remerciements, dédicace...) : hors sommaire par défaut
        ft = r["front_titles"]
        st = ensure_style("Titre preliminaire", normal)
        rpr = st.element.get_or_add_rPr()
        set_rpr_fonts(rpr, lat, ara, ft["size_latin"], ft["size_arabic"], bold=ft["bold"])
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st.paragraph_format.space_after = Pt(18)
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.page_break_before = False
        # légendes
        cap = r["caption"]
        cst = ensure_style("Caption", normal)
        rpr = cst.element.get_or_add_rPr()
        for t in ("color",):
            strip_rpr(rpr, [t])
        set_rpr_fonts(rpr, lat, ara, cap["size_latin"], cap["size_arabic"], bold=cap["bold"], italic=cap["italic"])
        cst.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cst.paragraph_format.space_after = Pt(6)
        cst.paragraph_format.line_spacing = 1.0
        # styles de sommaire (avec direction bidi si document majoritairement arabe)
        all_txt = " ".join(p_el_text(p) for p in doc.element.body.findall(qn("w:p"))[:150])
        self.doc_rtl = dominant_dir(all_txt) == "rtl"
        usable = pg["width_cm"] - pg["margin_left_cm"] - pg["margin_right_cm"]
        self.toc_text_width_twips = int(usable * 567)
        for lvl in (1, 2, 3, 4):
            name = "toc %d" % lvl
            st = ensure_style(name, normal)
            st.paragraph_format.left_indent = Cm(0.6 * (lvl - 1))
            st.paragraph_format.space_after = Pt(4)
            st.paragraph_format.line_spacing = 1.15
            if self.doc_rtl:
                ordered_set(st.element.get_or_add_pPr(), w_el("bidi"), PPR_ORDER)
            set_rpr_fonts(st.element.get_or_add_rPr(), lat, ara, b["size_latin"], b["size_arabic"], bold=(lvl == 1))

    def apply_page(self):
        pg = self.rules["page"]
        ml, mr = pg["margin_left_cm"], pg["margin_right_cm"]
        if getattr(self, "doc_rtl", False) and ml > mr:
            ml, mr = mr, ml
        for sect in self.doc.sections:
            sect.page_width, sect.page_height = Cm(pg["width_cm"]), Cm(pg["height_cm"])
            sect.top_margin, sect.bottom_margin = Cm(pg["margin_top_cm"]), Cm(pg["margin_bottom_cm"])
            sect.left_margin, sect.right_margin = Cm(ml), Cm(mr)

    # ----- 4. mise en forme des paragraphes
    def format_runs(self, p_el, role, level=0):
        para_rtl = dominant_dir(p_el_text(p_el)) == "rtl"
        r = self.rules
        lat, ara = self.body_font()
        runs = split_runs(p_el, para_rtl)
        for run, cl in runs:
            rpr = get_rpr(run)
            if role == "heading":
                strip_rpr(rpr, ["rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "color", "caps", "u", "highlight", "shd"])
            elif role == "front":
                strip_rpr(rpr, ["rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "color", "u", "highlight"])
            elif role == "caption":
                strip_rpr(rpr, ["rFonts", "sz", "szCs", "color"])
            elif role == "table":
                strip_rpr(rpr, ["rFonts", "sz", "szCs"])
                tb = r["table"]
                set_rpr_fonts(rpr, lat, ara, tb["size_latin"], tb["size_arabic"])
            else:  # body : le style Normal gouverne police et taille ; on garde gras/italique/souligné
                strip_rpr(rpr, ["rFonts", "sz", "szCs", "color", "highlight"])
            for old in rpr.findall(qn("w:rtl")):
                rpr.remove(old)
            if cl == "ar":
                ordered_set(rpr, w_el("rtl"), RPR_ORDER)
                ordered_set(rpr, w_el("lang", bidi=r["language"]["arabic"]), RPR_ORDER)
            else:
                ordered_set(rpr, w_el("lang", val=r["language"]["latin"]), RPR_ORDER)
            if cl == "ar":
                # le gras/italique doivent aussi s'appliquer à l'écriture complexe
                if rpr.find(qn("w:b")) is not None:
                    ordered_set(rpr, w_el("bCs"), RPR_ORDER)
                if rpr.find(qn("w:i")) is not None:
                    ordered_set(rpr, w_el("iCs"), RPR_ORDER)
            if not len(rpr):
                run.remove(rpr)
        return para_rtl

    def set_style(self, p_el, name):
        sid = None
        for st in self.doc.styles:
            if st.name == name:
                sid = st.style_id
                break
        ppr = get_ppr(p_el)
        ordered_set(ppr, w_el("pStyle", val=sid), PPR_ORDER)

    def clear_direct_ppr(self, p_el, keep=("pStyle", "sectPr", "numPr")):
        ppr = get_ppr(p_el)
        for child in list(ppr):
            if _local(child) not in keep:
                ppr.remove(child)

    def format_items(self, items, cover_end, body_start_idx):
        b = self.rules["body"]
        counters = {"headings": 0, "front": 0, "captions": 0, "paragraphs": 0, "tables": 0}
        hid = 0
        front_in_toc = self.rules["front_titles"].get("in_toc", False)
        for idx, it in enumerate(items):
            in_cover = cover_end is not None and idx <= cover_end
            if it["tag"] == "tbl":
                if in_cover and not self.rules["cover"]["restyle"]:
                    continue
                for p in it["el"].iter(qn("w:p")):
                    rtl = self.format_runs(p, "table")
                    set_bidi(p, rtl)
                    self.clear_direct_ppr(p)
                    ppr = get_ppr(p)
                    ordered_set(ppr, w_el("spacing", before=0, after=0, line=240, lineRule="auto"), PPR_ORDER)
                    set_jc(p, "left" if rtl else "left")
                counters["tables"] += 1
                continue
            p = it["el"]
            if it["kind"] == "empty":
                continue
            if in_cover:
                if self.rules["cover"]["restyle"]:
                    self.format_runs(p, "body")
                else:
                    rtl = dominant_dir(it["text"]) == "rtl"
                    for run, cl in split_runs(p, rtl):
                        rpr = get_rpr(run)
                        for old in rpr.findall(qn("w:rtl")):
                            rpr.remove(old)
                        if cl == "ar":
                            ordered_set(rpr, w_el("rtl"), RPR_ORDER)
                    set_bidi(p, rtl)
                continue
            kind = it["kind"]
            if kind == "heading" or (kind == "front" and front_in_toc):
                level = it["level"] if kind == "heading" else 1
                rtl = self.format_runs(p, "heading")
                self.clear_direct_ppr(p)
                self.set_style(p, "Heading %d" % level)
                set_bidi(p, rtl)

                if idx < body_start_idx and kind == "heading":
                    pass
                hid += 1
                it["hid"] = hid
                self.headings.append({"id": hid, "level": level, "text": it["text"], "rtl": rtl,
                                      "kind": kind, "uncertain": it["uncertain"] if kind == "heading" else False,
                                      "el": p})
                counters["headings"] += 1
            elif kind == "front":
                rtl = self.format_runs(p, "front")
                self.clear_direct_ppr(p)
                self.set_style(p, "Titre preliminaire")
                set_bidi(p, rtl)

                counters["front"] += 1
            elif kind == "caption":
                rtl = self.format_runs(p, "caption")
                self.clear_direct_ppr(p)
                self.set_style(p, "Caption")
                set_bidi(p, rtl)
                counters["captions"] += 1
            elif kind == "figure":
                ppr_keep = p.find(qn("w:pPr"))
                self.clear_direct_ppr(p)
                set_jc(p, "center")
                ordered_set(get_ppr(p), w_el("keepNext"), PPR_ORDER)
                set_bidi(p, False)
            else:
                ppr = p.find(qn("w:pPr"))
                has_num = ppr is not None and ppr.find(qn("w:numPr")) is not None
                orig_jc = ppr.find(qn("w:jc")).get(qn("w:val")) if ppr is not None and ppr.find(qn("w:jc")) is not None else None
                rtl = self.format_runs(p, "body")
                # Supprimer List Paragraph / styles hérités parasites et uniformiser sur Normal
                self.clear_direct_ppr(p, keep=("sectPr", "numPr"))
                self.set_style(p, "Normal")
                set_bidi(p, rtl)
                # Nettoyer les tabulations <w:tab/> et espaces manuels en début de paragraphe
                for r_el in list(p.findall(qn("w:r"))):
                    stopped = False
                    for child in list(r_el):
                        loc = _local(child)
                        if loc == "rPr":
                            continue
                        if loc == "tab":
                            r_el.remove(child)
                        elif loc == "t":
                            txt = child.text or ""
                            stripped = txt.lstrip(" \t\u00a0")
                            child.text = stripped
                            if stripped:
                                stopped = True
                                break
                        else:
                            stopped = True
                            break
                    if stopped:
                        break
                indent_twips = int((b["first_line_indent_cm"] or 0) * 567)
                short_center = orig_jc == "center" and len(it["text"]) <= 100 and not has_num
                if short_center:
                    set_jc(p, "center")
                    ordered_set(get_ppr(p), w_el("ind", left=0, right=0, firstLine=0), PPR_ORDER)
                elif has_num:
                    set_jc(p, ALIGN_MAP.get(b["align"], "both"))
                    ordered_set(get_ppr(p), w_el("ind", left=indent_twips if indent_twips else 567, right=0, hanging=284), PPR_ORDER)
                else:
                    set_jc(p, ALIGN_MAP.get(b["align"], "both"))
                    ordered_set(get_ppr(p), w_el("ind", left=0, right=0, firstLine=indent_twips), PPR_ORDER)
                counters["paragraphs"] += 1
        self.report["counts"].update(counters)

    # ----- 5. sommaire + sections
    def toc_levels(self):
        lvls = int(self.rules["toc"]["levels"])
        if getattr(self, "has_fara", False) and lvls < 4:
            lvls = 4
        return max(1, min(lvls, HEADING_MAX_LEVEL))

    def make_toc_entry(self, h, first=False):
        pg = self.page_map.get(h["id"])
        page_txt = str(pg) if pg is not None else ""
        lvl = h["level"]
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        p.append(ppr)
        ordered_set(ppr, w_el("pStyle", val=self._style_id("toc %d" % lvl)), PPR_ORDER)
        tabs = OxmlElement("w:tabs")
        tab = w_el("tab", val="right", leader="dot", pos=self.toc_text_width_twips)
        tabs.append(tab)
        ordered_set(ppr, tabs, PPR_ORDER)
        if h["rtl"]:
            ordered_set(ppr, w_el("bidi"), PPR_ORDER)
        runs = []
        if first:
            runs += [("fld", "begin"), ("instr", ' TOC \\o "1-%d" \\h \\z \\u ' % self.toc_levels()), ("fld", "separate")]
        runs += [("text", h["text"]), ("tab", None), ("text", page_txt)]
        for kind, val in runs:
            r = OxmlElement("w:r")
            if kind == "fld":
                fc = w_el("fldChar", fldCharType=val)
                if val == "begin":
                    fc.set(qn("w:dirty"), "true")
                r.append(fc)
            elif kind == "instr":
                it = OxmlElement("w:instrText")
                it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                it.text = val
                r.append(it)
            elif kind == "tab":
                r.append(OxmlElement("w:tab"))
            else:
                if val and kind == "text" and val is h["text"] and h["rtl"]:
                    rp = OxmlElement("w:rPr")
                    ordered_set(rp, w_el("rtl"), RPR_ORDER)
                    r.append(rp)
                t = OxmlElement("w:t")
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = val
                r.append(t)
            p.append(r)
        return p

    def make_section_break_paragraph(self, sect_template, end_field=False):
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        p.append(ppr)
        sp = copy.deepcopy(sect_template)
        for child in list(sp):
            if _local(child) in ("headerReference", "footerReference", "titlePg", "pgNumType", "type"):
                sp.remove(child)
        ordered_set(sp, w_el("type", val="nextPage"), SECT_ORDER)
        ordered_set(ppr, sp, PPR_ORDER)
        if end_field:
            r = OxmlElement("w:r")
            r.append(w_el("fldChar", fldCharType="end"))
            p.append(r)
        return p

    def _anchor_with_frame(self, items, idx, min_idx):
        """Si un cadre décoratif (figure sans texte) précède immédiatement idx, recule l'ancre sur ce cadre."""
        prev = idx - 1
        while prev >= min_idx and items[prev]["kind"] == "empty":
            prev -= 1
        if prev >= min_idx and items[prev]["kind"] == "figure" and not items[prev]["text"]:
            return prev
        return idx

    def remove_empty_paragraphs(self, items, cover_end):
        """Supprime toutes les lignes vides parasites après la page de garde."""
        start = (cover_end + 1) if cover_end is not None else 0
        removed = 0
        new_items = list(items[:start])
        for it in items[start:]:
            if it["tag"] == "p" and not it["text"] and not has_drawing(it["el"]):
                has_fld = it["el"].find(".//" + qn("w:fldChar")) is not None
                has_sect = it["el"].find(".//" + qn("w:sectPr")) is not None
                if not has_fld and not has_sect:
                    if it["el"].getparent() is not None:
                        it["el"].getparent().remove(it["el"])
                    removed += 1
                    continue
            new_items.append(it)
        self.report["counts"]["empty_paragraphs_removed"] = removed
        return new_items

    def build_toc_and_sections(self, items, cover_end, body_start_idx):
        body = self.doc.element.body
        final_sect = body.find(qn("w:sectPr"))
        ordered_set(final_sect, w_el("type", val="nextPage"), SECT_ORDER)
        toc_heads = [h for h in self.headings if h["level"] <= self.toc_levels()
                     and (h["kind"] == "heading" or self.rules["front_titles"].get("in_toc"))]
        toc_heads = [h for h in toc_heads if self._idx_of(items, h["el"]) >= body_start_idx] or toc_heads
        self.toc_heads = toc_heads
        if body_start_idx is None or body_start_idx >= len(items):
            self.report["warnings"].append("Aucun titre principal détecté : sommaire et sections non créés.")
            return

        doc_rtl = getattr(self, "doc_rtl", False)
        toc_title = self.rules["toc"]["title_arabic" if doc_rtl else "title_latin"]
        min_after_cover = (cover_end + 1) if cover_end is not None else 0

        # 1) Construire la liste des points de coupure de section (chaque section préliminaire, sommaire, et chaque Heading 1 du corps)
        section_starts = []  # liste de dict: {"idx": int, "el": elem, "kind": str, "title": str, "rtl": bool}
        for idx in range(min_after_cover, len(items)):
            it = items[idx]
            if it["tag"] != "p":
                continue
            if idx < body_start_idx and it["kind"] == "front":
                a_idx = self._anchor_with_frame(items, idx, min_after_cover)
                section_starts.append({
                    "idx": a_idx, "el": items[a_idx]["el"], "kind": "front",
                    "title": it["text"], "rtl": dominant_dir(it["text"]) == "rtl"
                })
            elif idx >= body_start_idx and it["kind"] == "heading" and it["level"] == 1:
                a_idx = self._anchor_with_frame(items, idx, body_start_idx if idx > body_start_idx else min_after_cover)
                kind = "annex" if ANNEX_PATTERN.match(norm(it["text"])) else "body"
                section_starts.append({
                    "idx": a_idx, "el": items[a_idx]["el"], "kind": kind,
                    "title": it["text"], "rtl": dominant_dir(it["text"]) == "rtl"
                })

        # 2) Insérer le sommaire juste avant le début du corps (avec un saut de section avant le sommaire et un après)
        body_anchor_idx = self._anchor_with_frame(items, body_start_idx, min_after_cover)
        body_anchor = items[body_anchor_idx]["el"]
        toc_inserted = False
        if self.rules["toc"]["enabled"] and toc_heads:
            tp = OxmlElement("w:p")
            tppr = get_ppr(tp)
            ordered_set(tppr, w_el("pStyle", val=self._style_id("Titre preliminaire")), PPR_ORDER)
            if doc_rtl:
                ordered_set(tppr, w_el("bidi"), PPR_ORDER)
            tr = OxmlElement("w:r")
            if doc_rtl:
                rp = OxmlElement("w:rPr")
                ordered_set(rp, w_el("rtl"), RPR_ORDER)
                tr.append(rp)
            tt = OxmlElement("w:t")
            tt.text = toc_title
            tr.append(tt)
            tp.append(tr)
            # Saut de section avant le sommaire (sépare le dernier préliminaire ou la couverture du sommaire)
            pre_toc_brk = self.make_section_break_paragraph(final_sect)
            # Saut de section après le sommaire (termine le champ TOC et démarre le corps)
            post_toc_brk = self.make_section_break_paragraph(final_sect, end_field=True)
            elems = [pre_toc_brk, tp] + [self.make_toc_entry(h, first=(i == 0)) for i, h in enumerate(toc_heads)] + [post_toc_brk]
            for el in elems:
                body_anchor.addprevious(el)
            toc_inserted = True
        else:
            brk = self.make_section_break_paragraph(final_sect)
            body_anchor.addprevious(brk)

        # 3) Insérer les sauts de section entre les sections préliminaires et entre chaque chapitre/section principale du corps
        #    Note : la transition juste avant body_anchor est déjà assurée par post_toc_brk (ou brk ci-dessus)
        inserted_before_els = {body_anchor}
        for s_idx, sec in enumerate(section_starts):
            el = sec["el"]
            if el in inserted_before_els or el.getparent() is None:
                continue
            # Ne pas insérer un saut avant le tout premier élément du document s'il n'y a pas de couverture
            if s_idx == 0 and cover_end is None:
                continue
            # Si c'est la 1ère section après la couverture, on nettoie un éventuel saut de page sur last_cover
            if s_idx == 0 and cover_end is not None:
                last_cover = items[cover_end]["el"]
                for br in list(last_cover.iter(qn("w:br"))):
                    if br.get(qn("w:type")) == "page":
                        br.getparent().remove(br)
            el.addprevious(self.make_section_break_paragraph(final_sect))
            inserted_before_els.add(el)

        # 4) Construire self.section_meta dans l'ordre exact des sections du document
        self.section_meta = []
        if cover_end is not None:
            self.section_meta.append({"kind": "cover", "title": "", "rtl": doc_rtl, "restart_num": False})
        first_front_seen = False
        first_body_seen = False
        for sec in section_starts:
            if sec["kind"] == "front":
                self.section_meta.append({
                    "kind": "front", "title": sec["title"], "rtl": sec["rtl"],
                    "restart_num": (not first_front_seen)
                })
                first_front_seen = True
        if toc_inserted:
            self.section_meta.append({
                "kind": "toc", "title": toc_title, "rtl": doc_rtl,
                "restart_num": (not first_front_seen)
            })
            first_front_seen = True
        for sec in section_starts:
            if sec["kind"] in ("body", "annex"):
                self.section_meta.append({
                    "kind": sec["kind"], "title": sec["title"], "rtl": sec["rtl"],
                    "restart_num": (sec["kind"] == "body" and not first_body_seen)
                })
                if sec["kind"] == "body":
                    first_body_seen = True
        self.has_sections = True

    def _style_id(self, name):
        for st in self.doc.styles:
            if st.name == name:
                return st.style_id
        return None

    def _idx_of(self, items, el):
        for i, it in enumerate(items):
            if it["el"] is el:
                return i
        return -1

    def _clear_hf(self, hf):
        for p in list(hf.paragraphs)[1:]:
            p._element.getparent().remove(p._element)
        fp = hf.paragraphs[0]
        for child in list(fp._element):
            if _local(child) != "pPr":
                fp._element.remove(child)
        ppr = fp._element.find(qn("w:pPr"))
        if ppr is not None:
            for child in list(ppr):
                ppr.remove(child)
        return fp

    def _write_page_number(self, hf, position, rtl=False):
        fp = self._clear_hf(hf)
        fp.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT,
                        "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(position, WD_ALIGN_PARAGRAPH.CENTER)
        lat, ara = self.body_font()
        for kind in ("begin", "instr", "separate", "text", "end"):
            r = OxmlElement("w:r")
            rp = OxmlElement("w:rPr")
            set_rpr_fonts(rp, lat, ara, 11, 12)
            r.append(rp)
            if kind in ("begin", "separate", "end"):
                r.append(w_el("fldChar", fldCharType=kind))
            elif kind == "instr":
                it = OxmlElement("w:instrText")
                it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                it.text = " PAGE "
                r.append(it)
            else:
                t = OxmlElement("w:t")
                t.text = "1"
                r.append(t)
            fp._element.append(r)

    def _write_header_title(self, hf, title, rtl=False):
        fp = self._clear_hf(hf)
        if not title:
            return
        ppr = get_ppr(fp._element)
        set_bidi(fp._element, rtl)
        ordered_set(ppr, w_el("jc", val="center"), PPR_ORDER)
        # Filet inférieur discret sous l'en-tête (style académique cahier.pdf)
        pbdr = OxmlElement("w:pBdr")
        pbdr.append(w_el("bottom", val="single", sz="6", space="4", color="444444"))
        ordered_set(ppr, pbdr, PPR_ORDER)
        lat, ara = self.body_font()
        r = OxmlElement("w:r")
        rp = OxmlElement("w:rPr")
        set_rpr_fonts(rp, lat, ara, 10, 13, bold=True)
        if rtl:
            ordered_set(rp, w_el("rtl"), RPR_ORDER)
        r.append(rp)
        t = OxmlElement("w:t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = title
        r.append(t)
        fp._element.append(r)

    def apply_footers(self):
        pn = self.rules["page_numbers"]
        sections = list(self.doc.sections)
        metas = getattr(self, "section_meta", [])
        for si, sect in enumerate(sections):
            meta = metas[si] if si < len(metas) else {"kind": "body", "title": "", "rtl": getattr(self, "doc_rtl", False), "restart_num": False}
            kind = meta["kind"]
            title = meta["title"]
            rtl = meta["rtl"]
            sp = sect._sectPr
            # Délier systématiquement en-têtes et pieds de page de la section précédente
            sect.header.is_linked_to_previous = False
            sect.footer.is_linked_to_previous = False
            sect.first_page_header.is_linked_to_previous = False
            sect.first_page_footer.is_linked_to_previous = False

            if kind == "cover":
                sect.different_first_page_header_footer = False
                self._clear_hf(sect.header)
                self._clear_hf(sect.footer)
                continue

            # Page 1 de chaque chapitre/section : en-tête vide (comme cahier.pdf) ; pages 2+ : titre de la section
            sect.different_first_page_header_footer = True
            self._clear_hf(sect.first_page_header)
            self._write_header_title(sect.header, title, rtl=rtl)

            # Pied de page / numérotation
            is_front = kind in ("front", "toc")
            is_annex = kind == "annex"
            fmt = "none" if is_annex else (pn["front_format"] if is_front else pn["body_format"])
            if not pn["enabled"] or fmt == "none":
                self._clear_hf(sect.first_page_footer)
                self._clear_hf(sect.footer)
                continue

            if meta.get("restart_num"):
                ordered_set(sp, w_el("pgNumType", fmt=fmt, start=1), SECT_ORDER)
            else:
                ordered_set(sp, w_el("pgNumType", fmt=fmt), SECT_ORDER)
            self._write_page_number(sect.first_page_footer, pn["position"], rtl=rtl)
            self._write_page_number(sect.footer, pn["position"], rtl=rtl)

    def add_markers(self):
        for h in self.headings:
            r = OxmlElement("w:r")
            rp = OxmlElement("w:rPr")
            ordered_set(rp, w_el("color", val="FFFFFF"), RPR_ORDER)
            ordered_set(rp, w_el("sz", val=2), RPR_ORDER)
            r.append(rp)
            t = OxmlElement("w:t")
            t.text = "ZZH%dZZ" % h["id"]
            r.append(t)
            ppr = h["el"].find(qn("w:pPr"))
            if ppr is not None:
                ppr.addnext(r)
            else:
                h["el"].insert(0, r)

    def set_update_fields(self):
        settings = self.doc.settings.element
        for old in settings.findall(qn("w:updateFields")):
            settings.remove(old)
        settings.append(w_el("updateFields", val="true"))

    # ----- orchestration
    def run(self):
        self.clean()
        self.apply_document_styles()
        self.apply_page()
        items = self.build_items()
        self.remove_manual_toc(items)
        cover_end = self.find_cover_end(items)
        items = self.remove_empty_paragraphs(items, cover_end)
        self.classify(items, cover_end)
        self.merge_chapter_titles(items)
        body_start = next((i for i, it in enumerate(items)
                           if it["kind"] == "heading" and (cover_end is None or i > cover_end)), None)
        self.report["cover_detected"] = cover_end is not None
        self.report["has_chapters"] = self.has_chapters
        self.format_items(items, cover_end, body_start if body_start is not None else len(items))
        self.has_sections = False
        if body_start is not None:
            self.build_toc_and_sections(items, cover_end, body_start)
        else:
            self.report["warnings"].append("Aucun titre principal détecté (chapitre/introduction).")
        if self.has_sections:
            self.apply_footers()
        else:
            self.apply_footers()
        if self.markers:
            self.add_markers()
        self.set_update_fields()
        self.report["outline"] = [{"level": h["level"], "text": h["text"], "uncertain": h["uncertain"]}
                                  for h in self.headings]
        self.report["uncertain_headings"] = sum(1 for h in self.headings if h["uncertain"])
        out = io.BytesIO()
        self.doc.save(out)
        return out.getvalue()


def p_el_text(p_el):
    return "".join(t.text or "" for t in p_el.iter(qn("w:t")))


# --------------------------------------------------------------------------- pagination via LibreOffice


def _render_page_map(docx_bytes, n_headings):
    """Convertit en PDF avec LibreOffice et renvoie {id_titre: n° de page logique du corps}."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return {}
    tmp = tempfile.mkdtemp()
    try:
        src = os.path.join(tmp, "m.docx")
        with open(src, "wb") as f:
            f.write(docx_bytes)
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, src],
                       check=True, timeout=180, capture_output=True,
                       env=dict(os.environ, HOME=tmp))
        txt = subprocess.run(["pdftotext", "-layout", os.path.join(tmp, "m.pdf"), "-"],
                             check=True, capture_output=True, timeout=60).stdout.decode("utf-8", "ignore")
        pages = txt.split("\f")
        found = {}
        for pi, pg in enumerate(pages):
            for m in re.finditer(r"ZZH(\d+)ZZ", pg):
                found.setdefault(int(m.group(1)), pi + 1)
        if not found:
            return {}
        first_phys = found[min(found)]
        return {hid: phys - first_phys + 1 for hid, phys in found.items()}
    except Exception:
        return {}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def format_docx(data, rules=None, with_pages=True):
    """Point d'entrée : renvoie (octets .docx, rapport)."""
    rules = _coerce(DEFAULT_RULES, normalize_rules(rules))
    first = Formatter(data, rules, markers=with_pages)
    out1 = first.run()
    page_map = _render_page_map(out1, len(first.headings)) if with_pages else {}
    second = Formatter(data, rules, page_map=page_map, markers=False)
    out2 = second.run()
    second.report["toc_pages_estimated"] = bool(page_map)
    return out2, second.report