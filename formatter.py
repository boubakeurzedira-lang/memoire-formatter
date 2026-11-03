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
            hs["size_arabic"] = float(v) + 2
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
    "ack": rx(r"^(remerciements?|شكر وتقدير|شكر و تقدير|شكر وعرفان|كلمه شكر|تشكرات|شكر|الشكر|avant[- ]propos)\s*[:\-]?$"),
    "abstract": rx(r"^(resume|abstract|ملخص( الدراسه| البحث| المذكره)?|الملخص|مستخلص|resume en francais|ملخص بالعربيه)\s*[:\-]?$"),
    "lists": rx(r"^(liste des (tableaux|figures|abreviations|sigles|acronymes|symboles|annexes|illustrations)|"
                r"table des (figures|tableaux|illustrations)|glossaire|"
                r"قايمه (الجداول|الاشكال|المختصرات|الرموز|الصور|الملاحق|الاختصارات)|فهرس (الجداول|الاشكال|الصور))\s*[:\-]?$"),
    "toc": rx(r"^(sommaire|table des matieres|فهرس المحتويات|جدول المحتويات|المحتويات|الفهرس|فهرس)\s*[:\-]?$"),
}
MAIN_PATTERN = rx(r"^(introduction( generale)?|conclusion( generale)?|bibliographie|references( bibliographiques)?|"
                  r"webographie|annexes?|perspectives|مقدمه( عامه)?|المقدمه( العامه)?|خاتمه( عامه)?|الخاتمه( العامه)?|"
                  r"قايمه (المراجع|المصادر والمراجع|المصادر)|المراجع|المصادر والمراجع|المصادر|الملاحق|ملحق( \d+)?)\s*[:\-]?$")
CHAPTER_PATTERN = rx(r"^(chapitre|chapter|partie|الفصل|الباب|الجزء)\s+\S+")
CHAPTER_LABEL_ONLY = rx(r"^(chapitre|chapter|partie|الفصل|الباب|الجزء)\s+\S+(\s+\S+)?\s*[:\-]?$")
SECTION_AR = [(rx(r"^المبحث\s"), 2), (rx(r"^المطلب\s"), 3), (rx(r"^الفرع\s"), 4)]
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
        # sauts de section d'origine (on reconstruit les sections)
        for sp in body.findall(".//" + qn("w:p") + "/" + qn("w:pPr") + "/" + qn("w:sectPr")):
            sp.getparent().remove(sp)
        self.report["counts"]["old_toc_elements_removed"] = removed

    def remove_manual_toc(self, items):
        """Supprime un sommaire tapé à la main : titre + lignes se terminant par un numéro de page."""
        i = 0
        while i < len(items):
            it = items[i]
            if it["tag"] == "p" and it["text"] and len(it["text"]) < 40 and FRONT_PATTERNS["toc"].match(norm(it["text"])):
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

    def classify(self, items, cover_end=None):
        has_chapters = any(it["tag"] == "p" and it["text"] and len(it["text"]) < 150 and
                           CHAPTER_PATTERN.match(norm(it["text"])) for it in items)
        self.has_chapters = has_chapters
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
            if has_drawing(it["el"]):
                it["kind"] = "figure"
                continue
            ppr = it["el"].find(qn("w:pPr"))
            has_num = ppr is not None and ppr.find(qn("w:numPr")) is not None
            short = len(t) <= 150 and not t.endswith((".", ";", ",")) or (len(t) <= 60)
            matched = False
            if len(t) <= 80:
                for cat, pat in FRONT_PATTERNS.items():
                    if pat.match(n):
                        it["kind"], it["cat"] = "front", cat
                        matched = True
                        break
            if matched:
                continue
            sname = style_name(self.doc, it["el"]).lower()
            m = re.match(r"^(heading|titre)\s*(\d)$", sname)
            if m and len(t) <= 200:
                it["kind"], it["level"] = "heading", min(int(m.group(2)), HEADING_MAX_LEVEL)
                if CHAPTER_PATTERN.match(n) or MAIN_PATTERN.match(n):
                    it["level"] = 1
                continue
            if not short or has_num:
                if CAPTION_PATTERN.match(n) and len(t) < 200:
                    it["kind"] = "caption"
                continue
            if CHAPTER_PATTERN.match(n):
                it["kind"], it["level"], it["chapter"] = "heading", 1, True
            elif MAIN_PATTERN.match(n):
                it["kind"], it["level"] = "heading", 1
            elif CAPTION_PATTERN.match(n):
                it["kind"] = "caption"
            else:
                for pat, lvl in SECTION_AR:
                    if pat.match(n):
                        it["kind"], it["level"] = "heading", lvl
                        break
                else:
                    words = len(t.split())
                    if ORDINAL_AR.match(n) and words <= 14:
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
                    elif all_bold(it["el"]) and words <= 12 and not t.endswith((".", ":", ";", ",")):
                        it["kind"], it["level"], it["uncertain"] = "heading", base + 1, True
            if it["kind"] == "heading":
                it["level"] = max(1, min(it["level"], HEADING_MAX_LEVEL))
        return items

    def merge_chapter_titles(self, items):
        """« Chapitre 1 » seul sur sa ligne + titre sur la ligne suivante -> un seul titre."""
        i = 0
        while i < len(items):
            it = items[i]
            if it.get("chapter") and CHAPTER_LABEL_ONLY.match(norm(it["text"])) and len(it["text"]) < 40:
                j = i + 1
                while j < len(items) and items[j]["kind"] == "empty":
                    j += 1
                if j < len(items) and items[j]["tag"] == "p" and items[j]["kind"] in ("body", "heading") \
                        and 0 < len(items[j]["text"]) <= 120 and not items[j]["text"].endswith("."):
                    nxt = items[j]
                    sep = w_el("r")
                    t = OxmlElement("w:t")
                    t.text = " : "
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    sep.append(t)
                    it["el"].append(sep)
                    for r in list(nxt["el"].findall(qn("w:r"))):
                        it["el"].append(r)
                    it["text"] = p_text(it["el"])
                    it["raw"] = it["text"]
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
        """Renvoie l'index (inclus) du dernier élément de la page de garde, ou None.
        Utilise uniquement les saut de page et les mots-clés (avant toute classification)."""
        first_struct = next((i for i, it in enumerate(items)
                             if it["tag"] == "p" and it["text"] and self.strict_structural(it["text"])), len(items))

        def content_count(upto):
            return sum(1 for x in items[:upto + 1] if x["tag"] == "tbl" or x["text"] or has_drawing(x["el"]))

        for i in range(min(first_struct + 1, len(items), 80)):
            it = items[i]
            if it["tag"] == "p" and has_page_break(it["el"]):
                return i if content_count(i) >= 2 else None
        if 0 < first_struct <= 40 and content_count(first_struct - 1) >= 3:
            return first_struct - 1
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
        normal = doc.styles["Normal"]
        npr = normal.element.get_or_add_rPr()
        set_rpr_fonts(npr, lat, ara, b["size_latin"], b["size_arabic"], bold=False, italic=False)
        pf = normal.paragraph_format
        pf.line_spacing = b["line_spacing"]
        pf.space_after = Pt(b["space_after_pt"])
        pf.space_before = Pt(0)
        # titres : certains fichiers Word (notamment modèles personnalisés/arabisés)
        # ne contiennent pas tous les styles intégrés Heading 1..4. Les créer au besoin
        # évite qu'un niveau absent (ex. Heading 4) fasse échouer toute la mise en forme.
        for lvl in range(1, HEADING_MAX_LEVEL + 1):
            h = r["headings"][str(lvl)]
            style_name = "Heading %d" % lvl
            try:
                st = doc.styles[style_name]
            except KeyError:
                st = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                st.base_style = doc.styles["Heading %d" % (lvl - 1)] if lvl > 1 else normal
                ppr = st.element.get_or_add_pPr()
                ordered_set(ppr, w_el("outlineLvl", val=lvl - 1), PPR_ORDER)
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
            p.page_break_before = bool(h.get("page_break_before"))
            p.alignment = getattr(WD_ALIGN_PARAGRAPH, "CENTER" if h["align"] == "center" else "LEFT")
        # titres préliminaires (remerciements, dédicace...) : hors sommaire par défaut
        ft = r["front_titles"]
        if "Titre preliminaire" not in [s.name for s in doc.styles]:
            st = doc.styles.add_style("Titre preliminaire", WD_STYLE_TYPE.PARAGRAPH)
            st.base_style = normal
        st = doc.styles["Titre preliminaire"]
        rpr = st.element.get_or_add_rPr()
        set_rpr_fonts(rpr, lat, ara, ft["size_latin"], ft["size_arabic"], bold=ft["bold"])
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st.paragraph_format.space_after = Pt(18)
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.page_break_before = bool(ft.get("page_break_before"))
        # légendes
        cap = r["caption"]
        cst = doc.styles["Caption"]
        rpr = cst.element.get_or_add_rPr()
        for t in ("color",):
            strip_rpr(rpr, [t])
        set_rpr_fonts(rpr, lat, ara, cap["size_latin"], cap["size_arabic"], bold=cap["bold"], italic=cap["italic"])
        cst.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cst.paragraph_format.space_after = Pt(6)
        cst.paragraph_format.line_spacing = 1.0
        # styles de sommaire
        usable = pg["width_cm"] - pg["margin_left_cm"] - pg["margin_right_cm"]
        self.toc_text_width_twips = int(usable * 567)
        for lvl in (1, 2, 3, 4):
            name = "toc %d" % lvl
            if name not in [s.name for s in doc.styles]:
                st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
                st.base_style = normal
            st = doc.styles[name]
            st.paragraph_format.left_indent = Cm(0.6 * (lvl - 1))
            st.paragraph_format.space_after = Pt(4)
            st.paragraph_format.line_spacing = 1.15
            if lvl == 1:
                set_rpr_fonts(st.element.get_or_add_rPr(), lat, ara, b["size_latin"], b["size_arabic"], bold=True)

    def apply_page(self):
        pg = self.rules["page"]
        for sect in self.doc.sections:
            sect.page_width, sect.page_height = Cm(pg["width_cm"]), Cm(pg["height_cm"])
            sect.top_margin, sect.bottom_margin = Cm(pg["margin_top_cm"]), Cm(pg["margin_bottom_cm"])
            sect.left_margin, sect.right_margin = Cm(pg["margin_left_cm"]), Cm(pg["margin_right_cm"])

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
                self.clear_direct_ppr(p)
                set_bidi(p, rtl)
                short_center = orig_jc == "center" and len(it["text"]) <= 100
                if short_center:
                    set_jc(p, "center")
                elif has_num or len(it["text"]) < 45 and not it["text"].endswith((".", ":", ";")) and orig_jc in ("left", "right", "start", "end"):
                    set_jc(p, "left")
                else:
                    set_jc(p, ALIGN_MAP.get(b["align"], "both"))
                    if b["first_line_indent_cm"] and not has_num:
                        ordered_set(get_ppr(p), w_el("ind", firstLine=int(b["first_line_indent_cm"] * 567)), PPR_ORDER)
                counters["paragraphs"] += 1
        self.report["counts"].update(counters)

    # ----- 5. sommaire + sections
    def toc_levels(self):
        return max(1, min(int(self.rules["toc"]["levels"]), HEADING_MAX_LEVEL))

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
            if _local(child) in ("headerReference", "footerReference", "titlePg"):
                sp.remove(child)
        ordered_set(ppr, sp, PPR_ORDER)
        if end_field:
            r = OxmlElement("w:r")
            r.append(w_el("fldChar", fldCharType="end"))
            p.append(r)
        return p

    def build_toc_and_sections(self, items, cover_end, body_start_idx):
        body = self.doc.element.body
        final_sect = body.find(qn("w:sectPr"))
        toc_heads = [h for h in self.headings if h["level"] <= self.toc_levels()
                     and (h["kind"] == "heading" or self.rules["front_titles"].get("in_toc"))]
        # le sommaire ne liste pas ce qui le précède (préliminaires)
        toc_heads = [h for h in toc_heads if self._idx_of(items, h["el"]) >= body_start_idx] or toc_heads
        self.toc_heads = toc_heads
        anchor = items[body_start_idx]["el"] if body_start_idx is not None and body_start_idx < len(items) else None
        if anchor is None:
            self.report["warnings"].append("Aucun titre principal détecté : sommaire et sections non créés.")
            return
        # 5a. sommaire
        if self.rules["toc"]["enabled"] and toc_heads:
            doc_rtl = sum(len(AR_LETTER.findall(h["text"])) for h in toc_heads) > sum(len(LAT_LETTER.findall(h["text"])) for h in toc_heads)
            title = self.rules["toc"]["title_arabic" if doc_rtl else "title_latin"]
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
            tt.text = title
            tr.append(tt)
            tp.append(tr)
            elems = [tp] + [self.make_toc_entry(h, first=(i == 0)) for i, h in enumerate(toc_heads)]
        else:
            elems = []
        # 5b. saut de section avant le corps (fin du bloc préliminaire)
        brk = self.make_section_break_paragraph(final_sect, end_field=bool(elems))
        elems.append(brk)
        for el in elems:
            anchor.addprevious(el)
        # 5c. saut de section après la page de garde
        if cover_end is not None:
            last_cover = items[cover_end]["el"]
            if last_cover.getparent() is not None:
                # retire le saut de page explicite (remplacé par le saut de section)
                for br in list(last_cover.iter(qn("w:br"))):
                    if br.get(qn("w:type")) == "page":
                        br.getparent().remove(br)
                cb = self.make_section_break_paragraph(final_sect)
                last_cover.addnext(cb)
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

    def apply_footers(self):
        pn = self.rules["page_numbers"]
        sections = list(self.doc.sections)
        n = len(sections)
        for si, sect in enumerate(sections):
            sp = sect._sectPr
            has_cover = n == 3
            is_cover = has_cover and si == 0
            is_front = (has_cover and si == 1) or (not has_cover and n == 2 and si == 0)
            fmt = pn["front_format"] if is_front else pn["body_format"]
            sect.footer.is_linked_to_previous = False
            footer = sect.footer
            for p in list(footer.paragraphs)[1:]:
                p._element.getparent().remove(p._element)
            fp = footer.paragraphs[0]
            for r in list(fp._element):
                if _local(r) != "pPr":
                    fp._element.remove(r)
            if is_cover or not pn["enabled"] or fmt == "none":
                continue
            if n > 1:
                ordered_set(sp, w_el("pgNumType", fmt=fmt, start=1), SECT_ORDER)
            fp.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT,
                            "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(pn["position"], WD_ALIGN_PARAGRAPH.CENTER)
            for kind in ("begin", "instr", "separate", "text", "end"):
                r = OxmlElement("w:r")
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
