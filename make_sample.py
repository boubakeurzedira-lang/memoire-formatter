"""Génère un mémoire factice mal formaté (arabe + français) pour tester le service."""
import sys
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK

d = Document()
def para(text, bold=False, center=False, size=None, pagebreak=False):
    p = d.add_paragraph()
    r = p.add_run(text); r.bold = bold
    if size: r.font.size = Pt(size)
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if pagebreak: p.add_run().add_break(WD_BREAK.PAGE)
    return p

# page de garde
para("République Algérienne Démocratique et Populaire", True, True, 14)
para("Université des Sciences et de la Technologie", True, True, 14)
para("Département de Recherche Opérationnelle", False, True)
para("Mémoire de fin d'études en vue de l'obtention du diplôme de Master", False, True)
para("Optimisation de la chaîne logistique : cas d'une entreprise de livraison", True, True, 18)
para("Réalisé par : Bob Zed", False, True)
para("Année universitaire 2025/2026", False, True, pagebreak=True)
# dédicace / remerciements
para("Dédicace", True)
para("À mes chers parents, pour leur soutien indéfectible.", pagebreak=True)
para("شكر وتقدير", True)
para("نتقدم بجزيل الشكر والامتنان إلى الأستاذ المشرف على توجيهاته القيمة وصبره طوال فترة إنجاز هذا العمل، كما نشكر جميع أساتذة القسم.", pagebreak=True)
para("Résumé", True)
para("Ce mémoire étudie l'optimisation des tournées de livraison à l'aide de la recherche opérationnelle.", pagebreak=True)
para("ملخص", True)
para("تتناول هذه المذكرة تحسين مسارات التوزيع باستخدام بحوث العمليات وأدوات دعم القرار.", pagebreak=True)
para("Sommaire", True)
para("Introduction générale .......... 1")
para("Chapitre 1 .......... 3", pagebreak=True)
# corps
para("Introduction générale", True)
para("La logistique de dernier kilomètre représente une part importante des coûts de distribution. Dans ce contexte, l'optimisation des tournées devient un enjeu majeur pour les entreprises de livraison comme ZR Express. " * 3)
para("مقدمة", True)
para("تعتبر اللوجستيات من أهم عوامل نجاح المؤسسات الاقتصادية، حيث تساهم في تخفيض التكاليف وتحسين جودة الخدمة المقدمة للزبائن. وفي هذا الإطار، نعتمد على نماذج مثل VRP و TSP لحل مشكلة التوزيع. " * 3)
para("Chapitre 1")
para("Etat de l'art sur le problème de tournées de véhicules")
para("1. Définitions et concepts de base")
para("Le problème de tournées de véhicules (VRP) consiste à déterminer un ensemble de routes de coût minimal. " * 4)
para("1.1 Historique du problème")
para("Introduit par Dantzig et Ramser en 1959, le VRP a donné lieu à une littérature abondante. " * 4)
para("1.1.1 Variantes classiques")
para("On distingue le CVRP, le VRPTW et le VRPPD selon les contraintes considérées. " * 3)
para("2. Méthodes de résolution")
para("Les méthodes exactes et les métaheuristiques sont comparées. " * 5)
para("Tableau 1 : Comparaison des méthodes")
t = d.add_table(rows=3, cols=3); t.style = "Table Grid"
for i, row in enumerate([("Méthode", "Type", "Complexité"), ("Branch & Cut", "Exacte", "Exponentielle"), ("Recuit simulé", "Heuristique", "Polynomiale")]):
    for j, v in enumerate(row): t.cell(i, j).text = v
para("الفصل الثاني")
para("الدراسة التطبيقية في مؤسسة التوصيل")
para("أولا: تقديم المؤسسة")
para("تأسست المؤسسة سنة 2015 وتنشط في مجال التوصيل السريع عبر مختلف ولايات الوطن، وتعتمد على نظام معلومات يسمى ERP لتسيير الطلبيات. " * 3)
para("1.2 نتائج النمذجة")
para("بعد تطبيق النموذج تم تخفيض المسافة الإجمالية بنسبة 18% مقارنة بالوضع الحالي حسب Table 2. " * 3)
para("Conclusion générale", True)
para("Ce travail a permis de proposer un modèle d'optimisation adapté au contexte algérien. " * 3)
para("خاتمة", True)
para("خلصت هذه الدراسة إلى أن استخدام بحوث العمليات يمكن أن يحسن الأداء اللوجستي للمؤسسة بشكل ملموس. " * 2)
para("Bibliographie", True)
para("Dantzig, G. B., & Ramser, J. H. (1959). The truck dispatching problem. Management Science, 6(1), 80-91.")
d.save(sys.argv[1] if len(sys.argv) > 1 else "sample.docx")
print("ok")
