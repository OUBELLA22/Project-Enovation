#!/usr/bin/env python3
"""
Builds a Word (.docx) 'Cahier des Charges' for the Wire Harness Routing & Generation project.
No external libraries required: a .docx is a ZIP of OOXML parts, which we assemble by hand.
"""
import zipfile
from xml.sax.saxutils import escape

# ---------------------------------------------------------------------------
# Helpers to produce WordprocessingML paragraphs/tables
# ---------------------------------------------------------------------------

def run(text, bold=False, italic=False, size=None, color=None):
    rpr = ""
    props = ""
    if bold:
        props += '<w:b/>'
    if italic:
        props += '<w:i/>'
    if size:  # half-points
        props += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    if color:
        props += f'<w:color w:val="{color}"/>'
    if props:
        rpr = f'<w:rPr>{props}</w:rPr>'
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(runs, style=None, align=None, spacing_after=120, spacing_before=0):
    if isinstance(runs, str):
        runs = [run(runs)]
    ppr = '<w:pPr>'
    if style:
        ppr += f'<w:pStyle w:val="{style}"/>'
    ppr += f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}"/>'
    if align:
        ppr += f'<w:jc w:val="{align}"/>'
    ppr += '</w:pPr>'
    return f'<w:p>{ppr}{"".join(runs)}</w:p>'


def heading(text, level=1):
    size = {1: 32, 2: 28, 3: 24}.get(level, 24)
    color = "1F4E79" if level == 1 else ("2E74B5" if level == 2 else "404040")
    return para([run(text, bold=True, size=size, color=color)],
                spacing_before=240, spacing_after=120)


def bullet(text, bold_prefix=None):
    runs = []
    if bold_prefix:
        runs.append(run(bold_prefix, bold=True))
    runs.append(run(text))
    ppr = ('<w:pPr><w:pStyle w:val="ListParagraph"/>'
           '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
           '<w:spacing w:after="60"/></w:pPr>')
    return f'<w:p>{ppr}{"".join(runs)}</w:p>'


def tcell(text, bold=False, width=2500, shade=None, color=None):
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>' if shade else ''
    tcpr = (f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shd}'
            '<w:vAlign w:val="center"/></w:tcPr>')
    return (f'<w:tc>{tcpr}'
            f'{para([run(text, bold=bold, color=color)], spacing_after=20)}'
            f'</w:tc>')


def table(rows, widths):
    """rows: list of lists of (text, bold, shade) ; first row treated as header."""
    grid = '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths) + '</w:tblGrid>'
    tblpr = ('<w:tblPr><w:tblStyle w:val="TableGrid"/>'
             '<w:tblW w:w="0" w:type="auto"/>'
             '<w:tblBorders>'
             '<w:top w:val="single" w:sz="4" w:color="BFBFBF"/>'
             '<w:left w:val="single" w:sz="4" w:color="BFBFBF"/>'
             '<w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/>'
             '<w:right w:val="single" w:sz="4" w:color="BFBFBF"/>'
             '<w:insideH w:val="single" w:sz="4" w:color="BFBFBF"/>'
             '<w:insideV w:val="single" w:sz="4" w:color="BFBFBF"/>'
             '</w:tblBorders></w:tblPr>')
    out = f'<w:tbl>{tblpr}{grid}'
    for i, r in enumerate(rows):
        out += '<w:tr>'
        for j, cell in enumerate(r):
            text, bold, shade = cell
            color = "FFFFFF" if i == 0 else None
            out += tcell(text, bold=bold or i == 0, width=widths[j],
                         shade=shade or ("1F4E79" if i == 0 else None), color=color)
        out += '</w:tr>'
    out += '</w:tbl>'
    # add a trailing empty paragraph (Word requires one after a table)
    out += para([run("")], spacing_after=120)
    return out


# ---------------------------------------------------------------------------
# Document body
# ---------------------------------------------------------------------------
body = []

# ---- Title page ----
body.append(para([run("CAHIER DES CHARGES", bold=True, size=56, color="1F4E79")],
                 align="center", spacing_before=1200, spacing_after=120))
body.append(para([run("Outil de Génération Automatique de Routage de Faisceau Électrique",
                      bold=True, size=32, color="2E74B5")], align="center", spacing_after=60))
body.append(para([run("(Automatic Wire Harness Routing & Generation Tool)", italic=True, size=24,
                      color="595959")], align="center", spacing_after=600))
body.append(para([run("Projet d'automatisation CAO / CATIA V5", size=24)], align="center", spacing_after=40))
body.append(para([run("Version 1.0  —  Document de spécification", size=22, color="595959")],
                 align="center", spacing_after=1600))
body.append(para([run("Date : _____________        Rédacteur : _____________        Validé par : _____________",
                      size=20, color="595959")], align="center", spacing_after=200))

# page break
body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

# ---- 0. Sommaire note ----
body.append(heading("Sommaire", 1))
toc = [
    "1. Contexte et objectif du projet",
    "2. Périmètre (ce qui est inclus / exclu)",
    "3. Définitions et glossaire",
    "4. Description fonctionnelle du besoin",
    "5. Exigences fonctionnelles",
    "6. Exigences techniques",
    "7. Architecture logicielle proposée",
    "8. Algorithme de génération du routage",
    "9. Rôle (optionnel) de l'IA",
    "10. Interface utilisateur",
    "11. Exigences non fonctionnelles",
    "12. Phases de développement",
    "13. Plan de tests et critères d'acceptation",
    "14. Livrables",
    "15. Hypothèses, contraintes et risques",
    "16. Questions ouvertes à clarifier",
]
for t in toc:
    body.append(para([run(t, size=22)], spacing_after=40))
body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

# ---- 1. Contexte ----
body.append(heading("1. Contexte et objectif du projet", 1))
body.append(para(
    "Le projet consiste à développer un outil capable de générer automatiquement le "
    "routage (le tracé) d'un faisceau électrique entre un connecteur de départ et un "
    "connecteur d'arrivée, en passant par un point de fixation choisi par l'utilisateur. "
    "L'environnement cible est un modèle CAO (CATIA V5) contenant déjà les objets de "
    "référence positionnés."))
body.append(para([run("Objectif principal : ", bold=True),
                  run("transformer une sélection simple de l'utilisateur (choix d'un point de "
                      "fixation) en une géométrie de faisceau générée de façon déterministe, "
                      "rapide, répétable et fiable.")]))
body.append(para([run("Principe directeur : ", bold=True),
                  run("la géométrie est calculée par un moteur déterministe (pas par une IA). "
                      "L'IA, si elle est utilisée, sert uniquement d'interface en langage "
                      "naturel qui produit des paramètres structurés, jamais à inventer des "
                      "coordonnées ou des règles d'ingénierie.")]))

# ---- 2. Périmètre ----
body.append(heading("2. Périmètre", 1))
body.append(heading("2.1 Inclus dans le périmètre", 2))
for t in [
    "Lecture des références existantes (connecteurs de départ/arrivée, points de fixation).",
    "Sélection par l'utilisateur du point de fixation à emprunter (F1 ou F2).",
    "Calcul déterministe du tracé A → fixation choisie → B.",
    "Génération de la géométrie du faisceau (courbe / spline puis balayage de section).",
    "Régénération propre du tracé en cas de changement de sélection (sans doublons).",
    "Contrôle que le tracé passe effectivement par le point de fixation (tolérance ε).",
]:
    body.append(bullet(t))
body.append(heading("2.2 Hors périmètre (version 1)", 2))
for t in [
    "Connectivité électrique (schéma, affectation des fils, continuité électrique).",
    "Évitement d'obstacles complexe (sauf si requis — voir §8.5).",
    "Génération de la planche de fabrication 2D à plat (formboard) — phase ultérieure.",
    "Gestion multi-faisceaux et branches multiples complexes — phase ultérieure.",
]:
    body.append(bullet(t))

# ---- 3. Glossaire ----
body.append(heading("3. Définitions et glossaire", 1))
body.append(table([
    [("Terme", False, None), ("Définition", False, None)],
    [("Connecteur", False, None), ("Point d'extrémité du faisceau (départ A ou arrivée B).", False, None)],
    [("Point de fixation (F1, F2)", False, None), ("Point d'ancrage/clip par lequel le faisceau doit passer.", False, None)],
    [("Point de passage obligatoire", False, None), ("Point que la ligne médiane (centerline) du faisceau doit traverser.", False, None)],
    [("Ligne médiane (centerline)", False, None), ("Courbe centrale du faisceau servant de guide au balayage.", False, None)],
    [("Diamètre de faisceau", False, None), ("Épaisseur du faisceau (section balayée le long du tracé).", False, None)],
    [("Rayon de courbure min.", False, None), ("Rayon minimal admissible dans un virage du faisceau.", False, None)],
    [("Tolérance ε", False, None), ("Écart maximal admis entre le tracé et le point de fixation.", False, None)],
    [("Set géométrique dédié", False, None), ("Ensemble (ex. AI_HARNESS_GENERATED) regroupant la géométrie générée.", False, None)],
], widths=[3200, 6300]))

# ---- 4. Description fonctionnelle ----
body.append(heading("4. Description fonctionnelle du besoin", 1))
body.append(para(
    "Le modèle CAO contient quatre objets de référence déjà positionnés : un connecteur de "
    "départ A, deux points de fixation possibles F1 et F2, et un connecteur d'arrivée B. "
    "L'utilisateur choisit le point de fixation à emprunter. L'outil calcule alors le tracé "
    "A → fixation choisie → B puis génère la géométrie correspondante."))
body.append(table([
    [("Référence", False, None), ("Rôle", False, None), ("Donnée requise", False, None)],
    [("Départ A", False, None), ("Début du faisceau", False, None), ("Position (+ direction de sortie)", False, None)],
    [("Fixation F1", False, None), ("Point de passage possible", False, None), ("Position (+ orientation si clip)", False, None)],
    [("Fixation F2", False, None), ("Point de passage alternatif", False, None), ("Position (+ orientation si clip)", False, None)],
    [("Arrivée B", False, None), ("Fin du faisceau", False, None), ("Position (+ direction d'approche)", False, None)],
], widths=[2200, 4000, 3300]))

# ---- 5. Exigences fonctionnelles ----
body.append(heading("5. Exigences fonctionnelles", 1))
reqs = [
    ("EF-01", "L'outil doit lire la position 3D des références A, F1, F2, B depuis le modèle."),
    ("EF-02", "L'utilisateur doit pouvoir sélectionner la fixation à emprunter (F1 ou F2)."),
    ("EF-03", "L'outil doit calculer le tracé A → fixation choisie → B."),
    ("EF-04", "Le tracé généré doit passer par la fixation choisie (||C(t_F) − F|| ≤ ε)."),
    ("EF-05", "L'outil doit générer la géométrie du faisceau (courbe + balayage de section)."),
    ("EF-06", "Toute la géométrie générée doit être placée dans un set dédié."),
    ("EF-07", "Au changement de sélection, l'ancienne géométrie doit être supprimée/remplacée sans doublon."),
    ("EF-08", "Les références d'origine (connecteurs, fixations, pièces) ne doivent jamais être modifiées."),
    ("EF-09", "En cas de référence manquante ou de tracé impossible, un message d'erreur clair doit être affiché."),
    ("EF-10", "La longueur totale du tracé doit être calculée et restituée."),
]
body.append(table(
    [[("ID", False, None), ("Exigence", False, None)]] +
    [[(i, False, None), (t, False, None)] for i, t in reqs],
    widths=[1400, 8100]))

# ---- 6. Exigences techniques ----
body.append(heading("6. Exigences techniques", 1))
for t in [
    "Environnement : CATIA V5 (atelier Generative Shape Design requis pour le balayage).",
    "Automatisation : CATScript / VBA (prototype) ou Python + COM (pywin32 / pycatia).",
    "Création de géométrie via HybridShapeFactory (spline, balayage circulaire) — à vérifier selon la version/licence installée.",
    "Système d'exploitation : Windows (requis pour l'automatisation COM de CATIA V5).",
    "Correspondance de l'architecture 32/64 bits entre Python et CATIA si automatisation Python.",
]:
    body.append(bullet(t))

# ---- 7. Architecture ----
body.append(heading("7. Architecture logicielle proposée", 1))
body.append(para("L'architecture sépare clairement trois responsabilités :"))
body.append(table([
    [("Couche", False, None), ("Rôle", False, None), ("Technologie", False, None)],
    [("Interface", False, None), ("Sélection F1/F2, saisie du diamètre, bouton Générer", False, None), ("Dialogue macro / app externe", False, None)],
    [("Moteur de routage", False, None), ("Calcul déterministe du tracé, contrôles, longueur", False, None), ("CATScript / VBA / Python", False, None)],
    [("Automatisation CAO", False, None), ("Création / mise à jour de la géométrie dans CATIA", False, None), ("API CATIA (COM)", False, None)],
    [("IA (optionnel)", False, None), ("Traduction langage naturel → paramètres JSON validés", False, None), ("Service LLM externe", False, None)],
], widths=[2400, 4600, 2500]))
body.append(para([run("Règle d'architecture : ", bold=True),
                  run("construire d'abord le moteur de routage, puis l'interface CATIA, puis "
                      "(optionnellement) la couche IA — afin de toujours savoir dans quelle "
                      "couche se situe un défaut.")]))

# ---- 8. Algorithme ----
body.append(heading("8. Algorithme de génération du routage", 1))
body.append(heading("8.1 Modélisation", 2))
body.append(para("Chaque référence est un point P = (x, y, z). Le tracé relie A, la fixation "
                 "choisie, et B dans cet ordre."))
body.append(heading("8.2 Longueur", 2))
body.append(para("Distance euclidienne entre deux points :  "
                 "d(P,Q) = √((xQ−xP)² + (yQ−yP)² + (zQ−zP)²)."))
body.append(para("Longueur du tracé simple :  L = d(A,F) + d(F,B)  "
                 "(la longueur réelle variera avec les transitions lissées)."))
body.append(heading("8.3 Méthodes de tracé", 2))
body.append(table([
    [("Méthode", False, None), ("Usage", False, None), ("Limite", False, None)],
    [("Polyligne", False, None), ("Premier test de connectivité", False, None), ("Angles vifs, non réaliste", False, None)],
    [("Spline", False, None), ("Prototype lissé", False, None), ("Ne garantit pas le rayon de courbure min.", False, None)],
    [("Routage contrôlé", False, None), ("Version production", False, None), ("Plus complexe (contraintes/obstacles)", False, None)],
], widths=[2400, 4000, 3100]))
body.append(heading("8.4 Contrainte de passage par la fixation", 2))
body.append(para("Vérifier explicitement que le tracé passe par la fixation : "
                 "||C(t_F) − F|| ≤ ε (ex. ε = 0,1 mm selon le besoin réel)."))
body.append(heading("8.5 Contraintes connecteur et obstacles", 2))
for t in [
    "Direction de sortie du connecteur + distance de dégagement.",
    "Rayon de courbure minimal du faisceau.",
    "Évitement d'obstacles (optionnel V1) : zones interdites / enveloppes de dégagement.",
    "Si plusieurs fixations/branches ou obstacles : revenir à un routage par graphe (Dijkstra / A*).",
]:
    body.append(bullet(t))

# ---- 9. IA ----
body.append(heading("9. Rôle (optionnel) de l'IA", 1))
body.append(para("L'IA ne calcule pas la géométrie. Elle traduit une instruction en langage "
                 "naturel en un objet JSON structuré et validé, par exemple :"))
body.append(para([run('{ "operation": "generate_harness", "start_reference": "START_CONNECTOR", '
                      '"end_reference": "END_CONNECTOR", "selected_fixation": "FIXATION_2", '
                      '"diameter_mm": 10.0, "routing_mode": "smooth", "bend_radius_mm": null, '
                      '"clearance_mm": null, "require_obstacle_check": true }',
                      italic=True, size=18, color="404040")]))
for t in [
    "Ne jamais inventer de coordonnées, dimensions ou exigences d'ingénierie.",
    "null = valeur non fournie → l'application la demande ou la lit dans la configuration.",
    "La sortie JSON doit être validée par un schéma strict avant toute opération CATIA.",
    "Ne jamais affirmer qu'un tracé est sans collision tant que le moteur ne l'a pas vérifié.",
]:
    body.append(bullet(t))

# ---- 10. Interface ----
body.append(heading("10. Interface utilisateur", 1))
for t in [
    "Affichage des connecteurs de départ et d'arrivée (lecture seule).",
    "Sélecteur de fixation : Fixation 1 / Fixation 2.",
    "Champ diamètre (ex. 10 mm) avec valeurs prédéfinies possibles.",
    "Bouton « Générer » et « Régénérer ».",
    "Zone de messages (succès / avertissement / erreur).",
    "Évolutions futures : détection auto des références, aperçu du tracé, réglage du rayon de courbure.",
]:
    body.append(bullet(t))

# ---- 11. Non fonctionnel ----
body.append(heading("11. Exigences non fonctionnelles", 1))
body.append(table([
    [("Catégorie", False, None), ("Exigence", False, None)],
    [("Fiabilité", False, None), ("Résultat identique et répétable pour une même entrée.", False, None)],
    [("Performance", False, None), ("Génération quasi instantanée pour un tracé simple.", False, None)],
    [("Robustesse", False, None), ("Gestion propre des erreurs (références manquantes, tracé impossible).", False, None)],
    [("Maintenabilité", False, None), ("Géométrie générée isolée dans un set dédié, facilement supprimable.", False, None)],
    [("Sécurité des données", False, None), ("Vérifier la politique interne avant tout envoi de données CAO à un service IA externe.", False, None)],
    [("Traçabilité", False, None), ("Journalisation des erreurs et des générations.", False, None)],
], widths=[2600, 6900]))

# ---- 12. Phases ----
body.append(heading("12. Phases de développement", 1))
phases = [
    ("Phase 1", "Comprendre le test", "Identifier A, F1, F2, B, leurs repères, le diamètre et la forme attendue."),
    ("Phase 2", "Prouver la géométrie", "Macro générant un tracé via F1 et un via F2, passant bien par la fixation."),
    ("Phase 3", "Améliorer le routage", "Transitions lissées, direction de sortie, rayon de courbure, collisions."),
    ("Phase 4", "Ajouter l'interface", "Sélection F1/F2 + Générer ; remplacement propre de la géométrie."),
    ("Phase 5", "Ajouter l'IA", "Prompt → paramètres JSON validés avant appel du moteur."),
    ("Phase 6", "Industrialiser", "Tests multi-assemblages, gestion d'erreurs, déploiement contrôlé."),
]
body.append(table(
    [[("Phase", False, None), ("Intitulé", False, None), ("Livrable / contenu", False, None)]] +
    [[(p, False, None), (n, False, None), (d, False, None)] for p, n, d in phases],
    widths=[1300, 2600, 5600]))

# ---- 13. Tests ----
body.append(heading("13. Plan de tests et critères d'acceptation", 1))
tests = [
    ("Sélection F1", "Le tracé passe par F1"),
    ("Sélection F2", "Le tracé passe par F2"),
    ("Générer deux fois", "Aucune géométrie en doublon"),
    ("Connecteur de départ manquant", "Erreur claire, aucun modèle invalide"),
    ("Fixation manquante", "Génération bloquée"),
    ("Tracé impossible", "Erreur signalée (pas de faux succès)"),
    ("Déplacement d'un connecteur", "Le tracé se met à jour correctement"),
    ("Tracé coupant une pièce", "Collision signalée si contrôle activé"),
    ("JSON IA invalide", "Données rejetées avant opérations CATIA"),
    ("Réouverture du CATPart", "Résultat attendu préservé"),
]
body.append(table(
    [[("Test", False, None), ("Résultat attendu", False, None)]] +
    [[(a, False, None), (b, False, None)] for a, b in tests],
    widths=[4200, 5300]))
body.append(para([run("Critères minimaux de livraison V1 : ", bold=True),
                  run("sélection correcte de la fixation, extrémités correctes, génération "
                      "stable et régénération propre.")]))

# ---- 14. Livrables ----
body.append(heading("14. Livrables", 1))
for t in [
    "Spécification des entrées/sorties du test (Phase 1).",
    "Macro CATIA de génération du tracé (Phase 2).",
    "Moteur de routage amélioré avec règles d'ingénierie (Phase 3).",
    "Interface utilisateur de sélection et génération (Phase 4).",
    "Couche IA optionnelle + schéma JSON de validation (Phase 5).",
    "Documentation d'utilisation et jeu de tests (Phase 6).",
]:
    body.append(bullet(t))

# ---- 15. Hypothèses / risques ----
body.append(heading("15. Hypothèses, contraintes et risques", 1))
body.append(heading("15.1 Hypothèses", 2))
for t in [
    "Les quatre références sont déjà présentes et correctement positionnées dans le modèle.",
    "Pour la V1, les deux tracés possibles sont supposés valides (sans collision).",
    "Le poste dispose de la licence Generative Shape Design.",
]:
    body.append(bullet(t))
body.append(heading("15.2 Risques", 2))
for t in [
    "Les API HybridShapeFactory varient selon la version/licence de CATIA.",
    "L'automatisation COM dépend de Windows et de la correspondance 32/64 bits.",
    "Une spline peut « manquer » le point de fixation → contrôle ε obligatoire.",
    "L'envoi de données CAO à une IA externe peut poser un problème de confidentialité.",
]:
    body.append(bullet(t))

# ---- 16. Questions ouvertes ----
body.append(heading("16. Questions ouvertes à clarifier", 1))
for t in [
    "Type de fixation : point de passage obligatoire 3D OU clip physique (avec orientation) ?",
    "Complexité du tracé : simple A → F → B lissé OU évitement d'obstacles requis ?",
    "Environnement : macro intégrée à CATIA V5 OU prototype moteur autonome (Python) d'abord ?",
    "Accès CATIA : assemblage disponible et licence Generative Shape Design présente ?",
    "Ordre des fixations imposé, ou le système doit-il choisir le meilleur ordre (si plusieurs) ?",
]:
    body.append(bullet(t))

body.append(para([run("")], spacing_after=200))
body.append(para([run("— Fin du Cahier des Charges —", italic=True, color="595959")],
                 align="center"))

# ---------------------------------------------------------------------------
# Assemble the body into a complete document.xml
# ---------------------------------------------------------------------------
sect = ('<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
        '</w:sectPr>')

document_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>' + ''.join(body) + sect + '</w:body></w:document>'
)

# ---------------------------------------------------------------------------
# Boilerplate parts
# ---------------------------------------------------------------------------
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    '</Types>'
)

root_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>'
)

doc_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    '</Relationships>'
)

styles_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="ListParagraph">'
    '<w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>'
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
    '<w:tblPr><w:tblBorders>'
    '<w:top w:val="single" w:sz="4" w:color="BFBFBF"/>'
    '<w:left w:val="single" w:sz="4" w:color="BFBFBF"/>'
    '<w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/>'
    '<w:right w:val="single" w:sz="4" w:color="BFBFBF"/>'
    '<w:insideH w:val="single" w:sz="4" w:color="BFBFBF"/>'
    '<w:insideV w:val="single" w:sz="4" w:color="BFBFBF"/>'
    '</w:tblBorders></w:tblPr></w:style>'
    '</w:styles>'
)

numbering_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
    '<w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/>'
    '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
    '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
    '</w:lvl></w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    '</w:numbering>'
)

# ---------------------------------------------------------------------------
# Write the .docx (zip)
# ---------------------------------------------------------------------------
out = "/projects/sandbox/Cahier_des_Charges_Faisceau_Routing.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", root_rels)
    z.writestr("word/document.xml", document_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels)
    z.writestr("word/styles.xml", styles_xml)
    z.writestr("word/numbering.xml", numbering_xml)

print("Written:", out)
