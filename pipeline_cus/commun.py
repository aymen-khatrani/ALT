"""Socle commun du pipeline Python -> agent Copilot -> Python.

Utilisé par prepare.py, controle.py et bilan.py. Seule dépendance externe : openpyxl.
"""
import csv
import hashlib
import re
import unicodedata
from pathlib import Path

VERSION = "1.0"

# Catalogue des règles (identique aux instructions de l'agent) : ID -> (nom, sévérité par défaut)
REGLES = {
    "AMB-01": ("Missing actor", "MAJEUR"),
    "AMB-02": ("Vague term", "MAJEUR"),
    "AMB-03": ("Indefinite article", "MINEUR"),
    "AMB-04": ("Unclear reference", "MAJEUR"),
    "VER-01": ("No measurable criterion", "BLOQUANT"),
    "VER-02": ("Unprovable negative", "BLOQUANT"),
    "VER-03": ("Absolute", "BLOQUANT"),
    "VER-04": ("Limit not stated as <= / >=", "MINEUR"),
    "RSK-01": ("Escape clause", "BLOQUANT"),
    "RSK-02": ("Open-ended clause", "BLOQUANT"),
    "RSK-03": ("Placeholder", "BLOQUANT"),
    "RSK-04": ("Binding status unclear", "MAJEUR"),
    "RSK-05": ("Several requirements in one", "MAJEUR"),
    "RSK-06": ("Reference to standards", "BLOQUANT"),
    "RSK-07": ("Imposed design solution", "MAJEUR"),
    "STY-01": ("Wrong imperative", "MINEUR"),
    "STY-02": ("Rationale or detail inside", "MINEUR"),
    "STY-03": ("Sentence order", "MINEUR"),
}

VERDICTS = ["A_CLARIFIER", "OK", "HORS_PERIMETRE"]
SEVERITES = ["BLOQUANT", "MAJEUR", "MINEUR"]
CONFIANCES = ["Haute", "Moyenne", "Basse"]
# Décisions des relecteurs dans le classeur de validation :
#   Acceptée     : problème réel, question et formulation reprises telles quelles
#   Modifiée     : problème réel, texte corrigé dans "Question finale"
#   Rejetée      : problème réel, mais pas de question au client (traité en interne, sans enjeu…)
#   Faux positif : l'agent a signalé un problème qui n'en est pas un
#   Oubli agent  : ligne OK sur laquelle le relecteur trouve un vrai problème (mesure les oublis)
DECISIONS = ["Acceptée", "Modifiée", "Rejetée", "Faux positif", "Oubli agent"]

# Colonnes des fichiers échangés avec l'agent
COLONNES_LOT = ["ID", "Texte", "Signaux_lexique"]
COLONNES_RESULTAT = ["ID", "Verdict", "Sévérité", "Règles", "Signaux écartés", "Extrait",
                     "Explication", "Question client", "Formulation proposée", "Confiance"]

RE_ID_REGLE = re.compile(r"\b(?:AMB|VER|RSK|STY)-\d{2}\b")
RE_PLACEHOLDER = re.compile(r"\[\s*TBC\s*:[^\]]*\]", re.IGNORECASE)
RE_NOMBRE = re.compile(r"\d+(?:[.,]\d+)?")


# ---------------------------------------------------------------- texte
def sans_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def cle(s):
    """Clé de comparaison pour les noms de colonnes : minuscules, sans accents, alphanumérique seulement."""
    return re.sub(r"[^a-z0-9]", "", sans_accents(str(s or "")).lower())


def texte_normalise(s):
    """Pour comparer un extrait au texte : minuscules, guillemets droits, espaces compactés."""
    s = str(s or "").replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    s = s.replace("…", "...").replace(" ", " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def en_texte(v):
    """Valeur de cellule -> texte (12.0 -> '12')."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def ids_regles(s):
    return list(dict.fromkeys(RE_ID_REGLE.findall(str(s or ""))))


def nombres(s):
    return {n.replace(",", ".") for n in RE_NOMBRE.findall(str(s or ""))}


def nombres_hors_placeholders(s):
    return nombres(RE_PLACEHOLDER.sub(" ", str(s or "")))


def empreinte(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()[:8]


# ---------------------------------------------------------------- lexique
def charger_lexique(path):
    """Lit lexique.csv (séparateur ';') -> liste de (regle, regex compilée ou None, nom_special)."""
    entrees = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for i, ligne in enumerate(csv.DictReader(f, delimiter=";"), start=2):
            regle = (ligne.get("regle") or "").strip()
            motif = (ligne.get("motif") or "").strip()
            genre = (ligne.get("type") or "mot").strip().lower()
            if not regle or regle.startswith("#") or not motif:
                continue
            if regle not in REGLES:
                raise ValueError(f"lexique.csv ligne {i} : règle inconnue '{regle}'")
            if genre == "special":
                entrees.append((regle, None, motif))
                continue
            if genre == "mot":
                corps = re.escape(motif).replace(r"\ ", r"\s+")
                motif_re, drapeaux = r"(?<!\w)" + corps + r"(?!\w)", re.IGNORECASE
            elif genre == "regex":
                motif_re, drapeaux = motif, re.IGNORECASE
            elif genre == "regex_casse":
                motif_re, drapeaux = motif, 0
            else:
                raise ValueError(f"lexique.csv ligne {i} : type inconnu '{genre}' (mot, regex, regex_casse, special)")
            try:
                entrees.append((regle, re.compile(motif_re, drapeaux), None))
            except re.error as e:
                raise ValueError(f"lexique.csv ligne {i} : regex invalide ({e})")
    return entrees


RE_MODAL_FORT = re.compile(r"\b(?:shall|must)\b", re.IGNORECASE)


def _special(nom, texte):
    """Contrôles qui ne sont pas de simples mots-clés."""
    if nom == "plusieurs_shall":
        n = len(RE_MODAL_FORT.findall(texte))
        return [f"{n} x shall/must"] if n >= 2 else []
    if nom == "aucun_shall":
        return [] if RE_MODAL_FORT.search(texte) else ["(no shall)"]
    raise ValueError(f"contrôle spécial inconnu : {nom}")


def detecter_signaux(texte, lexique):
    """-> dict ordonné {regle: [termes trouvés]}."""
    trouves = {}
    for regle, rx, special in lexique:
        termes = _special(special, texte) if special else [m.group(0) for m in rx.finditer(texte)]
        for t in termes:
            t = re.sub(r"\s+", " ", t).strip()
            if t and t.lower() not in [x.lower() for x in trouves.setdefault(regle, [])]:
                trouves[regle].append(t)
    ordre = list(REGLES)
    return {r: trouves[r] for r in sorted(trouves, key=ordre.index) if trouves[r]}


def formater_signaux(d):
    return " | ".join(f"{r}: {', '.join(t)}" for r, t in d.items())


# ---------------------------------------------------------------- Excel
def ligne_entete(ws, max_lignes=15):
    """Numéro de la ligne d'en-tête : celle qui a le plus de cellules texte parmi les premières lignes."""
    meilleure, score_max = 1, -1
    for r, ligne in enumerate(ws.iter_rows(min_row=1, max_row=max_lignes, values_only=True), start=1):
        score = sum(1 for v in ligne if isinstance(v, str) and v.strip())
        if score > score_max:
            meilleure, score_max = r, score
    return meilleure


def lire_feuille(ws):
    """-> (entetes, lignes) ; lignes = liste de (numéro_ligne_excel, dict entete->valeur)."""
    h = ligne_entete(ws)
    entetes = [en_texte(v) or f"Colonne_{i + 1}" for i, v in
               enumerate(next(ws.iter_rows(min_row=h, max_row=h, values_only=True)))]
    lignes = []
    for r, valeurs in enumerate(ws.iter_rows(min_row=h + 1, values_only=True), start=h + 1):
        if all(v is None or str(v).strip() == "" for v in valeurs):
            continue
        lignes.append((r, dict(zip(entetes, valeurs))))
    return entetes, lignes


def colonne(entetes, *noms):
    """Trouve la colonne dont le nom normalisé correspond à l'un des noms donnés."""
    cles = {cle(e): e for e in entetes}
    for n in noms:
        if cle(n) in cles:
            return cles[cle(n)]
    return None
