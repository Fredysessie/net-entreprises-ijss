"""
repartition_analytique.py
=========================
Complement du module `net_entreprise_final.py`.

Objet
-----
Rattacher chaque paiement CPAM (IJSS) recu pour un salarie aux CENTRES
ANALYTIQUES de ce salarie, afin de faciliter la saisie dans QUADRA.

Sources de reference (lues EN LIGNE dans les classeurs Google Sheets)
--------------------------------------------------------------------
                     ANALYTIQUE 2025                ANALYTIQUE 2026
  repartition        Repartition_analytique_2025    Repartition_analytique_2026
  ajustement         Ajustement                     Ajustement
  personnel          PERSONNEL (QUADRA)             Base_Personnel_SILAE

Les quatre premieres colonnes des onglets de repartition sont :
    Mois (MM/AAAA) | Nom et prenom(s) | Dispositif | Taux
Le centre "Ajustement", dans son propre onglet, porte des taux NEGATIFS qui
ramenent chaque salarie a 100 % sur le mois.

A chaque execution le module telecharge les onglets, controle que chaque
couple (Mois, Salarie) totalise bien 100 %, puis ecrit un cache local
(`cache/repartition_analytique.csv` et `cache/personnel_dunes.csv`). Si le reseau
ou le partage du classeur est indisponible, le cache prend le relais.

Identite des salaries
---------------------
Le MATRICULE est la cle pivot : il est commun a QUADRA 2025 et a SILAE 2026,
ce qui permet de rattacher au meme salarie ses repartitions 2025 et 2026 meme
si le libelle change d'un classeur a l'autre.

Un salarie peut avoir plusieurs contrats sur une periode (nouvelle mission le
lendemain, changement d'etablissement) : autant de lignes que de contrats dans
SILAE. Le module les consolide -> date d'entree = la plus ancienne trouvee
(donc celle de QUADRA 2025 pour les anciens, la colonne SILAE "Date du
changement" valant 01/01/2026 pour tout le monde lors de la bascule) ;
presence / sortie determinees par le classeur LE PLUS RECENT, puisqu'un
salarie parti en 2026 n'a evidemment pas de date de sortie dans QUADRA 2025.

Rapprochement des noms CPAM
---------------------------
Les libelles CPAM et les libelles paie ne sont pas identiques
(exemples, noms fictifs) :
  'DUPONT-MARTIN ANNE'   <-> 'DUPONT MARTIN Anne'      (tiret)
  'LEFEVRE RENEE'        <-> 'LEFEVRE Renee'           (accents)
  'BERNARD CLAIRE'       <-> 'BERNARD-ROUX Claire'     (nom compose)
  'PETIT MARIE LOUISE'   <-> 'PETIT Marie'             (2e prenom)
  'DUBOIS MATHIEU'       <-> 'DUBOIS Mathieu'          (coquille)
  'LEROY MARIE'          <-> 'DUPONT Marie nee LEROY'  (nom de naissance)
Rapprochement par jeux de jetons normalises (sans accents, sans ponctuation,
ordre indifferent), avec inclusion, nom de naissance ("nee ...") et
similarite floue. Rien n'est ventile au hasard : un nom incertain est remonte
dans l'onglet "Controle matching" avec les 3 candidats les plus proches, et
se resout en ajoutant une ligne a `config/alias_salaries.csv`.

Auteur : Frederic (DUNES)
"""

from __future__ import annotations

import os
import re
import csv
import logging
import datetime
import unicodedata
from difflib import SequenceMatcher
from typing import Optional, Iterable

import pandas as pd

logger = logging.getLogger("net_entreprise.analytique")

# ---------------------------------------------------------------------------
# Fichiers locaux
# ---------------------------------------------------------------------------
DOSSIER_COURANT = os.path.dirname(os.path.abspath(__file__))
DOSSIER_CONFIG  = os.path.join(DOSSIER_COURANT, "config")
DOSSIER_CACHE   = os.path.join(DOSSIER_COURANT, "cache")
DOSSIER_DONNEES = os.path.join(DOSSIER_COURANT, "data")


def _chemin(dossier: str, nom: str) -> str:
    """Chemin dans le sous-dossier dedie ; repli sur la racine (ancienne version)."""
    cible = os.path.join(dossier, nom)
    if not os.path.exists(cible):
        ancien = os.path.join(DOSSIER_COURANT, nom)
        if os.path.exists(ancien):
            return ancien
    os.makedirs(dossier, exist_ok=True)
    return cible


CSV_ALIAS         = _chemin(DOSSIER_CONFIG, "alias_salaries.csv")
CACHE_REPARTITION = _chemin(DOSSIER_CACHE, "repartition_analytique.csv")
CACHE_PERSONNEL   = _chemin(DOSSIER_CACHE, "personnel_dunes.csv")

# ---------------------------------------------------------------------------
# Sources en ligne : classeur + gid de chaque onglet
# ---------------------------------------------------------------------------
SOURCES: dict[int, dict] = {
    2025: {
        "classeur":         "1MPxWzgi9zAiXtREocfpf75FU9f17X2Xvtdk0tv7aDfU",
        "repartition":      570305692,      # Répartition_analytique_2025
        "ajustement":       198171407,      # Ajustement
        "personnel":        1000484018,     # PERSONNEL  (export QUADRA)
        "format_personnel": "QUADRA",
    },
    2026: {
        "classeur":         "1CTIEwPS-hn208H6vXTeA3Sx2gTTNExa5QUKfvvsX8g0",
        "repartition":      570305692,      # Répartition_analytique_2026
        "ajustement":       1744654119,     # Ajustement
        "personnel":        2092613264,     # Base_Personnel_SILAE
        "format_personnel": "SILAE",
    },
}

ANNEES_DEFAUT = tuple(sorted(SOURCES))

CENTRE_AJUSTEMENT = "Ajustement"

# Seuils de rapprochement (score sur 100)
SCORE_AUTO        = 88     # >= : rapprochement applique automatiquement
SCORE_PROPOSITION = 75     # >= : candidat propose, mais NON applique
ECART_AMBIGUITE   = 2      # si les 2 meilleurs candidats sont a moins de 2 pts
SCORE_FUSION      = 92     # seuil pour rattacher un libelle de repartition
                           # a une fiche du personnel

# Alias en dur (complete par alias_salaries.csv, prioritaire).
ALIAS_SALARIES: dict[str, str] = {}

_MOIS = re.compile(r"^\d{2}/\d{4}$")
_RE_NEE = re.compile(r"^(?P<usage>.+?)\s+n[ée]+e?\s+(?P<naissance>.+?)$", re.IGNORECASE)


# ===========================================================================
# Normalisation et similarite des noms
# ===========================================================================
_PONCTUATION = re.compile(r"[^A-Z0-9 ]")
_ESPACES     = re.compile(r"\s+")


def normaliser_nom(valeur) -> str:
    """'Sabatier-Amat  Élisa' -> 'SABATIER AMAT ELISA'."""
    if valeur is None:
        return ""
    try:
        if pd.isna(valeur):
            return ""
    except (TypeError, ValueError):
        pass
    texte = unicodedata.normalize("NFD", str(valeur))
    texte = "".join(c for c in texte if unicodedata.category(c) != "Mn").upper()
    for sep in ("-", "'", "’", ".", "_", "/", ","):
        texte = texte.replace(sep, " ")
    texte = _PONCTUATION.sub(" ", texte)
    return _ESPACES.sub(" ", texte).strip()


def _jetons(nom_normalise: str) -> list[str]:
    return [j for j in nom_normalise.split(" ") if j]


def _score_flou(jetons_a: list[str], jetons_b: list[str]) -> float:
    """Appariement glouton jeton a jeton, moyenne penalisee par l'ecart de taille."""
    if not jetons_a or not jetons_b:
        return 0.0
    courts, longs = (jetons_a, jetons_b) if len(jetons_a) <= len(jetons_b) else (jetons_b, jetons_a)
    restants, total = list(longs), 0.0
    for jeton in courts:
        if not restants:
            break
        meilleur, indice = 0.0, 0
        for i, candidat in enumerate(restants):
            score = SequenceMatcher(None, jeton, candidat).ratio()
            if score > meilleur:
                meilleur, indice = score, i
        total += meilleur
        restants.pop(indice)
    penalite = 1.0 - 0.05 * (len(longs) - len(courts))
    return max(0.0, (total / len(courts)) * penalite)


def apparier_nom(nom_cpam: str, index_noms: dict[str, str],
                 alias: Optional[dict[str, str]] = None,
                 jetons_naissance: Optional[dict[str, set]] = None,
                 libelles: Optional[dict[str, str]] = None) -> dict:
    """
    Rapproche un nom CPAM d'un salarie du referentiel.

    index_noms       : {nom_normalise -> cle_salarie}   (tous libelles connus)
    jetons_naissance : {cle_salarie -> {jetons du nom de naissance}}
    libelles         : {cle_salarie -> libelle d'affichage}

    Retourne {cle, salarie_analytique, score, methode, statut, candidats}.
    statut : 'OK' | 'A CONFIRMER' | 'AMBIGU' | 'NON TROUVE'
    """
    alias = alias or {}
    jetons_naissance = jetons_naissance or {}
    libelles = libelles or {}
    cible = normaliser_nom(nom_cpam)
    vide = {"cle": None, "salarie_analytique": None, "score": 0,
            "methode": "", "statut": "NON TROUVE", "candidats": ""}
    if not cible:
        return vide

    # 1. Alias explicite -> prioritaire absolu
    for table in (alias, ALIAS_SALARIES):
        if cible in table:
            vise = normaliser_nom(table[cible])
            cle = index_noms.get(vise)
            if cle:
                return {"cle": cle, "salarie_analytique": libelles.get(cle, table[cible]),
                        "score": 100, "methode": "alias", "statut": "OK", "candidats": ""}
            logger.warning("Alias '%s' -> '%s' : libelle inconnu du referentiel",
                           nom_cpam, table[cible])

    ens_cible = set(_jetons(cible))
    jetons_cible = _jetons(cible)

    # meilleur score par cle de salarie (un salarie a plusieurs libelles)
    par_cle: dict[str, tuple[float, str]] = {}
    for norm, cle in index_noms.items():
        ens_ref = set(_jetons(norm))
        naissance = jetons_naissance.get(cle) or set()

        if norm == cible:
            score, methode = 100.0, "exact"
        elif ens_cible == ens_ref:
            score, methode = 99.0, "jetons"
        elif ens_cible and ens_ref and (ens_cible <= ens_ref or ens_ref <= ens_cible):
            commun = len(ens_cible & ens_ref)
            score = 88.0 + 10.0 * (commun / max(len(ens_cible), len(ens_ref)))
            methode = "inclusion"
        elif naissance and ens_cible and ens_cible <= (ens_ref | naissance) \
                and (ens_cible & naissance):
            etendu = ens_ref | naissance
            commun = len(ens_cible & etendu)
            score = 86.0 + 8.0 * (commun / max(len(ens_cible), len(etendu)))
            methode = "nom de naissance"
        else:
            score = _score_flou(jetons_cible, _jetons(norm)) * 100.0
            methode = "flou"

        precedent = par_cle.get(cle)
        if precedent is None or score > precedent[0]:
            par_cle[cle] = (score, methode)

    if not par_cle:
        return vide

    classement = sorted(((sc, me, cle) for cle, (sc, me) in par_cle.items()),
                        key=lambda x: (-x[0], libelles.get(x[2], x[2])))
    meilleur_score, meilleure_methode, meilleure_cle = classement[0]
    candidats = " | ".join(f"{libelles.get(cle, cle)} ({sc:.0f})"
                           for sc, _, cle in classement[:3])

    if len(classement) > 1 and (meilleur_score - classement[1][0]) < ECART_AMBIGUITE \
            and meilleur_score >= SCORE_PROPOSITION:
        return {"cle": None, "salarie_analytique": None,
                "score": round(meilleur_score), "methode": meilleure_methode,
                "statut": "AMBIGU", "candidats": candidats}

    if meilleur_score >= SCORE_AUTO:
        return {"cle": meilleure_cle,
                "salarie_analytique": libelles.get(meilleure_cle, meilleure_cle),
                "score": round(meilleur_score), "methode": meilleure_methode,
                "statut": "OK", "candidats": candidats}

    statut = "A CONFIRMER" if meilleur_score >= SCORE_PROPOSITION else "NON TROUVE"
    return {"cle": None, "salarie_analytique": None, "score": round(meilleur_score),
            "methode": "", "statut": statut, "candidats": candidats}


# ===========================================================================
# Lecture des sources
# ===========================================================================
def _to_float(valeur) -> float:
    if valeur is None:
        return 0.0
    if isinstance(valeur, (int, float)) and not isinstance(valeur, bool):
        try:
            return 0.0 if pd.isna(valeur) else float(valeur)
        except (TypeError, ValueError):
            return 0.0
    texte = str(valeur).strip().replace(" ", "").replace(" ", "")
    texte = texte.replace("€", "").replace("%", "").replace(",", ".")
    try:
        return float(texte)
    except ValueError:
        return 0.0


def _to_date(valeur) -> str:
    """Normalise une date en JJ/MM/AAAA. Chaine vide si illisible."""
    texte = "" if valeur is None else str(valeur).strip()
    if not texte or texte.lower() in ("nan", "nat", "none"):
        return ""
    for motif in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(texte, motif).strftime("%d/%m/%Y")
        except ValueError:
            continue
    try:
        return pd.to_datetime(texte, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return ""


def _cle_date(date_jjmmaaaa: str) -> tuple:
    try:
        j, m, a = date_jjmmaaaa.split("/")
        return (int(a), int(m), int(j))
    except Exception:
        return (0, 0, 0)


def _cle_mois(mois_mmaaaa: str) -> tuple:
    try:
        m, a = mois_mmaaaa.split("/")
        return (int(a), int(m))
    except Exception:
        return (0, 0)


def url_csv(classeur: str, gid: int) -> str:
    """URL d'export CSV d'un onglet (necessite un classeur partage par lien)."""
    return f"https://docs.google.com/spreadsheets/d/{classeur}/export?format=csv&gid={gid}"


def _lire_onglet(classeur: str, gid: int, libelle: str) -> "pd.DataFrame":
    """Telecharge un onglet en CSV brut (sans en-tete, tout en texte)."""
    url = url_csv(classeur, gid)
    try:
        df = pd.read_csv(url, header=None, dtype=str, keep_default_na=False,
                         on_bad_lines="skip")
    except Exception as exc:
        raise ConnectionError(f"Onglet '{libelle}' illisible ({exc}). "
                              "Verifiez le partage par lien du classeur.") from exc
    if df.empty:
        raise ConnectionError(f"Onglet '{libelle}' vide.")
    premiere = " ".join(str(v) for v in df.iloc[0].tolist())[:200].lower()
    if "<html" in premiere or "doctype" in premiere:
        raise PermissionError(
            f"Acces refuse a l'onglet '{libelle}' : le classeur n'est pas partage "
            "par lien (Partager > Tous les utilisateurs disposant du lien > Lecteur)."
        )
    return df


def _trouver_entete(df: "pd.DataFrame", motifs: Iterable[str]) -> int:
    """Indice de la premiere ligne contenant l'un des motifs (sinon 0)."""
    motifs = [normaliser_nom(m) for m in motifs]
    for i in range(min(len(df), 15)):
        ligne = [normaliser_nom(v) for v in df.iloc[i].tolist()]
        if any(m in ligne for m in motifs):
            return i
    return 0


def _colonne(entete: list[str], *motifs: str) -> Optional[int]:
    """Indice de la colonne dont l'en-tete correspond a l'un des motifs."""
    cibles = [normaliser_nom(m) for m in motifs]
    for i, valeur in enumerate(entete):
        norme = normaliser_nom(valeur)
        if norme and norme in cibles:
            return i
    for i, valeur in enumerate(entete):
        norme = normaliser_nom(valeur)
        if norme and any(norme.startswith(c) for c in cibles):
            return i
    return None


# ---------------------------------------------------------------------------
# Repartition analytique
# ---------------------------------------------------------------------------
def charger_repartition(annees: Optional[Iterable[int]] = None,
                        en_ligne: bool = True,
                        cache: Optional[str] = None) -> "pd.DataFrame":
    """
    Repartition consolidee : Mois | Salarie | Centre analytique | Taux | Annee.
    Fusionne, pour chaque annee, l'onglet de repartition ET l'onglet Ajustement.
    Bascule sur le cache local si la lecture en ligne echoue.
    """
    cache = cache or CACHE_REPARTITION
    annees = tuple(annees) if annees else ANNEES_DEFAUT
    colonnes = ["Mois", "Salarié", "Centre analytique", "Taux", "Année"]

    morceaux, erreurs = [], []
    if en_ligne:
        for annee in annees:
            source = SOURCES.get(annee)
            if not source:
                logger.warning("Annee %s inconnue dans SOURCES", annee)
                continue
            for role in ("repartition", "ajustement"):
                try:
                    brut = _lire_onglet(source["classeur"], source[role],
                                        f"{role} {annee}")
                except (ConnectionError, PermissionError) as exc:
                    erreurs.append(str(exc))
                    continue
                bloc = brut.iloc[:, :4].copy()
                bloc.columns = ["Mois", "Salarié", "Centre analytique", "Taux"]
                bloc["Mois"] = bloc["Mois"].astype(str).str.strip()
                bloc = bloc[bloc["Mois"].str.match(_MOIS, na=False)]
                bloc["Année"] = annee
                morceaux.append(bloc)
                logger.info("Repartition %s / %s : %d ligne(s)", annee, role, len(bloc))

    if morceaux:
        df = pd.concat(morceaux, ignore_index=True)
    else:
        if erreurs:
            logger.warning("Lecture en ligne impossible : %s", " | ".join(erreurs[:2]))
        if not os.path.exists(cache):
            raise FileNotFoundError(
                "Referentiel analytique indisponible : ni en ligne, ni en cache "
                f"({cache}). Verifiez la connexion ou le partage des classeurs."
            )
        logger.warning("Referentiel analytique : utilisation du cache local %s",
                       os.path.basename(cache))
        df = pd.read_csv(cache, sep=";", encoding="utf-8-sig", dtype=str)
        if "Année" not in df.columns:
            df["Année"] = ""

    for col in ("Salarié", "Centre analytique"):
        df[col] = df[col].astype(str).str.strip()
    df["Mois"] = df["Mois"].astype(str).str.strip()
    df["Taux"] = df["Taux"].map(_to_float)
    df = df[df["Mois"].str.match(_MOIS, na=False)]
    df = df[df["Salarié"].ne("") & df["Centre analytique"].ne("")]

    doublons = df[df.duplicated(subset=["Mois", "Salarié", "Centre analytique", "Taux"],
                                keep="first")]
    if len(doublons):
        detail = ", ".join(f"{d['Mois']} {d['Salarié']} / {d['Centre analytique']}"
                           for d in doublons.head(6).to_dict("records"))
        logger.warning("Referentiel : %d ligne(s) en double ignoree(s) -> %s%s",
                       len(doublons), detail, " ..." if len(doublons) > 6 else "")
        df = df.drop_duplicates(subset=["Mois", "Salarié", "Centre analytique", "Taux"])
    df = (df.groupby(["Mois", "Salarié", "Centre analytique"], as_index=False, sort=False)
            .agg({"Taux": "sum", "Année": "first"}))

    # Controle : chaque couple (Mois, Salarie) doit totaliser 100 %
    totaux = df.groupby(["Mois", "Salarié"])["Taux"].sum()
    hors_100 = totaux[(totaux - 1).abs() > 0.001]
    if len(hors_100):
        logger.warning("%d couple(s) (Mois, Salarie) ne totalisent pas 100%% : %s",
                       len(hors_100),
                       ", ".join(f"{m} {s} = {v:.3f}" for (m, s), v in hors_100.head(8).items()))
    else:
        logger.info("Controle OK : les %d couples (Mois, Salarie) totalisent 100%%",
                    len(totaux))

    if morceaux and cache:
        try:
            export = df.copy()
            export["Taux"] = export["Taux"].map(lambda v: f"{v:g}".replace(".", ","))
            export[colonnes].to_csv(cache, sep=";", index=False, encoding="utf-8-sig")
        except Exception as exc:                    # noqa: BLE001
            logger.debug("Cache repartition non ecrit : %s", exc)

    logger.info("Referentiel analytique : %d lignes | %d salaries | mois %s -> %s",
                len(df), df["Salarié"].nunique(),
                min(df["Mois"], key=_cle_mois), max(df["Mois"], key=_cle_mois))
    return df[colonnes]


# ---------------------------------------------------------------------------
# Personnel
# ---------------------------------------------------------------------------
def _personnel_silae(brut: "pd.DataFrame", annee: int) -> list[dict]:
    idx = _trouver_entete(brut, ["Matricule", "Salarié"])
    entete = [str(v) for v in brut.iloc[idx].tolist()]
    corps = brut.iloc[idx + 1:]
    c_mat = _colonne(entete, "Matricule") or 0
    c_nom = _colonne(entete, "Salarié", "Salarie", "Nom et prénom(s)")
    c_ent = _colonne(entete, "Date entrée", "Date d'entrée")
    c_sor = _colonne(entete, "Date sortie", "Date de sortie")
    c_chg = _colonne(entete, "Date du changement")
    if c_nom is None:
        c_nom = 1
    # Les en-tetes "Date entree"/"Date sortie" sont parfois fusionnes donc vides :
    # la colonne de date suit alors immediatement "Entree/Sortie dans periode".
    if c_ent is None:
        reference = _colonne(entete, "Entrée dans période")
        c_ent = reference + 1 if reference is not None else None
    if c_sor is None:
        reference = _colonne(entete, "Sortie dans période")
        c_sor = reference + 1 if reference is not None else None
    fiches = []
    for ligne in corps.itertuples(index=False):
        valeurs = list(ligne)
        def cell(i):
            return str(valeurs[i]).strip() if i is not None and i < len(valeurs) else ""
        matricule = cell(c_mat)
        nom = cell(c_nom)
        if not re.match(r"^\d+$", matricule) or not nom:
            continue
        fiches.append({
            "Matricule": matricule,
            "libelle": nom,
            "entree": _to_date(cell(c_ent)),
            "sortie": _to_date(cell(c_sor)),
            "changement": _to_date(cell(c_chg)),
            "annee": annee,
            "source": f"SILAE {annee}",
        })
    return fiches


def _personnel_quadra(brut: "pd.DataFrame", annee: int) -> list[dict]:
    idx = _trouver_entete(brut, ["Nom et prénom(s)", "ID_salarié", "Numéro"])
    entete = [str(v) for v in brut.iloc[idx].tolist()]
    corps = brut.iloc[idx + 1:]
    c_num = _colonne(entete, "Numéro", "Matricule") or 0
    c_nom = _colonne(entete, "Nom")
    c_pre = _colonne(entete, "Prénom")
    c_cpl = _colonne(entete, "Nom et prénom(s)")
    c_ent = _colonne(entete, "Date d'entrée", "Date entrée")
    c_sor = _colonne(entete, "Date de sortie", "Date sortie")
    # l'en-tete des colonnes de dates est souvent fusionne/vide : repli positionnel
    if c_cpl is None:
        c_cpl = 3
    if c_ent is None:
        c_ent = c_cpl + 2
    if c_sor is None:
        c_sor = c_ent + 1
    fiches = []
    for ligne in corps.itertuples(index=False):
        valeurs = list(ligne)
        def cell(i):
            return str(valeurs[i]).strip() if i is not None and i < len(valeurs) else ""
        matricule = cell(c_num)
        complet = cell(c_cpl)
        nom, prenom = cell(c_nom), cell(c_pre)
        if not re.match(r"^\d+$", matricule):
            continue
        libelle = complet or f"{nom} {prenom}".strip()
        if not libelle:
            continue
        autres = {x for x in (complet, f"{nom} {prenom}".strip()) if x and x != libelle}
        fiches.append({
            "Matricule": matricule,
            "libelle": libelle,
            "autres": autres,
            "entree": _to_date(cell(c_ent)),
            "sortie": _to_date(cell(c_sor)),
            "changement": "",
            "annee": annee,
            "source": f"QUADRA {annee}",
        })
    return fiches


def charger_personnel(annees: Optional[Iterable[int]] = None,
                      en_ligne: bool = True,
                      cache: Optional[str] = None) -> "pd.DataFrame":
    """
    Personnel consolide, une ligne par MATRICULE :
        Matricule | Salarié | Nom de naissance | Autres libellés |
        Date entrée | Date sortie | Nb contrats | Sources

    - date d'entree = la plus ancienne trouvee, tous classeurs confondus
      (donc la date QUADRA pour les salaries anterieurs a la bascule SILAE) ;
    - date de sortie = la plus recente, VIDE tant qu'un contrat reste ouvert ;
    - "Nb contrats" = nombre de lignes de contrat rencontrees sur la periode.
    Seules les colonnes utiles au rapprochement sont conservees.
    """
    cache = cache or CACHE_PERSONNEL
    annees = tuple(annees) if annees else ANNEES_DEFAUT
    colonnes = ["Matricule", "Salarié", "Nom de naissance", "Autres libellés",
                "Date entrée", "Date sortie", "Nb contrats", "Sources"]

    fiches: list[dict] = []
    if en_ligne:
        for annee in annees:
            source = SOURCES.get(annee)
            if not source or not source.get("personnel"):
                continue
            try:
                brut = _lire_onglet(source["classeur"], source["personnel"],
                                    f"personnel {annee}")
            except (ConnectionError, PermissionError) as exc:
                logger.warning("Personnel %s indisponible : %s", annee, exc)
                continue
            if source.get("format_personnel") == "QUADRA":
                fiches += _personnel_quadra(brut, annee)
            else:
                fiches += _personnel_silae(brut, annee)

    if not fiches:
        if os.path.exists(cache):
            logger.warning("Personnel : utilisation du cache local %s",
                           os.path.basename(cache))
            df = pd.read_csv(cache, sep=";", encoding="utf-8-sig", dtype=str).fillna("")
            for col in colonnes:
                if col not in df.columns:
                    df[col] = ""
            return df[colonnes]
        logger.info("Personnel indisponible : rapprochement par nom de naissance "
                    "et dates de sortie desactives")
        return pd.DataFrame(columns=colonnes)

    consolide: dict[str, dict] = {}
    for fiche in fiches:
        matricule = fiche["Matricule"]
        brut_libelle = fiche["libelle"]
        trouve = _RE_NEE.match(brut_libelle)
        usage = trouve.group("usage").strip() if trouve else brut_libelle
        naissance = trouve.group("naissance").strip() if trouve else ""

        cible = consolide.setdefault(matricule, {
            "Matricule": matricule, "libelles": [], "naissance": "",
            "autres": set(), "entree": "", "sortie": "",
            "contrats": 0, "sources": [], "par_annee": {},
        })
        cible["libelles"].append(usage)
        cible["autres"].update(fiche.get("autres") or set())
        if naissance and not cible["naissance"]:
            cible["naissance"] = naissance
        cible["contrats"] += 1
        if fiche["source"] not in cible["sources"]:
            cible["sources"].append(fiche["source"])

        entree = fiche["entree"]
        if entree and (not cible["entree"] or _cle_date(entree) < _cle_date(cible["entree"])):
            cible["entree"] = entree

        # Le statut (sorti / present) est donne par le classeur LE PLUS RECENT :
        # un salarie parti en 2026 n'a evidemment pas de date de sortie en 2025.
        sortie = fiche["sortie"]
        etat = cible["par_annee"].setdefault(fiche.get("annee", 0),
                                             {"ouvert": False, "sortie": ""})
        if not sortie:
            etat["ouvert"] = True
        elif not etat["sortie"] or _cle_date(sortie) > _cle_date(etat["sortie"]):
            etat["sortie"] = sortie
        if sortie and (not cible["sortie"] or _cle_date(sortie) > _cle_date(cible["sortie"])):
            cible["sortie"] = sortie

    lignes = []
    for fiche in consolide.values():
        # libelle d'affichage : le dernier rencontre (SILAE passe apres QUADRA)
        libelle = fiche["libelles"][-1]
        autres = {x for x in fiche["libelles"][:-1]} | fiche["autres"]
        autres = {x for x in autres if normaliser_nom(x) != normaliser_nom(libelle)}
        recent = fiche["par_annee"].get(max(fiche["par_annee"]), {}) if fiche["par_annee"] else {}
        sortie_finale = "" if recent.get("ouvert") else (recent.get("sortie") or fiche["sortie"])
        lignes.append({
            "Matricule": fiche["Matricule"],
            "Salarié": libelle,
            "Nom de naissance": fiche["naissance"],
            "Autres libellés": " | ".join(sorted(autres)),
            "Date entrée": fiche["entree"],
            "Date sortie": sortie_finale,
            "Nb contrats": fiche["contrats"],
            "Sources": " + ".join(fiche["sources"]),
        })

    df = pd.DataFrame(lignes, columns=colonnes).sort_values("Salarié").reset_index(drop=True)

    if cache:
        try:
            df.to_csv(cache, sep=";", index=False, encoding="utf-8-sig")
        except Exception as exc:                    # noqa: BLE001
            logger.debug("Cache personnel non ecrit : %s", exc)

    logger.info("Personnel : %d salaries | %d sorti(s) | %d nom(s) de naissance | "
                "%d multi-contrats",
                len(df), int(df["Date sortie"].ne("").sum()),
                int(df["Nom de naissance"].ne("").sum()),
                int((df["Nb contrats"] > 1).sum()))
    return df


# ---------------------------------------------------------------------------
# Alias manuels
# ---------------------------------------------------------------------------
def charger_alias(chemin: Optional[str] = None) -> dict[str, str]:
    """Charge alias_salaries.csv ('Nom CPAM;Nom analytique'). Absent = {}."""
    chemin = chemin or CSV_ALIAS
    if not os.path.exists(chemin):
        return {}
    table: dict[str, str] = {}
    with open(chemin, encoding="utf-8-sig", newline="") as fh:
        for ligne in csv.reader(fh, delimiter=";"):
            if len(ligne) < 2:
                continue
            gauche, droite = ligne[0].strip(), ligne[1].strip()
            if not gauche or not droite or gauche.startswith("#") \
                    or gauche.lower().startswith("nom cpam"):
                continue
            table[normaliser_nom(gauche)] = droite
    logger.info("Alias salaries charges : %d", len(table))
    return table


# ===========================================================================
# Referentiel consolide
# ===========================================================================
class Referentiel:
    """Repartition + personnel + index de noms, agreges par salarie."""

    def __init__(self, repartition: "pd.DataFrame", personnel: "pd.DataFrame",
                 alias: Optional[dict[str, str]] = None):
        self.repartition = repartition
        self.personnel = personnel
        self.alias = alias or {}

        self.index_noms: dict[str, str] = {}
        self.libelles: dict[str, str] = {}
        self.jetons_naissance: dict[str, set] = {}
        self.infos: dict[str, dict] = {}
        self.par_cle: dict[str, dict[str, list[tuple[str, float]]]] = {}

        self._indexer_personnel()
        self._indexer_repartition()

    # -- index des fiches du personnel ---------------------------------
    def _indexer_personnel(self) -> None:
        for fiche in self.personnel.to_dict("records"):
            cle = f"M{fiche['Matricule']}"
            libelle = str(fiche["Salarié"]).strip()
            self.libelles[cle] = libelle
            self.infos[cle] = {
                "matricule": str(fiche["Matricule"]),
                "entree": fiche.get("Date entrée", ""),
                "sortie": fiche.get("Date sortie", ""),
                "contrats": fiche.get("Nb contrats", ""),
                "sources": fiche.get("Sources", ""),
            }
            formes = [libelle] + [x.strip() for x in
                                  str(fiche.get("Autres libellés", "")).split("|") if x.strip()]
            for forme in formes:
                norme = normaliser_nom(forme)
                if norme:
                    self.index_noms.setdefault(norme, cle)
            naissance = normaliser_nom(fiche.get("Nom de naissance", ""))
            if naissance:
                self.jetons_naissance.setdefault(cle, set()).update(_jetons(naissance))

    # -- rattachement des libelles de repartition -----------------------
    def _resoudre_libelle(self, libelle: str) -> str:
        norme = normaliser_nom(libelle)
        if norme in self.index_noms:
            return self.index_noms[norme]
        if self.index_noms:
            trouve = apparier_nom(libelle, self.index_noms, {},
                                  self.jetons_naissance, self.libelles)
            if trouve["cle"] and trouve["score"] >= SCORE_FUSION:
                self.index_noms.setdefault(norme, trouve["cle"])
                return trouve["cle"]
        cle = f"L:{norme}"                      # salarie hors fichier du personnel
        self.index_noms.setdefault(norme, cle)
        self.libelles.setdefault(cle, libelle)
        self.infos.setdefault(cle, {"matricule": "", "entree": "", "sortie": "",
                                    "contrats": "", "sources": ""})
        return cle

    def _indexer_repartition(self) -> None:
        resolution: dict[str, str] = {}
        for libelle in self.repartition["Salarié"].unique():
            resolution[libelle] = self._resoudre_libelle(libelle)
        orphelins = sorted({lib for lib, cle in resolution.items() if cle.startswith("L:")})
        if orphelins:
            logger.warning("%d libelle(s) de repartition sans fiche du personnel : %s%s",
                           len(orphelins), ", ".join(orphelins[:10]),
                           " ..." if len(orphelins) > 10 else "")
        for ligne in self.repartition.to_dict("records"):
            cle = resolution[ligne["Salarié"]]
            self.par_cle.setdefault(cle, {}).setdefault(ligne["Mois"], []) \
                        .append((ligne["Centre analytique"], float(ligne["Taux"])))
            # le libelle de repartition le plus recent sert d'affichage de secours
            self.libelles.setdefault(cle, ligne["Salarié"])

    # -- acces ----------------------------------------------------------
    def mois_disponibles(self, cle: str) -> list[str]:
        return list(self.par_cle.get(cle, {}).keys())

    def centres(self, cle: str, mois: str) -> list[tuple[str, float]]:
        return self.par_cle.get(cle, {}).get(mois, [])


def construire_referentiel(annees: Optional[Iterable[int]] = None,
                           en_ligne: bool = True,
                           chemin_alias: Optional[str] = None) -> Referentiel:
    """Telecharge (ou relit en cache) la repartition et le personnel."""
    repartition = charger_repartition(annees=annees, en_ligne=en_ligne)
    personnel = charger_personnel(annees=annees, en_ligne=en_ligne)
    alias = charger_alias(chemin_alias)
    return Referentiel(repartition, personnel, alias)


# ===========================================================================
# Ventilation
# ===========================================================================
def _mois_depuis_cle(cle) -> str:
    """'AS_05/2026@06/2026_00117' -> '05/2026' (mois de DateDeb)."""
    trouve = re.search(r"(\d{2}/\d{4})", str(cle or ""))
    return trouve.group(1) if trouve else ""


def _choisir_mois(mois_voulu: str, mois_dispo: list[str]) -> tuple[str, str]:
    """Mois retenu + commentaire. Repli sur le mois disponible le plus proche."""
    if not mois_dispo:
        return "", "aucun mois disponible"
    if mois_voulu in mois_dispo:
        return mois_voulu, ""
    tries = sorted(mois_dispo, key=_cle_mois)
    if not mois_voulu:
        return tries[-1], f"mois indetermine -> {tries[-1]}"
    reference = _cle_mois(mois_voulu)
    anterieurs = [m for m in tries if _cle_mois(m) <= reference]
    retenu = anterieurs[-1] if anterieurs else tries[0]
    return retenu, f"{mois_voulu} absent -> {retenu}"


def ventiler_montant(montant: float,
                     lignes: list[tuple[str, float]]) -> list[tuple[str, float, float]]:
    """
    Repartit `montant` sur [(centre, taux), ...].
    Taux normalises sur leur somme, arrondi au centime, ecart d'arrondi reporte
    sur la plus grosse ligne : la somme ventilee egale STRICTEMENT le montant.
    """
    total_taux = sum(t for _, t in lignes)
    if not lignes or abs(total_taux) < 1e-9:
        return []
    resultat = [(centre, taux, round(montant * taux / total_taux, 2))
                for centre, taux in lignes]
    ecart = round(montant - sum(v for _, _, v in resultat), 2)
    if abs(ecart) >= 0.01:
        i = max(range(len(resultat)), key=lambda k: abs(resultat[k][2]))
        centre, taux, valeur = resultat[i]
        resultat[i] = (centre, taux, round(valeur + ecart, 2))
    return resultat


def construire_ventilation_analytique(
    recap: "pd.DataFrame",
    referentiel: Optional[Referentiel] = None,
    annees: Optional[Iterable[int]] = None,
    en_ligne: bool = True,
    chemin_alias: Optional[str] = None,
    colonne_salarie: str = "Salarié",
    colonne_montant: str = "Montant",
    colonne_cle: str = "Clé",
) -> tuple["pd.DataFrame", "pd.DataFrame"]:
    """
    Entree : le recapitulatif CPAM par salarie (Salarie, Montant, Cle).
    Sortie : (ventilation, controle)
    """
    colonnes_vent = ["Salarié", "Salarié analytique", "Matricule", "Mois analytique",
                     "Clé", "Centre analytique", "Taux", "Montant ventilé", "Montant total"]
    colonnes_ctrl = ["Salarié", "Salarié analytique", "Matricule", "Date entrée",
                     "Date sortie", "Statut", "Score", "Méthode", "Mois demandé",
                     "Mois utilisé", "Nb centres", "Total taux", "Montant CPAM",
                     "Montant ventilé", "Candidats proches", "Commentaire"]

    if recap is None or len(recap) == 0:
        return pd.DataFrame(columns=colonnes_vent), pd.DataFrame(columns=colonnes_ctrl)

    if referentiel is None:
        referentiel = construire_referentiel(annees=annees, en_ligne=en_ligne,
                                             chemin_alias=chemin_alias)

    cache_appariement: dict[str, dict] = {}
    lignes_vent, lignes_ctrl = [], []

    for valeurs in recap.to_dict("records"):
        nom_cpam = str(valeurs.get(colonne_salarie, "") or "").strip()
        if not nom_cpam:
            continue
        montant = _to_float(valeurs.get(colonne_montant, 0))
        cle_paiement = valeurs.get(colonne_cle, "")

        if nom_cpam not in cache_appariement:
            cache_appariement[nom_cpam] = apparier_nom(
                nom_cpam, referentiel.index_noms, referentiel.alias,
                referentiel.jetons_naissance, referentiel.libelles)
        appariement = cache_appariement[nom_cpam]

        mois_voulu = _mois_depuis_cle(cle_paiement)
        cle_salarie = appariement["cle"]

        if not cle_salarie:
            lignes_ctrl.append({
                "Salarié": nom_cpam, "Salarié analytique": "", "Matricule": "",
                "Date entrée": "", "Date sortie": "",
                "Statut": appariement["statut"], "Score": appariement["score"],
                "Méthode": appariement["methode"], "Mois demandé": mois_voulu,
                "Mois utilisé": "", "Nb centres": 0, "Total taux": 0,
                "Montant CPAM": round(montant, 2), "Montant ventilé": 0,
                "Candidats proches": appariement["candidats"],
                "Commentaire": "A rattacher manuellement (voir config/alias_salaries.csv)",
            })
            continue

        libelle = referentiel.libelles.get(cle_salarie, cle_salarie)
        infos = referentiel.infos.get(cle_salarie, {})
        matricule = infos.get("matricule", "")
        date_entree = infos.get("entree", "")
        date_sortie = infos.get("sortie", "")

        mois_retenu, commentaire = _choisir_mois(mois_voulu,
                                                 referentiel.mois_disponibles(cle_salarie))
        centres = referentiel.centres(cle_salarie, mois_retenu)

        # Salarie sorti : la CPAM arrive souvent apres le depart
        if date_sortie and mois_voulu:
            try:
                _, mois_s, annee_s = date_sortie.split("/")
                if (int(annee_s), int(mois_s)) < _cle_mois(mois_voulu):
                    commentaire = "; ".join(x for x in
                                            (f"sorti(e) le {date_sortie}", commentaire) if x)
            except ValueError:
                pass

        if not centres:
            lignes_ctrl.append({
                "Salarié": nom_cpam, "Salarié analytique": libelle,
                "Matricule": matricule, "Date entrée": date_entree,
                "Date sortie": date_sortie, "Statut": "SANS REPARTITION",
                "Score": appariement["score"], "Méthode": appariement["methode"],
                "Mois demandé": mois_voulu, "Mois utilisé": mois_retenu,
                "Nb centres": 0, "Total taux": 0,
                "Montant CPAM": round(montant, 2), "Montant ventilé": 0,
                "Candidats proches": appariement["candidats"],
                "Commentaire": commentaire or "aucun centre pour ce mois",
            })
            continue

        ventile = ventiler_montant(montant, centres)
        for centre, taux, valeur in ventile:
            lignes_vent.append({
                "Salarié": nom_cpam, "Salarié analytique": libelle,
                "Matricule": matricule, "Mois analytique": mois_retenu,
                "Clé": cle_paiement, "Centre analytique": centre,
                "Taux": round(taux, 6), "Montant ventilé": valeur,
                "Montant total": round(montant, 2),
            })

        total_taux = sum(t for _, t in centres)
        alerte = (f"ATTENTION total taux = {total_taux:.4f} (renormalise a 100%)"
                  if abs(total_taux - 1.0) > 0.001 else "")
        lignes_ctrl.append({
            "Salarié": nom_cpam, "Salarié analytique": libelle,
            "Matricule": matricule, "Date entrée": date_entree,
            "Date sortie": date_sortie, "Statut": appariement["statut"],
            "Score": appariement["score"], "Méthode": appariement["methode"],
            "Mois demandé": mois_voulu, "Mois utilisé": mois_retenu,
            "Nb centres": len(centres), "Total taux": round(total_taux, 4),
            "Montant CPAM": round(montant, 2),
            "Montant ventilé": round(sum(v for _, _, v in ventile), 2),
            "Candidats proches": "" if appariement["methode"] in ("exact", "alias")
                                 else appariement["candidats"],
            "Commentaire": "; ".join(x for x in (commentaire, alerte) if x),
        })

    df_vent = pd.DataFrame(lignes_vent, columns=colonnes_vent)
    df_ctrl = pd.DataFrame(lignes_ctrl, columns=colonnes_ctrl)

    if not df_ctrl.empty:
        ordre = {"NON TROUVE": 0, "AMBIGU": 1, "A CONFIRMER": 2,
                 "SANS REPARTITION": 3, "OK": 4}
        df_ctrl = (df_ctrl.assign(_o=df_ctrl["Statut"].map(ordre).fillna(9))
                          .sort_values(["_o", "Salarié"])
                          .drop(columns="_o").reset_index(drop=True))
        logger.info("Ventilation analytique : %d lignes | %d paiements | %d a verifier",
                    len(df_vent), len(df_ctrl), int((df_ctrl["Statut"] != "OK").sum()))

    return df_vent, df_ctrl


def recap_par_centre(df_vent: "pd.DataFrame") -> "pd.DataFrame":
    """Total des IJSS par centre analytique (controle global de saisie)."""
    colonnes = ["Centre analytique", "Montant ventilé", "Nb lignes"]
    if df_vent is None or df_vent.empty:
        return pd.DataFrame(columns=colonnes)
    recap = (df_vent.groupby("Centre analytique", as_index=False)
                    .agg(**{"Montant ventilé": ("Montant ventilé", "sum"),
                            "Nb lignes": ("Montant ventilé", "size")}))
    recap["Montant ventilé"] = recap["Montant ventilé"].round(2)
    return recap.sort_values("Montant ventilé", ascending=False).reset_index(drop=True)


# ===========================================================================
# Utilitaires
# ===========================================================================
def rafraichir_cache(annees: Optional[Iterable[int]] = None) -> tuple[str, str]:
    """Force le telechargement des sources et reecrit les deux caches locaux."""
    charger_repartition(annees=annees, en_ligne=True)
    charger_personnel(annees=annees, en_ligne=True)
    logger.info("Caches a jour : %s | %s",
                os.path.basename(CACHE_REPARTITION), os.path.basename(CACHE_PERSONNEL))
    return CACHE_REPARTITION, CACHE_PERSONNEL


def diagnostic(annees: Optional[Iterable[int]] = None, en_ligne: bool = True) -> None:
    """Affiche un etat des sources : utile pour verifier acces et coherence."""
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ref = construire_referentiel(annees=annees, en_ligne=en_ligne)
    rep = ref.repartition
    print(f"\nRepartition : {len(rep)} lignes")
    print(rep.groupby("Année")["Mois"].agg(["count", "nunique"]))
    print(f"\nPersonnel   : {len(ref.personnel)} salaries")
    print(f"Index noms  : {len(ref.index_noms)} libelles -> {len(ref.par_cle)} salaries")
    sans_fiche = [c for c in ref.par_cle if c.startswith("L:")]
    if sans_fiche:
        print(f"\n{len(sans_fiche)} salarie(s) de la repartition sans fiche du personnel :")
        for cle in sans_fiche[:15]:
            print("   -", ref.libelles.get(cle, cle))


__all__ = [
    "SOURCES", "Referentiel", "construire_referentiel",
    "normaliser_nom", "apparier_nom",
    "charger_repartition", "charger_personnel", "charger_alias",
    "ventiler_montant", "construire_ventilation_analytique", "recap_par_centre",
    "rafraichir_cache", "diagnostic", "url_csv",
    "ALIAS_SALARIES", "CENTRE_AJUSTEMENT",
]


if __name__ == "__main__":
    diagnostic()
