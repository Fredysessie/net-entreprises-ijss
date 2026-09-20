"""
net_entreprise_final.py
========================
Module d'automatisation Net-Entreprises -> DSIJ / Attestations de salaire.

Corrections v4 :
  - Onglet "Paiements" : colonnes QUADRA "Caisse_short", "Cle_long" (identifiant
    du virement : {Journée BPIJ}@{Date}#{Montant reçu} €_{3 derniers chiffres du
    SIRET}) et "Libellé_QUADRA" (caisse abrégée + salarié(s) de la journée)

Corrections v3 :
  - Onglet "Paiements" :
      * Colonnes "SIREN/SIRET", "Numéro de Sécurité Sociale", "identifiant" -> numériques
      * Colonnes "SIREN/SIRET", "Date", "Numéro de Sécurité Sociale", "Type", "identifiant" centrées
  - Onglet "Détail des paiements" :
      * Colonnes "Identifiant_BPIJ", "NIR" -> numériques
      * Colonne "identifiant" (issue de la fusion avec Paiements) supprimée
      * Colonnes "Identifiant_BPIJ", "NIR", "Prestation_CodeNature", "DateDeb",
        "DateFin", "NB_jours", "Type" centrées
      * Nouvelle colonne "Clé" = Type_"DateDeb(MM/AAAA)"@"DateFin(MM/AAAA)"_"5 derniers chiffres SIRET"
  - Onglet "Récap par salarié" : récapitulatif par Salarié / Clé
    (colonnes : Salarié, Montant, Clé)

Corrections v2 (conservées) :
  - Colonne fantôme supprimée dans l'onglet "Paiements"
  - DateDeb / DateFin formatées en JJ/MM/AAAA (plus de timestamp)
  - Charte graphique DUNES appliquée à tous les onglets Excel
    · Rouge  #AF321E  -> en-têtes
    · Orange #F09426  -> lignes alternées
    · Police Arial, en-têtes blancs gras

Flux réel (reconstitué depuis DevTools) :
  ┌─ PORTAIL ──────────────────────────────────────────────────────────────┐
  │ 1. POST /auth/pass                         → 302 (cookies de session) │
  │    allow_redirects=True → suit srv/ → acces → declarations?rs=1       │
  │ 2. Parse declarations?rs=1                 → extrait href service      │
  │ 3. GET  /priv/accesServiceDeclaration/TOKEN/ID  → 303                 │
  │    La session requests suit le 303 vers :                              │
  │ 4. GET  /priv/formAccesServiceDeclarations?urlAcces=...&jeton=...      │
  │         → 200  (page intermédiaire avec formulaire auto-submit)        │
  │    On extrait : urlAcces, jeton (= JTD décodé)                        │
  └────────────────────────────────────────────────────────────────────────┘
  ┌─ DSIJ ─────────────────────────────────────────────────────────────────┐
  │ 5. POST /IhmHF/accesNetEntreprisesFrontal  data={jtd: jeton}  → 200   │
  │         (établit JSESSIONID DSIJ)                                      │
  │ 6. GET  /IhmHF/accueil.go                                    → 200    │
  │ 7. POST /IhmHF/connect.go                  data={jtn: jeton}  → 302   │
  │         → suit redirection vers chercherPaiement.go                   │
  │ 8. GET  /IhmHF/chercherPaiement.go                           → 200 ✅ │
  │ 9. POST /IhmHF/afficherPaiement.go         (filtres)         → 200    │
  │10. POST /IhmHF/gererTableauPaiementSalarie.go (z=300)        → 200    │
  └────────────────────────────────────────────────────────────────────────┘

Usage minimal :
    from net_entreprise_final import NetEntreprise

    ne = NetEntreprise(
        siret="VOTRE_SIRET",          # 14 chiffres
        nom="nom_declarant",
        prenom="PRENOM_DECLARANT",
        password="VotreMotDePasse",   # a lire depuis .env, jamais en dur
    )
    ne.login()
    ne.open_attestation()
    df = ne.get_tableau(date_debut="01/01/2026")
    ne.export_excel(df, "paiements_2026.xlsx")

Auteur : Frederic (DUNES)
Version : 1.0
"""

from __future__ import annotations

import re
import time
import logging
from io import StringIO
from urllib.parse import urljoin, urlparse, unquote, parse_qs
from typing import Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup

try:
    import pandas as pd
    _PANDAS_OK = True
except ImportError:
    _PANDAS_OK = False

try:
    import openpyxl
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logger = logging.getLogger("net_entreprise")
if not logger.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(_h)
logger.setLevel(logging.INFO)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
BASE_PORTAIL = "https://portail.net-entreprises.fr"
BASE_DSIJ    = "https://dsij-bpij.net-entreprises.ameli.fr"

# Charte DUNES
DUNES_ROUGE  = "AF321E"   # en-têtes
DUNES_ORANGE = "F09426"   # lignes paires
DUNES_BLANC  = "FFFFFF"
DUNES_GRIS   = "F5F5F5"   # lignes impaires

HEADERS_BASE = {
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,"
        "application/signed-exchange;v=b3;q=0.7"
    ),
    "Accept-Language": "fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7",
    "Cache-Control":   "max-age=0",
    "Connection":      "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/149.0.0.0 Safari/537.36"
    ),
    "sec-ch-ua":          '"Google Chrome";v="149", "Chromium";v="149", "Not)A;Brand";v="24"',
    "sec-ch-ua-mobile":   "?0",
    "sec-ch-ua-platform": '"Windows"',
}

# ---------------------------------------------------------------------------
# Colonnes -> numériques / centrées par onglet
# ---------------------------------------------------------------------------

# Colonnes à convertir en valeurs numériques (entiers) par table de données
NUMERIC_COLUMNS_PAIEMENTS = ["SIREN/SIRET", "Numéro de Sécurité Sociale", "identifiant"]
NUMERIC_COLUMNS_DETAIL    = ["Identifiant_BPIJ", "NIR"]

# Colonnes de rapprochement bancaire ajoutees a l'onglet "Paiements"
COL_MONTANT_RECU     = "Montant reçu"       # "Total a payer" de la journee BPIJ
COL_MONTANT_VIREMENT = "Montant virement"   # net reellement vire (toutes journees du jour)
COL_JOURNEE_BPIJ     = "Journée BPIJ"       # identifiant de la journee sur le site

# Colonnes à centrer (par nom d'en-tête, comparaison insensible à la casse) par onglet Excel
SHEET_CENTER_COLS = {
    "Paiements": [
        "SIREN/SIRET", "Date", "Numéro de Sécurité Sociale", "Type", "identifiant",
        "Journée BPIJ", "Cle_long",
    ],
    "Saisie QUADRA": [
        "Siret", "Date",
    ],
    "Détail des paiements": [
        "Identifiant_BPIJ", "NIR", "Prestation_CodeNature",
        "DateDeb", "DateFin", "NB_jours", "Type",
    ],
    "Répartition analytique": [
        "Matricule", "Mois analytique", "Clé", "Centre analytique", "Taux",
    ],
    "Récap centres analytiques": [
        "Centre analytique", "Nb lignes",
    ],
    "Contrôle matching": [
        "Matricule", "Date entrée", "Date sortie", "Statut", "Score", "Méthode",
        "Mois demandé", "Mois utilisé", "Nb centres", "Total taux",
    ],
}

# Colonnes "identifiantes" numériques pour lesquelles on force un format entier
# (évite la notation scientifique sur les grands nombres de type SIRET/NIR)
SHEET_ID_NUMBER_FORMAT_COLS = {
    "Paiements": ["SIREN/SIRET", "Numéro de Sécurité Sociale", "identifiant",
                  "Journée BPIJ"],
    "Saisie QUADRA": ["Siret"],
    "Détail des paiements": ["Identifiant_BPIJ", "NIR"],
}

# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------
class NetEntrepriseError(Exception):      pass
class LoginError(NetEntrepriseError):     pass
class NavigationError(NetEntrepriseError):pass
class ServiceNotFoundError(NetEntrepriseError): pass


# ===========================================================================
# Classe principale
# ===========================================================================
class NetEntreprise:
    """Client automatisé Net-Entreprises -> DSIJ."""

    def __init__(
        self,
        siret:       str,
        nom:         str,
        prenom:      str,
        password:    str,
        timeout:     int   = 30,
        pause:       float = 0.3,
        max_workers: int   = 5,
    ) -> None:
        self.siret       = siret.replace(" ", "")
        self.nom         = nom
        self.prenom      = prenom
        self.password    = password
        self.timeout     = timeout
        self.pause       = pause
        self.max_workers = max_workers

        self._session        = requests.Session()
        self._session.headers.update(HEADERS_BASE)
        self._logged_in      = False
        self._dashboard_html = None
        self._dsij_ready     = False
        self._jeton          = None

    # ------------------------------------------------------------------
    # Utilitaires HTTP
    # ------------------------------------------------------------------
    # Rapprochement bancaire : "Montant reçu" et "Montant virement"
    # ------------------------------------------------------------------

    @staticmethod
    def _cle_siret(serie) -> "pd.Series":
        """SIRET normalise en chaine de chiffres (pour servir de cle de jointure)."""
        return serie.astype(str).str.replace(r"\D", "", regex=True)

    def ajouter_montants_recus(self, df, df_bpij=None):
        """
        Ajoute a l'onglet « Paiements » les colonnes de rapprochement bancaire.

        Regles validees sur les donnees du site (296/296 journees de controle) :

          * une JOURNEE BPIJ = (SIRET, caisse emettrice, date)
            « Total a payer » de la journee = somme des montants de ses lignes
            salarie, indus negatifs compris.
            -> colonne "Montant reçu", repetee sur chaque ligne de la journee.

          * un VIREMENT = toutes les journees d'une meme DATE (tous SIRET et
            toutes caisses confondus : la CPAM regroupe le paiement du jour sur
            un seul virement).
            -> colonne "Montant virement" = net reellement credite en banque.

          * le "Montant" affiche par le site dans la vue BPIJ est le montant
            BRUT du virement du jour (avant indus), identique sur chaque
            journee du meme jour : il sert ici de controle.

        `df_bpij` (optionnel, issu de get_tableau_bpij) ajoute la colonne
        "Journée BPIJ" (identifiant de la journee sur le site) et active le
        controle brut.
        """
        if df is None or df.empty:
            return df
        for col in ("Montant", "Date", "SIREN/SIRET", "Caisse émettrice"):
            if col not in df.columns:
                logger.warning("Colonne « %s » absente : montants recus non calcules", col)
                return df

        df = df.copy()
        montant = pd.to_numeric(df["Montant"], errors="coerce").fillna(0.0)
        siret   = self._cle_siret(df["SIREN/SIRET"])
        caisse  = df["Caisse émettrice"].astype(str).str.strip()
        date    = df["Date"].astype(str).str.strip()

        cle_journee  = siret + "|" + caisse + "|" + date
        df[COL_MONTANT_RECU]     = montant.groupby(cle_journee).transform("sum").round(2)
        df[COL_MONTANT_VIREMENT] = montant.groupby(date).transform("sum").round(2)

        nb_journees = cle_journee.nunique()
        nb_virements = date.nunique()
        logger.info("Rapprochement : %d journee(s) BPIJ regroupee(s) en %d virement(s)",
                    nb_journees, nb_virements)

        # --- Identifiant de journee + controle du brut -----------------------
        if df_bpij is not None and not df_bpij.empty:
            b = df_bpij.copy()
            b_cle = (self._cle_siret(b["SIREN/SIRET"]) + "|"
                     + b["Caisse émettrice"].astype(str).str.strip() + "|"
                     + b["Date"].astype(str).str.strip())
            b = b.assign(_cle=b_cle).drop_duplicates(subset=["_cle"], keep="first")

            df[COL_JOURNEE_BPIJ] = pd.to_numeric(
                cle_journee.map(dict(zip(b["_cle"], b["identifiant"]))),
                errors="coerce",
            ).astype("Int64")
            manquants = int(df[COL_JOURNEE_BPIJ].isna().sum())
            if manquants:
                logger.warning("%d ligne(s) sans journee BPIJ correspondante", manquants)

            # Controle : brut calcule (somme des montants positifs du jour)
            #            vs brut affiche par le site pour ce jour.
            brut_calcule = montant.where(montant > 0, 0.0).groupby(date).sum().round(2)
            brut_site = (
                b.assign(_d=b["Date"].astype(str).str.strip())
                 .groupby("_d")["Montant brut du jour"].first().round(2)
            )
            communs = brut_calcule.index.intersection(brut_site.index)
            ecarts = [
                (d, float(brut_site[d]), float(brut_calcule[d]))
                for d in communs
                if abs(float(brut_site[d]) - float(brut_calcule[d])) > 0.005
            ]
            if ecarts:
                logger.warning(
                    "Controle brut : %d jour(s) en ecart avec le site (ex. %s)",
                    len(ecarts),
                    "; ".join(f"{d} site={s:.2f} calcule={c:.2f}" for d, s, c in ecarts[:3]),
                )
            else:
                logger.info("Controle brut : %d/%d jour(s) concordant(s) avec le site",
                            len(communs), len(communs))
        else:
            df[COL_JOURNEE_BPIJ] = pd.NA

        # --- Ordre des colonnes : les montants a cote du montant de la ligne --
        cols = list(df.columns)
        for c in (COL_MONTANT_RECU, COL_MONTANT_VIREMENT):
            cols.remove(c)
        i = cols.index("Montant") + 1
        cols[i:i] = [COL_MONTANT_RECU, COL_MONTANT_VIREMENT]
        return self.ajouter_libelles_quadra(df[cols])

    # ------------------------------------------------------------------
    # Libelles QUADRA : Caisse_short, Cle_long, Libellé_QUADRA
    # ------------------------------------------------------------------

    @staticmethod
    def _caisse_short(caisse) -> str:
        """
        Abrege la caisse emettrice pour les libelles QUADRA :
          * "BOUCHES du RHÔNE" -> "BDR"
          * " de " / " du " / " des " / " la " -> " "
        Ex. "CPAM des BOUCHES du RHÔNE" -> "CPAM BDR",
            "CPAM de la DROME" -> "CPAM DROME"
        """
        s = str(caisse or "").replace("\xa0", " ").strip()
        s = re.sub(r"BOUCHES\s+du\s+RH[ÔO]NE", "BDR", s, flags=re.IGNORECASE)
        for particule in (" des ", " du ", " de ", " la "):
            s = s.replace(particule, " ")
        return re.sub(r"\s+", " ", s).strip()

    # Particules qui font partie du nom de famille (ex. "EL GUIDI", "BEN ALI")
    _PARTICULES_NOM = {
        "EL", "AL", "BEN", "AIT", "OULD", "BOU", "ABOU", "ABD",
        "DE", "DEL", "DELLA", "DA", "DI", "DU", "DOS", "DAS",
        "LE", "LA", "VAN", "VON", "DER", "DEN", "MC", "MAC",
        "SAINT", "STE", "ST",
    }

    @classmethod
    def _nom_famille(cls, salarie) -> str:
        """
        Nom de famille depuis le libelle site « NOM PRENOM(S) » : premier mot,
        en gardant les particules accolees (ex. "EL GUIDI ICHEM" -> "EL GUIDI",
        "SARR MAHMOUDOU" -> "SARR").
        """
        mots = str(salarie or "").split()
        if not mots:
            return ""
        i = 0
        while i < len(mots) - 1 and mots[i].upper().rstrip(".'") in cls._PARTICULES_NOM:
            i += 1
        return " ".join(mots[: i + 1])

    def ajouter_libelles_quadra(self, df):
        """
        Ajoute a l'onglet « Paiements » les colonnes d'aide au rapprochement
        bancaire dans QUADRA (appelee automatiquement par
        ajouter_montants_recus) :

          * "Caisse_short"   : caisse emettrice abregee (cf. _caisse_short) ;

          * "Cle_long"       : identifiant du virement recu en banque.
            Les virements se font par journee BPIJ :
              {Journée BPIJ}@{Date}#{Montant reçu} €_{3 derniers chiffres SIRET}
            Ex. '3983312175@08/09/2026#282,94 €_158'

          * "Libellé_QUADRA" : libelle d'ecriture propose pour QUADRA.
              - Cle_long unique (1 ligne)  : Caisse_short + Salarié complet
                Ex. 'CPAM du VAUCLUSE BOAKYE WILHEMINA'
              - Cle_long partagee (n > 1)  : Caisse_short + noms de famille
                distincts des salaries de la journee, separes par "/"
                Ex. 'CPAM BDR AMAR/SEBAA'
        """
        if df is None or df.empty:
            return df
        requises = ("Caisse émettrice", "Salarié", "Date",
                    COL_MONTANT_RECU, "SIREN/SIRET")
        for col in requises:
            if col not in df.columns:
                logger.warning("Colonne « %s » absente : libelles QUADRA non calcules", col)
                return df

        df = df.copy()
        df["Caisse_short"] = df["Caisse émettrice"].map(self._caisse_short)

        # --- Cle_long ------------------------------------------------------
        if COL_JOURNEE_BPIJ in df.columns:
            journee_txt = (df[COL_JOURNEE_BPIJ].astype(str)
                           .str.replace(r"\D", "", regex=True))   # <NA> -> ""
        else:
            journee_txt = pd.Series("", index=df.index)
        montant_txt = (
            pd.to_numeric(df[COL_MONTANT_RECU], errors="coerce").fillna(0.0)
            .map(lambda v: f"{v:.2f}".replace(".", ","))
        )
        siret3   = df["SIREN/SIRET"].map(lambda v: self._siret_lastn(v, 3))
        date_txt = df["Date"].astype(str).str.strip()
        df["Cle_long"] = journee_txt + "@" + date_txt + "#" + montant_txt + " €_" + siret3

        # --- Libellé_QUADRA ------------------------------------------------
        salarie = df["Salarié"].astype(str).str.strip()
        noms    = salarie.map(self._nom_famille)          # nom de famille

        effectif = df.groupby("Cle_long")["Cle_long"].transform("size")
        noms_par_cle = noms.groupby(df["Cle_long"]).transform(
            lambda s: "/".join(dict.fromkeys(n for n in s if n))
        )
        libelle = (df["Caisse_short"] + " " + salarie).where(
            effectif == 1, df["Caisse_short"] + " " + noms_par_cle
        )
        df["Libellé_QUADRA"] = (libelle.str.replace(r"\s+", " ", regex=True)
                                       .str.strip())

        logger.info("Libelles QUADRA : %d Cle_long distincte(s) dont %d partagee(s)",
                    df["Cle_long"].nunique(),
                    df.loc[effectif > 1, "Cle_long"].nunique())

        # --- Ordre : Caisse_short apres la caisse, cles en fin de tableau --
        cols = [c for c in df.columns
                if c not in ("Caisse_short", "Cle_long", "Libellé_QUADRA")]
        cols.insert(cols.index("Caisse émettrice") + 1, "Caisse_short")
        cols += ["Cle_long", "Libellé_QUADRA"]
        return df[cols]

    # ------------------------------------------------------------------
    # Répartition analytique (réutilise repartition_analytique.py)
    # ------------------------------------------------------------------

    @staticmethod
    def _mois_depuis_date(date_txt) -> str:
        """'08/09/2026' -> '09/2026'."""
        trouve = re.search(r"\d{2}/(\d{2}/\d{4})", str(date_txt or ""))
        return trouve.group(1) if trouve else ""

    @staticmethod
    def _texte_repartition(centres) -> str:
        """[(centre, taux)] -> 'CENTRE' (un seul) ou 'C1 (50%); C2 (30%)'."""
        centres = [(c, t) for c, t in centres if t]
        if not centres:
            return ""
        if len(centres) == 1 and abs(centres[0][1] - 1.0) < 1e-9:
            return str(centres[0][0])
        def pct(t):
            return f"{round(t * 100, 2):g}".replace(".", ",") + "%"
        return "; ".join(f"{c} ({pct(t)})" for c, t in centres)

    @staticmethod
    def _appliquer_regle_siege(centres):
        """Salarié du SIEGE : si l'un de ses centres est DSIEGE, toute sa
        répartition est remplacée par DSIEGE à 100 % (règle de gestion :
        les salariés du siège ne sont pas ventilés sur les autres centres)."""
        if any(str(c).strip().upper() == "DSIEGE" for c, _ in centres):
            return [("DSIEGE", 1.0)]
        return centres

    def _centres_salarie(self, referentiel, nom: str, mois: str,
                         cache: dict) -> list:
        """Centres analytiques [(centre, taux)] d'un salarié pour un mois
        (règle SIEGE appliquée)."""
        from repartition_analytique import apparier_nom, _choisir_mois
        if (nom, mois) in cache:
            return cache[(nom, mois)]
        centres = []
        appariement = apparier_nom(nom, referentiel.index_noms, referentiel.alias,
                                   referentiel.jetons_naissance, referentiel.libelles)
        cle = appariement.get("cle")
        if cle:
            mois_retenu, _ = _choisir_mois(mois, referentiel.mois_disponibles(cle))
            if mois_retenu:
                centres = self._appliquer_regle_siege(
                    referentiel.centres(cle, mois_retenu))
        cache[(nom, mois)] = centres
        return centres

    def ajouter_repartition_paiements(self, df, referentiel):
        """Colonne « Répartition analytique » de l'onglet Paiements.

        Pour chaque ligne : centres analytiques du salarié au mois de la date
        de paiement (repli sur le mois disponible le plus proche). Placée
        juste après « Montant ».
        """
        if df is None or df.empty or referentiel is None:
            return df
        if not {"Salarié", "Date"}.issubset(df.columns):
            logger.warning("Répartition Paiements : colonnes Salarié/Date absentes")
            return df
        df = df.copy()
        cache: dict = {}
        df["Répartition analytique"] = [
            self._texte_repartition(
                self._centres_salarie(referentiel, str(nom).strip(),
                                      self._mois_depuis_date(date), cache))
            for nom, date in zip(df["Salarié"], df["Date"])
        ]
        nb = int((df["Répartition analytique"] != "").sum())
        logger.info("Répartition analytique Paiements : %d/%d lignes renseignées",
                    nb, len(df))
        cols = [c for c in df.columns if c != "Répartition analytique"]
        pos = cols.index("Montant") + 1 if "Montant" in cols else len(cols)
        cols.insert(pos, "Répartition analytique")
        return df[cols]

    def construire_saisie_quadra(self, df, referentiel):
        """Onglet « Saisie QUADRA » : une ligne par virement (Cle_long).

        Colonnes : Siret, Libellé (= Libellé_QUADRA), Date, Montant (montant
        reçu de la journée BPIJ, à défaut somme des paiements) et Répartition
        analytique ventilée en euros par centre (« 1938,90 : MSUMAR (90,07%) »,
        ou le centre seul s'il est unique) — somme strictement égale au
        Montant grâce à ventiler_montant. Les paiements sans centres sont
        portés en « NON AFFECTE ».
        """
        if df is None or df.empty:
            return None
        requises = {"Cle_long", "Libellé_QUADRA", "SIREN/SIRET", "Date",
                    "Salarié", "Montant"}
        if not requises.issubset(df.columns):
            logger.warning("Saisie QUADRA : colonnes manquantes (%s)",
                           ", ".join(sorted(requises - set(df.columns))))
            return None
        from repartition_analytique import ventiler_montant

        def fmt_montant(x):
            s = f"{x:.2f}".replace(".", ",")
            return s[:-3] if s.endswith(",00") else s

        cache: dict = {}
        lignes = []
        for cle_long, grp in df.groupby("Cle_long", dropna=False, sort=False):
            montants = pd.to_numeric(grp["Montant"], errors="coerce").fillna(0.0)
            if COL_MONTANT_RECU in grp.columns:
                recu = pd.to_numeric(grp[COL_MONTANT_RECU], errors="coerce").dropna()
                montant_quadra = round(float(recu.iloc[0]), 2) if not recu.empty                     else round(float(montants.sum()), 2)
            else:
                montant_quadra = round(float(montants.sum()), 2)

            # Poids par centre : chaque paiement du virement ventilé sur les
            # centres de son salarié (mois de la date de paiement).
            poids: dict[str, float] = {}
            for (_, ligne), montant in zip(grp.iterrows(), montants):
                centres = [] if referentiel is None else self._centres_salarie(
                    referentiel, str(ligne["Salarié"]).strip(),
                    self._mois_depuis_date(ligne["Date"]), cache)
                total_taux = sum(t for _, t in centres)
                if centres and total_taux:
                    for centre, taux in centres:
                        poids[centre] = poids.get(centre, 0.0) + montant * taux / total_taux
                else:
                    poids["NON AFFECTE"] = poids.get("NON AFFECTE", 0.0) + montant

            ventile = ventiler_montant(montant_quadra, sorted(poids.items()))
            ventile = [(c, t, v) for c, t, v in ventile if abs(v) >= 0.005]
            if len(ventile) == 1:
                texte = ventile[0][0]
            else:
                total = sum(v for _, _, v in ventile)
                def pct(v):
                    return (f"{round(v * 100 / total, 2):g}".replace(".", ",") + "%")                         if total else ""
                texte = "\n".join(
                    f"{fmt_montant(v)} : {c}" + (f" ({pct(v)})" if total else "")
                    for c, _, v in sorted(ventile, key=lambda x: (-x[2], x[0])))

            lignes.append({
                "Siret": grp["SIREN/SIRET"].iloc[0],
                "Libellé": grp["Libellé_QUADRA"].iloc[0],
                "Date": grp["Date"].iloc[0],
                "Montant": montant_quadra,
                "Répartition analytique": texte,
            })

        df_quadra = pd.DataFrame(
            lignes, columns=["Siret", "Libellé", "Date", "Montant",
                             "Répartition analytique"])
        logger.info("Saisie QUADRA : %d virement(s)", len(df_quadra))
        return df_quadra

    # ------------------------------------------------------------------

    def _get(self, url: str, **kwargs) -> requests.Response:
        time.sleep(self.pause)
        kwargs.setdefault("allow_redirects", True)
        r = self._session.get(url, timeout=self.timeout, **kwargs)
        logger.debug("GET  %-70s -> %s", url[:70], r.status_code)
        return r

    def _post(self, url: str, data: dict, **kwargs) -> requests.Response:
        time.sleep(self.pause)
        kwargs.setdefault("allow_redirects", True)
        r = self._session.post(url, data=data, timeout=self.timeout, **kwargs)
        logger.debug("POST %-70s -> %s", url[:70], r.status_code)
        return r

    def _assert_ok(self, r: requests.Response, label: str = "") -> None:
        if r.status_code != 200:
            raise NavigationError(f"[{label}] Statut {r.status_code}  URL: {r.url}")

    # ------------------------------------------------------------------
    # ÉTAPE 1 — Connexion portail
    # ------------------------------------------------------------------

    def login(self) -> None:
        """POST /auth/pass -> suit les 302 jusqu'à declarations?rs=1."""
        logger.info("Connexion: %s  %s %s", self.siret, self.nom.upper(), self.prenom)

        self._session.headers.update({
            "Origin":         "https://www.net-entreprises.fr",
            "Referer":        "https://www.net-entreprises.fr/",
            "Content-Type":   "application/x-www-form-urlencoded",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-site",
            "Sec-Fetch-User": "?1",
        })

        r = self._post(
            f"{BASE_PORTAIL}/auth/pass",
            data={
                "j_siret":    self.siret,
                "j_nom":      self.nom,
                "j_prenom":   self.prenom,
                "j_password": self.password,
            },
        )
        self._assert_ok(r, "login")
        self._logged_in      = True
        self._dashboard_html = r.text
        logger.info("Connexion réussie")

    # ------------------------------------------------------------------
    # ÉTAPE 2 — Services du dashboard
    # ------------------------------------------------------------------

    def get_services(self) -> dict[str, str]:
        if not self._logged_in:
            raise LoginError("Appelez login() d'abord")

        soup  = BeautifulSoup(self._dashboard_html or "", "html.parser")
        found = {}
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if "accesServiceDeclaration" not in href:
                continue
            titre_tag = a.find(class_="item-nom")
            titre = (
                a.get("title")
                or (titre_tag.get_text(strip=True) if titre_tag else None)
                or a.get_text(strip=True)
            )
            found[titre] = urljoin(f"{BASE_PORTAIL}/priv/", href)

        logger.info("%d service(s) trouvé(s)", len(found))
        return found

    # ------------------------------------------------------------------
    # ÉTAPE 3-8 — Accès Attestation de salaire -> session DSIJ complète
    # ------------------------------------------------------------------

    def open_attestation(self) -> None:
        """
        Flux exact DevTools :
          3. GET  accesServiceDeclaration/TOKEN/ID   -> 303
          4. GET  formAccesServiceDeclarations        -> 200  (extrait jeton)
          5. POST accesNetEntreprisesFrontal  {jtd}  -> 200  (crée JSESSIONID DSIJ)
          6. GET  accueil.go                         -> 200
          7. POST connect.go  {jtn}                  -> 302  -> chercherPaiement.go
          8. GET  chercherPaiement.go                -> 200  OK
        """
        if not self._logged_in:
            raise LoginError("Appelez login() d'abord")

        # --- 3. Trouver le service ---
        services = self.get_services()
        url_service = next(
            (u for t, u in services.items()
             if "attestation" in t.lower() and "salaire" in t.lower()),
            None,
        )
        if url_service is None:
            raise ServiceNotFoundError(
                f"Service 'Attestation de salaire' introuvable. "
                f"Disponibles : {list(services.keys())}"
            )
        logger.info("Service: %s", url_service)

        # --- 3->4 : GET avec suivi du 303 ---
        self._session.headers.update({
            "Referer":        f"{BASE_PORTAIL}/priv/declarations?rs=1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",
        })
        r = self._get(url_service)   # suit 303 -> formAccesServiceDeclarations

        if "formAccesServiceDeclarations" not in r.url:
            raise NavigationError(f"Page intermédiaire non atteinte: {r.url}")

        # --- 4. Extraction du jeton ---
        jeton = self._extract_jeton(r.text, r.url)
        if not jeton:
            raise NavigationError("Impossible d'extraire le jeton depuis formAccesServiceDeclarations")
        self._jeton = jeton
        logger.info("Jeton extrait (longueur %d)", len(jeton))

        # --- 5. POST accesNetEntreprisesFrontal ---
        self._session.headers.update({
            "Origin":         BASE_PORTAIL,
            "Referer":        BASE_PORTAIL + "/",
            "Content-Type":   "application/x-www-form-urlencoded",
            "Sec-Fetch-Site": "cross-site",
        })
        r = self._post(
            f"{BASE_DSIJ}/IhmHF/accesNetEntreprisesFrontal",
            data={"jtd": jeton},
        )
        self._assert_ok(r, "accesNetEntreprisesFrontal")
        logger.info("Session DSIJ créée (JSESSIONID: %s)",
                    "oui" if "JSESSIONID" in self._session.cookies else "absent")

        # --- 6. GET accueil.go ---
        self._session.headers.update({
            "Origin":         "",
            "Referer":        f"{BASE_DSIJ}/IhmHF/accesNetEntreprisesFrontal",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",
        })
        r = self._get(f"{BASE_DSIJ}/IhmHF/accueil.go")
        self._assert_ok(r, "accueil.go")

        # --- 7. POST connect.go ---
        self._session.headers.update({
            "Origin":       BASE_DSIJ,
            "Referer":      f"{BASE_DSIJ}/IhmHF/accueil.go",
            "Content-Type": "application/x-www-form-urlencoded",
        })
        r = self._post(f"{BASE_DSIJ}/IhmHF/connect.go", data={"jtn": jeton})

        # --- 8. Vérification finale ---
        if "chercherPaiement" not in r.url:
            self._session.headers.update({
                "Origin":         "",
                "Referer":        f"{BASE_DSIJ}/IhmHF/accueil.go",
                "Sec-Fetch-User": "?1",
            })
            r = self._get(f"{BASE_DSIJ}/IhmHF/chercherPaiement.go")
            self._assert_ok(r, "chercherPaiement.go")

        self._dsij_ready = True
        logger.info("Session DSIJ opérationnelle - OK")

    def _extract_jeton(self, html: str, page_url: str) -> str:
        """Extrait le jeton : query-string > input hidden > cookie JTD."""
        qs = parse_qs(urlparse(page_url).query)
        if "jeton" in qs:
            return unquote(qs["jeton"][0])

        soup = BeautifulSoup(html, "html.parser")
        for inp in soup.find_all("input"):
            name = (inp.get("name") or "").lower()
            val  = inp.get("value") or ""
            if name in ("jeton", "jtd", "jtn") and val:
                return unquote(val)

        jtd = self._session.cookies.get("JTD", "")
        return unquote(jtd) if jtd else ""

    # ------------------------------------------------------------------
    # ÉTAPE 9 — Recherche paiements
    # ------------------------------------------------------------------

    def rechercher_paiements(
        self,
        date_debut:      str  = "01/01/2026",
        date_fin:        str  = "",
        nom:             str  = "",
        prenom:          str  = "",
        nir:             str  = "",
        type_recherche:  str  = "salarie",
        tous_siret:      bool = True,
        liste_siret:     Optional[list[str]] = None,
    ) -> str:
        if not self._dsij_ready:
            raise NavigationError("Appelez open_attestation() d'abord")

        self._session.headers.update({
            "Origin":         BASE_DSIJ,
            "Referer":        f"{BASE_DSIJ}/IhmHF/chercherPaiement.go",
            "Content-Type":   "application/x-www-form-urlencoded",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",
        })

        if tous_siret:
            data = {
                "_listCaisseSelectionnees": "1",
                "_listeSiretCombo":         "1",
                "dateDebut":     date_debut,
                "dateFin":       date_fin,
                "nom":           nom,
                "prenom":        prenom,
                "nir":           nir,
                "typeRecherche": type_recherche,
            }
        else:
            if not liste_siret:
                raise ValueError("liste_siret vide alors que tous_siret=False")
            data = {
                "_listCaisseSelectionnees": "1",
                "listeSiretCombo":          liste_siret,
                "dateDebut":     date_debut,
                "dateFin":       date_fin,
                "nom":           nom,
                "prenom":        prenom,
                "nir":           nir,
                "typeRecherche": type_recherche,
            }

        r = self._post(f"{BASE_DSIJ}/IhmHF/afficherPaiement.go", data=data)
        self._assert_ok(r, "afficherPaiement")
        logger.info("Recherche paiements OK (du %s au %s)", date_debut, date_fin or "aujourd'hui")
        return r.text

    # ------------------------------------------------------------------
    # ÉTAPE 10 — Tableau 300 lignes
    # ------------------------------------------------------------------

    # Le site ne propose que 10 / 50 / 100 / 200 / 300 lignes par page
    # (l'option "Toutes" vaut en realite 300) -> il FAUT paginer.
    PAGE_SIZE_MAX  = 300
    TABLE_ID       = "d-3926063"   # tableau "Paiements" (vue salarie)
    TABLE_ID_BPIJ  = "d-6833829"   # tableau "Paiements" (vue BPIJ / journee)

    @staticmethod
    def _nb_lignes_trouvees(html: str) -> Optional[int]:
        """
        Lit le bandeau « <b>N</b> ligne(s) trouvee(s) » renvoye par le site.
        Retourne None si le bandeau est absent.
        """
        soup = BeautifulSoup(html, "html.parser")
        bloc = soup.find("div", class_="resultat")
        textes = []
        if bloc is not None:
            textes.append(bloc.get_text(" ", strip=True))
        textes.append(soup.get_text(" ", strip=True))

        for txt in textes:
            txt = txt.replace("\xa0", " ").replace("\u202f", " ")
            m = re.search(r"([\d][\d ]*)\s*ligne\(s\)\s*trouv", txt)
            if m:
                digits = re.sub(r"\D", "", m.group(1))
                if digits:
                    return int(digits)
        return None

    # ------------------------------------------------------------------
    # Recuperation HTML d'une page de tableau (salarie OU bpij)
    # ------------------------------------------------------------------

    def _tableau_html(
        self,
        type_recherche: str,
        table_id:       str,
        endpoint:       str,
        date_debut:     str,
        date_fin:       str,
        page:           int,
        page_size:      int,
        relancer:       bool,
    ) -> str:
        if relancer:
            self.rechercher_paiements(
                date_debut=date_debut,
                date_fin=date_fin,
                type_recherche=type_recherche,
            )

        self._session.headers.update({
            "Origin":       BASE_DSIJ,
            "Referer":      f"{BASE_DSIJ}/IhmHF/afficherPaiement.go",
            "Content-Type": "application/x-www-form-urlencoded",
        })

        params = {
            f"{table_id}-p":             str(page),
            "typeRecherche":             type_recherche,
            "dateDebut":                 date_debut,
            "nir":                       "",
            "_listeSiretCombo":          "1",
            "dateFin":                   date_fin,
            "_listCaisseSelectionnees":  "1",
            "nom":                       "",
            "prenom":                    "",
        }
        data = params.copy()
        data[f"{table_id}-z"] = str(page_size)

        r = self._post(f"{BASE_DSIJ}/IhmHF/{endpoint}", data=data, params=params)
        self._assert_ok(r, endpoint)
        logger.debug("%s page %d (%d octets)", endpoint, page, len(r.content))
        return r.text

    def get_tableau_html(
        self,
        date_debut:     str = "01/01/2026",
        date_fin:       str = "",
        type_recherche: str = "salarie",
        page:           int = 1,
        page_size:      int = PAGE_SIZE_MAX,
        relancer_recherche: bool = True,
    ) -> str:
        """HTML d'une page du tableau « Paiements » (vue salarie)."""
        return self._tableau_html(
            type_recherche=type_recherche,
            table_id=self.TABLE_ID,
            endpoint="gererTableauPaiementSalarie.go",
            date_debut=date_debut, date_fin=date_fin,
            page=page, page_size=page_size, relancer=relancer_recherche,
        )

    # ------------------------------------------------------------------
    # Boucle de pagination generique
    # ------------------------------------------------------------------

    def _collecte_pagine(
        self,
        type_recherche: str,
        table_id:       str,
        endpoint:       str,
        parser,
        libelle:        str,
        date_debut:     str,
        date_fin:       str,
        page_size:      int,
        max_pages:      int,
        essais:         int = 3,
    ):
        """
        Parcourt TOUTES les pages d'un tableau du site et concatene le resultat.

        Le site trie par date decroissante et plafonne l'affichage a
        PAGE_SIZE_MAX lignes par page : sans pagination, seules les lignes
        les plus RECENTES sont renvoyees et le debut de periode disparait
        silencieusement.
        """
        page_size = max(1, min(int(page_size), self.PAGE_SIZE_MAX))

        def charger(page: int, relancer: bool):
            derniere = None
            for tentative in range(1, essais + 1):
                try:
                    html = self._tableau_html(
                        type_recherche, table_id, endpoint,
                        date_debut, date_fin, page, page_size, relancer,
                    )
                    return html, parser(html)
                except Exception as exc:
                    derniere = exc
                    logger.warning("%s page %d : echec %d/%d (%s)",
                                   libelle, page, tentative, essais, exc)
                    if tentative < essais:
                        time.sleep(2 * tentative)
            raise NavigationError(
                f"{libelle} page {page} illisible apres {essais} tentatives : {derniere}"
            )

        html, df_page = charger(1, True)
        total = self._nb_lignes_trouvees(html)

        if total is None:
            logger.warning("%s : bandeau « N ligne(s) trouvee(s) » introuvable, "
                           "resultat possiblement tronque", libelle)
            nb_pages_attendu = 1
        else:
            nb_pages_attendu = max(1, -(-total // page_size))
            logger.info("%s : %d ligne(s) trouvee(s) -> %d page(s) de %d",
                        libelle, total, nb_pages_attendu, page_size)

        frames = [df_page]
        vus    = set(df_page["identifiant"].astype(str)) if "identifiant" in df_page.columns else set()
        recuperees = len(df_page)
        page   = 1
        limite = min(nb_pages_attendu, max_pages) if total is not None else 1

        while page < limite and recuperees < (total or 0):
            page += 1
            _, df_page = charger(page, False)
            if df_page.empty:
                logger.warning("%s page %d vide : arret de la pagination", libelle, page)
                break

            # Garde-fou : si le site ignorait le parametre de page, on
            # recevrait indefiniment les memes lignes.
            if "identifiant" in df_page.columns:
                nouveaux = set(df_page["identifiant"].astype(str)) - vus
                if not nouveaux:
                    logger.warning("%s page %d : aucune ligne nouvelle, arret", libelle, page)
                    break
                vus |= nouveaux

            frames.append(df_page)
            recuperees += len(df_page)
            logger.info("%s page %d/%d : %d lignes (cumul %d/%s)",
                        libelle, page, nb_pages_attendu, len(df_page),
                        recuperees, total if total is not None else "?")

        df = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]

        # Une meme date peut etre a cheval sur deux pages
        if "identifiant" in df.columns:
            avant = len(df)
            df = df.drop_duplicates(subset=["identifiant"], keep="first").reset_index(drop=True)
            if len(df) != avant:
                logger.info("%s : %d doublon(s) inter-pages supprime(s)", libelle, avant - len(df))

        if total is not None and len(df) != total:
            logger.warning("ATTENTION %s : %d ligne(s) recuperee(s) pour %d annoncee(s)",
                           libelle, len(df), total)
        else:
            logger.info("%s complet : %d ligne(s) sur %d page(s)", libelle, len(df), page)

        return df

    def get_tableau(
        self,
        date_debut:     str = "01/01/2026",
        date_fin:       str = "",
        type_recherche: str = "salarie",
        page_size:      int = PAGE_SIZE_MAX,
        max_pages:      int = 200,
    ):
        """Tableau « Paiements » (une ligne par paiement salarie), toutes pages."""
        if not _PANDAS_OK:
            raise ImportError("pip install pandas")
        return self._collecte_pagine(
            type_recherche=type_recherche,
            table_id=self.TABLE_ID,
            endpoint="gererTableauPaiementSalarie.go",
            parser=self._parse_tableau,
            libelle="Paiements",
            date_debut=date_debut, date_fin=date_fin,
            page_size=page_size, max_pages=max_pages,
        )

    # ------------------------------------------------------------------
    # Tableau BPIJ : une ligne par journee de paiement
    # ------------------------------------------------------------------

    def get_tableau_bpij(
        self,
        date_debut: str = "01/01/2026",
        date_fin:   str = "",
        page_size:  int = PAGE_SIZE_MAX,
        max_pages:  int = 200,
    ):
        """
        Tableau « Paiements » en vue BPIJ : une ligne par journee de paiement
        (SIRET + caisse + date), avec l'identifiant de la journee sur le site
        et le montant BRUT du virement du jour affiche par le site.
        """
        if not _PANDAS_OK:
            raise ImportError("pip install pandas")
        return self._collecte_pagine(
            type_recherche="bpij",
            table_id=self.TABLE_ID_BPIJ,
            endpoint="gererTableauPaiement.go",
            parser=self._parse_tableau_bpij,
            libelle="Journees BPIJ",
            date_debut=date_debut, date_fin=date_fin,
            page_size=page_size, max_pages=max_pages,
        )

    def _parse_tableau_bpij(self, html: str):
        """Parse le tableau #tabPaiements (vue BPIJ)."""
        soup  = BeautifulSoup(html, "html.parser")
        table = soup.find("table", id="tabPaiements")
        colonnes = ["SIREN/SIRET", "Caisse émettrice", "Date", "Montant brut du jour",
                    "identifiant"]
        if table is None:
            logger.warning("Table #tabPaiements introuvable")
            return pd.DataFrame(columns=colonnes)

        lignes = []
        tbody = table.find("tbody") or table
        for tr in tbody.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 4:
                continue
            ident = ""
            for a in tr.find_all("a", href=True):
                href = a["href"]
                m = re.search(r"identifiant=(\d+)", href)
                if not m:
                    continue
                if "afficherDetailPaiementJournee" in href:
                    ident = m.group(1)
                    break
                if not ident:
                    ident = m.group(1)
            lignes.append([
                tds[0].get_text(" ", strip=True),
                tds[1].get_text(" ", strip=True),
                tds[2].get_text(" ", strip=True),
                tds[3].get_text(" ", strip=True),
                ident,
            ])

        df = pd.DataFrame(lignes, columns=colonnes)
        df = df.map(lambda x: x.replace("\xa0", " ").strip() if isinstance(x, str) else x)
        df["Montant brut du jour"] = self._to_euros(df["Montant brut du jour"])
        for col in ("SIREN/SIRET", "identifiant"):
            df[col] = self._to_numeric_clean(df[col])
        logger.info("Journees BPIJ parsees : %d ligne(s)", len(df))
        return df

    # ------------------------------------------------------------------
    # Helper : conversion numérique "propre" (supprime tout sauf chiffres)
    # ------------------------------------------------------------------

    @staticmethod
    def _to_euros(series: "pd.Series") -> "pd.Series":
        """Convertit une colonne texte « 1 234,56 € » en float."""
        return pd.to_numeric(
            series.astype(str)
            .str.replace("€", "", regex=False)
            .str.replace("\xa0", "", regex=False)
            .str.replace("\u202f", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False),
            errors="coerce",
        )

    @staticmethod
    def _to_numeric_clean(series: "pd.Series") -> "pd.Series":
        """
        Nettoie une série texte (espaces, NBSP, séparateurs) et la convertit
        en entier nullable (Int64). Les valeurs vides/non numériques -> NA.
        """
        cleaned = (
            series.astype(str)
            .str.replace("\xa0", "", regex=False)
            .str.replace(r"[^\d]", "", regex=True)
        )
        cleaned = cleaned.replace("", pd.NA)
        return pd.to_numeric(cleaned, errors="coerce").astype("Int64")

    # ------------------------------------------------------------------
    # Parsing tableau principal
    # CORRECTIF 1 : suppression de la colonne vide entre "Type" et "identifiant"
    # ------------------------------------------------------------------

    def _parse_tableau(self, html: str):
        soup  = BeautifulSoup(html, "html.parser")
        table = soup.find("table", id="paiementSalarietab")

        if table is None:
            logger.warning("Table #paiementSalarietab introuvable, tentative pd.read_html …")
            tables = pd.read_html(StringIO(html))
            if not tables:
                raise NavigationError("Aucun tableau trouvé dans la réponse DSIJ")
            df = tables[0]
            if "identifiant" not in df.columns:
                df["identifiant"] = ""
            return self._postprocess_paiements(df)

        # En-têtes depuis <thead>
        headers = []
        thead = table.find("thead")
        if thead:
            headers = [th.get_text(strip=True) for th in thead.find_all("th")]
        if not headers:
            headers = [th.get_text(strip=True) for th in table.find_all("th")]

        # -- CORRECTIF 1 ----------------------------------------------------
        # Le tableau HTML a une dernière colonne "Actions" (icônes) qui génère
        # une cellule vide dans les données.  On la supprime des en-têtes
        # avant de les utiliser, puis on extrait l'identifiant depuis les liens.
        # On s'assure aussi qu'il n'y a pas de doublon de nom vide.
        headers_clean = [h for h in headers if h != ""]   # retire les vides
        # Retire "Actions" (colonne boutons) si présente
        headers_clean = [h for h in headers_clean if h.lower() != "actions"]
        headers_clean.append("identifiant")
        # ---------------------------------------------------------------------

        rows  = []
        tbody = table.find("tbody") or table
        for tr in tbody.find_all("tr"):
            tds = tr.find_all("td")
            if not tds:
                continue

            row = [td.get_text(" ", strip=True) for td in tds]

            # Identifiant depuis les liens href
            identifiant = ""
            for a in tr.find_all("a", href=True):
                m = re.search(r"identifiant=(\d+)", a["href"])
                if m:
                    identifiant = m.group(1)
                    break

            # On ne garde que autant de cellules que headers_clean - 1
            # (la dernière colonne = identifiant, ajoutée manuellement)
            n_data = len(headers_clean) - 1
            row_data = row[:n_data]
            if len(row_data) < n_data:
                row_data += [""] * (n_data - len(row_data))
            row_data.append(identifiant)
            rows.append(row_data)

        df = pd.DataFrame(rows, columns=headers_clean)
        df = df.map(lambda x: x.replace("\xa0", " ").strip() if isinstance(x, str) else x)

        return self._postprocess_paiements(df)

    def _postprocess_paiements(self, df: "pd.DataFrame") -> "pd.DataFrame":
        """Conversions numériques pour le tableau Paiements (Montant + colonnes ID)."""

        # Montant -> float
        if "Montant" in df.columns:
            df["Montant"] = self._to_euros(df["Montant"])

        # -- Demande utilisateur : SIREN/SIRET, Numéro de Sécurité Sociale,
        #    identifiant -> colonnes numériques
        for col in NUMERIC_COLUMNS_PAIEMENTS:
            if col in df.columns:
                df[col] = self._to_numeric_clean(df[col])

        logger.info("Tableau parsé : %d lignes x %d colonnes", len(df), len(df.columns))
        return df

    # ------------------------------------------------------------------
    # Détail salarié (ongletPrestation.go)
    # CORRECTIF 2 : DateDeb / DateFin -> format JJ/MM/AAAA dans le DataFrame
    # ------------------------------------------------------------------

    def get_detail_salarie_html(self, identifiant: str) -> str:
        if not self._dsij_ready:
            raise NavigationError("Appelez open_attestation() d'abord")
        self._session.headers.update({
            "Referer": f"{BASE_DSIJ}/IhmHF/retourPaiement.go",
        })
        r = self._get(
            f"{BASE_DSIJ}/IhmHF/ongletPrestation.go",
            params={"identifiant": identifiant},
        )
        self._assert_ok(r, f"ongletPrestation({identifiant})")
        return r.text

    def parse_detail_salarie(
        self,
        html_content: str,
        identifiant:  str,
        type_paiement: str = "",
    ) -> "pd.DataFrame":
        soup = BeautifulSoup(html_content, "html.parser")

        # --- Extraction NIR / Nom / Prénom ---
        nir = nom = prenom = ""
        for ligne in soup.find_all("div", class_="ligne"):
            label = ligne.find("label")
            if label and "Salarié" in label.get_text(strip=True):
                texts = [
                    (e.get_text(strip=True) if hasattr(e, "get_text") else str(e).strip())
                    for e in ligne.contents
                    if e.name != "label"
                ]
                full_text = " ".join(t for t in texts if t and t != ":")
                parts     = full_text.split()
                if len(parts) >= 3:
                    nir, nom, prenom = parts[0], parts[1], " ".join(parts[2:])
                elif len(parts) == 2:
                    nir, nom = parts
                elif len(parts) == 1:
                    nir = parts[0]
                break

        # --- Tableau de prestations ---
        table = (
            soup.find("table", id="paiementSalarietab")
            or soup.find("table", class_="tableau")
        )
        if not table:
            logger.warning("Tableau détail introuvable pour %s", identifiant)
            return pd.DataFrame()

        def to_float(val: str) -> float:
            try:
                return float(val.replace(" ", "").replace("€", "").replace(",", ".").strip() or 0)
            except Exception:
                return 0.0

        # Première passe : collecte brute
        raw_rows = []
        tbody = table.find("tbody") or table
        for tr in tbody.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 6:
                continue
            date_deb_raw = tds[0].get_text(strip=True)
            date_fin_raw = tds[1].get_text(strip=True)
            libelle      = tds[2].get_text(strip=True)
            nb_jours     = to_float(tds[3].get_text(strip=True))
            prix_u       = to_float(tds[4].get_text(strip=True))
            montant      = to_float(tds[5].get_text(strip=True))

            code_nature = ""
            lib_up = libelle.upper()
            if "I.J. MAJOREES" in lib_up:
                code_nature = "IJMAJ"
            elif "I.J. NORMALES" in lib_up:
                code_nature = "IJNOR"
            elif "CONTRIBUTION SOCIALE GENERALISEE" in lib_up:
                code_nature = "CSJ"
            elif "REGUL.CONTRIB. SOCIALE GENERALISEE" in lib_up:
                code_nature = "REGRCJ"
            elif "CONTRIB. REMBOURSEMENT DETTE SOCIALE" in lib_up:
                code_nature = "RETCRD"
            elif "REGUL. REMBOURSEMENT DETTE SOCIALE" in lib_up:
                code_nature = "REGRRD"
            elif "CARENCE" in lib_up:
                code_nature = "CAR"

            raw_rows.append({
                "date_deb_raw": date_deb_raw,
                "date_fin_raw": date_fin_raw,
                "libelle":      libelle,
                "code_nature":  code_nature,
                "nb_jours":     nb_jours,
                "prix_u":       prix_u,
                "montant":      montant,
            })

        # Deuxième passe : propagation des dates (remplissage vers le bas)
        cur_deb = cur_fin = ""
        for row in raw_rows:
            if row["date_deb_raw"]:
                cur_deb = row["date_deb_raw"]
                cur_fin = row["date_fin_raw"] or cur_deb
            else:
                row["date_deb_raw"] = cur_deb
                row["date_fin_raw"] = cur_fin or cur_deb

        # -- CORRECTIF 2 ------------------------------------------------------
        # Conversion de la date brute (format portail = JJ/MM/AAAA en
        # général, mais peut arriver comme ISO "AAAA-MM-JJ") -> objet date
        # Python puis formatage en JJ/MM/AAAA pour Excel.
        def normalise_date(val: str) -> Optional[str]:
            if not val:
                return None
            val = val.strip()
            # Tente JJ/MM/AAAA
            try:
                d = pd.to_datetime(val, format="%d/%m/%Y", dayfirst=True)
                return d.strftime("%d/%m/%Y")
            except Exception:
                pass
            # Tente ISO AAAA-MM-JJ (ou avec timestamp)
            try:
                d = pd.to_datetime(val)
                return d.strftime("%d/%m/%Y")
            except Exception:
                pass
            return val   # en dernier recours, retourner tel quel
        # -----------------------------------------------------------------------

        final_rows = []
        for row in raw_rows:
            final_rows.append({
                "Identifiant_BPIJ":    identifiant,
                "NIR":                 nir,
                "Nom_Salarie":         nom,
                "Prenom_Salarie":      prenom,
                "Nom et prénom(s)":    f"{nom} {prenom}".strip(),
                "Prestation_CodeNature": row["code_nature"],
                "Prestation_Libelle":  row["libelle"],
                "DateDeb":             normalise_date(row["date_deb_raw"]),
                "DateFin":             normalise_date(row["date_fin_raw"]),
                "NB_jours":            row["nb_jours"],
                "Prix_unitaire":       row["prix_u"],
                "Montant":             row["montant"],
                "Type":                type_paiement,
            })

        df = pd.DataFrame(final_rows)

        # -- Demande utilisateur : Identifiant_BPIJ, NIR -> colonnes numériques
        for col in NUMERIC_COLUMNS_DETAIL:
            if col in df.columns:
                df[col] = self._to_numeric_clean(df[col])

        logger.info("Détail %s : %d lignes", identifiant, len(df))
        return df

    # ------------------------------------------------------------------
    # Parallèle
    # ------------------------------------------------------------------

    def get_all_details_parallel(self, df) -> dict[str, "pd.DataFrame"]:
        """Télécharge et parse les détails en parallèle."""
        ids_with_type = {}
        for _, row in df.iterrows():
            ident = str(row.get("identifiant", "")).strip()
            if ident and ident.lower() != "nan" and ident not in ids_with_type:
                ids_with_type[ident] = str(row.get("Type", ""))

        if not ids_with_type:
            return {}

        logger.info("Téléchargement parallèle : %d dossier(s) / %d workers",
                    len(ids_with_type), self.max_workers)

        def fetch_one(ident: str, type_paiement: str):
            try:
                s = requests.Session()
                s.headers.update(HEADERS_BASE)
                for c in self._session.cookies:
                    s.cookies.set(c.name, c.value)
                s.headers["Referer"] = f"{BASE_DSIJ}/IhmHF/retourPaiement.go"

                r = s.get(
                    f"{BASE_DSIJ}/IhmHF/ongletPrestation.go",
                    params={"identifiant": ident},
                    timeout=self.timeout,
                    allow_redirects=True,
                )
                if r.status_code == 200:
                    return ident, self.parse_detail_salarie(r.text, ident, type_paiement)
                return ident, pd.DataFrame()
            except Exception as e:
                logger.error("Erreur %s : %s", ident, e)
                return ident, pd.DataFrame()

        results = {}
        with ThreadPoolExecutor(max_workers=self.max_workers) as ex:
            futures = {
                ex.submit(fetch_one, ident, typ): ident
                for ident, typ in ids_with_type.items()
            }
            for future in as_completed(futures):
                ident, df_det = future.result()
                if not df_det.empty:
                    results[ident] = df_det
                    logger.info("  OK %s : %d lignes", ident, len(df_det))
                else:
                    logger.warning("  -- %s : aucune donnée", ident)

        logger.info("Terminé : %d/%d dossiers récupérés", len(results), len(ids_with_type))
        return results

    get_all_details = get_all_details_parallel

    # ------------------------------------------------------------------
    # Helpers pour la colonne "Clé"
    # ------------------------------------------------------------------

    @staticmethod
    def _date_to_mm_aaaa(date_str) -> str:
        """Convertit une date 'JJ/MM/AAAA' (ou ISO) en 'MM/AAAA'. '' si vide/invalide."""
        if date_str is None:
            return ""
        if isinstance(date_str, float) and pd.isna(date_str):
            return ""
        date_str = str(date_str).strip()
        if not date_str:
            return ""
        try:
            d = pd.to_datetime(date_str, format="%d/%m/%Y", dayfirst=True)
            return d.strftime("%m/%Y")
        except Exception:
            pass
        try:
            d = pd.to_datetime(date_str)
            return d.strftime("%m/%Y")
        except Exception:
            return ""

    @staticmethod
    def _siret_lastn(siret_val, n: int) -> str:
        """Retourne les n derniers chiffres du SIRET (chaîne, zéros conservés)."""
        try:
            if siret_val is None or (isinstance(siret_val, float) and pd.isna(siret_val)) or pd.isna(siret_val):
                return "0" * n
        except Exception:
            pass
        try:
            v = int(siret_val)
            return f"{v:014d}"[-n:]
        except Exception:
            digits = re.sub(r"\D", "", str(siret_val))
            return digits[-n:].rjust(n, "0") if digits else "0" * n

    @classmethod
    def _siret_last5(cls, siret_val) -> str:
        """Retourne les 5 derniers chiffres du SIRET (chaîne, zéros conservés)."""
        return cls._siret_lastn(siret_val, 5)

    @classmethod
    def _build_cle(cls, row) -> str:
        """
        Construit la clé : Type_DateDeb(MM/AAAA)@DateFin(MM/AAAA)_5derniersSIRET
        Exemple : 'AS_05/2024@05/2024_00117'
        """
        type_p   = str(row.get("Type", "") or "").strip()
        deb      = cls._date_to_mm_aaaa(row.get("DateDeb"))
        fin      = cls._date_to_mm_aaaa(row.get("DateFin"))
        siret5   = cls._siret_last5(row.get("SIREN/SIRET"))
        return f"{type_p}_{deb}@{fin}_{siret5}"

    # ------------------------------------------------------------------
    # EXPORT EXCEL — charte graphique DUNES
    # CORRECTIF 3 : mise en forme complète avec couleurs DUNES
    # CORRECTIF 4 : onglet "Récap par salarié" avec seulement 3 colonnes
    # ------------------------------------------------------------------

    def export_excel(
        self,
        df_paiements: "pd.DataFrame",
        details:      Optional[dict[str, "pd.DataFrame"]] = None,
        filepath:     str = "paiements_dsij.xlsx",
    ) -> str:
        if not _PANDAS_OK or not _OPENPYXL_OK:
            raise ImportError("pip install pandas openpyxl")

        # ---- 1. Construction de l'onglet "Détail des paiements" + "Clé" ----
        df_all = pd.DataFrame()
        if details:
            all_details = [d for d in details.values() if not d.empty]
            if all_details:
                df_all = pd.concat(all_details, ignore_index=True)

                # Fusion avec df_paiements pour récupérer le SIREN/SIRET
                # (nécessaire pour calculer la "Clé")
                if "identifiant" in df_paiements.columns and "Identifiant_BPIJ" in df_all.columns:
                    siret_map = (
                        df_paiements[["identifiant", "SIREN/SIRET"]]
                        .dropna(subset=["identifiant"])
                        .drop_duplicates(subset=["identifiant"])
                    )
                    df_all = df_all.merge(
                        siret_map,
                        how="left",
                        left_on="Identifiant_BPIJ",
                        right_on="identifiant",
                    )
                    # -- Demande utilisateur : enlever la colonne "identifiant"
                    #    (issue de la fusion, non significative dans cet onglet)
                    df_all.drop(columns=["identifiant"], inplace=True, errors="ignore")
                elif "SIREN/SIRET" not in df_all.columns:
                    df_all["SIREN/SIRET"] = pd.NA

                # -- Demande utilisateur : ajout de la colonne "Clé"
                df_all["Clé"] = df_all.apply(self._build_cle, axis=1)

        # ---- 1 bis. Répartition analytique : colonne Paiements + Saisie QUADRA
        referentiel = None
        df_quadra = None
        try:
            from repartition_analytique import construire_referentiel
            referentiel = construire_referentiel()
            df_paiements = self.ajouter_repartition_paiements(df_paiements, referentiel)
            df_quadra = self.construire_saisie_quadra(df_paiements, referentiel)
        except FileNotFoundError as exc:
            logger.warning("Répartition analytique Paiements ignorée : %s", exc)
        except ImportError as exc:
            logger.warning("Module repartition_analytique absent : %s", exc)
        except Exception as exc:              # noqa: BLE001
            logger.exception("Échec répartition analytique Paiements : %s", exc)

        # ---- 2. Écriture brute avec pandas (aucun formatage) ----------------
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            # Onglet 1 — Saisie QUADRA (une ligne par virement / Cle_long) :
            # placé en premier, c'est l'onglet de travail quotidien.
            if df_quadra is not None and not df_quadra.empty:
                df_quadra.to_excel(writer, sheet_name="Saisie QUADRA", index=False)

            # Onglet 2 — Paiements
            df_paiements.to_excel(writer, sheet_name="Paiements", index=False)

            # Onglet 2 — Détail des paiements
            if not df_all.empty:
                col_order = [
                    "Identifiant_BPIJ", "NIR", "Nom_Salarie", "Prenom_Salarie",
                    "Nom et prénom(s)", "Prestation_CodeNature", "Prestation_Libelle",
                    "DateDeb", "DateFin", "NB_jours", "Prix_unitaire", "Montant",
                    "Type", "Clé",
                ]
                df_detail_export = df_all[[c for c in col_order if c in df_all.columns]]
                df_detail_export.to_excel(writer, sheet_name="Détail des paiements", index=False)

            # Onglet 3 — Récap par salarié (Salarié, Montant, Clé)
            recap_salarie = None
            if not df_all.empty and "Nom et prénom(s)" in df_all.columns and "Montant" in df_all.columns:
                recap = (
                    df_all.groupby(
                        ["Nom et prénom(s)", "Clé"],
                        dropna=False,
                    )["Montant"]
                    .sum()
                    .reset_index()
                )
                recap.rename(columns={"Nom et prénom(s)": "Salarié"}, inplace=True)
                recap = recap[["Salarié", "Montant", "Clé"]]
                recap.to_excel(writer, sheet_name="Récap par salarié", index=False)
                recap_salarie = recap
            elif "identifiant" in df_paiements.columns:
                # Fallback (pas de détails disponibles) : ancien récapitulatif
                if "Montant" in df_paiements.columns:
                    recap = df_paiements.groupby("identifiant")["Montant"].sum().reset_index()
                    recap.rename(columns={"Montant": "Montant total (€)"}, inplace=True)
                else:
                    recap = df_paiements.groupby("identifiant").size().reset_index(name="Nb paiements")
                recap.to_excel(writer, sheet_name="Récap par salarié", index=False)

            # ---- Onglets 4/5/6 — Répartition analytique ---------------------
            # Rattachement de chaque IJSS reçue aux centres analytiques du
            # salarié (référentiel ANALYTIQUE 2026, centre "Ajustement" inclus).
            if recap_salarie is not None and not recap_salarie.empty:
                try:
                    from repartition_analytique import (
                        construire_ventilation_analytique, recap_par_centre,
                    )
                    # Réutilise le référentiel déjà téléchargé pour Paiements/QUADRA
                    df_vent, df_ctrl = construire_ventilation_analytique(
                        recap_salarie, referentiel=referentiel)
                    if df_vent is not None and not df_vent.empty:
                        df_vent.to_excel(writer, sheet_name="Répartition analytique",
                                         index=False)
                        recap_centres = recap_par_centre(df_vent)
                        if not recap_centres.empty:
                            recap_centres.to_excel(writer,
                                                   sheet_name="Récap centres analytiques",
                                                   index=False)
                    if df_ctrl is not None and not df_ctrl.empty:
                        df_ctrl.to_excel(writer, sheet_name="Contrôle matching",
                                         index=False)
                        a_verifier = int((df_ctrl["Statut"] != "OK").sum())
                        if a_verifier:
                            logger.warning(
                                "Répartition analytique : %d salarié(s) à vérifier "
                                "(onglet \"Contrôle matching\")", a_verifier)
                except FileNotFoundError as exc:
                    logger.warning("Répartition analytique ignorée : %s", exc)
                except ImportError as exc:
                    logger.warning("Module repartition_analytique absent : %s", exc)
                except Exception as exc:          # noqa: BLE001
                    logger.exception("Échec de la répartition analytique : %s", exc)

        # ---- 3. Ré-ouvrir et appliquer la charte DUNES -----------------------
        wb = openpyxl.load_workbook(filepath)
        for ws in wb.worksheets:
            center_cols = SHEET_CENTER_COLS.get(ws.title, [])
            id_number_cols = SHEET_ID_NUMBER_FORMAT_COLS.get(ws.title, [])
            self._apply_dunes_style(ws, center_cols=center_cols, id_number_format_cols=id_number_cols)
        wb.save(filepath)

        logger.info("Fichier Excel DUNES créé : %s", filepath)
        return filepath

    # ------------------------------------------------------------------
    # Formatage charte DUNES
    # ------------------------------------------------------------------

    def _apply_dunes_style(
        self,
        ws,
        center_cols: Optional[list[str]] = None,
        id_number_format_cols: Optional[list[str]] = None,
    ) -> None:
        """
        Applique la charte DUNES à un onglet openpyxl :
          - En-tête  : fond rouge #AF321E, texte blanc gras, Arial 10
          - Lignes paires  : fond orange très clair (#FEF0DC ~ orange 15%)
          - Lignes impaires: fond blanc cassé (#FAFAFA)
          - Bordure fine gris clair sur toutes les cellules
          - Colonnes auto-élargies
          - Ligne d'en-tête figée
          - `center_cols` : en-têtes (insensible à la casse) dont les colonnes
            doivent être centrées, quel que soit le type de valeur.
          - `id_number_format_cols` : en-têtes dont les colonnes numériques
            doivent garder un format entier simple (pas de séparateur de
            milliers, pas de notation scientifique) — ex : SIRET, NIR.
        """
        center_cols_lower = {c.lower() for c in (center_cols or [])}
        id_number_format_cols_lower = {c.lower() for c in (id_number_format_cols or [])}

        # Couleurs
        fill_header = PatternFill("solid", fgColor=DUNES_ROUGE)
        fill_pair   = PatternFill("solid", fgColor="FEF0DC")   # orange très pâle
        fill_impair = PatternFill("solid", fgColor="FAFAFA")

        font_header = Font(name="Arial", bold=True, color=DUNES_BLANC, size=10)
        font_data   = Font(name="Arial", size=10)
        font_total  = Font(name="Arial", bold=True, size=10)

        align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
        align_left   = Alignment(horizontal="left",   vertical="center", wrap_text=False)
        align_right  = Alignment(horizontal="right",  vertical="center", wrap_text=False)
        # Cellules multi-lignes (ex. « Répartition analytique » de Saisie
        # QUADRA) : renvoi à la ligne activé, sinon Excel affiche tout sur
        # une seule ligne tant qu'on n'a pas cliqué dans la cellule.
        align_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        max_col = ws.max_column
        max_row = ws.max_row

        # --- En-têtes (ligne 1) ---
        for cell in ws[1]:
            cell.fill      = fill_header
            cell.font      = font_header
            cell.alignment = align_center
            cell.border    = border

        # Carte colonne -> nom d'en-tête (en minuscule)
        headers_lower = {
            cell.column: (str(cell.value) if cell.value is not None else "").lower()
            for cell in ws[1]
        }

        # --- Données (lignes 2+) ---
        rows_multilignes: set[int] = set()
        for row_idx in range(2, max_row + 1):
            fill = fill_pair if row_idx % 2 == 0 else fill_impair
            for cell in ws[row_idx]:
                cell.fill   = fill
                cell.border = border

                col_name = headers_lower.get(cell.column, "")

                # Police
                if row_idx == max_row and "total" in str(cell.value or "").lower():
                    cell.font = font_total
                else:
                    cell.font = font_data

                val = cell.value

                # Format numérique
                if isinstance(val, (int, float)):
                    if col_name in id_number_format_cols_lower:
                        # Colonnes "identifiantes" (SIRET, NIR, identifiant...) :
                        # format entier simple, sans séparateur de milliers
                        cell.number_format = "0"
                    elif "montant" in col_name or "prix" in col_name or "total" in col_name:
                        cell.number_format = '#,##0.00 "€"'

                # Alignement : priorité aux colonnes demandées en centré
                if col_name in center_cols_lower:
                    cell.alignment = align_center
                elif isinstance(val, (int, float)):
                    cell.alignment = align_right
                elif isinstance(val, str) and "\n" in val:
                    cell.alignment = align_left_wrap
                    rows_multilignes.add(row_idx)
                else:
                    cell.alignment = align_left

        # --- Largeur des colonnes (auto) ---
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(1, min(max_row + 1, 200)):   # sample 200 lignes max
                val = ws.cell(row_idx, col_idx).value
                if val is not None:
                    # Multi-lignes : la largeur se base sur la plus longue LIGNE
                    max_len = max(max_len, *(len(l) for l in str(val).split("\n")))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)

        # --- Hauteur des lignes ---
        ws.row_dimensions[1].height = 22   # en-tête un peu plus haut
        for row_idx in range(2, max_row + 1):
            if row_idx in rows_multilignes:
                # Hauteur remise en automatique : Excel l'ajuste au nombre de
                # lignes de la cellule (wrap_text), au lieu des 16 px fixes
                # (None efface aussi une hauteur posée par un passage précédent).
                ws.row_dimensions[row_idx].height = None
                continue
            ws.row_dimensions[row_idx].height = 16

        # --- Figer la ligne d'en-tête ---
        ws.freeze_panes = "A2"

        # --- Filtre automatique sur les en-têtes ---
        if max_row > 1:
            ws.auto_filter.ref = ws.dimensions

    # ------------------------------------------------------------------
    # All-in-one
    # ------------------------------------------------------------------

    def run_with_details(
        self,
        date_debut:      str  = "01/01/2026",
        date_fin:        str  = "",
        output_excel:    str  = "paiements_dsij_details.xlsx",
        include_details: bool = True,
        montants_recus:  bool = True,
    ):
        """
        Flux complet :
          login -> open_attestation -> get_tableau
          -> [get_all_details] -> [journees BPIJ] -> export_excel

        `montants_recus=True` ajoute a l'onglet « Paiements » les colonnes
        "Montant reçu" (total a payer de la journee BPIJ) et "Montant virement"
        (net reellement credite en banque pour la date), plus l'identifiant
        "Journée BPIJ".

        Returns
        -------
        (df_paiements, details_dict)
        """
        self.login()
        self.open_attestation()
        df = self.get_tableau(date_debut=date_debut, date_fin=date_fin)

        details = None
        if include_details and not df.empty:
            details = self.get_all_details(df)

        # La vue BPIJ est chargee EN DERNIER : elle relance une recherche
        # cote serveur avec typeRecherche=bpij et changerait le contexte
        # utilise par get_all_details.
        if montants_recus and not df.empty:
            df_bpij = None
            try:
                df_bpij = self.get_tableau_bpij(date_debut=date_debut, date_fin=date_fin)
            except Exception as exc:
                logger.warning("Journees BPIJ indisponibles (%s) : les montants "
                               "recus sont calcules sans identifiant ni controle", exc)
            df = self.ajouter_montants_recus(df, df_bpij)

        self.export_excel(df, details, output_excel)
        return df, details
