"""
lancer_extraction.py
====================
Lance l'extraction complete Net-Entreprises / DSIJ et produit le classeur
Excel avec les onglets analytiques.

A executer depuis le dossier "NET ENTREPRISE" :

    python lancer_extraction.py
    python lancer_extraction.py --debut 20/12/2025
    python lancer_extraction.py --debut 20/12/2025 --fin 31/08/2026
    python lancer_extraction.py --hors-ligne          # referentiel = cache local
    python lancer_extraction.py --rafraichir          # MAJ des caches, sans extraction

L'identite du compte declarant et le mot de passe sont lus dans le fichier
.env (SIRET, NOM, PRENOM, PASSWORD) ; rien n'est ecrit en dur dans le code
ni dans le classeur produit.

Onglets produits
----------------
  Paiements                  (dont "Montant reçu" / "Montant virement" / "Journée BPIJ")
  Détail des paiements
  Récap par salarié
  Répartition analytique     <- ventilation des IJSS par centre analytique
  Récap centres analytiques
  Contrôle matching          <- salaries non rapproches, a traiter en premier
"""

from __future__ import annotations

import os
import sys
import logging
import argparse
import datetime

from net_entreprise_final import NetEntreprise

# ---------------------------------------------------------------------------
# Parametres (le compte declarant est dans .env)
# ---------------------------------------------------------------------------
DATE_DEBUT_DEFAUT = "20/12/2025"

DOSSIER = os.path.dirname(os.path.abspath(__file__))


def lire_env(chemin: str | None = None) -> dict[str, str]:
    """Lit le fichier .env (lignes CLE=valeur, # pour les commentaires)."""
    chemin = chemin or os.path.join(DOSSIER, ".env")
    if not os.path.exists(chemin):
        raise FileNotFoundError(
            f"Fichier .env introuvable : {chemin}\n"
            "Copiez .env.example en .env et renseignez-le."
        )
    valeurs: dict[str, str] = {}
    with open(chemin, encoding="utf-8") as fh:
        for ligne in fh:
            ligne = ligne.strip()
            if not ligne or ligne.startswith("#") or "=" not in ligne:
                continue
            cle, valeur = ligne.split("=", 1)
            valeurs[cle.strip().upper()] = valeur.strip().strip('"').strip("'")
    manquantes = [c for c in ("SIRET", "NOM", "PRENOM", "PASSWORD")
                  if not valeurs.get(c)]
    if manquantes:
        raise ValueError(f"Cle(s) absente(s) ou vide(s) dans .env : {manquantes}")
    return valeurs


def main() -> int:
    parseur = argparse.ArgumentParser(description="Extraction Net-Entreprises / DSIJ")
    parseur.add_argument("--debut", default=DATE_DEBUT_DEFAUT,
                         help=f"date de debut JJ/MM/AAAA (defaut : {DATE_DEBUT_DEFAUT})")
    parseur.add_argument("--fin", default="",
                         help="date de fin JJ/MM/AAAA (defaut : aujourd'hui)")
    parseur.add_argument("--sortie", default="",
                         help="nom du fichier Excel de sortie")
    parseur.add_argument("--sans-details", action="store_true",
                         help="ne pas telecharger le detail salarie par salarie")
    parseur.add_argument("--hors-ligne", action="store_true",
                         help="ne pas relire les classeurs Google : utiliser le cache local")
    parseur.add_argument("--rafraichir", action="store_true",
                         help="mettre a jour les caches analytiques puis quitter")
    args = parseur.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    # net_entreprise_final installe son propre handler : sans cela chaque
    # message serait affiche deux fois (le sien + celui de la racine).
    logging.getLogger("net_entreprise").propagate = False

    # -- Rafraichissement seul du referentiel analytique --------------------
    if args.rafraichir:
        from repartition_analytique import rafraichir_cache
        rafraichir_cache()
        return 0

    if args.hors_ligne:
        import repartition_analytique as ra
        _construire = ra.construire_ventilation_analytique

        def hors_ligne(recap, *a, **kw):
            kw["en_ligne"] = False
            return _construire(recap, *a, **kw)

        ra.construire_ventilation_analytique = hors_ligne
        logging.info("Referentiel analytique : mode hors ligne (cache local)")

    dossier_sorties = os.path.join(DOSSIER, "data", "sorties")
    os.makedirs(dossier_sorties, exist_ok=True)
    sortie = args.sortie or os.path.join(
        dossier_sorties,
        "paiements_analytique_{}.xlsx".format(
            args.debut.replace("/", "")[-8:] or datetime.date.today().strftime("%d%m%Y")
        ),
    )

    env = lire_env()
    ne = NetEntreprise(
        siret=env["SIRET"],
        nom=env["NOM"],
        prenom=env["PRENOM"],
        password=env["PASSWORD"],
    )

    logging.info("Extraction du %s au %s", args.debut, args.fin or "aujourd'hui")
    df, details = ne.run_with_details(
        date_debut=args.debut,
        date_fin=args.fin,
        output_excel=sortie,
        include_details=not args.sans_details,
        montants_recus=True,
    )

    logging.info("Termine : %d paiement(s) -> %s", len(df), sortie)

    # --- Rappel des points a verifier -------------------------------------
    try:
        import pandas as pd
        controle = pd.read_excel(sortie, sheet_name="Contrôle matching")
        a_voir = controle[controle["Statut"] != "OK"].drop_duplicates(subset=["Salarié"])
        print()
        if len(a_voir):
            print("=" * 78)
            print(f"{len(a_voir)} salarie(s) non rattache(s) a un centre analytique :")
            for ligne in a_voir.itertuples(index=False):
                salarie = getattr(ligne, "Salarié", "")
                statut = getattr(ligne, "Statut", "")
                candidats = getattr(ligne, "_14", "") or ""
                print(f"  - {salarie:<28} {statut:<16} {candidats}")
            print("\nAjoutez-les dans config/alias_salaries.csv puis relancez l'export.")
            print("=" * 78)
        else:
            print("Tous les salaries sont rattaches a un centre analytique.")
    except Exception as exc:                      # noqa: BLE001
        logging.debug("Resume du controle indisponible : %s", exc)

    return 0


if __name__ == "__main__":
    sys.exit(main())
