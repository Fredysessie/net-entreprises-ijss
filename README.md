# net-entreprises-ijss

Automatisation du traitement des **indemnités journalières de sécurité sociale
(IJSS)** pour l'Association DUNES : récupération des paiements CPAM sur
Net-Entreprises / DSIJ, rapprochement bancaire, puis ventilation de chaque
remboursement sur les **centres analytiques** du salarié, prête à saisir dans
QUADRA.

---

## 1. Le contexte

Quand un salarié est en arrêt de travail, l'employeur maintient tout ou partie
du salaire et la CPAM lui rembourse ensuite les indemnités journalières. Trois
difficultés rendent le rapprochement comptable pénible à la main :

**Les paiements arrivent groupés.** La CPAM règle par lots. Ce qui est
réellement crédité en banque est le *total à payer* de la journée BPIJ — après
retenues et indus — et non la somme des remboursements individuels. Il faut
donc distinguer le montant dû par dossier, le total de la journée et le
virement réel de la date.

**Un remboursement n'appartient pas à un seul budget.** Un salarié est réparti
entre plusieurs dispositifs (le *dispositif* est le centre analytique), dans
des proportions qui changent chaque mois. Le remboursement doit suivre la même
clé que le salaire qu'il compense, sinon les comptes de résultat par dispositif
sont faux. Cette répartition est arrêtée en réunion avec les directeurs
territoriaux : **tant que la réunion n'a pas eu lieu, le mois n'existe pas** —
d'où les replis de mois décrits plus bas.

**Les noms ne concordent pas.** La CPAM désigne les salariés autrement que la
paie : nom de naissance au lieu du nom d'usage, prénom tronqué, accents perdus,
second prénom ajouté, coquilles. Deux libellés différents peuvent viser la même
personne — et deux libellés proches, deux personnes différentes.

S'ajoutent deux particularités de gestion :

- **Salariés sortis.** La CPAM rembourse souvent après le départ. La
  répartition analytique est alors réactivée sur un mois postérieur à la sortie
  pour pouvoir imputer le remboursement.
- **Bascule de paie.** Jusqu'à fin 2025 la paie était gérée sous QUADRA, depuis
  janvier 2026 sous SILAE. Les exports SILAE portent un
  `Date du changement = 01/01/2026` pour presque tout le monde : la vraie date
  d'entrée des anciens est celle de QUADRA. Le **matricule**, commun aux deux
  systèmes, sert de clé pivot.

---

## 2. Installation

```bash
pip install -r requirements.txt
cp .env.example .env          # puis renseigner les quatre clés
cp config/alias_salaries.exemple.csv config/alias_salaries.csv
```

`.env` porte l'identité du compte déclarant Net-Entreprises et rien d'autre :

```
SIRET=
NOM=
PRENOM=
PASSWORD=
```

Aucune de ces valeurs n'apparaît dans le code ni dans le classeur produit, et
`.env` n'est pas versionné.

---

## 3. Utilisation

```bash
python lancer_extraction.py                      # depuis le 20/12/2025
python lancer_extraction.py --debut 01/01/2026
python lancer_extraction.py --debut 20/12/2025 --fin 31/08/2026
python lancer_extraction.py --hors-ligne         # référentiel = cache local
python lancer_extraction.py --rafraichir         # MAJ des caches, sans extraction
python repartition_analytique.py                 # diagnostic des sources
```

Le classeur produit va dans `data/sorties/paiements_analytique_<ddmmaaaa>.xlsx`.

---

## 4. Les six onglets produits

| Onglet | Contenu |
|---|---|
| **Paiements** | Un paiement par ligne, avec `Montant reçu` (total à payer de la journée BPIJ), `Montant virement` (net réellement crédité pour la date) et `Journée BPIJ` |
| **Détail des paiements** | Le détail prestation par prestation, avec `Clé` = `Type_MM/AAAA@MM/AAAA_5 derniers du SIRET` |
| **Récap par salarié** | Salarié, Montant, Clé |
| **Répartition analytique** | Une ligne par salarié × clé × centre analytique, avec taux et montant ventilé |
| **Récap centres analytiques** | Total des IJSS par centre — contrôle global avant saisie |
| **Contrôle matching** | Une ligne par paiement : statut du rapprochement, matricule, dates d'entrée/sortie, mois utilisé, commentaires. **Les anomalies sont en tête de tableau.** |

Deux invariants tenus à chaque exécution :

- la somme des lignes ventilées est **strictement** égale au montant CPAM —
  l'écart d'arrondi est reporté sur la plus grosse ligne ;
- chaque couple (mois, salarié) du référentiel totalise **100 %**, vérifié au
  chargement ; tout écart est signalé dans le journal.

---

## 5. D'où viennent les données analytiques

Tout est lu **en ligne**, à chaque exécution, dans les classeurs Google Sheets.
Aucun référentiel à maintenir à la main.

| | ANALYTIQUE 2025 | ANALYTIQUE 2026 |
|---|---|---|
| Répartition | `Répartition_analytique_2025` | `Répartition_analytique_2026` |
| Ajustement | `Ajustement` | `Ajustement` |
| Personnel | `PERSONNEL` (export QUADRA) | `Base_Personnel_SILAE` |

Les quatre premières colonnes des onglets de répartition sont lues :
`Mois (MM/AAAA) | Nom et prénom(s) | Dispositif | Taux`.

Le centre **« Ajustement »** vit dans son propre onglet et porte des **taux
négatifs** qui ramènent chaque salarié à 100 % sur le mois. Il est fusionné avec
la répartition : il apparaît donc normalement dans la ventilation, avec un
montant négatif. L'omettre reviendrait à sur-imputer les salariés répartis à
110 % ou 120 %.

Les identifiants de classeurs et les `gid` de chaque onglet sont réunis dans le
dictionnaire `SOURCES`, en tête de `repartition_analytique.py`. **Ajouter 2027
se résume à y ajouter une entrée.**

### Cache local

Après chaque lecture réussie, deux fichiers sont réécrits dans `cache/` :
`repartition_analytique.csv` et `personnel_dunes.csv`. Si les classeurs sont
inaccessibles — pas de réseau, partage retiré — le module bascule dessus et le
signale. Ce sont des fichiers **générés** : ne pas les éditer, ils ne sont pas
versionnés.

> L'accès en ligne suppose que les classeurs restent partagés par lien
> (*Partager → Tous les utilisateurs disposant du lien → Lecteur*).

---

## 6. Le rapprochement des noms

Le module compare des **jeux de jetons normalisés** — majuscules, accents et
ponctuation supprimés, ordre indifférent — et applique dans l'ordre :

| Méthode | Exemple *(noms fictifs)* | Score |
|---|---|---|
| `alias` | ligne manuelle dans `config/alias_salaries.csv` | 100 |
| `exact` | `DUPONT MARIE` = `DUPONT Marie` | 100 |
| `jetons` | `MARIE DUPONT` = `DUPONT Marie` | 99 |
| `inclusion` | `DUPONT MARIE CLAIRE` ⊃ `DUPONT Marie` | 88-98 |
| `nom de naissance` | `LEROY Marie` → `DUPONT Marie née LEROY` | 86-94 |
| `flou` | `DUPOND Marie` ≈ `DUPONT Marie` | variable |

Règles de sécurité : en dessous de **88** rien n'est ventilé ; si les deux
meilleurs candidats sont à moins de 2 points, le cas est marqué `AMBIGU` et
laissé de côté. **Aucun montant n'est imputé sur une supposition.** Les cas non
tranchés arrivent en tête de l'onglet *Contrôle matching*, avec les trois
candidats les plus proches et leur score.

Le **matricule** réunit sous une seule identité les répartitions 2025 et 2026
d'un même salarié, même si son libellé a changé entre QUADRA et SILAE. C'est ce
qui permet à un arrêt de décembre 2025 d'être ventilé sur la répartition 2025.

### Corriger un cas non résolu

Ajouter une ligne à `config/alias_salaries.csv` :

```csv
Nom CPAM;Nom analytique
LEROY MARIE;DUPONT Marie
```

Colonne 1 : le libellé exact vu dans *Récap par salarié*. Colonne 2 : le libellé
exact du référentiel analytique. Les lignes commençant par `#` sont ignorées. Un
alias est prioritaire sur toute détection automatique.

Ce fichier contient des données nominatives : il n'est **pas versionné**.
`config/alias_salaries.exemple.csv` sert de modèle. Pensez à le sauvegarder
ailleurs, c'est le seul fichier qui ne se régénère pas tout seul.

---

## 7. Lire l'onglet Contrôle matching

| Commentaire | Signification |
|---|---|
| `sorti(e) le JJ/MM/AAAA` | Le salarié avait quitté DUNES quand l'arrêt a commencé. Normal : la CPAM rembourse après le départ. |
| `MM/AAAA absent -> MM/AAAA` | Pas de répartition saisie pour ce mois ; repli sur le mois disponible le plus proche, antérieur de préférence. |
| `aucun centre pour ce mois` | Salarié identifié mais sans aucune répartition : rien n'est ventilé. |
| `ATTENTION total taux = ...` | Le mois ne totalise pas 100 % dans le classeur. Le montant est ventilé au prorata — **à corriger à la source**. |

Un repli massif sur le dernier mois disponible est attendu en fin d'année :
la répartition d'un mois n'existe qu'après la réunion avec les directeurs
territoriaux.

Le journal signale aussi, en les nommant, les lignes en double dans les onglets
de répartition et les salariés présents dans la répartition sans fiche du
personnel.

---

## 8. Arborescence

```
net-entreprises-ijss/
├─ net_entreprise_final.py      client Net-Entreprises / DSIJ, parsing, export Excel charté DUNES
├─ repartition_analytique.py    référentiel analytique, rapprochement des noms, ventilation
├─ lancer_extraction.py         lanceur en ligne de commande
├─ config/
│   ├─ alias_salaries.exemple.csv   modèle versionné
│   └─ alias_salaries.csv           correspondances réelles — non versionné
├─ cache/                       référentiels téléchargés — généré, non versionné
├─ data/
│   ├─ entrees/                 exports XML / CSV de la CPAM — non versionné
│   └─ sorties/                 classeurs produits — non versionné
└─ archives/                    sauvegardes horodatées du code — non versionné
```

---

## 9. Données et confidentialité

Le dépôt ne contient **que du code**. Restent hors dépôt, via `.gitignore` :

- `.env` — mot de passe Net-Entreprises ;
- `data/entrees/` — exports CPAM, qui portent les NIR et les périodes d'arrêt ;
- `data/sorties/` — classeurs nominatifs avec montants par personne ;
- `cache/` et `config/alias_salaries.csv` — noms, noms de naissance, matricules,
  dates d'entrée et de sortie.

Ces données sont des données de santé et des données RH : elles restent dans le
système d'information de DUNES. Tout est régénérable en une commande à partir
des classeurs Google et du portail, sauf le fichier d'alias.

---

## 10. Dépendances

```
pandas  openpyxl  requests  beautifulsoup4  lxml
```
