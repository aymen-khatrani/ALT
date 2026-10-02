# Pipeline qualité des exigences client : Python → agent Copilot → Python

Python fait tout ce qui demande de l'exactitude (lire l'export, repérer les mots à risque, découper, vérifier). L'agent Copilot fait seulement ce qui demande de comprendre la phrase : juger, expliquer, reformuler. Python ne fait pas confiance à l'agent : il revérifie tout ce que l'agent renvoie.

```
Export .CUS.xlsx ─► prepare.py ─► lots/lot_XX.xlsx ─► Agent Copilot ─► resultats/resultats_lot_XX.xlsx ─► controle.py ─► _validation.xlsx ─► ingénieurs ─► bilan.py
                    (lexique,                          (jugement,                                         (8 contrôles,                      (décisions)     (indicateurs)
                     découpage)                         reformulation)                                     classeur)
```

## Contenu

| Fichier | Rôle |
|---|---|
| `prepare.py` | Lit l'export, détecte les colonnes, écarte titres et informations, applique le lexique, découpe en lots |
| `controle.py` | Fusionne les résultats de l'agent, contrôle chaque ligne, produit le classeur de validation |
| `bilan.py` | Calcule les indicateurs du PoC à partir des décisions des relecteurs |
| `commun.py` | Code partagé (catalogue des règles, lecture Excel, lexique) |
| `lexique.csv` | Mots et motifs à repérer, par règle. Modifiable dans Excel |
| `instructions_agent_CUS_lots.txt` | Instructions de l'agent pour ce mode (7 728 caractères sur 8 000) |
| `exemple/` | Faux export, faux résultats d'agent (avec erreurs volontaires) et référence annotée, pour la démo |

## Installation (une fois)

Python 3.8 ou plus récent, et une seule bibliothèque :

```
pip install openpyxl
```

Avec Anaconda, `openpyxl` est déjà installé.

## Démo en 2 minutes (données fictives)

```
python prepare.py exemple/exemple.CUS.xlsx --taille-lot 8
python controle.py exemple/exemple_pipeline --resultats exemple/resultats_simules --reference exemple/reference.xlsx
python bilan.py exemple/exemple_pipeline/exemple_validation.xlsx
```

Les faux résultats contiennent 8 erreurs volontaires : un ID inventé, un signal non traité, un chiffre ajouté, un extrait inexistant, une question en français, un verdict incohérent, une ligne oubliée et un lot entier absent. `controle.py` doit toutes les signaler, et lister les lots 2 et 3 à relancer. Pour voir le bilan, remplis quelques cellules « Décision » du classeur, puis relance `bilan.py`.

## Utilisation réelle

### 1. Préparer
```
python prepare.py "C:\...\Projet.CUS.xlsx"
```
Vérifie les trois colonnes affichées (ID, texte, type). Si l'une est fausse :

| Option | Effet |
|---|---|
| `--col-texte "Object Text"` | force la colonne texte (idem `--col-id`, `--col-type`) |
| `--col-type aucune` | analyse toutes les lignes qui ont du texte |
| `--types-exigence "Requirement,Constraint"` | seules ces valeurs de type sont analysées |
| `--feuille "Export"` | choisit la feuille |
| `--taille-lot 20` | exigences par lot (20 par défaut) |

Le script crée `Projet_pipeline/` avec le fichier maître, les lots et un dossier `resultats/` vide. Il affiche aussi une ligne `SIGNATURE` (nombre de lignes, lots, empreinte du lexique) : garde-la pour la traçabilité.

### 2. Analyser avec l'agent
Pour chaque fichier `lots/lot_XX.xlsx` :
1. Joins le fichier dans l'agent et écris « Analyse ce lot ».
2. Télécharge `resultats_lot_XX.xlsx` dans `Projet_pipeline/resultats/`.
3. Si l'agent ne parvient pas à créer le fichier : tape `json`, puis colle sa réponse dans `resultats/resultats_lot_XX.txt`. `controle.py` sait la lire.

Ouvre une nouvelle conversation tous les 5 lots environ : la qualité baisse quand la conversation s'allonge.

### 3. Contrôler
```
python controle.py "C:\...\Projet_pipeline"
```
Le script produit `Projet_validation.xlsx` et affiche les lots à relancer. Relance ces lots dans l'agent, dépose les nouveaux fichiers dans `resultats/`, puis relance `controle.py`. Pour un même ID, c'est le fichier le plus récent qui l'emporte. Les décisions déjà saisies par les relecteurs sont reprises, et l'ancienne version est sauvegardée en `.bak.xlsx`.

### 4. Faire relire par les ingénieurs
Dans l'onglet « Validation », les couleurs d'en-tête indiquent qui remplit quoi :
- gris : la source ;
- bleu : l'agent ;
- orange : les contrôles ;
- vert : le relecteur.

Les valeurs de « Décision » :

| Décision | Sens |
|---|---|
| Acceptée | Vrai problème : question et formulation reprises telles quelles |
| Modifiée | Vrai problème : texte corrigé dans « Question finale » |
| Rejetée | Vrai problème, mais pas de question au client |
| Faux positif | Ce n'est pas un problème |
| Oubli agent | Sur une ligne OK : le relecteur trouve un problème que l'agent a manqué |

### 5. Mesurer
```
python bilan.py "C:\...\Projet_pipeline\Projet_validation.xlsx"
```
Le script calcule l'avancement, la précision de l'agent, les taux de reprise, les faux positifs, les oublis et un rappel estimé, avec le détail par règle et par sévérité. Il écrit le résultat dans `Projet_bilan.xlsx`.

## Ce que vérifie controle.py

| Anomalie | Signification | Action |
|---|---|---|
| `MANQUANT` | Aucun résultat pour cette exigence | Relancer le lot |
| `VERDICT_INVALIDE` | Verdict hors A_CLARIFIER / OK / HORS_PERIMETRE | Relancer le lot |
| `ID_INCONNU`, `DOUBLON` | ID inventé, ou présent deux fois dans un fichier | Ignoré, à surveiller |
| `SIGNAL_NON_ARBITRE` | Un signal du lexique n'est ni confirmé ni écarté | Vérifier la ligne |
| `EXTRAIT_INTROUVABLE` | L'extrait cité n'existe pas dans le texte | Vérifier la ligne |
| `CHIFFRE_AJOUTE` | La formulation contient un chiffre absent du texte d'origine | Vérifier : valeur inventée ? |
| `INCOHERENCE`, `CHAMP_MANQUANT`, `CHAMP_EN_TROP` | Verdict, sévérité et champs ne concordent pas | Vérifier la ligne |
| `LANGUE` | La question ou la formulation semble en français | Vérifier la ligne |
| `REGLE_INCONNUE`, `CONFIANCE_INVALIDE` | Valeur hors catalogue | Vérifier la ligne |

L'onglet « Synthèse » donne, pour chaque règle, combien de signaux du lexique l'agent a confirmés et combien de problèmes il a trouvés seul. C'est ce qui permet d'ajuster le lexique.

## Configurer l'agent pour ce mode

Même procédure que `configuration_agent_CUS.md`, avec ces différences :

| Champ | Valeur |
|---|---|
| Nom | Qualité Exigences CUS – lots |
| Instructions | `instructions_agent_CUS_lots.txt` |
| Connaissances | `Requirements_writing_guide.docx` |
| Créer des documents, graphiques et code | Activé |
| Amorces | « Analyse ce lot. » · `détail <ID>` · `variante <ID>` · `json` |

## Modifier le lexique

`lexique.csv` s'ouvre dans Excel (séparateur `;`). Il a quatre colonnes : `regle`, `motif`, `type`, `commentaire`.

Le champ `type` peut prendre ces valeurs :
- `mot` : mot ou expression exacte, sans tenir compte des majuscules ;
- `regex` : expression régulière ;
- `regex_casse` : expression régulière sensible aux majuscules ;
- `special` : contrôle codé dans `commun.py` (`aucun_shall`, `plusieurs_shall`).

Chaque modification change l'empreinte du lexique affichée dans la SIGNATURE. Pour ajuster le lexique, appuie-toi sur l'onglet Synthèse :
- un mot dont le taux de confirmation reste sous 30 % fait surtout du bruit ;
- les problèmes que l'agent trouve seul donnent des mots à ajouter.

## Limites connues

- L'agent ne donne pas toujours le même résultat. Pour mesurer sa stabilité, passe deux fois le même lot dans deux dossiers séparés, puis compare avec `--resultats`.
- Deux contrôles reposent sur des heuristiques :
  - la langue : quelques mots français courants ;
  - les chiffres : seuls les chiffres écrits en nombres sont vérifiés, pas ceux écrits en lettres.
- Le lexique est en anglais uniquement, comme les exigences.
