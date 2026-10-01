# Agent Copilot : qualité des exigences client (.CUS)

Fiche de configuration pour **Agent Builder** (Microsoft 365 Copilot → Créer un agent → onglet *Configurer*).
Les instructions sont dans `instructions_agent_CUS.txt` : copie-colle le fichier entier, tel quel.

## 1. Réglages

| Champ | Valeur |
|---|---|
| Nom (30 car. max) | Qualité Exigences CUS |
| Description | Voir ci-dessous |
| Instructions | Contenu de `instructions_agent_CUS.txt` (7 813 caractères sur 8 000) |
| Connaissances | `Requirements_writing_guide.docx` (version réécrite pour l'agent, une section par règle) |
| Recherche web | Désactivée |
| Créer des documents, graphiques et code | **Activé** (c'est ce qui lit et écrit les Excel) |
| Créer des images | Désactivé |
| Partage | Privé pendant les tests |

**Description**

> Analyse un export d'exigences client (.CUS.xlsx) et repère les exigences ambiguës, non vérifiables ou à risque contractuel, selon le guide de rédaction des exigences. Pour chaque exigence à clarifier, l'agent propose une question au client et une formulation à faire confirmer, en anglais. Il travaille par lots de 20, vérifie ses résultats par code et produit un classeur Excel où l'ingénieur système valide chaque proposition.

## 2. Amorces de conversation

| Titre | Message |
|---|---|
| Analyser un fichier .CUS | Analyse le fichier .CUS.xlsx joint. |
| Lot suivant | suite |
| Expliquer une exigence | détail <ID> |
| Autre formulation | variante <ID> |
| Exporter le classeur de validation | export |

## 3. Pourquoi ce format

- **Structure recommandée par Microsoft** pour les agents de ce type : OBJECTIVE → RESPONSE RULES → WORKFLOW (chaque étape en *Goal / Action / Transition*) → OUTPUT FORMAT → EXAMPLES → FINAL CHECK.
- **Instructions en anglais.** Les exigences et les mots-clés à repérer sont en anglais, et l'anglais prend environ 15 % de caractères en moins, ce qui aide à tenir dans les 8 000. L'agent répond quand même en français.
- **Catalogue de règles dans les instructions, pas en pièce jointe.** Les fichiers de connaissance ne sont pas relus en entier : l'agent n'en consulte que des extraits trouvés par recherche. Une règle placée uniquement là serait appliquée de façon irrégulière. Microsoft déconseille aussi de déplacer des instructions vers SharePoint pour contourner la limite.
- **Sévérité définie par l'impact, pas par règle.** Bloquant = vérification impossible ; majeur = deux lectures possibles ; mineur = style seul. L'agent peut ainsi juger un cas que le catalogue n'avait pas prévu.
- **Exemples 2 et 3 = cas à ne pas signaler.** Ils empêchent l'agent de tout signaler (une spec client tolère des écarts de style).
- **Étape 3 = contrôles exécutés par code.** L'agent vérifie qu'aucun ID ne manque, que chaque extrait existe vraiment dans le texte et qu'aucun chiffre n'a été inventé. Il ne se contente pas de l'affirmer.
- **`[TBC: …]`** : quand une valeur manque, c'est au client de la fixer, pas à l'agent.

## 4. Premier test (10 min)

Crée un petit `test.CUS.xlsx` avec les colonnes `ID | Object Type | Text`. Mets-y un titre, une ligne d'information et les exigences ci-dessous.

**Attention :** ces exigences sont analysées en entier dans le fichier de connaissance (`Requirements_writing_guide.docx`, section 7). Ce test vérifie donc que l'agent **retrouve et applique le guide**, pas son jugement sur des cas nouveaux. Pour mesurer ce jugement, ajoute 10 exigences anonymisées qui ne sont pas dans le guide.

| Exigence (tirée du guide) | Résultat attendu |
|---|---|
| GIVEN the train is in normal operation WHEN requested THEN the bodyside doors shall open. | A_CLARIFIER · MAJEUR · AMB-01 (+ AMB-02 « normal operation », confiance basse) |
| RATP will determine the safe speed limit of train and speed profile that VATP will use to vitally manage speed. | A_CLARIFIER · MAJEUR · RSK-04, RSK-05, STY-02 |
| Communicate with ATS through RATO via Data Transmission System (DTS) connection. | A_CLARIFIER · MAJEUR · AMB-01, RSK-07, AMB-04 (confiance basse sur les acronymes) |
| The Report shall be sent within 1 hour of it being approved. | A_CLARIFIER · MAJEUR · AMB-01, AMB-04 |
| Titre et ligne d'information | HORS_PERIMETRE |

Ce qu'il faut vérifier pendant le test :
1. À l'étape 1, l'agent identifie les bonnes colonnes et demande confirmation.
2. Le message « Contrôles : OK » s'affiche et le lien de téléchargement fonctionne.
3. Au message `suite`, le fichier `Resultats_CUS.xlsx` existe encore. Sinon, l'agent doit basculer sur un fichier par lot.
4. Avec `export`, le classeur contient bien la liste déroulante « Décision » et l'onglet « Synthèse ».
5. Les questions et formulations sont en anglais, et aucun chiffre n'a été ajouté.

Note chaque écart : c'est sur ces écarts qu'on ajustera les instructions.

## 5. Limites connues

- Pas de réglage de température : passe deux fois le même lot pour mesurer la stabilité.
- Les fichiers générés ne restent disponibles que pendant la conversation. Télécharge-les tout de suite.
- Un fichier .CUS protégé par une étiquette de confidentialité qui bloque l'extraction ne pourra pas être lu par l'agent.
- Il reste environ 190 caractères de marge : toute nouvelle règle devra en remplacer une autre.
