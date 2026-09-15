<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

## D-027 — Documentation de dépôt à jour pour J05 (CHANGELOG et CONTRIBUTING)

**Question.** Faut-il une part dédiée qui écrive les changements de comportement de J05 dans `CHANGELOG.md` (`[Unreleased]`) et corrige l'instruction devenue fausse de `CONTRIBUTING.md` ?

**Contexte.** Ni `CHANGELOG.md` ni `CONTRIBUTING.md` ne sont dans les `files` d'une part de J05, et J05-P6 est une revue (elle ne change rien) — exactement la situation que D-015 a tranchée pour J03 en créant J03-P11. Sept changements de comportement visibles par un appelant, tous confirmés dans le diff par la revue J05-P6 :
- `insert_numbered_list_near_text` alloue une définition de liste neuve à chaque appel au lieu des `numId` 1/2 en dur : deux appels successifs ne produisent plus une numérotation continue.
- Le même outil n'applique plus le repli sur le style `Normal` quand `List Paragraph` est absent.
- `insert_line_or_paragraph_near_text` répond `with style 'Normal'.` au lieu du `repr` d'un objet Python, qui changeait à chaque exécution.
- `core.styles.ensure_heading_style` et `core.styles.create_style` écrivent réellement au lieu d'être des no-op ; `create_style` rend un `StyleInfo` au lieu d'un `Style` python-docx.
- `create_custom_style` écrit vraiment le style et refuse une couleur illisible au lieu de l'ignorer.
- `add_heading` et `create_document` peuvent désormais écrire `word/styles.xml` (jusqu'à 9 styles de titre créés) — relevé par la revue, absent de l'inventaire initial.
- `add_heading` peut désormais échouer sur un document sans style `Normal` — c'est l'objet de D-032, mais le CHANGELOG doit en porter la trace quelle que soit l'issue.

Second point, même famille : `CONTRIBUTING.md` demande « Update the tool count in README.md badges and text » alors que ce bloc est généré par `scripts/gen_tools_md.py` depuis J04-P3 et gardé par `tests/test_docs_sync.py` — suivre l'instruction casse `gen_tools_md.py --check`. Relevé par le worker de J05-P7.

**Options.**
1. *(Recommandée)* Une part T5 dédiée en fin de J05, `CHANGELOG.md` et `CONTRIBUTING.md` seuls, dispatchable en parallèle de J06. Motif : c'est la solution que D-015 a retenue pour J03 et qui a marché ; un changement de comportement non écrit est une rupture d'API silencieuse pour l'appelant, et `main` hérite de J05 dès la clôture de cette session (D-026).
2. Reporter à J06-P3, qui touche déjà la documentation.
3. Ne rien écrire : ce sont des corrections de défauts, pas des ruptures.

**Bloque.** Rien.

**Source.** Rapports de J05-P4, J05-P7 et de la revue J05-P6 (s8, 2026-09-15).

## Queue

- D-027 Documentation de dépôt à jour pour J05 : CHANGELOG `[Unreleased]` (7 changements de comportement) + instruction fausse de CONTRIBUTING (options: part T5 dédiée en fin de J05 — reporter à J06-P3 — ne rien écrire ; reco: part dédiée, comme D-015 pour J03 ; bloque: rien ; source: J05-P4, J05-P7, revue J05-P6, s8)
- D-032 Findings de la revue J05-P6 : (1) **should-fix**, `core/styles.py:106` — `ensure_heading_style` ne rattrape que `already_exists`, donc sur un document dont la feuille ne définit pas `Normal`, `_reference_id` lève `LocatorError("not_found")` et l'outil historique `add_heading` **échoue** (« Failed to add heading… ») là où l'ancien code (`except Exception: pass`) insérait le titre en formatage par défaut ; reproduit par exécution ; un paquet sans partie `styles` échoue de même par `PackageError` jamais rattrapée. (2) **optional**, `engine/effective.py:194` — la constante `_OFF` est recopiée de `engine/styles.py:246` au lieu d'être importée : aucune incohérence aujourd'hui, mais élargir le vocabulaire de bascule d'un seul côté ferait diverger les six propriétés extras des cinq du modèle (options: part corrective couvrant les deux — corriger le should-fix seul — laisser en l'état ; reco: part corrective couvrant les deux, « ensure » est au mieux-effort par contrat et la régression casse un outil historique sur un document réel ; bloque: rien ; source: revue J05-P6, s8)
