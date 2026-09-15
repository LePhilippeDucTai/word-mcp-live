<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

## D-027 — CHANGELOG des changements de comportement de J05

**Question.** Les changements de comportement visibles par un appelant introduits par J05 doivent-ils être écrits dans `CHANGELOG.md` (`[Unreleased]`) par une part dédiée, comme J03-P11 l'a fait pour J03 (D-015) ?

**Contexte.** `CHANGELOG.md` n'est dans les `files` d'aucune part de J05, et J05-P6 est une revue (elle ne change rien) — exactement la situation que D-015 a tranchée pour J03. Changements relevés à ce stade, tous rapportés par le worker de J05-P4 :
- `insert_numbered_list_near_text` alloue désormais une définition de liste neuve à chaque appel (nouveaux `abstractNum` + `num` dans `numbering.xml`) au lieu de rejoindre les `numId` 1 ou 2 codés en dur : deux appels successifs ne produisent plus une numérotation continue, et aucun `numId` existant du document n'est modifié.
- Le même outil n'applique plus le repli sur le style `Normal` quand `List Paragraph` est absent du document.
- `insert_line_or_paragraph_near_text` termine sa réponse par `with style 'Normal'.` au lieu du `repr` de l'objet `_ParagraphStyle` (D-021) — la chaîne était différente à chaque exécution, donc non parsable.

Les signatures et la forme des réponses restent inchangées : l'invariant de compatibilité du PLAN.md est tenu. J05-P1..P3, P5 et P7 peuvent en ajouter d'autres d'ici la clôture ; la question se pose une fois pour l'ensemble de J05.

**Options.**
1. *(Recommandée)* Part `CHANGELOG.md` dédiée en fin de J05 (T5, S, `CHANGELOG.md` seul), dispatchable en parallèle de la revue, couvrant tous les changements de J05 relevés au moment où elle part. Motif : c'est la solution que D-015 a retenue pour J03 et qui a marché ; un changement de comportement non écrit est une rupture d'API silencieuse pour l'appelant, et `main` hérite de J05 dès la clôture de cette session (D-026).
2. Reporter à J06, qui touche déjà la documentation (J06-P3).
3. Ne rien écrire : ces trois changements sont des corrections de défauts, pas des ruptures.

**Bloque.** Rien. La décision ajoute une part, elle n'en empêche aucune.

**Source.** Rapport de J05-P4 (s8, 2026-09-15).

## Queue

- D-027 CHANGELOG des changements de comportement de J05 (options: part `CHANGELOG.md` dédiée en fin de J05 — reporter à J06-P3 — ne rien écrire ; reco: part dédiée, comme D-015 pour J03 ; bloque: rien ; source: rapport J05-P4, s8)
