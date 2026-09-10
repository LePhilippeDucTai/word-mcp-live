<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

## D-020 — Point d'intégration sur `main` maintenant que la revue J03-P8 est verte

D-014 (user, 2026-09-10) posait : « `main` reste intact jusqu'à la revue J03-P8 — la session suivante
lance la revue puis merge sur `main` ». La revue a tourné en s6 : round 1 en échec sur un bloquant
(`delete_paragraph` hors espace V2, corrigé par J03-P5 review fix 1), round 2 **sans finding
bloquant**, 24 acceptations et 5 Checks verts, 1091 passed / 1 xfailed, 120 outils MCP, signatures et
formats de retour inchangés. La condition posée par D-014 est donc remplie.

Ce que D-014 ne pouvait pas prévoir : la même revue a établi que 5 points d'appel de la surface
publique rendent encore un résultat faux dès qu'un `w:sdt` de bloc est présent — `get_paragraph_text`,
`set_paragraph_spacing` et les 3 `insert_*_near_text` composent l'index python-docx avec l'index V2
que `find_text_in_document` rend désormais. J04-P7 les corrige, en tête de J04, et embarque la 5e
puce du CHANGELOG, la borne d'`add_bookmark` et le séparateur de `search_and_replace`.

`main` a 66 commits d'avance sur la branche et 0 de retard (mesure de s5, à revérifier).

- Options : (a) merger après J04-P7, quand la surface publique est cohérente sur un seul espace
  d'index ; (b) merger maintenant, D-014 étant remplie à la lettre, et laisser J04-P7 arriver
  ensuite ; (c) ouvrir une pull request depuis `v2-semantic-engine` au lieu de merger en local.
- Reco : (a) — le motif écrit de D-014 est que « J03 réécrit la surface des outils existants (120
  noms, signatures et formats de retour censés être inchangés) » ; or c'est précisément cette
  surface que J04-P7 finit de rendre cohérente, et un merge intermédiaire livrerait sur `main` deux
  sens de `paragraph_index` sous le même nom de paramètre.
- Bloque : rien de ⬜ (J04-P7 et la suite de J04 se font sur la branche).
- Source : D-014 arrivée à échéance, revue J03-P8 round 2.

## Queue

- D-021 — message non déterministe d'`insert_line_or_paragraph_near_text` (`utils/document_utils.py`) : appelé sans `line_style`, l'outil interpole l'objet style dans son message de retour (`style '_ParagraphStyle('Normal') id: 1407…'`), donc une chaîne qui change à chaque exécution ; relevé par le planner en s6, hors du trigger du plan update, laissé tel quel. Options : corriger dans J04-P7 (déjà propriétaire de `document_utils.py`) — part corrective séparée — laisser tel quel ; reco : corriger dans J04-P7, un `id:` d'objet Python dans un message d'outil n'est parsable par personne et rend le message intestable. Bloque : rien. Source : plan update s6 (planner)
