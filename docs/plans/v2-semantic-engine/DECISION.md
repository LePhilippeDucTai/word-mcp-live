<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

No pending decision.

## Queue

- D-021 — message non déterministe d'`insert_line_or_paragraph_near_text` (`utils/document_utils.py`) : appelé sans `line_style`, l'outil interpole l'objet style dans son message de retour (`style '_ParagraphStyle('Normal') id: 1407…'`), donc une chaîne qui change à chaque exécution ; relevé par le planner en s6, hors du trigger du plan update, laissé tel quel. Options : corriger dans J04-P7 (déjà propriétaire de `document_utils.py`) — part corrective séparée — laisser tel quel ; reco : corriger dans J04-P7, un `id:` d'objet Python dans un message d'outil n'est parsable par personne et rend le message intestable. Bloque : rien. Source : plan update s6 (planner)
- D-024 — 2 should-fix de la revue J04-P6 : (1) `capabilities.py` n'exporte pas de `TOOLS`, `doc_capabilities` n'est jamais enregistré (125 outils, aucun `doc_capabilities` mesuré) alors que PLAN.md/DESIGN-REVIEW.md le listent comme livrable J04 ; (2) `doc_apply_edits` (`tools/v2/batch.py:91-97`) lit chaque charge par `.get()` sans valider le jeu de clés — `{"locator": {...}, "content": "…"}` (typo pour `text`) devient un `replace` avec `text=""`, efface le span visé et sauvegarde en répondant `status: ok` (prouvé par exécution ; `doc_edit_text` refuse le même typo sans toucher au fichier). Options : part corrective dédiée avant J05 — élargir le scope d'une part J05 déjà en `files` proches — laisser en l'état ; reco : part corrective dédiée avant J05, (2) est une perte de données silencieuse sur simple faute de frappe d'un agent appelant, plus grave que tout ce que J04 a corrigé jusqu'ici. Bloque : rien de ⬜ (J05 n'a pas commencé). Source : revue J04-P6 (reviewer)
