<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

No pending decision.

## Queue

- D-021 — message non déterministe d'`insert_line_or_paragraph_near_text` (`utils/document_utils.py`) : appelé sans `line_style`, l'outil interpole l'objet style dans son message de retour (`style '_ParagraphStyle('Normal') id: 1407…'`), donc une chaîne qui change à chaque exécution ; relevé par le planner en s6, hors du trigger du plan update, laissé tel quel. Options : corriger dans J04-P7 (déjà propriétaire de `document_utils.py`) — part corrective séparée — laisser tel quel ; reco : corriger dans J04-P7, un `id:` d'objet Python dans un message d'outil n'est parsable par personne et rend le message intestable. Bloque : rien. Source : plan update s6 (planner)
