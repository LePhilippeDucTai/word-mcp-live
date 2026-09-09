<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

No pending decision.

## Queue

- CHANGELOG des changements de comportement de J03 : D-011 confiait la rédaction à « J03-P7 ou J03-P8 », mais `CHANGELOG.md` n'est dans les `files` d'aucune part et J03-P8 est une revue (elle ne change rien) ; trois changements attendent d'être écrits — `merge_documents` fusionne désormais dans le premier source (plus un `Document()` vierge) et refuse une source à commentaires ou notes (J03-P6), un `end_anchor_text` introuvable est signalé au lieu de supprimer jusqu'à la fin du document (J03-P5), un bloc portant plus de sauts de section qu'il ne reste de paragraphes porteurs est refusé au lieu d'en perdre un (J03-P5). Options : nouvelle part J03-P11 (T5, S) sur `CHANGELOG.md` avant la revue — reporter en J04 — renoncer au CHANGELOG. Reco : J03-P11 avant la revue, la revue J03-P8 est le dernier moment où le jalon est en mémoire et un changement de comportement non écrit est une rupture d'API silencieuse pour l'appelant. Bloque : rien. Source : orchestrateur s5, dispatch de J03-P7.
