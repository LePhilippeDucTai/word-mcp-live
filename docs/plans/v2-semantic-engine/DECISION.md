<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

No pending decision. Proceed with implementation.

## Queue

- D-005 — Résultat en cache d'un champ imbriqué (`PAGE` dans l'instruction d'un `IF`) : `w:t` ordinaire, donc lu comme visible ; `fields()` rend `"1first page"` là où Word n'afficherait que `"first page"`. Options : (a) lecture mécanique, `fields()` signale les deux spans, `ranges`/`find` refusent d'y couper — déjà implémentée et testée en J02-P2 ; (b) masquer le résultat des champs imbriqués sous une instruction, ce qui demande d'évaluer la structure des champs. Reco : (a) — le moteur rapporte ce que le paquet stocke ; (b) revient à simuler le moteur de champs de Word. Bloque : rien (J02-P3/P5 s'appuient sur le comportement actuel ; (b) les rouvrirait). Source : REPORT J02-P2.
