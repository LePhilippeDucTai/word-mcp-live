<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

No pending decision.

## Queue

- D-036 — `rPrChange` et `pPrChange` ne sont jamais retirés par `accept_tracked_changes`/`reject_tracked_changes` : mesuré par le planner le 2026-09-15 sur les fixtures `tracked_changes` et `combined`. Le document `docs/audit/destructive-ops.md` le note et c'est toujours vrai ; aucune part ni décision du plan ne l'a couvert — J02-P6 refuse explicitement les genres de révision non gérés plutôt que de les ignorer, donc ce n'est pas une régression, mais un trou de périmètre. Conséquence pour un appelant : accepter toutes les révisions laisse le document porteur de marques de changement de formatage que Word affichera encore comme révisions en attente. Options : nouvelle part (étendre `accept`/`reject` aux deux genres, avec fixtures et tests d'invariant) — ligne LONG_TERM_RECO datée, le plan étant clos par ailleurs — ne rien faire, le refus typé de J02-P6 valant déjà avertissement. Reco : **ligne LONG_TERM_RECO**, le plan atteint son objectif sans cela et les deux genres demandent leur propre travail de fidélité (un `rPrChange` porte l'ancien `rPr` complet), mais l'écart mérite d'être écrit plutôt que découvert. Bloque : rien. Source : plan update s9 (planner) — **surfacée par la mise à jour de plan elle-même, donc à poser au démarrage de s10, pas à la clôture de s9**.
