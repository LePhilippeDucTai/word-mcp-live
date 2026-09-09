<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

## D-009 — Should-fix de la revue J02 : conteneur perdu par `_split_around`, index V2 faussé par les zones de texte

Question : la revue J02 (J02-P7, round 1, aucun finding bloquant) laisse deux défauts prouvés par sonde,
dans des zones qu'aucune des 19 fixtures ne couvre. Faut-il une part corrective avant d'attaquer J03 ?

1. `engine/revisions.py:603` — `_free_of_insertions` appelle `_split_around(container, child)` avec un `child`
   qui n'est pas forcément enfant direct du `w:ins` englobant (`_enclosing` remonte au `w:ins` le plus proche,
   quel que soit le nombre de conteneurs intermédiaires). Sur
   `<w:ins A><w:hyperlink><w:r>AAA</w:r><w:r>BBB</w:r></w:hyperlink></w:ins>`,
   `tracked_insert(…, offset=3, author="Bob")` sort « BBB » du `w:hyperlink` : le texte survit, le lien est perdu.
   Idem pour `w:sdt` et `w:smartTag`. Correctif proposé par la revue : remonter `run` jusqu'à l'enfant direct
   d'`enclosing` avant l'appel, plus un test sur cette forme.
2. `engine/find.py:141` — `_v2_index_map` ne filtre que `w:tc`, alors qu'`iter_paragraphs` descend dans
   `w:txbxContent`/`mc:AlternateContent` : un paragraphe de zone de texte consomme un index V2 et décale tous
   les suivants (`[p0, textbox, p1]` → `p1.index == 2` au lieu de `1`). C'est le contrat d'adressage
   `{"paragraph": i}` de J03 et J04 : dans tout document à zone de texte, l'édition tomberait sur le mauvais
   paragraphe. Correctif proposé : étendre le filtre d'ancêtres comme pour `w:tc`, avec une fixture le prouvant.
   (Un troisième finding, `optional` : `LIVE_CONTENT_TYPES` déclarée publique par J02-P8 mais non réexportée par
   `engine/__init__.py` ; à embarquer dans le correctif s'il a lieu.)

Options :
- **Part corrective J02-P10 (T3, S) avant J03** *(recommandée)* — même schéma que D-004 et D-007 : l'espace
  d'index V2 est le contrat que J03 et J04 vont consommer 120 fois ; un index faux ne se voit pas en test et se
  paie en édition du mauvais paragraphe. Les deux correctifs sont petits, localisés et déjà spécifiés par la revue.
- Élargir le scope de J03-P1 (index) et J03-P2 (conteneur) — évite une part, mais mêle un correctif de moteur à une
  migration d'outils et rend le diff de J03 illisible pour sa propre revue.
- Limites documentées en Risks, corrigées si J06 les rencontre — laisse J03 et J04 bâtir sur un adressage faux.

Bloque : J03-P1, J03-P2 (⏸). Source : J02-P7 (revue J02).

## Queue

(vide)
