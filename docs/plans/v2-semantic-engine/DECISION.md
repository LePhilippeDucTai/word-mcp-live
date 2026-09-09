<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

## D-004 — Zones aveugles restantes du harnais J01

La revue J01 round 2 rend `status: done` (aucun finding bloquant, toutes les acceptations
et tous les Checks verts). Elle laisse cinq findings non bloquants, tous mesurés, tous
portant sur ce que l'instrument **ne voit pas** :

- `snapshot.py:1089` (should-fix) — les relations sont indexées sur la partie décrite, que
  toute allowance de paragraphe met déjà dans `allowed_parts` : supprimer la relation image
  `rId14 → media/image1.png` passe `assert_unchanged_except` **et** `validate_package() == []`.
  Un `manage_hyperlinks` réécrit en J03 qui reconstruit les `.rels` casserait toutes les
  images sans faire rougir un test. Correctif proposé : code `REL-REF-UNRESOLVED` dans
  `validate_package` (tout `r:id`/`r:embed` d'une story résout dans ses `.rels`).
- `snapshot.py:862` (should-fix) — le `w:sectPr` de corps n'entre dans aucune signature :
  retirer les `headerReference`/`footerReference` (le document perd en-tête et pied) passe.
  `add_header_footer` est une zone sensible des Risks. Correctif : `sect_pr` par story.
- `test_package_check.py:192` (should-fix) — seule la variante *signet* de `ID-DUPLICATE`
  est testée ; neutraliser la boucle des ids de révision ou de commentaire laisse la suite
  verte, alors que « ids d'annotation uniques » est un invariant des Risks et que
  `_generate_id` est ce que J03 réécrit. Correctif : deux tests de corruption ciblée.
- `snapshot.py:718` (optional) — `TableSignature` ne retient du `w:tblGrid` que le nombre de
  `w:gridCol` : écraser toutes les largeurs passe.
- `test_package_check.py:305` (optional) — `XML-MALFORMED` et `CT-PART-MISSING` sans test
  positif.

Question : traiter ces zones aveugles maintenant, avant d'écrire le cœur OOXML de J02 ?

- Option A (reco) — **Nouvelle part J01-P7 (T3) traitant les 3 should-fix**, dispatchée en
  début de session suivante avant J02-P1. Le harnais est l'instrument dont dépendent J02 à
  J06 ; une perte de relation ou d'en-tête non vue à J03 se paie beaucoup plus cher que la
  part. Les deux `optional` restent en note.
- Option B — **Tout traiter (3 should-fix + 2 optional) dans J01-P7.** Couverture maximale,
  part un peu plus longue.
- Option C — **Reporter en J06** (`J06-P2` reprend déjà `snapshot.py` dans `engine/compare.py`).
  J02 à J05 se construisent alors sur un instrument dont on connaît les trous.

Blocks: aucune part (J02-P1 est prête) · Source: revue J01-P6 round 2

## File d'attente

(vide)
