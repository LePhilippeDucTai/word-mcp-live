<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

## D-016 — Les 10 points d'appel restés sur `doc.paragraphs` avec un `paragraph_index` public

J03-P1 a fait passer `find_text_in_document` de l'index python-docx à l'index V2 (qui compte les
paragraphes d'un `w:sdt` de bloc). J03-P5 a suivi pour `delete_paragraph` et `add_bookmark` après le
bloquant de la revue. Les autres consommateurs d'un `paragraph_index` public sont restés sur
`doc.paragraphs`, mesuré par la revue round 2 sur un document à sommaire :
`find_text_in_document("Charlie") → 5` puis `add_footnote_to_document(5, "note")` pose la note sur
« Echo » et répond `Footnote added to paragraph 5` ; `get_paragraph_text(5) → "Echo"`.

Points d'appel : `utils/extended_document_utils.py:63` (`get_paragraph_text`) ;
`tools/footnote_tools.py:58,129,319,465,625` (5 outils de notes) ; `tools/layout_tools.py:351`
(`set_paragraph_spacing`) ; `utils/document_utils.py:422,471,565` (`target_paragraph_index` des 3
outils `insert_*_near_text`). Aucun n'est destructeur, mais tous rendent un résultat faux dès qu'un
`w:sdt` de bloc est présent — et `add_table_of_contents` en insère un.

PLAN.md prescrit que « la migration J03 aligne sur l'espace V2 sans renommer les paramètres » ; les
notes de bas de page sont explicitement hors périmètre du plan (« outils de notes de bas de page
conservés tels quels ») et R-003 leur unification est déjà en réserve pour après J03.

- Options : (a) part corrective en tête de J04 qui aligne les 10 points d'appel sur
  `indexed_paragraphs()`, sauf les 5 outils de notes qui partent en R-003 ; (b) part corrective
  couvrant les 10, y compris les notes, en élargissant le périmètre ; (c) acter la divergence,
  n'aligner rien et l'écrire au CHANGELOG (fusionne alors avec D-017).
- Reco : (a) — l'espace d'index est le contrat que J04-P1 va réutiliser, le laisser à deux valeurs
  sous le même nom de paramètre garantit qu'un appelant compose deux outils et corrompt son
  document ; les notes de bas de page sont hors scope du plan et déjà couvertes par R-003, les y
  laisser évite d'ouvrir dans J04 le refactor que R-003 décrit.
- Bloque : rien de ⬜ formellement, mais J04-P1 (locators et inspection) bâtit sur cet espace.
- Source : J03-P8 revue round 2 (should-fix).

## Queue

- D-017 — 5e puce `[Unreleased]` sur le changement d'espace d'index (`CHANGELOG.md:11`) : les 4 puces couvrent `merge_documents`, `add_header_footer` et 2 garde-fous, mais pas le changement le plus visible pour l'appelant — `paragraph_index` de `find_text_in_document`/`delete_paragraph`/`add_bookmark` compte désormais les paragraphes d'un `w:sdt` de bloc, ce qui casse les index qu'un appelant avait enregistrés ; D-015 a été créée sur l'argument qu'un changement de comportement non écrit est une rupture d'API silencieuse. Options : écrire la puce dans la part corrective de D-016 — part `CHANGELOG.md` dédiée — ne rien écrire ; reco : écrire la puce dans la part corrective de D-016, son contenu final dépend de ce que D-016 aligne. Bloque : rien. Source : J03-P8 round 2 (should-fix) · defer: D-016
- D-018 — `add_bookmark` sans borne inférieure (`tools/layout_tools.py:472`) : la garde ne teste que `paragraph_index >= len(paragraphs)`, donc `add_bookmark(path, -1, "Neg")` signe le dernier paragraphe et répond `{"success": true, "paragraph_index": -1}`, là où `delete_paragraph(-1)` refuse avec `Invalid paragraph index. Document has 3 paragraphs (0-2).` ; défaut antérieur au jalon, mais les deux outils partagent maintenant le même espace d'index. Options : ajouter `paragraph_index < 0 or` à la garde dans la part corrective de D-016 — laisser tel quel ; reco : ajouter la garde, une ligne, et deux outils du même espace qui divergent sur les bornes est un piège d'appelant. Bloque : rien. Source : J03-P8 round 2 (optional)
- D-019 — ponctuation doublée de `search_and_replace` (`tools/content_tools.py:647`) : quand toutes les occurrences tombent dans un champ, le message est `No occurrences of 'X' found., 1 skipped (inside fields)`, forme figée par `tests/tools/test_search_replace.py:194`. Touche un format de retour d'outil, donc l'invariant de compatibilité. Options : corriger le séparateur pour le cas « 0 remplacement » et mettre le test à jour — laisser tel quel (format figé) ; reco : corriger, un format de retour ne se fige que sur ce qu'un appelant peut parser et cette virgule n'est décrite nulle part. Bloque : rien. Source : J03-P8 round 1 (optional)
