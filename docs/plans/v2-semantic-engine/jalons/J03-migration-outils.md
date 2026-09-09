# J03 — Migration des outils existants sur le cœur

Goal: Les outils cross-platform existants délèguent au moteur sans changer nom, paramètres ni format de retour ; les opérations destructrices prouvées deviennent des modifications locales ; le monkey-patch de sauvegarde disparaît ; toute écriture est atomique ; les xfail de caractérisation correspondants deviennent des tests verts.
Depends on: J02 · Orchestrator: opus/high

## J03-P1 — search_and_replace et find_text_in_document

```yaml
id: J03-P1
kind: implement
tier: T3
size: M
depends_on: []
files:
  - word_document_server/utils/document_utils.py
  - word_document_server/utils/extended_document_utils.py
  - word_document_server/tools/content_tools.py
  - tests/tools/test_search_replace.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_search_replace.py tests/characterization -q"
```

### Scope
`find_and_replace_text` → `engine.find` (toutes les stories de `DocxPackage.stories()` passées dans `stories`) + `replace_range` : correspondances à cheval sur des runs, hyperliens, `w:ins`, sdt, tableaux, en-têtes, pieds, notes ; `count` = occurrences ; correspondances chevauchant un champ → ignorées et signalées en fin de chaîne (`, N skipped (inside fields)`) ; règle de saut des paragraphes de style `TOC` conservée ; plus aucun remplacement par run.
`find_text_in_document` → `engine.find` : résultat avec story (`Match.story`, id réel rendu par `DocxPackage.stories()` — `document`, `header1`, `footnotes`… ; `"body"` n'est qu'un alias d'entrée du filtre `stories` de `find`, jamais rendu), index V2 du paragraphe (`Match.index`, `None` pour un paragraphe en cellule de tableau ou en zone de texte, hors de l'espace V2), offsets, contexte ; format JSON existant conservé, clés ajoutées seulement.
`tools/content_tools.py` : `search_and_replace` conserve son message ; sauvegarde via `DocxPackage.save`.
Retirer les xfail correspondants dans `tests/characterization`.

### Context
`utils/document_utils.py:246-285`, `utils/extended_document_utils.py:44-90`, `tools/content_tools.py:440-472`. Engine : `find.py`, `ranges.py`, `package.py`.

## J03-P2 — Tracked changes sur le moteur

```yaml
id: J03-P2
kind: implement
tier: T2
size: M
depends_on: []
files:
  - word_document_server/core/tracked_changes.py
  - tests/tools/test_tracked_changes.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_tracked_changes.py -q --timeout=60"
```

### Scope
Réécrire `track_replace_in_doc`, `track_insert_in_doc`, `track_delete_in_doc`, `list_tracked_changes_in_doc`, `accept_changes_in_doc`, `reject_changes_in_doc` sur `DocxPackage` + `engine.find` + `engine.revisions` ; clés des dicts de retour inchangées ; sémantique actuelle conservée (insert : première occurrence ; replace et delete : toutes).
Supprimer l'accès `zipfile`, les helpers privés dupliqués et l'écriture non atomique.
Tests : le cas `new_text ⊇ old_text` termine ; `rPr` de chaque morceau conservé ; texte sous `w:del` jamais apparié ; accept/reject par id et par auteur ; instantané des parties non ciblées inchangé.

### Context
`core/tracked_changes.py`, `tools/tracked_changes_tools.py` (ne pas modifier). Engine : `revisions.py`, `find.py`.

## J03-P3 — Commentaires et hyperliens sur le moteur

```yaml
id: J03-P3
kind: implement
tier: T3
size: M
depends_on: []
files:
  - word_document_server/core/comment_writer.py
  - word_document_server/core/comments.py
  - word_document_server/core/hyperlink_writer.py
  - word_document_server/tools/hyperlink_tools.py
  - word_document_server/tools/comment_tools.py
  - tests/tools/test_comments.py
  - tests/tools/test_hyperlinks.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_comments.py tests/tools/test_hyperlinks.py -q"
```

### Scope
`add_comment_to_doc` → `engine.find` + `format.wrap(comment)` + `ensure_part` pour `comments.xml`, `commentsExtended.xml`, `commentsIds.xml`, `people.xml` (fil Word fonctionnel, `paraId` unique, ids via `ids.py`) ; sauvegarde atomique.
Lecture (`core/comments.py`) : ancres retrouvées (story = id réel de `DocxPackage.stories()`, index V2 ou `None` en cellule de tableau ou en zone de texte, texte ancré) ; suppression du repli par `str(element)` ; erreurs remontées, plus d'`except Exception` silencieux.
Hyperliens : action `add` via `wrap(hyperlink)` + relation externe dédoublonnée ; nouvelle action `remove` (`unwrap`, relation retirée si plus référencée) ; action `list` ; formats de retour JSON existants conservés.
Supprimer les trois copies du matcher et des allocateurs de `rId`.

### Context
`core/comment_writer.py`, `core/comments.py`, `core/hyperlink_writer.py`, `tools/hyperlink_tools.py`, `tools/comment_tools.py`. Engine : `format.py`, `package.py`, `ids.py`.

## J03-P4 — format_text et format_cell_text sans reconstruction

```yaml
id: J03-P4
kind: implement
tier: T3
size: S
depends_on: []
files:
  - word_document_server/tools/format_tools.py
  - word_document_server/core/tables.py
  - tests/tools/test_format_text.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_format_text.py -q"
```

### Scope
`format_text(paragraph_index, start_pos, end_pos, ...)` → `ranges.resolve` + `format.apply_rpr` ; offsets sur le texte visible du moteur ; index de paragraphe = espace V2 (un paragraphe en cellule de tableau ou en zone de texte n'y a pas d'index ; `format_cell_text` adresse la cellule) ; couleur invalide → message d'erreur, jamais noir ; plus aucun `run.clear()` ni `add_run`.
`core/tables.py::format_cell_text` → `apply_rpr` sur les runs de la cellule ; `text_content` remplace le texte du premier paragraphe via `replace_range`, autres paragraphes et tableaux imbriqués conservés.
Retirer le préfixe `DESTRUCTIVE` de la docstring de `format_text`.

### Context
`tools/format_tools.py:27-138`, `core/tables.py:656-739`. Preuve : duplication du texte d'hyperlien avec le code actuel.

## J03-P5 — Opérations destructrices : TOC, en-têtes, suppression, signets, blocs

```yaml
id: J03-P5
kind: implement
tier: T3
size: M
depends_on: [J03-P1]
files:
  - word_document_server/tools/content_tools.py
  - word_document_server/tools/layout_tools.py
  - word_document_server/utils/document_utils.py
  - tests/tools/test_toc.py
  - tests/tools/test_layout_blocks.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_toc.py tests/tools/test_layout_blocks.py tests/characterization -q"
```

### Scope
`add_table_of_contents` : insérer un vrai sommaire (`w:sdt` galerie « Table of Contents » contenant le champ complexe `TOC \o "1-3" \h \z \u` avec résultat de substitution) en tête du corps ou après le premier titre, `w:updateFields` posé dans `settings.xml` ; jamais de reconstruction ; message de retour conservé.
`add_header_footer` : remplacer le texte visible du premier paragraphe via `replace_range` ; paragraphe contenant champs ou images → erreur explicite sauf nouveau paramètre optionnel `replace_content=False`.
`delete_paragraph` : `w:sectPr` final déplacé sur le paragraphe précédent. `add_bookmark` : id via `ids.py`, marqueurs placés après `w:pPr`, nom échappé.
`replace_paragraph_block_below_header` et `replace_block_between_manual_anchors` : opérations sur les enfants du corps (tableaux compris), `sectPr` jamais supprimé, sauvegarde unique, ancre de fin déterminée par style de titre ou `end_anchor_text` explicite.
Retirer les xfail correspondants.

### Context
`tools/content_tools.py:313-400`, `tools/layout_tools.py:112-172, 405-441`, `utils/document_utils.py:549-726`. Engine : `ranges.py`, `format.py`, `ids.py`, `package.py`.

## J03-P6 — merge_documents par import d'éléments

```yaml
id: J03-P6
kind: implement
tier: T2
size: M
depends_on: [J03-P4]
files:
  - word_document_server/engine/merge.py
  - word_document_server/tools/document_tools.py
  - word_document_server/core/tables.py
  - tests/tools/test_merge.py
  - tests/engine/test_merge.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_merge.py tests/tools/test_merge.py -q"
```

### Scope
`engine/merge.py::append_document(target, source, page_break) -> Report` : copie profonde des blocs du corps source (paragraphes, tableaux, sdt) avant le `sectPr` final ; réécriture des `r:id`/`r:embed` (parties média copiées, relations externes recréées) ; styles manquants copiés avec leur chaîne `basedOn` et définitions de numérotation référencées renumérotées ; source contenant commentaires ou notes → refus explicite ; en-têtes et pieds de la source ignorés avec avertissement.
`merge_documents` utilise `append_document` ; message de retour conservé, avertissements ajoutés. `core/tables.py::copy_table` supprimé s'il n'a plus d'appelant.
Retirer les xfail correspondants.

### Context
`tools/document_tools.py:141-214`, `core/tables.py:110-142`. Engine : `package.py` (`ensure_part`, relations).

## J03-P7 — Retrait du hook de sauvegarde, écritures atomiques, annotations

```yaml
id: J03-P7
kind: implement
tier: T4
size: M
depends_on: [J03-P1, J03-P2, J03-P3, J03-P4, J03-P5, J03-P6]
files:
  - word_document_server/utils/save_utils.py
  - word_document_server/main.py
  - word_document_server/tools/protection_tools.py
  - tests/tools/test_atomic_writes.py
acceptance:
  - "test ! -f word_document_server/utils/save_utils.py"
  - "! grep -q install_save_hook word_document_server/main.py"
  - "! grep -q 'DESTRUCTIVE' word_document_server/main.py"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_atomic_writes.py tests/characterization -q"
```

### Scope
Supprimer `utils/save_utils.py` et son installation dans `run_server` ; conserver `utils/path_utils.py`.
`protection_tools.py` : chiffrement et déchiffrement écrits via `engine.package.atomic_write_bytes` ; original intact en cas d'échec ; plus de `open(filename, "wb")`.
`main.py` : retirer les préfixes `DESTRUCTIVE` de `add_table_of_contents` et `merge_documents`, remettre `destructiveHint=False` sur `format_text`.
Tests : écriture interrompue (exception injectée) laisse l'original intact et aucun temporaire ; `test_no_op_roundtrip` reste vert sans le hook ; plus aucun xfail dans `tests/characterization` pour les outils migrés en J03.

### Context
`utils/save_utils.py`, `main.py:1937-1998`, `tools/protection_tools.py:26-80, 240-280`.

## J03-P8 — Review J03

```yaml
id: J03-P8
kind: review
tier: T2
size: S
depends_on: [J03-P1, J03-P2, J03-P3, J03-P4, J03-P5, J03-P6, J03-P7]
files: []
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools tests/characterization tests/engine -q --timeout=60"
  - "test ! -f word_document_server/utils/save_utils.py"
  - "! grep -q install_save_hook word_document_server/main.py"
  - "! grep -q 'DESTRUCTIVE' word_document_server/main.py"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv sync && PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/ -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv build"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run ruff check ."
```

### Scope
Review the merged milestone diff with verify-before-done, code-review, test-design. Report; change nothing.
Relit aussi J02-P10 (corrective D-009, hors du périmètre de J02-P7 close) : son diff et ses acceptations rejouées.
