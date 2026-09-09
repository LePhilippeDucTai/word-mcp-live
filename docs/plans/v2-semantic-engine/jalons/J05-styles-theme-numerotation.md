# J05 — Styles, thème, format effectif, numérotation, styles de tableau

Goal: Le style, le thème et la numérotation deviennent des objets de premier rang du moteur : lecture des quatre familles avec chaîne d'héritage et références de thème, format effectif avec provenance, création et mise à jour de styles paragraphe et caractère, listes correctement définies dans `numbering.xml`, application de styles de tableau existants.
Depends on: J04 · Orchestrator: opus/high

## J05-P1 — Thème et modèle de style en lecture

```yaml
id: J05-P1
kind: implement
tier: T3
size: M
depends_on: []
files:
  - word_document_server/engine/theme.py
  - word_document_server/engine/styles.py
  - word_document_server/tools/v2/styles.py
  - tests/engine/test_theme.py
  - tests/engine/test_styles_read.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_theme.py tests/engine/test_styles_read.py -q"
```

### Scope
`theme.py` : lecture de `word/theme/theme1.xml` → polices majeure et mineure (latin, ea, cs), schéma de couleurs (dk1…accent6, hlink, folHlink, `sysClr lastClr`), `resolve_color(theme_color, tint, shade) -> hex`, `resolve_font(theme_font) -> name`.
`styles.py` : `list_styles(pkg, family=None) -> list[StyleInfo(style_id, name, family, builtin, based_on, next, link, ui_priority, q_format, hidden, semi_hidden)]`, `get_style(pkg, id_or_name)` avec `chain` (basedOn résolu), `run_props` et `paragraph_props` décodés par niveau, `resolved` (chaîne + docDefaults), références de thème conservées (`{"value": "Aptos", "theme": "minorHAnsi"}`), `find_style_usage(pkg, style_id)` (paragraphes, runs, tableaux avec locators).
`tools/v2/styles.py` : `doc_list_styles`, `doc_get_style`, `doc_find_style_usage`.
Tests sur `themes`, `style_inheritance`, `character_styles`, fixtures LibreOffice.

### Context
`engine/locators.py` pour les usages. Modèle minimal : identité, héritage, police, paragraphe, lien de numérotation, métadonnées ; pas de styles de liste en écriture.

## J05-P2 — Format effectif avec provenance

```yaml
id: J05-P2
kind: implement
tier: T3
size: M
depends_on: [J05-P1]
files:
  - word_document_server/engine/effective.py
  - word_document_server/tools/v2/effective.py
  - tests/engine/test_effective.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_effective.py -q"
```

### Scope
`effective_format(pkg, target) -> dict` pour un paragraphe ou une plage : chaque propriété → `{"value", "source"}`, `source` ∈ `direct`, `style:<id>` (caractère), `style:<id>` (chaîne de paragraphe), `table_style:<id>` (valeur `unresolved` en v1), `docDefaults`, `theme`.
Propriétés de run : polices résolues par script via le thème, couleurs via thème + tint/shade, taille, propriétés bascule (b, i, caps, smallCaps, strike, dstrike, outline, shadow, emboss, imprint, vanish) selon ECMA-376 §17.7.3 (XOR entre couches de style, formatage direct prioritaire). Propriétés de paragraphe : alignement, retraits, espacements, keep, niveau de plan, numérotation. Valeurs différentes sur une plage → `"mixed"`.
`tools/v2/effective.py` : `doc_get_effective_format(filename, locator, start, end)`.

### Context
`engine/styles.py`, `engine/theme.py`, `engine/textmodel.py`. Approximation documentée : styles de tableau non résolus.

## J05-P3 — Écriture de styles paragraphe et caractère

```yaml
id: J05-P3
kind: implement
tier: T3
size: M
depends_on: [J05-P1]
files:
  - word_document_server/engine/styles.py
  - word_document_server/core/styles.py
  - word_document_server/tools/format_tools.py
  - word_document_server/tools/v2/styles.py
  - tests/engine/test_styles_write.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_styles_write.py tests/characterization -q"
```

### Scope
`create_style(pkg, spec)` : `spec` = `{name, family: paragraph | character, based_on, next, link, q_format, run_props, paragraph_props}` ; `styleId` dérivé, collision → erreur ; enfants en ordre de schéma ; références de thème acceptées dans `run_props`. `update_style`, `clone_style`, `delete_style(reassign_to=None)` refusé si utilisé sans réaffectation.
`core/styles.py` : `ensure_heading_style` crée de vrais styles de titre (`basedOn Normal`, `next Normal`, `outlineLvl`, `qFormat`, `uiPriority`) ; `create_style` n'est plus un no-op ; `create_custom_style` garde sa signature et accepte `style_type` `paragraph` ou `character`.
`tools/v2/styles.py` : `doc_create_style`, `doc_update_style`, `doc_delete_style`. Retirer le xfail de `create_custom_style`.

### Context
`core/styles.py:53-134` (no-op prouvé : `get_by_id` ne lève jamais), `tools/format_tools.py:141-196`.

## J05-P4 — Numérotation

```yaml
id: J05-P4
kind: implement
tier: T3
size: M
depends_on: []
files:
  - word_document_server/engine/numbering.py
  - word_document_server/tools/v2/numbering.py
  - word_document_server/utils/document_utils.py
  - tests/engine/test_numbering.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_numbering.py tests/characterization -q"
```

### Scope
`numbering.py` : partie `numbering.xml` créée si absente ; `create_list_definition(pkg, kind: bullet | decimal | multilevel, levels) -> num_id` (nouveaux `abstractNum` et `num`, ids uniques, `nsid`) ; `apply_list(paragraph, num_id, level)` avec `w:numPr` à sa position de schéma ; `restart_list(pkg, num_id) -> num_id` (`lvlOverride`/`startOverride`) ; `list_definitions(pkg)` ; `paragraph_list_info(paragraph)`.
`utils/document_utils.py` : `add_bullet_numbering` et `insert_numbered_list_near_text` allouent une définition au lieu de `numId` 1/2 ; style `List Paragraph` appliqué seulement s'il existe.
`tools/v2/numbering.py` : `doc_apply_list(filename, locators, kind, level, dry_run)`. Retirer le xfail de `insert_numbered_list_near_text`.

### Context
`utils/document_utils.py:406-525`. Fixture `complex_numbering`. ECMA-376 §17.9.

## J05-P5 — Styles de tableau

```yaml
id: J05-P5
kind: implement
tier: T4
size: S
depends_on: [J05-P1]
files:
  - word_document_server/engine/table_styles.py
  - word_document_server/tools/v2/tables.py
  - tests/engine/test_table_styles.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_table_styles.py -q"
```

### Scope
`list_table_styles(pkg)` (famille `table`, genres de formats conditionnels présents), `apply_table_style(table, style_id, look)` avec `look` = `{first_row, last_row, first_column, last_column, no_h_band, no_v_band}` → `w:tblStyle` + `w:tblLook` (masque `w:val` et attributs explicites) ; style inexistant → erreur ; `clear_table_style(table)`.
`tools/v2/tables.py` : `doc_apply_table_style(filename, table_index, style, look, dry_run)`. Pas de création de style de tableau.

### Context
`engine/styles.py`. ECMA-376 §17.4.63 et §17.4.56.

## J05-P6 — Review J05

```yaml
id: J05-P6
kind: review
tier: T2
size: S
depends_on: [J05-P1, J05-P2, J05-P3, J05-P4, J05-P5]
files: []
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine tests/tools tests/characterization -q --timeout=60"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python scripts/gen_tools_md.py --check"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv sync && PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/ -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv build"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run ruff check ."
```

### Scope
Review the merged milestone diff with verify-before-done, code-review, test-design. Report; change nothing.
