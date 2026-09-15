# J05 — Styles, thème, format effectif, numérotation, styles de tableau

Goal: Le style, le thème et la numérotation deviennent des objets de premier rang du moteur : lecture des quatre familles avec chaîne d'héritage et références de thème, format effectif avec provenance, création et mise à jour de styles paragraphe et caractère, listes correctement définies dans `numbering.xml`, application de styles de tableau existants. Ouvre par la corrective D-024 : `doc_capabilities` enregistré, charges de `doc_apply_edits` validées ; D-021 (réponse déterministe de `insert_line_or_paragraph_near_text`) embarquée dans J05-P4 ; se clôt par la réécriture complète du README sur le MCP réellement implémenté (D-025, J05-P7), relue par la revue.
Depends on: J04 · Orchestrator: opus/high

## J05-P0 — Correctif D-024 : `doc_capabilities` enregistré, charges de `doc_apply_edits` validées

```yaml
id: J05-P0
kind: implement
tier: T4
size: S
depends_on: []
files:
  - word_document_server/tools/v2/capabilities.py
  - word_document_server/tools/platforms.py
  - word_document_server/tools/v2/batch.py
  - TOOLS.md
  - README.md
  - tests/tools/test_v2_capabilities.py
  - tests/tools/test_v2_batch.py
  - tests/test_registry_consistency.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_v2_capabilities.py tests/tools/test_v2_batch.py tests/test_registry_consistency.py tests/test_docs_sync.py -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"import asyncio; from word_document_server.main import mcp, register_tools; register_tools(); names={t.name for t in asyncio.run(mcp.list_tools())}; assert 'doc_capabilities' in names and len(names) >= 126, sorted(names)\""
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python scripts/gen_tools_md.py --check"
  - "grep -q 'def test_an_edit_with_an_unknown_key_is_refused_and_writes_nothing' tests/tools/test_v2_batch.py && grep -q 'def test_capabilities_is_discovered_and_registered' tests/tools/test_v2_capabilities.py && grep -q 'def test_every_v2_module_exports_tools' tests/test_registry_consistency.py"
```

### Scope
D-024 (revue J04-P6), deux findings prouvés par exécution le 2026-09-15 :
1. `tools/v2/capabilities.py` définit `doc_capabilities()` (J04-P3) sans exporter `TOOLS` ; `discover_tool_specs()` (`tools/v2/registry.py:269-300`, « a module without a `TOOLS` list is simply not a tool module ») l'ignore en silence : 125 outils enregistrés, `doc_capabilities` absent de `mcp.list_tools()`, de `_V2_TOOLS`, de `TOOLS.md` et du README, alors que PLAN.md et DESIGN-REVIEW.md le comptent livré.
2. `doc_apply_edits` (`tools/v2/batch.py:91-97,152-155`) lit chaque charge par `.get()` sans contrôler ses clés : `{"locator": {...}, "content": "…"}` (faute pour `text`) passe pour un `replace` de `text=""`, efface la plage visée, sauvegarde et répond `status: ok` ; `doc_edit_text` refuse la même faute en `invalid_argument` sans toucher le fichier.
Correctifs :
- `capabilities.py` : `TOOLS = [ToolSpec(fn=doc_capabilities, annotations=ToolAnnotations(title="Capabilities", readOnlyHint=True), tags=frozenset({"v2", "read"}))]` (forme de `doc_inspect`, `tools/v2/text.py:344`), `__all__` ; `platforms.py` : `"doc_capabilities"` dans `_V2_TOOLS` (docstring « The four `doc_*` V2 tools » → six) ; `TOOLS.md` et `README.md` régénérés par `scripts/gen_tools_md.py` (126 outils, 6 `doc_*`), jamais édités à la main.
- `batch.py` : listes blanches explicites `_TEXT_EDIT_KEYS = frozenset({"locator", "action", "text", "start", "end", "track_changes", "author"})` (paramètres de `doc_edit_text` moins `filename` et `dry_run`) et `_FORMAT_EDIT_KEYS = frozenset({"locator", "patch", "start", "end"})` (idem `doc_format_range`) ; avant toute lecture, une charge portant une clé hors de sa liste (liste choisie, comme aujourd'hui, par la présence de `patch`) ou sans `locator` lève `ValueError` (→ `invalid_argument` par `error_code`, index prépendu par `_named` comme les autres échecs), message nommant la clé fautive et les clés admises. Les défauts de `doc_edit_text` sur une clé admise absente (`action="replace"`, `text=""`…) sont conservés : `test_an_edit_without_an_action_defaults_to_replace` reste vert ; le contrôle précède `resolve`, rien n'est écrit.
Tests, rouges sur le code actuel : `tests/tools/test_v2_batch.py` — `test_an_edit_with_an_unknown_key_is_refused_and_writes_nothing`, paramétré sur la charge exacte de la revue `{"locator": {"paragraph": 1}, "content": "…"}`, une charge `patch` portant `action`, une charge sans `locator` : `status == "error"`, `code == "invalid_argument"`, `"edit 0"` et la clé fautive dans le message, `read_bytes()` identiques, `assert_unchanged_except(before, snapshot(path))`. `tests/tools/test_v2_capabilities.py` (nouveau, `call` redéfini comme dans `test_v2_batch.py`) — `test_capabilities_is_discovered_and_registered` (modèle `test_apply_edits_is_discovered_and_registered`) et `test_the_report_counts_what_platforms_lists` (`status == "ok"`, `tool_count == len(PLATFORMS)`, `tool_families["v2_semantic"]` = nombre de `doc_*` de `PLATFORMS`, `note` présent). `tests/test_registry_consistency.py` — `test_every_v2_module_exports_tools` : chaque module de `word_document_server.tools.v2` hors `registry` et `__init__` (`_modules()` de `registry.py` ou `pkgutil.iter_modules`) exporte un `TOOLS` non vide — garde contre la récidive dans chaque module que J05-P1..P5 et J06 ajoutent.

### Context
`tools/v2/registry.py:122-142` (`ToolSpec`), `:157-173` (`error_code` : `ValueError`/`TypeError` → `invalid_argument`), `:269-300` (`discover_tool_specs`) ; `tools/v2/text.py:186-196, 281-288` (signatures de référence des deux listes), `:344-362` (`TOOLS`) ; `tools/platforms.py:32, 179-185` ; `tests/test_registry_consistency.py:60-65` (`PLATFORMS` ⇔ registre, aveugle à un outil absent des deux) ; `tests/test_docs_sync.py`. Vérifié le 2026-09-15 : `register_tools()` → 125 noms, `'doc_capabilities' in names` → `False`.

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
  - tests/tools/test_index_space.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_numbering.py tests/tools/test_index_space.py tests/characterization -q"
  - "grep -q 'def test_the_insert_line_reply_names_the_style_not_the_object' tests/tools/test_index_space.py && ! grep -qF \"with style '{style}'\" word_document_server/utils/document_utils.py"
```

### Scope
`numbering.py` : partie `numbering.xml` créée si absente ; `create_list_definition(pkg, kind: bullet | decimal | multilevel, levels) -> num_id` (nouveaux `abstractNum` et `num`, ids uniques, `nsid`) ; `apply_list(paragraph, num_id, level)` avec `w:numPr` à sa position de schéma ; `restart_list(pkg, num_id) -> num_id` (`lvlOverride`/`startOverride`) ; `list_definitions(pkg)` ; `paragraph_list_info(paragraph)`.
`utils/document_utils.py` : `add_bullet_numbering` et `insert_numbered_list_near_text` allouent une définition au lieu de `numId` 1/2 ; style `List Paragraph` appliqué seulement s'il existe.
`tools/v2/numbering.py` : `doc_apply_list(filename, locators, kind, level, dry_run)`. Retirer le xfail de `insert_numbered_list_near_text`.
D-021 : `insert_line_or_paragraph_near_text` (`:481-541`, même fichier, même région) interpole l'objet style dans sa réponse quand `line_style` est omis (`with style '_ParagraphStyle('Normal') id: 1407…'`, chaîne différente à chaque exécution). `style_name = line_style or para.style.name` (`:529`), passé à `add_paragraph` et interpolé : `… with style 'Normal'.` dans les deux cas, `line_style` donné ou style de la cible repris ; reste du message, clés et comportement inchangés (un nom de style inconnu échoue comme aujourd'hui). Test `test_the_insert_line_reply_names_the_style_not_the_object` dans `tests/tools/test_index_space.py` (mêmes documents et helpers, rien d'autre dans ce fichier) : sans `line_style`, la réponse se termine par `with style 'Normal'.` et deux appels sur deux copies rendent la même chaîne ; rouge sur le code actuel.

### Context
`utils/document_utils.py:481-541` (`insert_line_or_paragraph_near_text`, style et message `:529-539`), `:544-580` (`add_bullet_numbering`, `numId` en dur `:573`), `:583-673` (`insert_numbered_list_near_text`, `num_id = 1 if … else 2` `:638`, candidats `List Paragraph` `:642`). Fixture `complex_numbering`. ECMA-376 §17.9. Modèle de test : `tests/tools/test_index_space.py` (J04-P7, `_document_with_a_table_of_contents`, assertion par sous-chaîne posée là faute de ce correctif).

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

## J05-P7 — Réécriture complète du README sur le MCP réel

```yaml
id: J05-P7
kind: implement
tier: T3
size: M
depends_on: [J05-P0]
files:
  - README.md
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python scripts/gen_tools_md.py --check"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/test_docs_sync.py -q"
  - "grep -qF '<!-- tool-counts:start -->' README.md && grep -qF '<!-- tool-counts:end -->' README.md && [ \"$(grep -v '^[[:space:]]*$' README.md | tail -n 1)\" = '<!-- mcp-name: io.github.ykarapazar/word-mcp-live -->' ]"
  - "grep -q 'doc_inspect' README.md && grep -q 'doc_find' README.md && grep -q 'doc_edit_text' README.md && grep -q 'doc_format_range' README.md && grep -q 'doc_apply_edits' README.md && grep -q 'doc_capabilities' README.md && grep -q 'word_document_server/engine' README.md && ! grep -q '124 tools' README.md"
```

### Scope
D-025 (user, s8). `README.md` (396 lignes) est pour l'essentiel le texte hérité de l'amont (`io.github.ykarapazar/word-mcp-live`) : pitch « live editing », `124 tools` en dur, installation `uvx word-mcp-live` et deeplinks Cursor/VS Code (paquet PyPI de l'amont, sans le moteur), image `ghcr.io`, vidéo, Star History ; rien sur `engine/`, la surface V2, le harnais ni la non-dégradation, sujet réel du dépôt depuis J01. Réécriture complète, en anglais, contre le code présent dans l'arbre au moment de la rédaction : aucune affirmation qui n'y soit vérifiable (un outil cité existe dans `mcp.list_tools()`, une garantie citée a son test) ; pas de badge de CI — le seul workflow, `.github/workflows/publish.yml`, publie sur PyPI à la release et ne teste rien, et le plan interdit d'en ajouter ; le paquet PyPI de l'amont n'est pas présenté comme ce serveur.
Contenu : (1) ce qu'est le serveur — MCP de manipulation de documents Word pour agents, cœur OOXML pur (Linux, sans Word) ; (2) architecture en couches réelle de `word_document_server/engine/` d'après les modules présents (`package` → `textmodel` → `ranges` → `format` → `find`/`revisions`/`locators`/`inspect`/`merge` ; pas la liste cible du PLAN.md), politique de texte (visible, caché, opaque, marqueurs), erreurs typées ; (3) les trois familles telles que `tools/platforms.py` et `TOOLS.md` les classent : outils cross-platform historiques migrés sur le moteur (noms, paramètres et formats de retour inchangés ; fichier fermé), surface V2 `doc_*` sans état (locators `paragraph`/`find`/`bookmark`/`heading`/`table` + `story`/`expect_text` de `engine/locators.py`, `dry_run`, rapport `{status, dry_run, changes, warnings}` ou `{status: "error", code, message}`, lot atomique `doc_apply_edits`, `doc_capabilities`), outils live COM/JXA dépendants de la plateforme (Windows + macOS / Windows seul, comptes du bloc généré) ; un exemple d'appel `doc_*` réel avec sa réponse, rejoué avant d'être écrit ; (4) garanties de fidélité (texte et formatage hors plage inchangés, signets et plages de commentaire conservés, champs atomiques, `w:del` jamais modifié, parties non ciblées identiques, ids d'annotation uniques, ordre de schéma, sauvegarde atomique) et écarts assumés (champ imbriqué D-005, outils de notes sur `doc.paragraphs` D-016, conteneur vidé D-009, changements de comportement listés dans `CHANGELOG.md` `[Unreleased]`) ; (5) harnais : fixtures générées (`tests/fixtures/builders.py`, ODT LibreOffice), instantanés canoniques (`tests/support/snapshot.py`), validateur de paquet (`tests/support/package_check.py`), caractérisation (`tests/characterization`), LibreOffice optionnel ; (6) installation depuis le dépôt (`uv sync` ou `pip install -e .`, entrée `word_mcp_server` de `pyproject.toml [project.scripts]`) et configuration MCP (commande du type `uv run --directory <clone> word_mcp_server`, vérifiée ; `MCP_AUTHOR`/`MCP_AUTHOR_INITIALS` lues dans `defaults.py:5-6`, `MCP_TRANSPORT`/`MCP_HOST`/`MCP_PORT` dans `main.py:59-70`) ; (7) renvois vers `TOOLS.md` (liste exhaustive), `CHANGELOG.md`, `CONTRIBUTING.md`, `RENDER_DEPLOYMENT.md`, `PRIVACY.md`, `LICENSE`.
Partie générée : le bloc `<!-- tool-counts:start -->` … `<!-- tool-counts:end -->` (`scripts/gen_tools_md.py:38-39,161-195`) n'est **jamais** édité à la main — conservé tel quel ou régénéré par `uv run python scripts/gen_tools_md.py` (qui réécrit aussi `TOOLS.md`, hors `files` : si `--check` est rouge sur la base avant toute édition, rapport `blocked`, la dérive se corrige sur la base, pas ici) ; le trailer `<!-- mcp-name: io.github.ykarapazar/word-mcp-live -->` reste la dernière ligne (registre MCP, `server.json`). Seul `README.md` change : pas de renommage du paquet, `pyproject.toml`, `TOOLS.md`, `CHANGELOG.md`, `CONTRIBUTING.md`, `RENDER_DEPLOYMENT.md` intacts. Les outils de J05-P1..P5 ne sont décrits que s'ils sont dans l'arbre ; J06-P3 complète le README avec ceux de J05 et J06.

### Context
`README.md:1-58` (en-tête, badges et pitch hérités), `:60-212` (installation `uvx`, deeplinks), `:216-245` (« Two Modes »), `:247-257` (variables), `:337-348` (bloc généré), `:396` (trailer) ; `scripts/gen_tools_md.py:38-39` (marqueurs), `:161-185` (`render_readme_counts`), `:187-195` (`apply_readme_counts`, `ValueError` « README.md is missing … markers ») ; `tests/test_docs_sync.py` (`test_check_flag_passes_on_the_committed_docs`, `test_readme_keeps_its_tool_count_markers`) ; `tools/platforms.py:1-33` (trois familles) ; `TOOLS.md:87-98` (section `doc_*`) ; `tools/v2/registry.py` (enveloppe de rapport) ; `engine/locators.py:13-28,132-143` ; PLAN.md `## Architecture` et `## Risks` (garanties, écarts D-005/D-009/D-016). Vérifié le 2026-09-15 : 396 lignes, marqueurs lignes 339/348, trailer ligne 396, `gen_tools_md.py --check` → 0 sur la base (125 outils ; 126 après J05-P0) ; la 4e acceptation est rouge sur le README actuel.

## J05-P6 — Review J05

```yaml
id: J05-P6
kind: review
tier: T2
size: S
depends_on: [J05-P0, J05-P1, J05-P2, J05-P3, J05-P4, J05-P5, J05-P7]
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
Relit aussi J05-P0 (corrective D-024) : `doc_capabilities` dans `mcp.list_tools()`, `_V2_TOOLS`, `TOOLS.md` et README régénérés ; listes blanches de `doc_apply_edits` alignées sur les signatures de `doc_edit_text` et `doc_format_range`, la charge de la revue refusée sans écriture ; et D-021 dans J05-P4 : réponse de `insert_line_or_paragraph_near_text` sans objet Python interpolé.
Relit aussi le README réécrit par J05-P7 (D-025) contre le code réel : aucun outil décrit qui n'est pas dans `mcp.list_tools()`, aucune garantie, option (`dry_run`, locators) ou commande d'installation qui ne soit vérifiable dans l'arbre, aucune promesse non tenue ni texte hérité de l'amont qui décrit un autre serveur ; bloc de compteurs et trailer intacts (`gen_tools_md.py --check`, `tests/test_docs_sync.py`).
