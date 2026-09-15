# J05 — Styles, thème, format effectif, numérotation, styles de tableau

Goal: Le style, le thème et la numérotation deviennent des objets de premier rang du moteur : lecture des quatre familles avec chaîne d'héritage et références de thème, format effectif avec provenance, création et mise à jour de styles paragraphe et caractère, listes correctement définies dans `numbering.xml`, application de styles de tableau existants. Ouvre par la corrective D-024 : `doc_capabilities` enregistré, charges de `doc_apply_edits` validées ; D-021 (réponse déterministe de `insert_line_or_paragraph_near_text`) embarquée dans J05-P4 ; se clôt par la réécriture complète du README sur le MCP réellement implémenté (D-025, J05-P7), relue par la revue. Ajoutées après la clôture de J05-P6 et relues par J06-P4 : J05-P9 (D-032, corrective de J05-P3 et J05-P2) et J05-P8 (D-027, `CHANGELOG.md` et `CONTRIBUTING.md`).
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

## J05-P9 — Correctif D-032 : `ensure_heading_style` au mieux-effort, bascules d'`effective` par le décodeur de `styles`

```yaml
id: J05-P9
kind: implement
tier: T4
size: S
depends_on: []
files:
  - word_document_server/core/styles.py
  - word_document_server/engine/effective.py
  - tests/engine/test_styles_write.py
  - tests/engine/test_effective.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_styles_write.py tests/engine/test_styles_read.py tests/engine/test_effective.py tests/characterization -q --timeout=60"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"import asyncio, pathlib, tempfile; from tests.fixtures.builders import build; from word_document_server.engine.package import DocxPackage; from word_document_server.engine.styles import get_style, styles_root; from word_document_server.engine.xmlns import qn; from word_document_server.tools.content_tools import add_heading; pkg = DocxPackage.open(build('simple')); root = styles_root(pkg); [root.remove(s) for s in list(root.findall(qn('w:style'))) if s.get(qn('w:styleId')) in {'Normal', 'Heading7', 'Heading8', 'Heading9'}]; path = pathlib.Path(tempfile.mkdtemp()) / 'no-normal.docx'; pkg.save(path); one = asyncio.run(add_heading(str(path), 'Probe one', 1)); seven = asyncio.run(add_heading(str(path), 'Probe seven', 7)); assert not one.startswith('Failed') and not seven.startswith('Failed'), (one, seven); info = get_style(DocxPackage.open(path), 'Heading7').info; assert info.based_on is None and info.next_style is None, info\""
  - "grep -q '^def test_ensure_heading_style_tolerates_a_style_sheet_without_normal' tests/engine/test_styles_write.py && grep -q '^def test_add_heading_still_inserts_a_heading_when_normal_is_missing' tests/engine/test_styles_write.py && grep -q '^def test_the_extra_toggles_read_off_the_way_the_model_toggles_do' tests/engine/test_effective.py"
  - "! grep -q '^_OFF = ' word_document_server/engine/effective.py && grep -q '_toggle_value' word_document_server/engine/effective.py"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run ruff check ."
```

### Scope
Part corrective de J05-P3 (finding 1) et J05-P2 (finding 2), D-032 (revue J05-P6) ; aucune n'est réécrite, J05-P6 (✅) n'est pas rejouée : J06-P4 relit cette part.
1. **should-fix**, reproduit par exécution le 2026-09-15 : `core/styles.py:106`, `ensure_heading_style` ne rattrape que `already_exists`. Sur `simple` privé de `Normal`, `Heading7`, `Heading8`, `Heading9`, `engine_create_style` lève pour `Heading7` `LocatorError("not_found")` depuis `_reference_id` (`engine/styles.py:1293-1319`, `'based_on' names 'Normal', which this document does not define…`), qui traverse `ensure_heading_style` et fait répondre à `add_heading` (`tools/content_tools.py:126`, appel hors du `try` de repli `:129-144`) `Failed to add heading: 'based_on' names 'Normal'…` — au niveau 1 aussi, alors que `Heading1` existe. Avant J05 (`except Exception: pass`, « we'll just use default formatting ») le titre était inséré. Un paquet dont la partie styles n'est pas chargée en XML vivant échoue de même par `PackageError` (`_live_styles_root`, `:1126-1132`), jamais rattrapé. Correctif, dans `ensure_heading_style` seul : (a) sonder `Normal` une fois par `style_info(doc, "Normal")` (déjà importé) — si elle lève une `EngineError` (`LocatorError not_found` : pas de `Normal` ; `PackageError` : pas de feuille), la spec des neuf titres omet `based_on` et `next` (le reste — `outline_level`, `q_format`, `ui_priority`, `keep_next`, `spacing`, gras, taille — inchangé : ce sont toujours de vrais titres), sinon elle les garde ; (b) le `except` autour d'`engine_create_style` s'élargit à `EngineError` : `already_exists` continue comme aujourd'hui, toute autre `EngineError` (`PackageError`, `not_found` résiduel) saute le niveau sans lever — « ensure » est au mieux-effort par contrat, il ne lève jamais pour ce que le document n'a pas ; `ValueError` (spec constante fausse) propage toujours. `core.create_style` (`:241-250`) n'est pas touché : un `base_style` inconnu donné par l'appelant doit toujours faire répondre `Failed to create style: …` à `create_custom_style`. Docstring d'`ensure_heading_style` mise à jour (feuille sans `Normal`, mieux-effort). `add_heading` et `create_document` ne changent pas.
Tests dans `tests/engine/test_styles_write.py`, section « The callers: core/styles.py », rouges sur le code actuel, sur un document réellement privé de `Normal` (pas de mock) : helper `_without_normal(pkg)` retirant le `w:style` d'id `Normal` de `styles_root(pkg)` (`docDefaults` intact), forme de `_without_headings` (`:780-787`). `test_ensure_heading_style_tolerates_a_style_sheet_without_normal` : sur `_without_normal(_without_headings(_pkg("simple")))`, `ensure_heading_style(pkg)` ne lève pas ; `get_style(pkg, "Heading3")` → `info.based_on is None`, `info.next_style is None`, `info.q_format is True`, `info.ui_priority == 9`, `paragraph_props["outline_level"] == 2`, `run_props["bold"] is True` ; les neuf `HeadingN` sont dans `list_styles(pkg)`. `test_add_heading_still_inserts_a_heading_when_normal_is_missing(tmp_path)` : la reproduction de la revue — `simple` privé de `Normal`, `Heading7`, `Heading8`, `Heading9`, sauvegardé ; `_run(add_heading(str(path), "Probe one", 1))` puis `(…, "Probe seven", 7)` répondent `Heading '…' (level N) added to …` (`tools/content_tools.py:177`) ; réouvert : les deux derniers paragraphes du corps portent ces textes avec `w:pStyle` `Heading1` et `Heading7`, `get_style(…, "Heading7").info.based_on is None`, `get_style(…, "Heading1")` a les mêmes `run_props`/`paragraph_props` qu'avant (modèle `test_ensure_heading_style_leaves_the_headings_a_document_already_has`), `validate_package(path.read_bytes()) == []`. Pas de test dédié au `PackageError` (il faudrait une partie styles hors `LIVE_CONTENT_TYPES`) : couvert par le même `except`.
2. **optional** : `engine/effective.py:194` recopie `_OFF = frozenset({"0", "false", "off"})` de `engine/styles.py:246` ; `_extra_toggles` (`:288-297`) refait à la main (`child.get(_W_VAL) not in _OFF`) ce que `styles._toggle_value` (`:438-448`, trois états : absent → `None`) fait pour les cinq bascules du modèle. Correctif : `effective.py` importe `_toggle_value` depuis `engine.styles` (l'import de noms privés entre modules du moteur est la pratique du dépôt : `styles.py:105-110` importe `_RPR_RANK`, `_drop_children`, `_ensure_child`, `_parse_patch` de `format`) et `_extra_toggles` devient `value = _toggle_value(properties, tag)` ; `if value is not None: found[name] = value` ; la constante `_OFF` locale disparaît (`_W_VAL` reste, utilisé `:437`) ; un seul vocabulaire et une seule sémantique pour les onze bascules de §17.7.3. Exposer un `toggle_value()` public demanderait `engine/styles.py`, hors `files`. Test `test_the_extra_toggles_read_off_the_way_the_model_toggles_do` dans `tests/engine/test_effective.py` (helpers `_package(body=[…])`, `_report`), paramétré sur `w:val` ∈ `"0"`, `"false"`, `"off"` (→ `False`), `"1"`, `"true"`, `"on"` et attribut absent (→ `True`) : un run direct portant `<w:b w:val="…"/>` et `<w:vanish w:val="…"/>` → `report["run"]["bold"]` et `["vanish"]` valent tous deux `{"value": <attendu>, "source": "direct"}`. Vert avant comme après (aucun changement de comportement) : il gèle l'accord des deux décodeurs que la revue a établi par lecture.
Rien d'autre : ni `engine/styles.py`, ni les outils, ni `tests/characterization` (les xfail retirés par J05-P3 restent verts).

### Context
`core/styles.py:62-111` (`ensure_heading_style`, spec `:83-105`, `except LocatorError` `:106-111`), `:20-22` (imports `LocatorError`, `engine_create_style`, `style_info` ; `EngineError` à importer de `engine.errors`), `:241-250` (`create_style`, intact) ; `engine/errors.py` (`EngineError`, base de `PackageError` et `LocatorError`) ; `engine/styles.py:1293-1319` (`_reference_id`), `:1126-1132` (`_live_styles_root`), `:1177-1190` (`_require_styles_root`, `_ensure_styles_root`), `:1804-1900` (`create_style`, `based_on`/`next` optionnels), `:317` (`StyleInfo.based_on`, `.next_style`), `:438-448` (`_toggle_value`), `:246` (`_OFF`) ; `engine/effective.py:134-149` (`EXTRA_RUN_TOGGLES`), `:194` (`_OFF`), `:288-297` (`_extra_toggles`), `:102-108` (bloc d'import depuis `styles`) ; `tools/content_tools.py:121-179` (`add_heading` : `ensure_heading_style` `:126`, repli `:129-144`, réponse `:177`, `except` `:178-179`) ; `tests/engine/test_styles_write.py:66-115` (helpers `_pkg`, `_element`, `_run`), `:780-840` (`_without_headings`, trois tests `ensure_heading_style`), `:901-925` (tests outil `create_custom_style`, forme `tmp_path` + `build`) ; `tests/engine/test_effective.py:73-110` (`_package`, `_report`), `:294-320` (tests des bascules hors modèle). Reproduit le 2026-09-15 : la 2e acceptation sort en 1 (`Failed to add heading: 'based_on' names 'Normal'…` aux niveaux 1 et 7), les 3e et 4e aussi ; `python-docx` `Document()` définit `Normal` et les neuf titres, donc `create_document` n'est pas concerné.

## J05-P8 — Correctif D-027 : CHANGELOG des changements de comportement de J05, CONTRIBUTING sans édition manuelle des compteurs

```yaml
id: J05-P8
kind: implement
tier: T5
size: S
depends_on: [J05-P9]
files:
  - CHANGELOG.md
  - CONTRIBUTING.md
acceptance:
  - "grep -q -x -F '## [Unreleased]' CHANGELOG.md && test \"$(grep -c '^## ' CHANGELOG.md)\" -eq 10"
  - "test \"$(head -n 7 CHANGELOG.md | sha256sum | cut -d' ' -f1)\" = e3f2ec4b33cfdeabaf535e8e00174ddc3e762efb0b41d2ea1591c247fc4033a9"
  - "test \"$(sed -n '/^## .1.6.0. - 2026-04-29$/,$p' CHANGELOG.md | sha256sum | cut -d' ' -f1)\" = 6e7f8c4252d8a2b6d1a498f36f0e7e6262764dba59ef7943303119c2d3298fe5"
  - "grep -q 'first source document' CHANGELOG.md && grep -q 'End anchor' CHANGELOG.md && grep -q 'w:sectPr' CHANGELOG.md && grep -q 'replace_content=False' CHANGELOG.md && grep -q 'block content control' CHANGELOG.md"
  - "grep -q 'abstractNum' CHANGELOG.md && grep -q 'List Paragraph' CHANGELOG.md && grep -qF \"with style 'Normal'.\" CHANGELOG.md"
  - "grep -q 'ensure_heading_style' CHANGELOG.md && grep -q 'StyleInfo' CHANGELOG.md && grep -q 'create_custom_style' CHANGELOG.md && grep -q 'add_heading' CHANGELOG.md && grep -q 'create_document' CHANGELOG.md && grep -q 'styles.xml' CHANGELOG.md && grep -q 'basedOn' CHANGELOG.md"
  - "! grep -q 'Update the tool count' CONTRIBUTING.md && grep -q 'scripts/gen_tools_md.py' CONTRIBUTING.md && grep -qF '<!-- tool-counts:start -->' CONTRIBUTING.md"
  - "test \"$(head -n 128 CONTRIBUTING.md | sha256sum | cut -d' ' -f1)\" = 20c7b42e65799b66a5830849e2acdd4f14bff17c0b1c4992e2358a38c73442a7 && test \"$(tail -n 2 CONTRIBUTING.md | sha256sum | cut -d' ' -f1)\" = ace960b6545d1daf27d7e432cac58f7546967e8fc442707c6edece067ec031be"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python scripts/gen_tools_md.py --check && PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/test_docs_sync.py -q"
```

### Scope
D-027 (user, s8) : sept changements de comportement visibles par un appelant, introduits par J05 et confirmés dans le diff par la revue J05-P6, n'ont de part propriétaire nulle part — même situation que D-015 (J03-P11) ; et `CONTRIBUTING.md:129` demande « Update the tool count in `README.md` badges and text » alors que ce bloc est généré depuis J04-P3. Écrire les deux, en anglais ; aucun autre fichier. Dépend de J05-P9 pour une seule raison, tranchée ici : la puce d'`add_heading` décrit la feuille sans `Normal` **après** correction (D-032, « quelle que soit l'issue »), et une puce vraie dans les deux états tairait précisément ce que D-027 demande d'écrire ; elle se rédige contre `core/styles.py` tel que mergé — relire `ensure_heading_style` avant d'écrire. Coût nul : J05-P9 est S sans dépendance, la part reste dispatchable en parallèle de J06.
Forme (J03-P11) : entrée `## [Unreleased]` existante (lignes 8–17 : `### Changed` trois puces, `### Fixed` deux puces), une puce par outil — nom en code en tête, tiret cadratin, ancien et nouveau comportement ; les cinq puces présentes restent telles quelles, les nouvelles s'ajoutent à la suite dans chaque sous-section ; pas de numéro de version, pas de ligne `[Unreleased]:` dans le bloc de liens ; préambule (lignes 1–7) et tout ce qui suit `## [1.6.0]` byte à byte intacts (hashes ci-dessus). Cinq puces pour les sept changements, contenu (anglais) :
`### Changed` —
1. `insert_numbered_list_near_text` : now allocates a list definition of its own on every call — a new `w:abstractNum` and `w:num` in `word/numbering.xml`, the part being created when the document has none — instead of pointing the items at the hard-coded `numId` 1 (bullets) or 2 (numbers): two successive calls no longer continue one numbering, and no definition the document already had is modified. The `List Paragraph` style is applied only when the document defines it; the items no longer fall back to `Normal` when it is missing.
2. `insert_line_or_paragraph_near_text` : the reply now ends with `with style 'Normal'.` — the name of the style applied, `line_style` or the target paragraph's — where it used to interpolate the `repr` of a python-docx `_ParagraphStyle` object, different on every run and therefore unparsable.
3. `add_heading` / `create_document` : now write `word/styles.xml` when heading styles are missing — up to nine real heading styles (`Heading1`…`Heading9`: `basedOn Normal`, `next Normal`, `w:outlineLvl`, `qFormat`, `uiPriority` 9) are created, so a heading added to a document that lacks them appears in the navigation pane and in a table of contents; styles the document already defines are left as they are; a style sheet that does not define `Normal` gets them without `basedOn`/`next`, and what cannot be created is skipped rather than failing the tool.
4. `core.styles.ensure_heading_style` / `core.styles.create_style` (Python API, `word_document_server.core.styles`) : no longer no-ops — `ensure_heading_style` swallowed every failure and `create_style` guarded its creation with `get_by_id`, which never raises, so neither wrote anything; both now write through `word_document_server.engine.styles`, and `create_style` returns the engine's `StyleInfo` instead of a python-docx `Style`.
`### Fixed` —
5. `create_custom_style` : now really writes the style into `word/styles.xml` (it used to answer `created successfully` without creating anything, see above) and refuses an unreadable colour (`Invalid color '…'. Use a hex value such as 'FF0000', 'auto', or one of: …`) instead of ignoring it.
`CONTRIBUTING.md` : remplacer la seule ligne 129 (« 2. Update the tool count in `README.md` badges and text if you add/remove tools ») par une consigne (une ligne ou plusieurs, même numéro 2) : after adding or removing a tool, run `uv run python scripts/gen_tools_md.py` — it regenerates `TOOLS.md` and the block between `<!-- tool-counts:start -->` and `<!-- tool-counts:end -->` in `README.md` from the registry — and never edit that block by hand: `uv run python scripts/gen_tools_md.py --check` and `tests/test_docs_sync.py` fail on a stale copy. Lignes 1–128 et 130–131 byte à byte intactes (hashes). Rien d'autre dans ce fichier (le bloc « Project Structure » sans `engine/` ni `tools/v2/` revient à J06-P3).
Si, avant toute édition, `sha256sum CHANGELOG.md` ≠ `71cad7a961f46f81e506fd6f4ce27107ea8992e388a51f52cf08169ffc748851` ou `sha256sum CONTRIBUTING.md` ≠ `a2e2f9dff763108b86447ab695a94d90e37a208e1252d26c660bd814ba9ac42a`, la base du worktree n'est pas celle du plan : rapporter `blocked`, ne pas adapter les hashes. La part ne régénère rien : `gen_tools_md.py --check` doit être vert avant comme après.

### Context
`CHANGELOG.md` (155 lignes : préambule 1–7, `## [Unreleased]` 8–17 — `### Changed` 10–13, `### Fixed` 15–17 —, `## [1.6.0]` ligne 19, bloc de liens 149–155) ; `CONTRIBUTING.md` (131 lignes, ligne 129 fautive). Code, tel que mergé : `utils/document_utils.py:481-542` (`insert_line_or_paragraph_near_text`, `style_name` `:528`, réponses `:540-542`), `:569-660` (`insert_numbered_list_near_text`, allocation `:623-634`, candidats `List Paragraph`/`ListParagraph` `:636-647`), `engine/numbering.py` (`create_list_definition`) ; `core/styles.py:62-111` (`ensure_heading_style`, corrigée par J05-P9), `:185-250` (`create_style`, docstring « Until J05 this function created nothing ») ; `tools/format_tools.py:154-192` (`create_custom_style`) ; `tools/content_tools.py:121-179` (`add_heading`), `tools/document_tools.py:29-47` (`create_document`) ; `scripts/gen_tools_md.py:38-39,161-195`, `tests/test_docs_sync.py`. Avant J05 : `git show c9980c1^:word_document_server/utils/document_utils.py` (`num_id = 1 if … else 2` `:638`, candidats `['List Paragraph', 'ListParagraph', 'Normal']` `:642`) ; `git show fb32c98^:word_document_server/core/styles.py` (`except Exception: pass`, `get_by_id` puis `return style`). Hashes calculés le 2026-09-15 ; les 5e, 6e et 7e acceptations sont rouges sur l'arbre actuel, les six autres vertes.

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
