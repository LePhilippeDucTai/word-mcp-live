# J04 — Surface V2 : adressage, inspection, dry-run, capacités, docs

Goal: Nouveaux outils `doc_*` sur le moteur : locators hybrides sans état, inspection, édition de texte avec `dry_run` et rapport structuré, lot atomique, capacités de plateforme, `TOOLS.md` et compteurs README générés avec garde-fous de dérive, registre corrigé, trois plantages macOS prouvés par lecture corrigés.
Depends on: J03 · Orchestrator: opus/high

## J04-P1 — Locators et inspection

```yaml
id: J04-P1
kind: implement
tier: T3
size: M
depends_on: []
files:
  - word_document_server/engine/locators.py
  - word_document_server/engine/inspect.py
  - tests/engine/test_locators.py
  - tests/engine/test_inspect.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_locators.py tests/engine/test_inspect.py -q"
```

### Scope
Schéma de locator (dict) : `{"paragraph": i, "expect_text": s}` (i = index V2 : `w:p` du corps en ordre de document, contenu des sdt de bloc inclus, cellules exclues, base 0 ; `expect_text` = préfixe ou texte exact), `{"find": s, "occurrence": n, "within": locator}`, `{"bookmark": name}`, `{"heading": s}`, `{"table": t, "row": r, "col": c, "paragraph": k}`, clé optionnelle `"story"` (`body`, `header:i`, `footer:i`, `footnotes`, `endnotes`).
`resolve(pkg, locator) -> Target(story, paragraph, index, start, end)` ; erreurs `LocatorError` avec code `not_found`, `ambiguous`, `stale_anchor` (candidats les plus proches en détail).
`inspect(pkg) -> dict` : blocs (index, genre, style, préfixe de texte 80 caractères, niveau de liste, présence de champs, commentaires, révisions), tableaux (index, lignes, colonnes, première cellule), sections (format, orientation, marges, en-têtes et pieds), stories, compteurs, styles utilisés, signets, champs.
Tests sur les fixtures : chaque forme de locator, ancre périmée, ambiguïté, indices V2 stables face aux cellules et sdt.

### Context
`engine/find.py::iter_paragraphs`, `engine/textmodel.py`. Le résultat de `inspect` est ce que l'agent lit avant de construire un locator.

## J04-P2 — Outils `doc_*` texte, enregistrement, rapport structuré

```yaml
id: J04-P2
kind: implement
tier: T3
size: M
depends_on: [J04-P1]
files:
  - word_document_server/tools/v2/__init__.py
  - word_document_server/tools/v2/registry.py
  - word_document_server/tools/v2/text.py
  - word_document_server/main.py
  - tests/tools/test_v2_text.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_v2_text.py -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"import asyncio; from word_document_server.main import mcp, register_tools; register_tools(); names={t.name for t in asyncio.run(mcp.list_tools())}; assert {'doc_inspect','doc_find','doc_edit_text','doc_format_range'} <= names and len(names) >= 124, sorted(names)\""
```

### Scope
`registry.py` : `register_v2_tools(mcp)` découvre les modules `tools/v2/*.py` exposant `TOOLS: list[ToolSpec(fn, annotations, tags)]`, enregistre chaque fonction telle quelle (schéma dérivé de la signature, docstring = description), convertit les `EngineError` en `{"status": "error", "code", "message"}` ; résultat de succès `{"status": "ok", "dry_run", "changes": [{"story", "paragraph", "before", "after"}], "warnings"}` ; verrou de fichier via `utils.file_utils.get_file_lock`.
`text.py` : `doc_inspect(filename)`, `doc_find(filename, pattern, regex, case, whole_word, stories, max_results)`, `doc_edit_text(filename, locator, action="replace" | "insert" | "delete", text, start, end, track_changes, author, dry_run)`, `doc_format_range(filename, locator, start, end, patch, dry_run)` ; `dry_run` ne sauvegarde rien mais renvoie le rapport complet.
`main.py` : un seul appel `register_v2_tools(mcp)` dans `register_tools` ; aucun wrapper existant modifié.

### Context
fastmcp 4.0.3 : `mcp.tool(fn, annotations=..., tags=...)`, retour dict → contenu structuré. Locators : `engine/locators.py`. Formats d'erreur et de succès repris par J04-P5, J05, J06.

## J04-P3 — Plateformes, TOOLS.md généré, garde-fous de dérive, registre corrigé

```yaml
id: J04-P3
kind: implement
tier: T4
size: M
depends_on: [J04-P2]
files:
  - word_document_server/tools/platforms.py
  - word_document_server/tools/v2/capabilities.py
  - scripts/gen_tools_md.py
  - TOOLS.md
  - README.md
  - tests/test_docs_sync.py
  - tests/test_registry_consistency.py
  - word_document_server/main.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python scripts/gen_tools_md.py --check"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/test_docs_sync.py tests/test_registry_consistency.py -q"
```

### Scope
`platforms.py` : `PLATFORMS: dict[str, frozenset[str]]` (valeurs parmi `linux`, `windows`, `macos`) pour les 120 outils existants selon la table de parité de l'audit (30 outils live avec branche macOS fonctionnelle, 15 Windows seulement, cross-platform pour le reste) et pour les outils `doc_*`.
`capabilities.py` : `doc_capabilities()` → plateforme, LibreOffice détecté, familles d'outils et compteurs, note « Word availability is only known to live tools ».
`scripts/gen_tools_md.py` : génère `TOOLS.md` depuis `mcp.list_tools()` + `PLATFORMS` (nom, première ligne de description, plateformes, lecture seule / destructif) et remplace les compteurs du README entre marqueurs `<!-- tool-counts:start -->` / `<!-- tool-counts:end -->` ; `--check` sort en 1 en cas de dérive.
`test_registry_consistency.py` : chaque outil enregistré a une entrée `PLATFORMS` ; vérification AST que chaque wrapper de `main.py` transmet tous les paramètres de son implémentation. Corrections dans `main.py` : `word_live_modify_table` transmet `scrub_orphans` ; `config['debug']` remplacé par `config.get('debug', False)` ; `load_dotenv()` appelé avant l'import de `defaults`.

### Context
Compteurs actuels faux : README 124/80/44/40, TOOLS.md 115, manifest 114 ; réel 120 (75 + 45, dont 30 avec macOS). `main.py:39-71, 1064-1088, 1993`.

## J04-P4 — Plantages macOS prouvés par lecture

```yaml
id: J04-P4
kind: implement
tier: T4
size: M
depends_on: []
files:
  - word_document_server/tools/live_layout_tools.py
  - word_document_server/core/word_mac.py
  - word_document_server/tools/live_tools.py
  - tests/live/test_mac_paths.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/live/test_mac_paths.py -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -W error -c \"import pathlib; compile(pathlib.Path('word_document_server/tools/live_tools.py').read_text(), 'live_tools.py', 'exec')\""
```

### Scope
`live_layout_tools.py:47` : la branche macOS de `word_live_set_page_layout` transmet les paramètres `*_inches` existants convertis en points, comme la branche COM.
`core/word_mac.py` : `import re` au niveau module (usage à `:1761`).
`live_tools.py:195` : `highlight_color` entier converti en nom de couleur JXA avant `mac_format_text` ; docstring de `live_tools.py:544` contenant `^\d` passée en chaîne brute (`SyntaxWarning` ; l'acceptation compile depuis la source, un `.pyc` en cache masque l'avertissement à l'import).
Tests : `_MAC_AVAILABLE` forcé à vrai, `_run_jxa`/`_run_applescript` bouchonnés pour capturer le script généré ; aucune exception, valeurs attendues présentes dans le script.
Aucune autre modification des modules live.

### Context
Preuves : rapport d'audit live (NameError `page_width`, NameError `re`, AttributeError `int.replace`). Décision D-001 : périmètre live limité à ces trois corrections.
Bouchonnage : `_MAC_AVAILABLE` est défini par module (`tools/live_layout_tools.py:12`, `tools/live_tools.py:14`, `sys.platform == 'darwin'`) ; `_run_jxa` et `_run_applescript` vivent dans `core/word_mac.py:18,47` et les branches macOS importent `word_mac` à l'appel.

## J04-P5 — Lot atomique `doc_apply_edits`

```yaml
id: J04-P5
kind: implement
tier: T4
size: S
depends_on: [J04-P2]
files:
  - word_document_server/tools/v2/batch.py
  - tests/tools/test_v2_batch.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_v2_batch.py -q"
```

### Scope
`doc_apply_edits(filename, edits, dry_run)` : `edits` = liste de charges identiques à `doc_edit_text` et `doc_format_range` ; application en mémoire sur un seul `DocxPackage`, locators résolus après chaque édition ; toute erreur → rien n'est écrit, rapport avec l'index de l'édition fautive ; une seule sauvegarde atomique ; rapport agrégé.
Tests : succès, échec au milieu (fichier inchangé, instantané identique), `dry_run`.

### Context
`tools/v2/text.py`, `tools/v2/registry.py`, `engine/package.py`.

## J04-P6 — Review J04

```yaml
id: J04-P6
kind: review
tier: T2
size: S
depends_on: [J04-P1, J04-P2, J04-P3, J04-P4, J04-P5]
files: []
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine tests/tools tests/live tests/test_docs_sync.py tests/test_registry_consistency.py -q --timeout=60"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python scripts/gen_tools_md.py --check"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -W error -c \"import pathlib; compile(pathlib.Path('word_document_server/tools/live_tools.py').read_text(), 'live_tools.py', 'exec')\""
  - "PATH=\"$HOME/.local/bin:$PATH\" uv sync && PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/ -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv build"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run ruff check ."
```

### Scope
Review the merged milestone diff with verify-before-done, code-review, test-design. Report; change nothing.
