# J04 — Surface V2 : adressage, inspection, dry-run, capacités, docs

Goal: Nouveaux outils `doc_*` sur le moteur : locators hybrides sans état, inspection, édition de texte avec `dry_run` et rapport structuré, lot atomique, capacités de plateforme, `TOOLS.md` et compteurs README générés avec garde-fous de dérive, registre corrigé, trois plantages macOS prouvés par lecture corrigés. Ouvre par la corrective D-016 : un seul espace d'index pour tout `paragraph_index` public, notes exceptées.
Depends on: J03 · Orchestrator: opus/high

## J04-P7 — Correctif D-016 : un seul espace d'index pour `paragraph_index` (D-017, D-018, D-019 embarquées)

```yaml
id: J04-P7
kind: implement
tier: T3
size: M
depends_on: []
files:
  - word_document_server/utils/extended_document_utils.py
  - word_document_server/utils/document_utils.py
  - word_document_server/tools/layout_tools.py
  - word_document_server/tools/content_tools.py
  - tests/tools/test_index_space.py
  - tests/tools/test_search_replace.py
  - CHANGELOG.md
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/tools/test_index_space.py tests/tools/test_search_replace.py tests/tools/test_layout_blocks.py tests/characterization -q"
  - "grep -q 'def test_the_paragraph_read_is_the_one_find_text_reported' tests/tools/test_index_space.py && grep -q 'def test_the_paragraph_spaced_is_the_one_find_text_reported' tests/tools/test_index_space.py && grep -q 'def test_the_anchor_of_each_insert_near_text_is_the_one_find_text_reported' tests/tools/test_index_space.py && grep -q 'def test_a_negative_bookmark_index_is_refused' tests/tools/test_index_space.py"
  - "! grep -q 'doc.paragraphs\\[paragraph_index\\]' word_document_server/utils/extended_document_utils.py && ! grep -q 'doc\\.paragraphs' word_document_server/tools/layout_tools.py && ! grep -q 'doc.paragraphs\\[target_paragraph_index\\]' word_document_server/utils/document_utils.py"
  - "grep -q 'paragraph_index < 0 or paragraph_index >= len(paragraphs)' word_document_server/tools/layout_tools.py"
  - "! grep -qF ', {report.skipped} skipped' word_document_server/tools/content_tools.py && ! grep -q 'found\\., ' tests/tools/test_search_replace.py"
  - "test \"$(sed -n '/^## .Unreleased.$/,/^## .1.6.0. - 2026-04-29$/p' CHANGELOG.md | grep -c '^- ')\" -eq 5 && test \"$(grep -c '^## ' CHANGELOG.md)\" -eq 10"
  - "grep -q 'find_text_in_document' CHANGELOG.md && grep -q 'get_paragraph_text_from_document' CHANGELOG.md && grep -q 'set_paragraph_spacing' CHANGELOG.md && grep -q 'insert_line_or_paragraph_near_text' CHANGELOG.md && grep -q 'w:sdt' CHANGELOG.md"
  - "test \"$(head -n 7 CHANGELOG.md | sha256sum | cut -d' ' -f1)\" = e3f2ec4b33cfdeabaf535e8e00174ddc3e762efb0b41d2ea1591c247fc4033a9 && test \"$(sed -n '/^## .1.6.0. - 2026-04-29$/,$p' CHANGELOG.md | sha256sum | cut -d' ' -f1)\" = 6e7f8c4252d8a2b6d1a498f36f0e7e6262764dba59ef7943303119c2d3298fe5"
```

### Scope
D-016 (revue J03-P8 round 2) : J03-P1 et J03-P5 ont mis `find_text_in_document`, `delete_paragraph` et `add_bookmark` dans l'espace V2, les autres consommateurs d'un `paragraph_index` public lisent encore `doc.paragraphs` (python-docx : enfants directs du corps, contenu d'un `w:sdt` de bloc exclu). Mesuré le 2026-09-10 sur un document à sommaire (`add_table_of_contents` insère un `w:sdt` de bloc de 2 paragraphes) : `find_text_in_document("Charlie") → 4`, puis `get_paragraph_text_from_document(4) → "Echo"`, `set_paragraph_spacing(paragraph_index=4)` espace « Echo », chacun des trois `insert_*_near_text(target_paragraph_index=4)` insère près de « Echo » en répondant `(index 4)`. Aucun n'est destructeur, tous sont faux dès qu'un `w:sdt` de bloc précède.
Aligner sur `indexed_paragraphs()` (`utils/document_utils.py:748`, enveloppe de `_v2_index_map(iter_paragraphs(root))` : la réutiliser, jamais réécrire ni dupliquer le filtre) :
- `utils/extended_document_utils.py::get_paragraph_text` — index et borne (`Document has N paragraphs`, N = taille de l'espace V2) ; lecture inchangée (`get_effective_text`, `paragraph.style.name`, clés du dict).
- `tools/layout_tools.py::set_paragraph_spacing` — `total`, `paragraph_index` et la plage `start_paragraph`/`end_paragraph` dans le même espace (sans index = tout l'espace V2, paragraphes d'un sdt de bloc compris) ; `paragraph_format` conservé.
- `utils/document_utils.py::insert_header_near_text`, `insert_line_or_paragraph_near_text`, `insert_numbered_list_near_text` — `target_paragraph_index` et sa borne ; la recherche par texte parcourt le même espace (règle de saut des styles `TOC…` conservée) pour que l'`(index N)` du message soit compté dans l'espace où `target_paragraph_index` est lu.
Chemin vérifié : python-docx conservé, `[Paragraph(p, doc) for p in indexed_paragraphs(doc.element.body)]` (`docx.text.paragraph.Paragraph` ; `CT_Body` et `CT_P` sont des éléments lxml ; `.style`, `.paragraph_format`, `._element` fonctionnent) ; messages, clés et formats de retour inchangés, seuls N et le paragraphe atteint changent. `main.py` inchangé (signatures identiques). Rien dans `tools/footnote_tools.py` ni `core/footnotes.py` : les 5 outils de notes restent sur `doc.paragraphs`, écart assumé (Scope « Out », R-003, Risks du PLAN.md).
D-018 : `tools/layout_tools.py:472` → `if paragraph_index < 0 or paragraph_index >= len(paragraphs):`, message `Paragraph -1 does not exist.` (forme existante).
D-019 : `tools/content_tools.py:662`, site unique d'assemblage du suffixe → `message += f" {report.skipped} skipped (inside fields)."`, soit `No occurrences of 'Page 1' found. 1 skipped (inside fields).` et `Replaced 1 occurrence(s) of 'Fixture: fields' with 'Fixture: champs'. 1 skipped (inside fields).` (la forme `'., 1 skipped` du cas `Replaced` sort du même site et de J03-P1 : corrigée avec) ; `tests/tools/test_search_replace.py:177-180,194` mis à jour, rien d'autre dans ce fichier.
D-017 : 5e puce de `## [Unreleased]`, dernière de `### Changed`, en anglais, format des 4 existantes (outils en code et gras, tiret cadratin, ancien puis nouveau comportement) : `find_text_in_document`, `delete_paragraph`, `add_bookmark`, `get_paragraph_text_from_document`, `set_paragraph_spacing`, `insert_header_near_text`, `insert_line_or_paragraph_near_text`, `insert_numbered_list_near_text` — `paragraph_index` / `target_paragraph_index` count every paragraph of the body in document order, the content of a block content control (`w:sdt`, the one `add_table_of_contents` inserts for instance) included, table cells and text boxes excluded ; they used to count python-docx's direct body paragraphs only, so an index recorded before this change may point at another paragraph in a document holding a block content control ; the footnote and endnote tools are unchanged and still count the direct body paragraphs. Préambule et historique ≥ 1.6.0 byte à byte intacts (hashes de J03-P11).
Tests, `tests/tools/test_index_space.py` (nouveau ; helpers à recopier de `tests/tools/test_layout_blocks.py` — `_document_with_a_table_of_contents`, `_v2_index_of`, `_all_paragraph_texts` —, aucun import d'un module de test) : `test_the_paragraph_read_is_the_one_find_text_reported` (`get_paragraph_text_from_document(path, i)`, `i` rendu par `find_text` pour « Charlie » → `text == "Charlie"`, `index == i`) ; `test_the_paragraph_spaced_is_the_one_find_text_reported` (formes `paragraph_index` et `start_paragraph`/`end_paragraph` ; seul « Charlie » porte le `w:spacing`, `assert_unchanged_except(…, paragraphs=[i])` — sans tableau, l'espace de l'instantané est l'espace V2) ; `test_the_anchor_of_each_insert_near_text_is_the_one_find_text_reported` (paramétré sur les trois fonctions, `position="after"` : dans `_all_paragraph_texts` le nouveau paragraphe suit « Charlie », `(index i)` dans le message — assertion par sous-chaîne, le message de `insert_line_or_paragraph_near_text` interpole l'objet style sans `line_style`, défaut hors trigger à laisser) ; `test_a_negative_bookmark_index_is_refused` (`add_bookmark(path, -1, "Neg")` → pas de `"success"`, aucun `w:bookmarkStart` nommé `Neg`, octets du fichier intacts). Les quatre sont rouges sur le code actuel ; `validate_package(path) == []` après chaque écriture.

### Context
`utils/extended_document_utils.py:44-75` ; `tools/layout_tools.py:351-431` (`set_paragraph_spacing`, `doc.paragraphs` à `:395,410`), `:434-496` (`add_bookmark`, garde `:472`) ; `utils/document_utils.py:422-468, 471-523, 565-646` (`doc.paragraphs[target_paragraph_index]` à `:435,488,589`), `:736-766` (`body_paragraphs`, `indexed_paragraphs`) ; `tools/content_tools.py:655-663` ; `tools/extended_document_tools.py:18-38` (wrapper, refuse déjà `paragraph_index < 0`) ; `engine/find.py:93,134-172` (`_UNINDEXED_ANCESTORS`, `iter_paragraphs`, `_v2_index_map`). Hors `files`, à laisser : `tools/footnote_tools.py:58,129,319,465,625`, `core/footnotes.py:803,822`, `main.py`. Modèle de test : `tests/tools/test_layout_blocks.py:245-292` (J03-P5, mêmes documents, mêmes helpers). `CHANGELOG.md` : 154 lignes, `## [Unreleased]` ligne 8 (2 `Changed` + 2 `Fixed`), `## [1.6.0]` ligne 18 ; `sha256sum CHANGELOG.md` = `4862174bc4a1e7ceba5bf39ded135b75e589740f2ea3a5abf392a284121815d8` avant édition, sinon la base du worktree n'est pas la branche : `blocked`, ne pas adapter les hashes.

## J04-P1 — Locators et inspection

```yaml
id: J04-P1
kind: implement
tier: T3
size: M
depends_on: [J04-P7]
files:
  - word_document_server/engine/locators.py
  - word_document_server/engine/inspect.py
  - tests/engine/test_locators.py
  - tests/engine/test_inspect.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_locators.py tests/engine/test_inspect.py -q"
```

### Scope
Schéma de locator (dict) : `{"paragraph": i, "expect_text": s}` (i = index V2 : `w:p` du corps en ordre de document, contenu des sdt de bloc inclus, cellules et zones de texte exclues, base 0 ; `expect_text` = préfixe ou texte exact), `{"find": s, "occurrence": n, "within": locator}`, `{"bookmark": name}`, `{"heading": s}`, `{"table": t, "row": r, "col": c, "paragraph": k}`, clé optionnelle `"story"` = id de story tel que `DocxPackage.stories()` le rend (`document`, `header1`, `footer2`, `footnotes`, `endnotes`… ; `"body"` accepté en entrée comme alias de `document`, comme dans `find`, jamais rendu).
`resolve(pkg, locator) -> Target(story, paragraph, index, start, end)`, même forme que `Match` : `story` = id réel, `paragraph` = `w:p` vivant, `index` = index V2 ou `None` pour un paragraphe en cellule (atteint par le locator `table`) ou en zone de texte (atteint par `find` seulement), jamais par `paragraph` ; erreurs `LocatorError` avec code `not_found`, `ambiguous`, `stale_anchor` (candidats les plus proches en détail).
`inspect(pkg) -> dict` : blocs (index, genre, style, préfixe de texte 80 caractères, niveau de liste, présence de champs, commentaires, révisions), tableaux (index, lignes, colonnes, première cellule), sections (format, orientation, marges, en-têtes et pieds), stories, compteurs, styles utilisés, signets, champs.
Tests sur les fixtures : chaque forme de locator, ancre périmée, ambiguïté, indices V2 stables face aux cellules, aux zones de texte (`text_boxes`) et aux sdt.

### Context
`engine/find.py::iter_paragraphs` et `_v2_index_map` (seul filtre de l'espace V2, partagé par `find` et `list_revisions` : le réutiliser, pas le réécrire), `engine/textmodel.py`. `indexed_paragraphs()` (`utils/document_utils.py`) est l'enveloppe côté outils du même filtre, alignée par J04-P7 sur les cinq derniers points d'appel : l'espace que `resolve` numérote est exactement celui que `find_text_in_document` et `get_paragraph_text_from_document` rendent — un filtre, deux enveloppes (`engine`, `utils`), jamais une troisième. Le résultat de `inspect` est ce que l'agent lit avant de construire un locator.

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
`registry.py` : `register_v2_tools(mcp)` découvre les modules `tools/v2/*.py` exposant `TOOLS: list[ToolSpec(fn, annotations, tags)]`, enregistre chaque fonction telle quelle (schéma dérivé de la signature, docstring = description), convertit les `EngineError` en `{"status": "error", "code", "message"}` ; résultat de succès `{"status": "ok", "dry_run", "changes": [{"story", "paragraph", "before", "after"}], "warnings"}` (`story` = id réel de story, `paragraph` = index V2 ou `null` en cellule de tableau ou en zone de texte) ; verrou de fichier via `utils.file_utils.get_file_lock`.
`text.py` : `doc_inspect(filename)`, `doc_find(filename, pattern, regex, case, whole_word, stories, max_results)` (`stories` accepte l'alias `"body"` en entrée ; chaque résultat rapporte `Match.story` et `Match.index`, `null` en cellule ou en zone de texte), `doc_edit_text(filename, locator, action="replace" | "insert" | "delete", text, start, end, track_changes, author, dry_run)`, `doc_format_range(filename, locator, start, end, patch, dry_run)` ; `dry_run` ne sauvegarde rien mais renvoie le rapport complet.
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
depends_on: [J04-P1, J04-P2, J04-P3, J04-P4, J04-P5, J04-P7]
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
Relit aussi J04-P7 (corrective D-016, D-017, D-018, D-019) : les cinq points d'appel alignés sur `indexed_paragraphs()` sans second filtre ; les 5 outils de notes de `tools/footnote_tools.py` restés délibérément sur `doc.paragraphs` (hors périmètre, R-003 — pas un finding) ; la 5e puce `[Unreleased]` contre le code mergé, historique intact ; la borne inférieure d'`add_bookmark` ; le séparateur de `search_and_replace` dans ses deux formes.
