# J01 — Harnais de fidélité documentaire

Goal: Rendre la dégradation mesurable avant toute modification du code : fixtures générées par deux producteurs, instantanés canoniques, validateur de paquet, caractérisation des outils existants ; outillage pytest, ruff et groupe de dépendances de développement.
Depends on: — · Orchestrator: opus/high

## J01-P1 — Outillage : pytest, groupe dev, ruff, script de vérification, avertissements destructifs

```yaml
id: J01-P1
kind: implement
tier: T5
size: M
depends_on: []
files:
  - pyproject.toml
  - uv.lock
  - scripts/check.sh
  - tests/
  - word_document_server/main.py
  - word_document_server/tools/format_tools.py
  - CONTRIBUTING.md
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv sync && PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/ -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run ruff check ."
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"import tomllib;d=tomllib.load(open('pyproject.toml','rb'));deps=' '.join(d['project']['dependencies']);assert 'pytest' not in deps and 'lxml' in deps and 'python-dotenv' in deps;assert 'dev' in d['dependency-groups']\""
  - "grep -q 'DESTRUCTIVE' word_document_server/main.py && grep -q 'DESTRUCTIVE' word_document_server/tools/format_tools.py"
  - "bash scripts/check.sh"
```

### Scope
Déplacer `pytest` dans `[dependency-groups] dev` avec `pytest-timeout` et `ruff` ; déclarer `lxml>=5` et `python-dotenv` dans les dépendances runtime ; ne rien retirer d'autre.
`[tool.pytest.ini_options]` : `testpaths = ["tests"]`, marqueurs `libreoffice` et `characterization`, `timeout = 120`.
`[tool.ruff]` : `target-version = "py311"`, `include` limité à `word_document_server/engine/**`, `tests/**`, `scripts/**` (le code hérité n'est pas linté).
`scripts/check.sh` : `uv sync`, `ruff check .`, `pytest tests/ -q`, avec `PATH="$HOME/.local/bin:$PATH"`.
Sous `tests/`, créer uniquement les `__init__.py` vides de `tests/`, `tests/support/`, `tests/fixtures/`, `tests/engine/`, `tests/tools/`, `tests/characterization/`, `tests/live/`, `tests/e2e/` ; ne pas toucher à `tests/test_convert_to_pdf.py`. `CONTRIBUTING.md` : `uv sync` remplace `pip install -e ".[dev]"`, section tests.
Dans `main.py`, préfixer les descriptions de `add_table_of_contents` et `merge_documents` par `DESTRUCTIVE: rebuilds the document from plain text (formatting, images, fields, comments, footnotes and sections are lost); scheduled for rewrite.` et passer `format_text` en `destructiveHint=True` ; préfixer la docstring de `format_tools.format_text` (source de sa description) par `DESTRUCTIVE: clears every run of the paragraph; scheduled for rewrite.`.
Ne pas toucher aux autres outils ni au code hérité.

### Context
`pyproject.toml`, `CONTRIBUTING.md`, `main.py:100-1934` (registre), `tools/format_tools.py:27-138`. `uv` est en `~/.local/bin`, hors PATH. Aucun workflow GitHub Actions.

## J01-P2 — Fixtures générées : constructeur OOXML déterministe

```yaml
id: J01-P2
kind: implement
tier: T3
size: M
depends_on: [J01-P1]
files:
  - tests/fixtures/builders.py
  - tests/fixtures/test_builders.py
  - tests/conftest.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/fixtures -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"from tests.fixtures.builders import ALL_FIXTURES; assert len(ALL_FIXTURES) >= 18, ALL_FIXTURES.keys()\""
```

### Scope
`builders.py` : fonctions `build_<name>() -> bytes` produisant un `.docx` depuis le gabarit python-docx puis des fragments lxml injectés (dates, ids, rsids fixes → sortie déterministe) ; `ALL_FIXTURES: dict[str, Callable]`.
Fixtures : `simple`, `mixed_runs` (runs coupés par rsid et propriétés), `paragraph_styles`, `character_styles` (`rStyle`), `style_inheritance` (chaîne `basedOn`, `next`, `link`), `themes` (`asciiTheme`, `themeColor` + tint/shade), `complex_numbering` (abstractNum multi-niveaux, restart), `tables` (cellules fusionnées, tableau imbriqué), `comments` (comments.xml + commentsExtended fil de réponse), `tracked_changes` (ins, del, ins>del, rPrChange, pPrChange, marque de paragraphe supprimée), `hyperlinks` (externe et ancre), `bookmarks`, `fields` (fldSimple PAGE, TOC complexe fldChar, champ imbriqué), `footnotes` (+ endnotes), `headers_footers` (champs, image), `sections` (deux sectPr), `content_controls` (sdt bloc et inline), `drawings` (image inline), `combined`.
`conftest.py` : fixture `fixture_docx(name, tmp_path) -> Path` écrivant le fichier ; `test_builders.py` : chaque fixture s'ouvre avec python-docx et son XML contient les balises attendues.
Ne pas utiliser LibreOffice ici ; ne pas commiter de binaire.

### Context
Gabarit : `.venv/lib/python3.12/site-packages/docx/templates/default.docx`. Ajouter une partie via `docx.opc.part.Part` + `doc.part.relate_to`. Les contrats de nommage sont repris par J01-P5, J02 et J03.

## J01-P3 — Producteur LibreOffice et validateur de paquet

```yaml
id: J01-P3
kind: implement
tier: T4
size: M
depends_on: [J01-P1]
files:
  - tests/support/libreoffice.py
  - tests/support/package_check.py
  - tests/support/test_libreoffice.py
  - tests/support/test_package_check.py
  - tests/fixtures/odt/rich.fodt
  - tests/fixtures/odt/tables_lists.fodt
  - tests/fixtures/odt/notes_fields.fodt
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/support/test_libreoffice.py tests/support/test_package_check.py -q"
```

### Scope
`libreoffice.py` : `soffice_path()`, `convert(src, fmt, outdir, timeout=120)`, `fodt_to_docx(name, cache_dir)` (une conversion par session, `tmp_path_factory`), `opens_in_libreoffice(path) -> bool` (conversion PDF, code 0 et PDF non vide), marqueur `requires_libreoffice = pytest.mark.skipif(...)` + `pytest.mark.libreoffice`.
Trois `.fodt` couvrant : styles et runs mixtes, révisions, commentaire, signet, note, champs page, hyperlien, listes, tableaux.
`package_check.py` : `validate_package(path) -> list[Issue(code, part, message)]` : `testzip`, type de contenu pour chaque partie, cibles de relations existantes (paquet et parties), partie principale présente, XML bien formé, ids uniques par famille (signets, commentaires, ins/del), ids de commentaires et de notes cohérents avec leurs parties, `numId` résolus, aucun fichier temporaire à côté du document.
Tests : fixtures LibreOffice valides ; échantillons corrompus (cible de relation manquante, id dupliqué, partie sans type) détectés.

### Context
`soffice` est en `/usr/bin` sur cette machine (LibreOffice 25.2), ≈2 s par conversion. Le modèle `/tmp/lo-probe/rich.fodt` de la session de planification n'est pas dans le dépôt : réécrire les fodt. `core/footnotes.py:613-735` (`validate_document_footnotes`) sert de référence.

## J01-P4 — Instantané canonique et comparaison

```yaml
id: J01-P4
kind: implement
tier: T3
size: S
depends_on: [J01-P1]
files:
  - tests/support/snapshot.py
  - tests/support/test_snapshot.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/support/test_snapshot.py -q"
```

### Scope
`snapshot(path_or_bytes) -> Snapshot` : empreinte par partie (c14n pour XML et rels, octets bruts sinon ; `[Content_Types].xml` comparé comme ensemble `(partie, type)`), signatures de paragraphes par story (story, index, style, texte visible par parcours complet des `w:t` hors `w:del`, runs `(texte, rPr c14n)`, marqueurs, numPr), tableaux (dimensions, texte des cellules), relations `(type, cible)`, compteurs (commentaires, révisions, signets, champs, notes).
`diff(a, b) -> Diff` (parties, paragraphes ajoutés/supprimés/modifiés avec détail), `assert_unchanged_except(a, b, paragraphs=set(), parts=set())`.
lxml et zipfile uniquement, sans python-docx. Tests sur des documents construits en mémoire.

### Context
Sera déplacé dans `engine/compare.py` en J06-P2 ; garder une API fonctionnelle simple.

## J01-P5 — Caractérisation des outils existants

```yaml
id: J01-P5
kind: implement
tier: T4
size: M
depends_on: [J01-P2, J01-P3, J01-P4]
files:
  - tests/characterization/test_no_op_roundtrip.py
  - tests/characterization/test_existing_tools.py
  - docs/audit/destructive-ops.md
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/characterization -q"
  - "test -f docs/audit/destructive-ops.md && grep -q add_table_of_contents docs/audit/destructive-ops.md"
```

### Scope
`test_no_op_roundtrip.py` : pour chaque fixture (constructeur et LibreOffice), `Document(path).save(out)` sans le hook `save_utils` : aucune partie perdue, parties XML canoniquement identiques, parties binaires identiques, `validate_package` vide, `comments.xml` et `footnotes.xml` conservés.
`test_existing_tools.py` : pour chaque outil cross-platform mutateur (au moins : `search_and_replace`, `format_text`, `add_table_of_contents`, `merge_documents`, `add_header_footer`, `delete_paragraph`, `add_bookmark`, `insert_numbered_list_near_text`, `add_comment`, `track_replace`, `track_insert`, `track_delete`, `accept_tracked_changes`, `reject_tracked_changes`, `manage_hyperlinks`, `create_custom_style`, `add_heading`, `add_paragraph`, `set_paragraph_spacing`, `format_table_cell_text`, `replace_paragraph_block_below_header`, `replace_block_between_manual_anchors`, `add_footnote_robust`, `protect_document`/`unprotect_document`) : opération minimale sur `combined`, puis `assert_unchanged_except` et `validate_package` ; comportement destructeur connu → `xfail(strict=True, reason=...)` ; `track_replace("Risk", "Risk Risk")` sous `timeout(10)` en xfail.
`docs/audit/destructive-ops.md` : tableau outil, symptôme, test témoin, généré à la main depuis les raisons de xfail.

### Context
Preuves de la planification : `content_tools.py:313-400` (TOC), `document_tools.py:141-214` (merge), `format_tools.py:27-138`, `core/styles.py:53-134`, `core/tracked_changes.py:200-271`, `utils/document_utils.py:246-285`, `utils/save_utils.py`. Les outils sont `async` : appeler via `asyncio.run`.

## J01-P6 — Review J01

```yaml
id: J01-P6
kind: review
tier: T2
size: S
depends_on: [J01-P1, J01-P2, J01-P3, J01-P4, J01-P5]
files: []
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv sync && PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/ -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run ruff check ."
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"from tests.fixtures.builders import ALL_FIXTURES; assert len(ALL_FIXTURES) >= 18\""
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/characterization -q"
  - "test -f docs/audit/destructive-ops.md && grep -q add_table_of_contents docs/audit/destructive-ops.md"
  - "bash scripts/check.sh"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv build"
```

### Scope
Review the merged milestone diff with verify-before-done, code-review, test-design. Report; change nothing.
