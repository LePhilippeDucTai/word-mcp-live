# J06 — Audit, comparaison et validation de bout en bout

Goal: Outils d'audit et de comparaison sémantique en lecture seule, scénarios agent de bout en bout sur les fixtures des deux producteurs, documentation alignée sur le registre ; valide l'objectif du plan.
Depends on: J05 · Orchestrator: opus/high

## J06-P1 — Audit documentaire

```yaml
id: J06-P1
kind: implement
tier: T3
size: M
depends_on: []
files:
  - word_document_server/engine/audit.py
  - word_document_server/tools/v2/audit.py
  - tests/engine/test_audit.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_audit.py -q"
```

### Scope
`audit(pkg) -> list[Finding(kind, severity, locator, message, data)]` : paragraphes à allure de titre en formatage direct ; formatage direct qui écrase le style (propriétés comptées) ; styles personnalisés inutilisés ; noms de styles quasi identiques (normalisation casse, espaces, ponctuation) ; `numId` ou `abstractNumId` pendants ; signets et plages de commentaire sans début ou sans fin ; commentaires sans ancre ; révisions par auteur ; champs présents (TOC, REF) avec indication `updateFields` ; suites de paragraphes vides ; polices mixtes dans un paragraphe.
Lecture seule, aucune normalisation. `tools/v2/audit.py` : `doc_audit(filename)`.

### Context
`engine/effective.py`, `engine/styles.py`, `engine/numbering.py`, `engine/locators.py`.

## J06-P2 — Comparaison sémantique

```yaml
id: J06-P2
kind: implement
tier: T4
size: M
depends_on: []
files:
  - word_document_server/engine/compare.py
  - word_document_server/tools/v2/compare.py
  - tests/support/snapshot.py
  - tests/engine/test_compare.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_compare.py tests/support/test_snapshot.py -q"
```

### Scope
Déplacer `snapshot` et `diff` de `tests/support/snapshot.py` vers `engine/compare.py` ; `tests/support/snapshot.py` ne fait plus que réexporter.
`diff` : parties modifiées, paragraphes ajoutés, supprimés, modifiés (texte, style, formatage des runs), tableaux, deltas de commentaires, révisions, signets. `tools/v2/compare.py` : `doc_compare(filename_a, filename_b)`.

### Context
`tests/support/snapshot.py` (J01-P4), `tests/support/test_snapshot.py`.

## J06-P3 — Scénarios de bout en bout et documentation

```yaml
id: J06-P3
kind: implement
tier: T4
size: M
depends_on: [J06-P1, J06-P2, J05-P8]
files:
  - tests/e2e/test_agent_flows.py
  - README.md
  - CHANGELOG.md
  - CONTRIBUTING.md
  - TOOLS.md
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/e2e -q --timeout=120"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python scripts/gen_tools_md.py --check"
  - "grep -q doc_edit_text README.md && grep -q doc_audit CHANGELOG.md"
```

### Scope
`test_agent_flows.py` sur `combined` et les fixtures LibreOffice : `doc_inspect` → `doc_find` → `doc_edit_text` en révision → commentaire → style de caractère → `doc_apply_list` → `doc_apply_table_style` → `doc_audit` → `doc_compare` avec l'original (seuls les changements voulus) → `validate_package` vide → ouverture LibreOffice (skip si absent) ; un scénario par refus (chevauchement de champ, ancre périmée, recherche ambiguë, révision non supportée).
README (réécrit par J05-P7, D-025 ; marqueurs de compteurs et trailer `<!-- mcp-name: … -->` conservés) : y ajouter les outils de J05 et J06 (styles, format effectif, numérotation, styles de tableau, audit, comparaison) et les scénarios de bout en bout, compteurs régénérés, table des modes toujours d'après `PLATFORMS` ; CHANGELOG `[Unreleased]` couvrant J01 à J06 — les puces de J03 (J03-P11), J04 (J04-P7) et J05 (J05-P8) restent telles quelles, seules celles de J06 s'ajoutent (d'où `J05-P8` dans `depends_on` : mêmes `files`) ; CONTRIBUTING : conventions moteur (pas de mutation via `paragraph.runs`, jamais de `Document()` de reconstruction, sauvegardes atomiques, erreurs typées, enregistrement `tools/v2`) ; TOOLS.md régénéré.

### Context
`scripts/gen_tools_md.py`, `tests/support/libreoffice.py`, `docs/audit/destructive-ops.md` (à mettre à jour : outils corrigés).

## J06-P4 — Review J06

```yaml
id: J06-P4
kind: review
tier: T2
size: S
depends_on: [J06-P1, J06-P2, J06-P3, J05-P8, J05-P9]
files: []
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/ -q --timeout=120"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python scripts/gen_tools_md.py --check"
  - "grep -q doc_edit_text README.md && grep -q doc_audit CHANGELOG.md"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv sync && PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/ -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv build"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run ruff check ."
```

### Scope
Review the merged milestone diff with verify-before-done, code-review, test-design. Report; change nothing.
Relit aussi J05-P9 (corrective D-032, hors du périmètre de J05-P6 close — précédent J02-P10 relue par J03-P8) : son diff et ses acceptations rejouées ; `ensure_heading_style` au mieux-effort sur une feuille sans `Normal` (reproduction de la revue J05-P6 rejouée : `add_heading` niveaux 1 et 7 sur `simple` privé de `Normal` et `Heading7-9`), `core.create_style` toujours strict sur un `base_style` inconnu, `_extra_toggles` par `styles._toggle_value` sans `_OFF` local.
Relit aussi J05-P8 (D-027) : les cinq puces ajoutées à `[Unreleased]` contre le code mergé (`utils/document_utils.py`, `core/styles.py` tel que corrigé par J05-P9, `tools/format_tools.py`, `tools/content_tools.py`, `tools/document_tools.py`), préambule et historique intacts (hashes de la part), consigne CONTRIBUTING conforme à `scripts/gen_tools_md.py`.
Les deux sont dans `depends_on` : S toutes deux, J05-P9 sans dépendance et J05-P8 derrière elle seule, elles sont fusionnées bien avant que J06-P3 (deux parts M en amont) ne soit prête — aucun blocage — et sans cela la revue pourrait partir avant leur merge, et personne ne les relirait.

## J06-P5 — Correctif D-033/D-034/D-035 : enveloppe `package_error` de `doc_compare`, styles référencés depuis les commentaires, gardes inertes, inventaire avant/après

```yaml
id: J06-P5
kind: implement
tier: T4
size: M
depends_on: []
files:
  - word_document_server/tools/v2/compare.py
  - word_document_server/engine/audit.py
  - word_document_server/tools/v2/audit.py
  - tests/engine/test_compare.py
  - tests/engine/test_audit.py
  - tests/e2e/test_agent_flows.py
  - docs/audit/destructive-ops.md
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_compare.py tests/engine/test_audit.py tests/e2e -q --timeout=120"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"import asyncio, pathlib, tempfile; from tests.fixtures.builders import build; from word_document_server.tools.v2.registry import v2_tools; d = pathlib.Path(tempfile.mkdtemp()); bad = d / 'not-a-zip.docx'; bad.write_bytes(b'this is not a zip'); cut = d / 'truncated.docx'; cut.write_bytes(build('simple')[:1024]); good = d / 'simple.docx'; good.write_bytes(build('simple')); compare = v2_tools()['doc_compare']; reports = [asyncio.run(compare(str(a), str(b))) for a, b in ((bad, good), (good, bad), (cut, good))]; assert all(r['status'] == 'error' and r['code'] == 'package_error' for r in reports), reports; assert asyncio.run(compare(str(good), str(good)))['identical'] is True\""
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"from lxml import etree; from tests.fixtures.builders import build; from word_document_server.engine.audit import audit; from word_document_server.engine.package import DocxPackage; from word_document_server.engine.styles import styles_root; from word_document_server.engine.xmlns import qn; pkg = DocxPackage.open(build('comments')); root = styles_root(pkg); [etree.SubElement(etree.SubElement(root, qn('w:style'), {qn('w:type'): 'paragraph', qn('w:customStyle'): '1', qn('w:styleId'): sid}), qn('w:name'), {qn('w:val'): sid}) for sid in ('OnlyInComment', 'NowhereAtAll')]; comment = pkg.root_of(pkg.part('/word/comments.xml')).find(qn('w:comment')); etree.SubElement(etree.SubElement(etree.SubElement(comment, qn('w:p')), qn('w:pPr')), qn('w:pStyle'), {qn('w:val'): 'OnlyInComment'}); unused = {f.data['style_id'] for f in audit(pkg) if f.kind == 'unused_custom_style'}; assert 'NowhereAtAll' in unused and 'OnlyInComment' not in unused, unused\""
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"from word_document_server.engine.audit import FINDING_KINDS, audit; from word_document_server.tools.v2.audit import doc_audit; assert len(FINDING_KINDS) == 15, FINDING_KINDS; assert 'fifteen finding kinds across twelve checks' in audit.__doc__.lower(), audit.__doc__[:120]; assert 'fifteen finding kinds across twelve checks' in doc_audit.__doc__.lower(), doc_audit.__doc__[:120]; missing = [k for k in FINDING_KINDS if k not in doc_audit.__doc__]; assert not missing, missing\""
  - "! grep -rnE '(difference|delta|diff[(].*[)])[.]is_empty([^(]|$)' tests/ word_document_server/"
  - "grep -qF 'word/theme/theme1.xml' tests/e2e/test_agent_flows.py && ! grep -qF 'word/theme1.xml' tests/e2e/test_agent_flows.py && PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"import io, zipfile; from tests.fixtures.builders import build; names = set(zipfile.ZipFile(io.BytesIO(build('combined'))).namelist()); assert {'word/styles.xml', 'word/theme/theme1.xml', 'word/settings.xml'} <= names, sorted(names)\""
  - "grep -q '^def test_a_file_that_is_not_a_package_is_reported_by_doc_compare' tests/engine/test_compare.py && grep -q '^def test_a_style_applied_only_in_a_comment_is_not_unused' tests/engine/test_audit.py"
  - "for t in test_format_text_keeps_untouched_run_formatting test_add_table_of_contents_drops_ancillary_parts test_merge_documents_drops_ancillary_parts test_add_header_footer_clears_existing_header_content test_create_custom_style_never_creates_the_style test_replace_block_between_manual_anchors_never_finds_the_anchor test_track_replace_terminates_when_replacement_contains_original test_accept_tracked_changes_merges_a_deleted_paragraph_mark test_reject_tracked_changes_merges_an_inserted_paragraph_mark test_insert_numbered_list_near_text_no_longer_points_at_numid_1_or_2 test_protect_then_unprotect_roundtrips_bytes; do grep -q \"$t\" docs/audit/destructive-ops.md && grep -rq \"^def $t\" tests/ || exit 1; done"
  - "grep -q add_table_of_contents docs/audit/destructive-ops.md && grep -q 'tests/characterization/test_existing_tools.py' docs/audit/destructive-ops.md && grep -q 'J03-P2' docs/audit/destructive-ops.md && grep -q 'J03-P4' docs/audit/destructive-ops.md && grep -q 'J03-P5' docs/audit/destructive-ops.md && grep -q 'J03-P6' docs/audit/destructive-ops.md && grep -q 'J05-P3' docs/audit/destructive-ops.md && grep -q 'J05-P4' docs/audit/destructive-ops.md && ! grep -q 'pour le jalon J03' docs/audit/destructive-ops.md"
  - "! grep -qE 'mark[.]xfail|pytest[.]xfail[(]' tests/characterization/test_existing_tools.py"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/ -q --timeout=120"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python scripts/gen_tools_md.py --check"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run ruff check ."
  - "PATH=\"$HOME/.local/bin:$PATH\" uv build"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv lock --check"
```

### Scope
Part corrective sans revue derrière elle : J06-P4 est close et dernière du plan, personne ne relira ce diff. Ses acceptations sont sa seule garantie — toutes se rejouent avant le rapport, aucune ne se retire ni ne s'affaiblit. Rien des parts ✅ n'est réécrit ; les fichiers de test existants s'étendent, ils ne se réorganisent pas.

D-035 (revue J06-P4, quatre findings, tous reproduits par exécution le 2026-09-15) :
1. `tools/v2/compare.py` : `engine_snapshot` lit le zip avant tout `DocxPackage.open`, et `zipfile.BadZipFile` n'est ni dans `_REPORTED` ni une `OSError` : un `.docx` tronqué ou renommé lève hors de l'outil, là où `doc_inspect`/`doc_audit` répondent `{"status": "error", "code": "package_error"}` et où le docstring « Paths » de `tools/v2/registry.py` promet « a path that is not a package fails saying so ». Ouvrir `DocxPackage.open(filename_a)` puis `DocxPackage.open(filename_b)` **avant** les deux `engine_snapshot` et réutiliser ces deux paquets pour `_v2_positions` (une seule ouverture par fichier) — ou traduire `BadZipFile` en `PackageError` dans l'outil ; jamais élargir `_REPORTED`. Test `test_a_file_that_is_not_a_package_is_reported_by_doc_compare` dans `tests/engine/test_compare.py` (foyer existant des tests de `doc_compare`, section « registration »), par `v2_tools()['doc_compare']`, sur un fichier non-zip et un `simple` tronqué, le mauvais fichier en première puis en seconde position ; modèle `test_a_file_that_is_not_a_package_is_reported` (`tests/tools/test_v2_text.py`).
2. `tests/e2e/test_agent_flows.py` : `untouched_parts` nomme `word/theme1.xml`, qui n'est pas une partie du paquet (`word/theme/theme1.xml` l'est) ; ce tiers de la garde est inerte et une réécriture du thème à la sauvegarde — la classe de régression D-007 — le traverserait. Corriger le nom et rendre la garde auto-vérifiante : le test échoue si un nom de `untouched_parts` n'est pas une partie du paquet (noms lus par `zipfile` sur `path`). Le test reste vert, la partie n'étant pas touchée aujourd'hui.
3. `engine/audit.py`, `_referenced_styles` : n'itère que `state.places` (`pkg.stories()`, commentaires exclus) ; un style personnalisé appliqué seulement dans `word/comments.xml` est rapporté `unused_custom_style`, et un agent qui agit sur le finding le supprime. Ajouter la racine des commentaires (`_comments_root(state.pkg)`, déjà présente) au balayage de `w:pStyle`/`w:rStyle`/`w:tblStyle`. Test `test_a_style_applied_only_in_a_comment_is_not_unused` dans `tests/engine/test_audit.py` : sur `comments`, deux styles `customStyle` ajoutés à `styles.xml`, l'un appliqué dans un paragraphe ajouté au premier `w:comment`, l'autre nulle part — le premier absent des `unused_custom_style`, le second présent (un test qui ne prouve que l'absence serait vacueux).
4. Garde grep D-007 : depuis J06-P2, `Diff` vit dans `word_document_server/engine/compare.py` et son porteur de production est `delta.is_empty()` (`tools/v2/compare.py`) ; l'acceptation de J02-P8 (✅, `jalons/J02-coeur-ooxml.md`, intouchée) ne balaie que `tests/`. Cette part porte la garde élargie aux deux arbres, en acceptation ; elle passe sur `main` au 2026-09-15, rien à changer dans le code pour elle.

D-034 : `engine/audit.py` (`audit()`, « Twelve checks run, in the order of FINDING_KINDS ») et `tools/v2/audit.py` (`doc_audit`, « The twelve checks: ») comptent douze là où `len(FINDING_KINDS) == 15` (signets et plages de commentaire donnent chacun un genre `without_start` et un `without_end`). Réécrire les deux phrases en « fifteen finding kinds across twelve checks » ; la liste à puces de `doc_audit` nomme déjà les 15 genres, la garder complète. `TOOLS.md` ne reprend pas ces phrases (descriptions inchangées) ; `gen_tools_md.py --check` en acceptation le confirme.

D-033 : réécrire `docs/audit/destructive-ops.md` en inventaire « état avant le plan → état après », en français, sans historique ni changelog. En tête : photo de l'état initial prise par J01-P5 le 2026-09-09 ; l'inventaire vivant est `tests/characterization/test_existing_tools.py`, qui n'a plus aucun `xfail`. Tableau à quatre colonnes — outil ; symptôme avant (reprendre le texte actuel, il est exact) ; correctif : part et ce que fait le code désormais ; test témoin, sous son nom actuel — une ligne par entrée actuelle : `format_text` (J03-P4 ; `test_format_text_keeps_untouched_run_formatting`, `tests/tools/test_format_text.py`), `add_table_of_contents` (J03-P5 ; `test_add_table_of_contents_drops_ancillary_parts`, nom conservé et assertion inversée, `tests/tools/test_toc.py`), `merge_documents` (J03-P6, D-015 ; `test_merge_documents_drops_ancillary_parts`, `tests/tools/test_merge.py`), `add_header_footer` (J03-P5 ; `test_add_header_footer_clears_existing_header_content`), `create_custom_style` (J05-P3, D-027 ; `test_create_custom_style_never_creates_the_style`, `tests/engine/test_styles_write.py`), `replace_block_between_manual_anchors` (J03-P5, D-015 ; `test_replace_block_between_manual_anchors_never_finds_the_anchor`, `tests/tools/test_layout_blocks.py`), `track_replace` (J03-P2 ; `test_track_replace_terminates_when_replacement_contains_original`), marque de paragraphe d'`accept`/`reject_tracked_changes` (J03-P2, `engine/revisions.py` ; `test_accept_tracked_changes_merges_a_deleted_paragraph_mark`, `test_reject_tracked_changes_merges_an_inserted_paragraph_mark`). La colonne « après » se lit dans les commentaires `# Was xfail(strict=True) until …` du fichier de tests et dans `CHANGELOG.md` `[Unreleased]`, jamais de mémoire. « Notes complémentaires » traitées de même : `numId` 1/2 en dur — réglé par J05-P4 (D-027 : définition de liste neuve à chaque appel), témoin `test_insert_numbered_list_near_text_no_longer_points_at_numid_1_or_2` (`tests/engine/test_numbering.py`) ; marqueurs `w:rPrChange`/`w:pPrChange` jamais retirés par `accept`/`reject` — toujours vrai, mesuré le 2026-09-15 sur `tracked_changes` et `combined` (compte inchangé après chacun des deux outils), à dater ainsi, hors périmètre de cette part ; `protect`/`unprotect` — inchangé, témoin `test_protect_then_unprotect_roundtrips_bytes`. `tests/characterization/test_existing_tools.py` n'est pas dans `files` : on le cite, on ne le renomme pas.

### Context
`tools/v2/registry.py` (`_REPORTED`, `error_code`, docstring « Paths »), `engine/package.py` (`DocxPackage.open` traduit tout échec de lecture en `PackageError`), `engine/audit.py` (`_comments_root`, `_referenced_styles`, `FINDING_KINDS`), `engine/compare.py` (`snapshot` lit par `zipfile`), `tests/tools/test_v2_text.py` (tests d'enveloppe), `tests/engine/test_audit.py` (`_pkg`, `_append_styles`, `_of_kind`), `CHANGELOG.md` `[Unreleased]`, PLAN.md Risks (D-007, D-033, D-035). Lancée en s9, où J06-P1..P3 ont été fusionnées : worktree créé à la main sur `main` et dispatch sans `isolation` (Convention Worktrees) ; dès s10, `isolation: "worktree"` convient.
