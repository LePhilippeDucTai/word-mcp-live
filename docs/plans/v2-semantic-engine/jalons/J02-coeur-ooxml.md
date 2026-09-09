# J02 — Cœur OOXML : paquet, flux de texte, plages, révisions

Goal: Un paquet `word_document_server/engine/` pur (OPC de python-docx + lxml, synchrone, sans MCP ni Word) : ouverture et sauvegarde atomique, flux de texte d'un paragraphe, plages sûres multi-runs, formatage direct, recherche, révisions ; invariants vérifiés par tests de propriété sur les fixtures.
Depends on: J01 · Orchestrator: opus/high

## J02-P1 — Paquet, ids, espaces de noms, erreurs

```yaml
id: J02-P1
kind: implement
tier: T3
size: M
depends_on: []
files:
  - word_document_server/engine/__init__.py
  - word_document_server/engine/package.py
  - word_document_server/engine/ids.py
  - word_document_server/engine/xmlns.py
  - word_document_server/engine/errors.py
  - tests/engine/test_package.py
  - tests/engine/test_ids.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_package.py tests/engine/test_ids.py -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"import sys, word_document_server.engine.package, word_document_server.engine.ids; assert 'fastmcp' not in sys.modules and 'win32com' not in sys.modules and 'zipfile' not in sys.modules\""
```

### Scope
`DocxPackage.open(path | bytes)` ; `.document` (racine lxml de `word/document.xml`) ; `.part(partname)` ; `.stories()` → `(name, root)` pour body, en-têtes, pieds, notes de bas de page, notes de fin ; `.ensure_part(partname, content_type, reltype, initial_xml) -> root` idempotent via `docx.opc.part.Part` + `relate_to` ; `.rel_target(root_part, rId)` ; `.add_external_rel(part, url) -> rId` avec dédoublonnage par cible ; `.save(path)` atomique (temporaire dans le même dossier, `fsync`, `os.replace`, temporaire supprimé en cas d'erreur) ; `.to_bytes()` ; `atomic_write_bytes(path, data)`.
`ids.py` : `next_annotation_id(pkg)` = max des `w:id` de ins/del/moveFrom/moveTo/bookmarkStart/commentRangeStart/commentReference sur toutes les stories + 1 ; `next_comment_id`, `next_footnote_id`, `next_para_id` (w14, unique).
`xmlns.py` : constantes W, R, W14, W15, MC, XML, `qn()`. `errors.py` : `EngineError` → `PackageError`, `LocatorError(code)`, `UnsupportedRange(reason)`, `InvalidText`, `UnsupportedRevision`, `IdExhausted`.
Aucun `zipfile` ni réécriture manuelle de `[Content_Types].xml` ou des rels. Tests : chaque fixture ouverte puis sauvée sans modification → instantané inchangé et `validate_package` vide ; `ensure_part` idempotent ; échec pendant `save` → fichier d'origine intact, aucun temporaire.

### Context
`tests/support/snapshot.py`, `tests/support/package_check.py`, `tests/fixtures/builders.py`. Références utiles : `core/footnotes.py:141-203` (idempotence), `:451-473` (temp + `os.replace`).

## J02-P2 — Flux de texte d'un paragraphe

```yaml
id: J02-P2
kind: implement
tier: T2
size: M
depends_on: [J02-P1]
files:
  - word_document_server/engine/textmodel.py
  - tests/engine/test_textmodel.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_textmodel.py -q"
```

### Scope
`segments(paragraph) -> list[Segment(kind, element, containers, start, end, text)]`, `kind` ∈ `text | tab | break | opaque | marker | hidden`.
Texte : `w:t`, `w:tab`→`\t`, `w:br`/`w:cr`→`\n`, `w:sym`, `w:noBreakHyphen`→`‑`, `w:softHyphen`→`­`. Opaques de largeur nulle : `w:drawing`, `w:pict`, `w:object`, `w:footnoteReference`, `w:endnoteReference`, `w:commentReference`, `w:fldChar`, `w:instrText`, `w:delInstrText`. Cachés : `w:delText` et tout contenu sous `w:del`/`w:moveFrom`. Marqueurs : `bookmarkStart/End`, `commentRangeStart/End`, `permStart/End`, `proofErr`, `moveFrom/ToRangeStart/End`. Conteneurs traversés : `w:hyperlink`, `w:ins`, `w:moveTo`, `w:sdt/w:sdtContent`, `w:smartTag`, `w:fldSimple` (résultat visible), `w:customXml`, `w:dir`, `w:bdo`.
Champs complexes : machine à états begin/separate/end (imbrication) → `fields(paragraph) -> list[Field(instr, start, end, atomic=True)]`.
`visible_text(paragraph)`, `position(paragraph, offset) -> (segment, local_offset)`. Tests sur chaque fixture : texte attendu par paragraphe (table dans le test), segments contigus et couvrants, champs et instructions détectés, texte supprimé absent, texte inséré présent.

### Context
`utils/document_utils.py:13-45` (`get_effective_text`) est la lecture actuelle ; python-docx `paragraph.text` ignore `w:ins` (vérifié). Politique de texte fixée dans PLAN.md, Architecture.

## J02-P3 — Plages : découpe, suppression, insertion, remplacement

```yaml
id: J02-P3
kind: implement
tier: T2
size: M
depends_on: [J02-P2]
files:
  - word_document_server/engine/ranges.py
  - tests/engine/test_ranges.py
  - tests/engine/test_ranges_properties.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_ranges.py tests/engine/test_ranges_properties.py -q"
```

### Scope
`split_run(run, offset) -> (left, right)` : clone du `rPr`, `xml:space="preserve"`, ordre des enfants conservé, idempotent aux bornes.
`resolve(paragraph, start, end) -> Pieces` : runs entiers couverts après découpe + marqueurs intérieurs ; `UnsupportedRange` si la plage chevauche partiellement un champ, touche du contenu caché, ou contient un opaque alors que l'opération est `delete`/`replace` ; plage vide autorisée pour `insert`.
`delete_range(paragraph, start, end)`, `insert_text(paragraph, offset, text, rpr_from="left" | "right" | element)`, `replace_range(paragraph, start, end, text)` ; marqueurs d'une plage supprimée conservés à sa position de début ; runs vides retirés ; `\n` → `w:br`, `\t` → `w:tab` ; caractères de contrôle interdits → `InvalidText` (même politique que `utils/text_safety.py`).
Tests de propriété : plages aléatoires sur toutes les fixtures → texte et `rPr` hors plage inchangés (instantané), nombre de marqueurs inchangé, `validate_package` vide, `visible_text` conforme ; refus explicites testés.

### Context
`core/comment_writer.py:283-325` est la seule découpe correcte existante ; ne pas reproduire `parent = first_run.getparent()` + index capturé avant mutation (`core/tracked_changes.py:238-269`).

## J02-P4 — Formatage direct et enveloppes

```yaml
id: J02-P4
kind: implement
tier: T3
size: S
depends_on: [J02-P3]
files:
  - word_document_server/engine/format.py
  - tests/engine/test_format.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_format.py -q"
```

### Scope
`apply_rpr(pieces, patch)` : `patch` = dict `bold, italic, underline, strike, caps, small_caps, superscript, subscript, size_pt, font, color, highlight, char_style` ; valeur `None` retire la propriété ; enfants de `w:rPr` insérés dans l'ordre du schéma ; `font` explicite retire `asciiTheme`/`hAnsiTheme` et pose `ascii`/`hAnsi`/`cs` ; `color` explicite retire `themeColor`/`themeTint`/`themeShade` ; `char_style` vérifie l'existence du style (`styles.xml`).
`wrap(pieces, factory)` : runs contigus frères requis sinon `UnsupportedRange` ; fabriques `hyperlink(rId)` et `comment(id)` (`commentRangeStart`, `commentRangeEnd`, run `commentReference` avec `rStyle CommentReference` si présent) ; `unwrap(element)` conserve les runs.
`set_ppr_child(paragraph, tag, attrs)` en ordre du schéma `CT_PPr`. Tests : sous-plage formatée sans toucher les autres runs, propriétés de thème, aller-retour wrap/unwrap, ordre des enfants.

### Context
Ordre `CT_RPr` et `CT_PPr` : ECMA-376 §17.3.2.28 et §17.3.1.26. Défaut existant : `utils/document_utils.py:440` (`pPr.append(numPr)`).

## J02-P5 — Recherche

```yaml
id: J02-P5
kind: implement
tier: T4
size: S
depends_on: [J02-P2]
files:
  - word_document_server/engine/find.py
  - tests/engine/test_find.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_find.py -q"
```

### Scope
`find(pkg, pattern, *, regex=False, case=True, whole_word=False, stories=("body",), max_results=None) -> list[Match(story, paragraph, index, start, end, text, context)]` sur le texte visible, paragraphe par paragraphe (corps y compris tableaux et sdt, en-têtes, pieds, notes) ; `index` = index V2 du paragraphe dans sa story ; pas de correspondance inter-paragraphes ; ordre déterministe ; motif vide ou regex à correspondance vide → `ValueError`.
`iter_paragraphs(story_root)` en ordre de document (cellules et sdt inclus) partagé avec J04-P1.
Tests : correspondance à cheval sur des runs, texte sous `w:del` ignoré, texte sous hyperlien et `w:ins` trouvé, `whole_word`, regex, stories.

### Context
`engine/textmodel.py`. Le comportement actuel à remplacer : `utils/document_utils.py:246-285`.

## J02-P6 — Révisions

```yaml
id: J02-P6
kind: implement
tier: T2
size: M
depends_on: [J02-P3]
files:
  - word_document_server/engine/revisions.py
  - tests/engine/test_revisions.py
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine/test_revisions.py -q --timeout=60"
```

### Scope
`tracked_delete(pieces, author, date)` : chaque run enveloppé dans `w:del` (`w:t` → `w:delText`) ; sous un `w:ins` d'un autre auteur → `w:ins/w:del` ; sous un `w:ins` du même auteur → suppression simple. `tracked_insert(paragraph, offset, text, author, date)` : run cloné dans `w:ins`. `tracked_replace` = suppression puis insertion adjacente. Ids via `ids.py`, date ISO UTC.
`list_revisions(pkg) -> list[Revision(id, kind, author, date, story, paragraph_index, text)]` pour ins, del, moveFrom, moveTo, rPrChange, pPrChange, marque de paragraphe.
`accept(pkg, ids=None, author=None)` et `reject(...)` pour ins, del et marque de paragraphe supprimée (fusion des paragraphes à l'acceptation) sur toutes les stories ; autres genres → `UnsupportedRevision` listant les ids.
Tests : `new_text ⊇ old_text` termine ; `rPr` de chaque morceau conservé ; imbrication ins>del ; accepter puis rejeter restaure l'instantané ; stories multiples.

### Context
Défauts à ne pas reproduire : `core/tracked_changes.py:200-271` (boucle infinie reproduite le 2026-09-09), `:518-660` (accept/reject partiel). ECMA-376 §17.13.5.

## J02-P7 — Review J02

```yaml
id: J02-P7
kind: review
tier: T2
size: S
depends_on: [J02-P1, J02-P2, J02-P3, J02-P4, J02-P5, J02-P6]
files: []
acceptance:
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/engine -q --timeout=60"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run python -c \"import sys, word_document_server.engine.package, word_document_server.engine.ids; assert 'fastmcp' not in sys.modules and 'win32com' not in sys.modules and 'zipfile' not in sys.modules\""
  - "PATH=\"$HOME/.local/bin:$PATH\" uv sync && PATH=\"$HOME/.local/bin:$PATH\" uv run pytest tests/ -q"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv build"
  - "PATH=\"$HOME/.local/bin:$PATH\" uv run ruff check ."
```

### Scope
Review the merged milestone diff with verify-before-done, code-review, test-design. Report; change nothing.
