# PLAN — v2-semantic-engine

Goal: Faire de word-mcp-live un moteur de manipulation sémantique de documents Word pour agents IA : un cœur OOXML pur (Linux, sans Word) qui garantit la non-dégradation (styles, relations, champs, signets, commentaires, révisions, numérotation, mise en page), les outils existants migrés dessus sans rupture d'API, une surface V2 `doc_*` (adressage sans état, dry-run, styles/thème, audit, comparaison), les outils live COM/JXA préservés.
Base branch: v2-semantic-engine · Remote: origin · Language: fr · Created: 2026-09-09 · Planned with: best/max

## Scope

- In: paquet `word_document_server/engine/` ; harnais de tests (fixtures générées python-docx+lxml et LibreOffice, instantanés canoniques, validateur de paquet, caractérisation) ; migration des outils cross-platform texte, révisions, commentaires, hyperliens, formatage, TOC, fusion, en-têtes ; retrait du monkey-patch de sauvegarde ; écritures atomiques ; outils `doc_*` (inspect, find, edit_text, format_range, apply_edits, capabilities, styles, effective format, list, table style, audit, compare) ; TOOLS.md et compteurs README générés ; garde-fous de dérive ; trois plantages macOS prouvés par lecture.
- Out: tout autre changement des modules COM/JXA ; outils de notes de bas de page (conservés tels quels) ; protection, signature, sidecar `.protection` ; création de styles de tableau ; normalisation automatique ; publication PyPI et versionnage ; CI hébergée (validation locale uniquement).

## Architecture

- `engine/` : Python synchrone, lxml sur l'OPC de python-docx (`Document.part`, `Part`, `relate_to`) ; aucune dépendance MCP ni Word ; exceptions typées (`PackageError`, `LocatorError`, `UnsupportedRange`, `InvalidText`, `UnsupportedRevision`).
- Couches : `package` (parties, stories, ids d'annotation, sauvegarde atomique temp+`os.replace`) → `textmodel` (flux de texte d'un paragraphe : segments texte, opaques, marqueurs, cachés ; champs complexes atomiques) → `ranges` (split, delete, insert, replace) → `format` (rPr en ordre de schéma, enveloppes hyperlink/commentaire) → `find`, `revisions`, `locators`, `inspect`, `styles`/`theme`/`effective`, `numbering`, `table_styles`, `merge`, `audit`, `compare`.
- Politique de texte : visible = `w:t` sous `w:r`, `w:hyperlink`, `w:ins`, `w:sdt`, `w:smartTag`, `w:fldSimple` ; `w:tab`→`\t`, `w:br`/`w:cr`→`\n` ; `w:del` et `w:instrText` cachés ; `drawing`, `footnoteReference`, `fldChar` opaques ; signets et plages de commentaire = marqueurs jamais supprimés.
- Outils existants : noms, paramètres et formats de retour inchangés ; leurs implémentations délèguent au moteur (J03). `main.py` n'est touché que pour enregistrer `tools/v2` et corriger le registre.
- Outils V2 : `tools/v2/*.py` découverts dynamiquement, schéma dérivé de la signature, retour dict `{status, dry_run, changes, warnings}` ou `{status: "error", code, message}`.
- Adressage V2 : locators sans état — `{"paragraph": i, "expect_text"}` (i = `w:p` en ordre du corps, sdt de bloc inclus, cellules exclues, base 0), `{"find", "occurrence", "within"}`, `{"bookmark"}`, `{"heading"}`, `{"table", "row", "col", "paragraph"}`, `{"story"}` ; aucun identifiant de session.
- Pas d'abstraction backend : COM/JXA restent une famille d'outils distincte ; `tools/platforms.py` porte la disponibilité par plateforme et alimente `TOOLS.md`, README et `doc_capabilities`.
- Tests : fixtures générées et déterministes (jamais de binaires commités) ; instantané = c14n par partie + signature des runs par paragraphe ; LibreOffice = producteur de fixtures et validateur d'ouverture optionnel, jamais source de vérité.

## Risks

- Invariants moteur : texte et rPr hors plage inchangés ; marqueurs conservés ; champs atomiques ; `w:del` jamais modifié ; parties non ciblées canoniquement identiques, parties binaires identiques octet à octet ; ids d'annotation uniques ; enfants pPr/rPr/tblPr en ordre de schéma ; sauvegarde atomique ; LibreOffice rouvre le fichier.
- Zones sensibles : `core/tracked_changes.py` (boucle infinie, splice multi-parents), `format_text`, `add_table_of_contents`, `merge_documents`, `add_header_footer`, `utils/save_utils.py`, numérotation (`numId` 1/2 en dur).
- Compatibilité : 120 noms d'outils, signatures et formats de retour inchangés ; live COM/JXA intact hors J04-P4 ; nouveaux paramètres optionnels seulement.
- Portabilité : LibreOffice optionnel (`skip`), jamais requis par les tests unitaires ; aucun test ne dépend de Word ; `uv` hors PATH des shells (préfixer `PATH="$HOME/.local/bin:$PATH"`).
- Espaces d'index : trois espaces de `paragraph_index` coexistent (python-docx corps, `//w:p`, `body//w:p`) ; la migration J03 aligne sur l'espace V2 sans renommer les paramètres.
- Zones aveugles du harnais laissées telles quelles (D-004) : `TableSignature` ne retient du `w:tblGrid` que le nombre de `w:gridCol` — des largeurs `w:gridCol/@w:w` réécrites passent `paragraphs=[...]` (mesuré sur `combined` ; `w:tcW` reste couvert par `cell_properties`) ; `XML-MALFORMED` et `CT-PART-MISSING` de `validate_package` sont fonctionnels mais sans test positif. Un test qui touche la grille d'un tableau (J03-P4, J05-P5) ne peut pas s'en remettre à `assert_unchanged_except` pour elle.
- Lint partiel : `[tool.ruff] include` ne couvre que `word_document_server/engine/**`, `tests/**`, `scripts/**` ; un `ruff check .` vert ne dit rien de `word_document_server/tools|core|utils` (95 RUF013 et 9 W605 sur `live_tools.py` seul), que J03 réécrit.

## Checks

- Setup: `PATH="$HOME/.local/bin:$PATH" uv sync` · marker: `.venv/`
- Test: `PATH="$HOME/.local/bin:$PATH" uv run pytest tests/ -q` · Build: `PATH="$HOME/.local/bin:$PATH" uv build` · Lint: `PATH="$HOME/.local/bin:$PATH" uv run ruff check .`
- Baseline 2026-09-09: green — test : 1 passed (3 s, rejoué à la révision D-002) ; lint et build utilisables à partir de J01-P1 (ruff absent avant).
- Lock: `PATH="$HOME/.local/bin:$PATH" uv lock --check` → 0 ; `uv.lock` régénéré le 2026-09-09 (périmé depuis v1.1.3, il se réécrivait à chaque `uv run`) ; un `uv run` ne doit plus modifier l'arbre.

## Milestones

| Id  | Milestone | Parts | Depends on | Orchestrator | File |
|-----|-----------|-------|------------|--------------|------|
| J01 | Harnais de fidélité documentaire | 7 | — | opus/high | jalons/J01-harnais-fidelite.md |
| J02 | Cœur OOXML : paquet, flux de texte, plages, révisions | 7 | J01 | opus/high | jalons/J02-coeur-ooxml.md |
| J03 | Migration des outils existants sur le cœur | 8 | J02 | opus/high | jalons/J03-migration-outils.md |
| J04 | Surface V2 : adressage, inspection, dry-run, capacités, docs | 6 | J03 | opus/high | jalons/J04-surface-v2.md |
| J05 | Styles, thème, format effectif, numérotation, styles de tableau | 6 | J04 | opus/high | jalons/J05-styles-theme-numerotation.md |
| J06 | Audit, comparaison et validation de bout en bout | 4 | J05 | opus/high | jalons/J06-audit-validation.md |

## Tiers

| Tier | Model  | Effort | Use for | Executor |
|------|--------|--------|---------|----------|
| T1 | best   | max    | planning, architecture, cross-cutting decisions, plan updates | orchestrator inline; philippele-skills:planner |
| T2 | opus   | xhigh  | complex refactor, tricky logic, concurrency, migrations, milestone review | philippele-skills:worker-t2 / reviewer |
| T3 | opus   | high   | standard implementation with design latitude | philippele-skills:worker-t3 |
| T4 | sonnet | high   | well-specified implementation, tests | philippele-skills:worker-t4 |
| T5 | sonnet | medium | mechanical changes, docs, renames, config | philippele-skills:worker-t5 |
| T6 | haiku  | low    | trivial edits, search, inventory | philippele-skills:worker-t6 |

## Conventions

- Writers: orchestrator → PROGRESS.md, DECISION.md, DECISIONS_LOG.md, LONG_TERM_RECO.md; planner → PLAN.md, jalons/; workers → source only.
- Every milestone ends with a `kind: review` part run by philippele-skills:reviewer; the reviewer also runs the Checks and reads Risks.
- Wave cap: 4 parts in flight; review, size L, and T1 parts run alone.
- Git: workers commit in their worktree; one merge commit per validated part `<id>: <title>`; push v2-semantic-engine to origin at session end (none when Remote is none); never force, never amend.
- Failed part: try 2 at the same tier with the failure evidence, try 3 one tier higher, then ❌ and a queued decision; a `blocked` report or a first merge conflict does not count.
- Plan updates: only when remaining parts are invalidated or the user asks for a change; planner rewrites files in place; ids never renumbered; a dropped part is ⏭ and its dependents get their Scope rewritten.
- Code, docstrings, descriptions d'outils, TOOLS.md et CHANGELOG en anglais ; documents de plan en français. Aucun workflow GitHub Actions n'est ajouté.
