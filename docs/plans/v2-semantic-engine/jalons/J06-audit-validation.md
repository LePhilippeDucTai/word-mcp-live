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
depends_on: [J06-P1, J06-P2]
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
README : table des modes honnête (Linux, Windows, macOS d'après `PLATFORMS`), section outils V2, compteurs générés ; CHANGELOG `[Unreleased]` couvrant J01 à J06 ; CONTRIBUTING : conventions moteur (pas de mutation via `paragraph.runs`, jamais de `Document()` de reconstruction, sauvegardes atomiques, erreurs typées, enregistrement `tools/v2`) ; TOOLS.md régénéré.

### Context
`scripts/gen_tools_md.py`, `tests/support/libreoffice.py`, `docs/audit/destructive-ops.md` (à mettre à jour : outils corrigés).

## J06-P4 — Review J06

```yaml
id: J06-P4
kind: review
tier: T2
size: S
depends_on: [J06-P1, J06-P2, J06-P3]
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
