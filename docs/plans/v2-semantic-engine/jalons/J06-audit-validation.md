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
