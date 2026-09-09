# LONG_TERM_RECO — v2-semantic-engine

One line per item. Remove an item when done and note it in DECISIONS_LOG.md.

- R-001 (D-001, 2026-09-09): raccorder les locators V2, `dry_run` et le rapport structuré aux outils `word_live_*` (COM et JXA) ; when: session sur Windows ou macOS avec Word installé ; est. T3/M
- R-002 (D-001, 2026-09-09): déplacer la branche macOS des outils live à l'intérieur du `try` (erreurs JSON comme sous Windows), ajouter un timeout aux appels COM, unifier les messages « only available on Windows » ; when: même condition que R-001 ; est. T4/S
- R-003 (D-001, 2026-09-09): unifier les 10 outils de notes de bas de page sur le moteur (`ensure_part`, ids, `w:sz` dans `rPr`, `XML_NS` défini avant usage) ; when: après J03 ; est. T3/M
- R-004 (D-001, 2026-09-09): supprimer `add_restricted_editing`, `add_digital_signature`, `verify_document` et le sidecar `.protection` (protection en nom seulement, signature auto-invalidée) ou les remplacer par `w:documentProtection` ; when: décision produit ; est. T4/S
- R-005 (D-001, 2026-09-09): le nom PyPI `word-mcp-live` appartient à l'upstream (`ykarapazar`, dernier push 2026-05-29, fork à jour) : choisir un nom de fork, ou proposer J01–J06 en PR upstream ; `publish.yml` hérité ne se déclenche que sur release ; when: avant toute release ; est. T5/S
- R-006 (D-001, 2026-09-09): `merge_documents` avec notes et commentaires (import des parties avec renumérotation) ; when: besoin réel après J03-P6 ; est. T2/M
- R-007 (D-001, 2026-09-09): création de styles de tableau (formats conditionnels) et outil `normalize_styles` séparé de l'audit ; when: retours d'usage de `doc_audit` ; est. T3/M
- R-008 (D-001, 2026-09-09): enregistrement automatique des 120 wrappers de `main.py` depuis les signatures des implémentations (réconcilier les 104 descriptions divergentes) ; when: après J04 ; est. T4/M
- R-009 (D-001, 2026-09-09): image Docker avec LibreOffice pour `convert_to_pdf` ; `docx2pdf` en extra optionnel ; `manifest.json`, `server.json`, `__version__` alignés sur `pyproject.toml` ; when: avant toute release ; est. T5/S
- R-010 (D-001, 2026-09-09): évaluer python-docx ≥ 1.2 (partie commentaires native) et pinner `python-docx>=1.1.2,<2` d'ici là ; when: après J03 ; est. T4/S
- R-011 (D-001, 2026-09-09): `word_live_apply_list` et `word_live_setup_heading_numbering` divergent entre COM et JXA (`outline_numbered` ignoré côté COM, `heading_map` macOS seulement) : réaligner ou documenter la surface macOS ; when: R-001 ; est. T3/M
