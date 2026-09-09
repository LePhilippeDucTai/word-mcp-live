# PROGRESS — v2-semantic-engine

Updated 2026-09-09 · s3 · base v2-semantic-engine · ⬜ todo 🔄 wip ✅ done ❌ failed ⏸ blocked ⏭ dropped

## J01 — Harnais de fidélité documentaire (7/7)

| Part | Title | Tier | Status | Tries | Note |
|------|-------|------|--------|-------|------|
| J01-P1 | Outillage : pytest, groupe dev, ruff, script de vérification, avertissements destructifs | T5 | ✅ | 1 | conflit uv.lock (base de worktree périmée, D-003) : try rendu |
| J01-P2 | Fixtures générées : constructeur OOXML déterministe | T3 | ✅ | 1 | 19 fixtures, `build(name)`, `fixture_docx(name, directory=None)` |
| J01-P3 | Producteur LibreOffice et validateur de paquet | T4 | ✅ | 2 | review fix 1 : notes réservées par `w:type`, 6 tests positifs de corruption |
| J01-P4 | Instantané canonique et comparaison | T3 | ✅ | 2 | review fix 1 : pPr, enfants de run, tblPr/trPr/tcPr, sdtPr ; index `story_root.iter(w:p)` |
| J01-P5 | Caractérisation des outils existants | T4 | ✅ | 2 | review fix 1 : 4 bugs caractérisés (voir docs/audit/destructive-ops.md) |
| J01-P6 | Review J01 | T2 | ✅ | 1 | round 1 failed → 3 corrections ; round 2 done, 3 should-fix restants (D-004) |
| J01-P7 | Zones aveugles du harnais : références de relation, `sectPr` de corps, ids dupliqués | T3 | ✅ | 1 | corrective D-004 ; `REL-REF-UNRESOLVED`, `Snapshot.sect_pr`, 4 tests positifs ; suite 176→183 passed |

## J02 — Cœur OOXML : paquet, flux de texte, plages, révisions (9/9)

| Part | Title | Tier | Status | Tries | Note |
|------|-------|------|--------|-------|------|
| J02-P1 | Paquet, ids, espaces de noms, erreurs | T3 | ✅ | 1 | `DocxPackage`, `ids`, `xmlns`, `errors` ; PartFactory XmlPart (sinon notes/commentaires en blobs) |
| J02-P2 | Flux de texte d'un paragraphe | T2 | ✅ | 1 | `segments`/`visible_text`/`fields`/`position`, scan unique ; opaque par défaut ; 190 tests ; D-005 en attente |
| J02-P3 | Plages : découpe, suppression, insertion, remplacement | T2 | ✅ | 1 | `split_run`/`resolve`/`delete`/`insert`/`replace` ; largeur nulle couverte si strictement intérieure ; 139 tests ; a levé D-007 |
| J02-P4 | Formatage direct et enveloppes | T3 | ✅ | 1 | `apply_rpr(…, pkg=)`, `wrap`/`unwrap`, `set_ppr_child` ; refus typés ; 71 tests |
| J02-P5 | Recherche | T4 | ✅ | 1 | `find`, `iter_paragraphs` ; `"body"` alias de `"document"`, `Match.index int \| None` (D-006) |
| J02-P6 | Révisions | T2 | ✅ | 1 | `tracked_*(pkg, …)`, `list_revisions`, `accept`/`reject` ; genres non gérés refusés, pas ignorés ; 57 tests ; a levé D-008 |
| J02-P8 | Correctif D-007 : aller-retour `open`/`save` neutre, assertions `is_empty()` | T3 | ✅ | 1 | `LIVE_CONTENT_TYPES` ; aller-retour neutre 0/19 → 19/19 ; garde grep `is_empty()` ; suite 775 passed |
| J02-P9 | Marque de paragraphe insérée dans accept/reject | T3 | ✅ | 1 | suivi D-008 ; `_joins_paragraphs(kind, accepting)` partagé par le pré-contrôle et `_apply_one` ; 780 passed, 9 xfailed |
| J02-P7 | Review J02 | T2 | ✅ | 1 | round 1 sans finding bloquant ; 780 passed / 9 xfailed, build et lint verts ; 2 should-fix + 1 optional → D-009 |

## J03 — Migration des outils existants sur le cœur (0/8)

| Part | Title | Tier | Status | Tries | Note |
|------|-------|------|--------|-------|------|
| J03-P1 | search_and_replace et find_text_in_document | T3 | ⬜ | 0 | |
| J03-P2 | Tracked changes sur le moteur | T2 | ⬜ | 0 | |
| J03-P3 | Commentaires et hyperliens sur le moteur | T3 | ⬜ | 0 | |
| J03-P4 | format_text et format_cell_text sans reconstruction | T3 | ⬜ | 0 | |
| J03-P5 | Opérations destructrices : TOC, en-têtes, suppression, signets, blocs | T3 | ⬜ | 0 | |
| J03-P6 | merge_documents par import d'éléments | T2 | ⬜ | 0 | |
| J03-P7 | Retrait du hook de sauvegarde, écritures atomiques, annotations | T4 | ⬜ | 0 | |
| J03-P8 | Review J03 | T2 | ⬜ | 0 | |

## J04 — Surface V2 : adressage, inspection, dry-run, capacités, docs (0/6)

| Part | Title | Tier | Status | Tries | Note |
|------|-------|------|--------|-------|------|
| J04-P1 | Locators et inspection | T3 | ⬜ | 0 | |
| J04-P2 | Outils `doc_*` texte, enregistrement, rapport structuré | T3 | ⬜ | 0 | |
| J04-P3 | Plateformes, TOOLS.md généré, garde-fous de dérive, registre corrigé | T4 | ⬜ | 0 | |
| J04-P4 | Plantages macOS prouvés par lecture | T4 | ⬜ | 0 | |
| J04-P5 | Lot atomique `doc_apply_edits` | T4 | ⬜ | 0 | |
| J04-P6 | Review J04 | T2 | ⬜ | 0 | |

## J05 — Styles, thème, format effectif, numérotation, styles de tableau (0/6)

| Part | Title | Tier | Status | Tries | Note |
|------|-------|------|--------|-------|------|
| J05-P1 | Thème et modèle de style en lecture | T3 | ⬜ | 0 | |
| J05-P2 | Format effectif avec provenance | T3 | ⬜ | 0 | |
| J05-P3 | Écriture de styles paragraphe et caractère | T3 | ⬜ | 0 | |
| J05-P4 | Numérotation | T3 | ⬜ | 0 | |
| J05-P5 | Styles de tableau | T4 | ⬜ | 0 | |
| J05-P6 | Review J05 | T2 | ⬜ | 0 | |

## J06 — Audit, comparaison et validation de bout en bout (0/4)

| Part | Title | Tier | Status | Tries | Note |
|------|-------|------|--------|-------|------|
| J06-P1 | Audit documentaire | T3 | ⬜ | 0 | |
| J06-P2 | Comparaison sémantique | T4 | ⬜ | 0 | |
| J06-P3 | Scénarios de bout en bout et documentation | T4 | ⬜ | 0 | |
| J06-P4 | Review J06 | T2 | ⬜ | 0 | |

## Next session

- Ready: J02-P9 (suivi D-008), puis J02-P7 (revue J02, coupée en s2 : à relancer)
- Orchestrator: opus/high
- Decision: none (D-005, D-006, D-008 tranchées à la clôture de s2)

## Log

- 2026-09-09 s2 (clôture): D-005 (lecture mécanique des champs imbriqués), D-006 (alias "body", index V2 None en cellule ; vocabulaire J03/J04 aligné), D-008 (+J02-P9) tranchées
- 2026-09-09 s2: J01-P7, J02-P1..P6 et J02-P8 ✅ ; D-007 tranchée en séance (assertions `is_empty` vides + aller-retour `open`/`save` non neutre) → +J02-P8, aller-retour neutre 19/19 ; suite 1 → 775 passed ; J02-P7 coupée par une limite Opus, try rendu

- 2026-09-09 s1: J01-P1..P6 ✅ ; revue round 1 en échec → corrections P3/P4/P5, round 2 sans finding bloquant ; D-003 (worktrees manuels), D-004 (+J01-P7 corrective)
- 2026-09-09 s0: plan created
- 2026-09-09 s0: plan revised (D-002) — acceptations J02-P1/P7 (`zipfile`) et J04-P4/P6 (`SyntaxWarning`) corrigées ; J03-P7 dépend de J03-P1..P6 ; `uv.lock` régénéré (lock périmé qui se réécrivait à chaque `uv run`)
