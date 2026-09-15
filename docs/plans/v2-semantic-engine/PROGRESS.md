# PROGRESS — v2-semantic-engine

Updated 2026-09-10 · s6 · base v2-semantic-engine · ⬜ todo 🔄 wip ✅ done ❌ failed ⏸ blocked ⏭ dropped

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

## J02 — Cœur OOXML : paquet, flux de texte, plages, révisions (10/10)

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
| J02-P10 | Correctif D-009 : index V2 sans zones de texte, tracked_insert sans sortie de conteneur, réexports | T3 | ✅ | 1 | corrective D-009 ; fixtures `text_boxes` et `tracked_containers` ; repousse du `w:ins` dans le conteneur ; hors périmètre de J02-P7 (close), relue par J03-P8 ; suite 780 → 830 passed, 9 xfailed |
| J02-P7 | Review J02 | T2 | ✅ | 1 | round 1 sans finding bloquant ; 780 passed / 9 xfailed, build et lint verts ; 2 should-fix + 1 optional → D-009 |

## J03 — Migration des outils existants sur le cœur (11/11)

| Part | Title | Tier | Status | Tries | Note |
|------|-------|------|--------|-------|------|
| J03-P1 | search_and_replace et find_text_in_document | T3 | ✅ | 1 | `replace_text_everywhere` + `ReplacementReport` ; `, N skipped (inside fields)` ; aucun xfail de caractérisation ne visait ces 2 outils (D-010) |
| J03-P2 | Tracked changes sur le moteur | T2 | ✅ | 1 | droite→gauche (boucle infinie supprimée par construction), toutes les stories ; **laisse 5 rouges dans `tests/characterization/test_existing_tools.py`** (3 XPASS strict + 2 instantanés périmés par D-008) → corrigé par J03-P9 (D-012) |
| J03-P3 | Commentaires et hyperliens sur le moteur | T3 | ✅ | 1 | fil Word complet (`commentsExtended`/`commentsIds`/`people`), hyperliens `add`/`remove`/`list` ; hors `files` : `parts=` élargi sur 1 assertion de `tests/characterization/test_existing_tools.py` |
| J03-P4 | format_text et format_cell_text sans reconstruction | T3 | ✅ | 1 | `resolve`+`apply_rpr`, plus de `run.clear()` ; `resolve_color`/`run_patch` dans `core/tables.py` ; hors `files` : 1 xfail retiré dans `tests/characterization/test_existing_tools.py` |
| J03-P5 | Opérations destructrices : TOC, en-têtes, suppression, signets, blocs | T3 | ✅ | 3 | **review fix 1** (T2, findings J03-P8 round 1) : `delete_paragraph`/`add_bookmark` passés à l'espace V2 via `indexed_paragraphs()` (réutilise `_v2_index_map`), retrait par `getparent().remove()`, `w:sectPr` reporté sur le précédent *de cet espace* ; `replace_content` transmis par le wrapper `main.py` ; 4e puce `[Unreleased]` ; 4 tests rouges-puis-verts ; hors `files` : `tests/characterization/test_existing_tools.py` (helper `_v2_index_of`) ; 1087 → 1091 passed. Reste : `get_paragraph_text` et 9 autres points d'appel sur `doc.paragraphs` (D-016) · essai 1 coupé en vol (s4, aucun commit) ; patch de reprise relu et corrigé (4 défauts réels : `ensure_part`, double `w:sectPr` sur l'ancre, `sectPr` hissé hors d'une cellule, message de retour modifié), `delete_block_under_header` réintroduit ; 116 passed, 1 xfailed ; hors `files` : 3 xfail retirés dans `tests/characterization/test_existing_tools.py` (prévu par D-012) ; 2 changements de comportement pour le CHANGELOG de J03-P7/P8 |
| J03-P6 | merge_documents par import d'éléments | T2 | ✅ | 1 | `append_document(target, source, page_break) -> Report` ; document fusionné = premier source (garde styles/sections/en-têtes) ; `copy_table` supprimé (plus d'appelant) ; refus explicite si commentaires/notes dans la source |
| J03-P7 | Retrait du hook de sauvegarde, écritures atomiques, annotations | T4 | ✅ | 1 | `save_utils.py` supprimé, `protection_tools` sur `atomic_write_bytes`, préfixes `DESTRUCTIVE` retirés de `main.py` ; bug latent corrigé : `msoffcrypto.exceptions.InvalidFormatError` n'existe pas dans la version installée (`FileFormatError`) ; 59 passed, 1 xfailed |
| J03-P9 | Correctif D-012 : caractérisation des tracked changes alignée sur J03-P2 | T4 | ✅ | 1 | 3 tests renommés (XPASS→verts) ; `_reindex_after_merge` sur les 2 `touches_only…` ; `tests/characterization` : 49 passed, 4 xfailed |
| J03-P10 | Correctif D-013 : profil `soffice` jetable par conversion | T5 | ✅ | 1 | `-env:UserInstallation` jetable par appel, `rmtree` en `finally` ; 3 tests dont 1 de concurrence réelle (LibreOffice 25.2 présent, non skip) ; 14 passed |
| J03-P11 | Correctif D-015 : CHANGELOG des changements de comportement de J03 | T5 | ✅ | 1 | entrée `[Unreleased]` : 1 `Changed` (fusion) + 2 `Fixed` (garde-fous) ; préambule et historique ≥ 1.6.0 byte à byte intacts (hashes) ; 7/7 acceptations rejouées sur la base |
| J03-P8 | Review J03 | T2 | ✅ | 1 | round 1 **failed** (1 bloquant : `delete_paragraph`/`add_bookmark` hors espace V2, `find_text → 5` supprimait « Echo ») → J03-P5 review fix 1 ; round 2 sans finding bloquant ; 24 acceptations + 5 Checks verts, 1091 passed / 1 xfailed, 120 outils MCP, `uv lock --check` 0 ; 2 should-fix (D-016, D-017) + 2 optional (D-018, D-019) |

## J04 — Surface V2 : adressage, inspection, dry-run, capacités, docs (6/7)

| Part | Title | Tier | Status | Tries | Note |
|------|-------|------|--------|-------|------|
| J04-P7 | Correctif D-016 : un seul espace d'index pour `paragraph_index` (D-017, D-018, D-019 embarquées) | T3 | ✅ | 2 | essai 1 `blocked` (try rendu) : `isolation: "worktree"` a branché depuis `main` (D-003) ; essai 2 en worktree manuel, 8/8 acceptations vertes, 1101 passed / 1 xfailed sur la suite complète ; hors `files` nécessaire : ancre de `test_insert_numbered_list_near_text_only_appends_at_the_target` recalculée dans l'espace V2 (`tests/characterization/test_existing_tools.py`) |
| J04-P1 | Locators et inspection | T3 | ✅ | 1 | `Target`/`resolve`/`inspect`, 102 tests ; codes d'erreur `invalid`/`not_found`/`ambiguous`/`stale_anchor` (D-022) ; suite complète 1049 → au moins autant passed ; `ruff check .` casse déjà sur `tests/live/test_mac_paths.py` (I001, hérité de J04-P4, hors `files`) |
| J04-P2 | Outils `doc_*` texte, enregistrement, rapport structuré | T3 | ✅ | 1 | `registry.py`/`text.py` : 4 outils, 55 tests, 120 → 124 outils ; suite complète 1258 passed, 1 xfailed, ruff vert ; écart assumé signalé pour J04-P6 : la surface V2 n'applique pas `ensure_docx_extension` (verrou pris sur le fichier réellement ouvert) |
| J04-P3 | Plateformes, TOOLS.md généré, garde-fous de dérive, registre corrigé | T4 | ✅ | 1 | `PLATFORMS` classé par lecture du code réel (31 macOS / 14 Windows-only sur les 45 live, écart avec l'estimation de brief documenté, D-023) ; branché avant le merge de J04-P5 : `doc_apply_edits` manquant de `PLATFORMS`/`TOOLS.md`, corrigé par l'orchestrateur après merge (1 ligne + régénération) ; suite complète 1383 passed, 1 xfailed, ruff et build verts |
| J04-P4 | Plantages macOS prouvés par lecture | T4 | ✅ | 1 | `depends_on: []` ; worktree branché par erreur depuis `main` (D-003) mais fichiers jamais touchés par `v2-semantic-engine` (`git log v2-semantic-engine ^main -- <files>` vide) : diff-stat vérifié limité aux 4 `files`, merge propre, acceptations rejouées vertes sur la base |
| J04-P5 | Lot atomique `doc_apply_edits` | T4 | ✅ | 1 | `doc_apply_edits`, réutilise les helpers `text.py`/`registry.py` ; 13 tests ; index de l'édition fautive préfixé au message (`edit {index}: …`) faute de place dans l'enveloppe générique |
| J04-P6 | Review J04 | T2 | 🔄 | 1 | `depends_on` étendu à J04-P7 ; relit l'alignement, les notes laissées à l'écart, la 5e puce, la borne d'`add_bookmark` et les 2 formes de `search_and_replace` ; range `c0e98c8..HEAD` |

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

- Ready: **J04-P7 en premier** (T3, `depends_on: []`, tourne seule — J04-P1 en dépend et J04-P2..P5 partagent `content_tools.py`/`layout_tools.py` avec elle). Puis J04-P1..P5 (cap de vague 4, `files` disjoints), J04-P6 en revue.
- D-020 (merge de `v2-semantic-engine` sur `main`) est à trancher au démarrage : la revue J03-P8 est verte, ce qui lève la condition posée par D-014, mais J04-P7 corrige encore la surface publique de 5 outils.
- Orchestrator: best/max (D-020 et D-021 en attente)
- Decision: D-020 (détaillée dans DECISION.md), puis D-021 en file

## Log

- 2026-09-10 s6: J03-P11 ✅ (entrée `[Unreleased]`, 7/7 acceptations) ; revue J03-P8 round 1 **failed** sur 1 bloquant — `find_text_in_document` rendait un index V2 depuis J03-P1 tandis que `delete_paragraph` et `add_bookmark` indexaient les enfants du corps, donc `find_text → 5` puis `delete_paragraph(5)` supprimait le mauvais paragraphe en répondant « deleted successfully », dès qu'un `w:sdt` de bloc était présent (celui qu'`add_table_of_contents` insère lui-même) → J03-P5 review fix 1 (T2, essai 3) : `indexed_paragraphs()` réutilise `_v2_index_map`, retrait par `getparent().remove()`, `w:sectPr` reporté sur le précédent de l'espace V2, `replace_content` transmis par le wrapper `main.py`, 4e puce au CHANGELOG, 4 tests rouges-puis-verts ; revue round 2 sans finding bloquant ; **J03 complet (11/11)** ; suite 1087 → 1091 passed, 1 xfailed ; lint, build et `uv lock --check` verts ; 4 décisions tranchées à la clôture (D-016 espace d'index sans les notes ; D-017 5e puce du CHANGELOG ; D-018 borne inférieure d'`add_bookmark` ; D-019 séparateur de `search_and_replace`) → **+J04-P7** en tête de J04, J04-P1 et J04-P6 en dépendent ; D-020 (merge sur `main`) et D-021 (message non déterministe d'`insert_line_or_paragraph_near_text`) laissées en file
- 2026-09-10 s5: reprise de la s4 coupée en vol (PROGRESS.md non commité, worktree J03-P5 sans commit → essai 1 compté, 775 insertions sauvées en patch) ; D-013 tranchée au démarrage → +J03-P10 ✅ (profil `soffice` jetable, faux rouge sous parallélisme corrigé) ; J03-P5 ✅ en essai 2 (patch de reprise relu, 4 défauts réels corrigés) ; J03-P7 ✅ (hook de sauvegarde supprimé, écritures atomiques, `msoffcrypto.exceptions.InvalidFormatError` inexistant corrigé) ; **J03 à 9/10, seule la revue J03-P8 reste** ; suite 830 → 1087 passed, 1 xfailed ; lint, build et `uv lock --check` verts ; D-014 (user) : `main` intact jusqu'à la revue
- 2026-09-10 s5 (clôture): D-015 tranchée — le CHANGELOG des 3 changements de comportement de J03 n'était dans les `files` d'aucune part (D-011 le confiait à J03-P7, dispatchée sans, ou à J03-P8, qui est une revue) → +J03-P11 (T5, `CHANGELOG.md` seul), J03-P8 en dépend
- 2026-09-09 s3: J02-P9 et J02-P7 (revue J02, round 1 sans finding bloquant) ✅ ; D-009 tranchée en séance sur les 2 should-fix de la revue (run sur conteneur sorti de son `w:hyperlink`/`w:sdt` par `tracked_insert`, index V2 faussé par les zones de texte) → +J02-P10 ✅ ; **J02 complet (10/10)** ; suite 775 → 830 passed, 9 xfailed
- 2026-09-09 s2 (clôture): D-005 (lecture mécanique des champs imbriqués), D-006 (alias "body", index V2 None en cellule ; vocabulaire J03/J04 aligné), D-008 (+J02-P9) tranchées
- 2026-09-09 s2: J01-P7, J02-P1..P6 et J02-P8 ✅ ; D-007 tranchée en séance (assertions `is_empty` vides + aller-retour `open`/`save` non neutre) → +J02-P8, aller-retour neutre 19/19 ; suite 1 → 775 passed ; J02-P7 coupée par une limite Opus, try rendu

- 2026-09-09 s1: J01-P1..P6 ✅ ; revue round 1 en échec → corrections P3/P4/P5, round 2 sans finding bloquant ; D-003 (worktrees manuels), D-004 (+J01-P7 corrective)
- 2026-09-09 s0: plan created
- 2026-09-09 s0: plan revised (D-002) — acceptations J02-P1/P7 (`zipfile`) et J04-P4/P6 (`SyntaxWarning`) corrigées ; J03-P7 dépend de J03-P1..P6 ; `uv.lock` régénéré (lock périmé qui se réécrivait à chaque `uv run`)
