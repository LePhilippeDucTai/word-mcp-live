# DESIGN-REVIEW — v2-semantic-engine

Revue contradictoire du brief « moteur Word robuste pour agents IA », après audit complet du dépôt (2026-09-09, commit `c6c7617`, fork à jour sur l'upstream). Les positions ci-dessous sont celles que le plan met en œuvre ; le détail exécutable est dans `PLAN.md` et `jalons/`.

## 1. État réel du projet

| Domaine | Constat vérifié |
|---|---|
| Registre | 120 outils (75 cross-platform, 45 live), pas 124. README, TOOLS.md, manifest, server.json, CHANGELOG et `__version__` donnent six chiffres différents. `main.py` = 120 wrappers écrits à la main ; un wrapper omet déjà un paramètre (`scrub_orphans`). |
| Tests | Couverture automatisée ≈ 0 : un test qui s'auto-saute sans LibreOffice, un script async non collectable. Aucune fixture. Aucun test ajouté dans l'historique du fork. |
| Chemins d'écriture | python-docx pour 75 outils ; quatre modules lxml qui réécrivent le ZIP à la main (`tracked_changes`, `comment_writer`, `hyperlink_writer`, `footnotes`) avec trois politiques d'ids, quatre allocateurs de `rId`, six copies de la boucle de réécriture, une seule sauvegarde atomique (`footnotes`). 50 `doc.save()` en place, non atomiques. |
| Destructif prouvé | `add_table_of_contents` et `merge_documents` reconstruisent le document depuis `paragraph.text` (perte totale : runs, images, champs, signets, commentaires, notes, sections, styles, numérotation, ordre). `format_text` vide tous les runs du paragraphe et duplique le texte des hyperliens (reproduit). `add_header_footer` efface logos et champs. `create_custom_style` est un no-op (`get_by_id` ne lève jamais) qui répond « created successfully ». `track_replace` boucle à l'infini si le remplacement contient le texte cherché (reproduit, timeout). `search_and_replace` rate toute correspondance à cheval sur deux runs et répond « No occurrences ». Les listes réutilisent `numId` 1/2 en dur. `delete_paragraph` peut supprimer le `sectPr` final. |
| Prémisses périmées | Le monkey-patch de `Document.save` (« python-docx perd comments.xml ») est faux en 1.1.2 : sur une sauvegarde sans modification, aucune partie n'est perdue, toutes sont canoniquement identiques sauf `[Content_Types].xml` régénéré (équivalent), les parties binaires sont identiques octet à octet, LibreOffice rouvre le fichier. |
| Live | Dispatch = 44 blocs `if _MAC_AVAILABLE:` placés avant le `try` (toute erreur JXA remonte brute). Parité réelle 30/45, pas 40/44. Six bases d'index coexistent. Trois plantages macOS démontrables par lecture. `mac_apply_list` génère du VBA via AppleScript avec un fichier `/tmp` fixe. |
| Ce qui est bon | L'OPC de python-docx comme couche paquet ; la découpe de runs de `comment_writer` ; le pattern temp+`os.replace` et le validateur de cohérence de `footnotes.py` ; `text_safety.py` ; le verrou par fichier ; LibreOffice headless disponible (≈2 s par document sur le Pi). |

## 2. Critique du brief

| Directive | Verdict | Justification |
|---|---|---|
| Architecture cible en quatre couches avec abstraction backend | **REJECT** (couches) / **MODIFY** (cœur) | Une interface commune aux trois backends serait une seconde implémentation de Word : find COM a des jokers, JXA n'a ni réponses ni résolution de commentaires, le live adresse en offsets absolus, python-docx en index de corps. Le cœur OOXML devient bien la référence ; le live reste une famille d'outils distincte reliée par le vocabulaire (locators, rapport) et les métadonnées de plateforme. |
| Backend OOXML « de première classe » | **ACCEPT** | C'est le seul chemin développable et testable ici ; il porte déjà la moitié des outils. |
| P0 Style Engine complet | **MODIFY → J05** | Les styles ne sont pas la cause des pertes de données ; le moteur de plages l'est. Modèle minimal retenu : identité, héritage, police, paragraphe, lien de numérotation, métadonnées ; écriture limitée aux familles paragraphe et caractère ; `compare_styles` abandonné (trivial sur `get_style`). |
| P0 Cascade `get_effective_format` | **MODIFY** | Gardé en lecture seule comme approximation documentée : direct → style de caractère → chaîne de paragraphe → docDefaults → thème, avec sémantique XOR des propriétés bascule ; styles de tableau signalés `unresolved`. Reproduire toute la cascade Word (tableaux conditionnels, `w:rPr` de marque de paragraphe, `webSettings`) n'a pas de valeur pour un agent. |
| P0 Theme-aware formatting | **MODIFY** | Pas un chantier séparé : deux règles du moteur (les lecteurs renvoient la référence de thème et la valeur résolue ; les écritures copient le `rPr` entier et ne retirent `themeColor`/`asciiTheme` que sur demande explicite) plus un lecteur de thème en J05. |
| P0 Safe cross-run editing | **ACCEPT** — c'est le cœur (J02) | Primitive unique : découpe de runs aux offsets + morceaux entiers, avec invariants formels et refus explicites (champ chevauché, contenu opaque supprimé, texte caché). Elle rend inutiles les trois matchers dupliqués et sert commentaires, révisions, hyperliens, formatage, signets. |
| P0 Éliminer les opérations destructrices | **ACCEPT** (J03) | Liste prouvée en §1 ; aucune reconstruction n'est jugée sûre : même `copy_table` perd les paragraphes multiples des cellules. |
| P0 Invariants de non-dégradation | **ACCEPT** — premier jalon (J01) | Les invariants sont testables (§5) et l'outillage qui les mesure doit exister avant le moteur ; la caractérisation des outils existants transforme l'audit en tests témoins. |
| P0 Tests round-trip et 17 fixtures binaires | **MODIFY** | Aucun binaire commité : fixtures générées (constructeur python-docx+lxml déterministe, et LibreOffice depuis des `.fodt` pour une forme OOXML étrangère), cache par session. LibreOffice valide l'ouverture, jamais la vérité Word. |
| P1 Semantic Document Model avec IDs de session | **REJECT** (IDs) / **MODIFY** (adressage) | Un registre d'IDs est invalidé par toute modification externe et n'a pas de sens pour des appels d'outil sans état sur un fichier. Retenu : locators hybrides sans état (index V2 + ancre textuelle, recherche, signet, titre, cellule, story) avec erreurs `stale_anchor`/`ambiguous` et `doc_inspect` qui fournit les index. |
| P1 Transactions et diff sémantique | **REJECT** (transactions) / **MODIFY** | Incompatible avec l'undo natif, l'autosave et l'absence d'état côté serveur. Retenu : `dry_run` + rapport de changements sur chaque outil V2, lot `doc_apply_edits` tout-ou-rien en mémoire avec une seule sauvegarde atomique, `doc_compare` en J06. |
| P1 Character styles | **ACCEPT** (J05) | Dépend du moteur de plages ; `char_style` fait partie du patch de formatage dès J02-P4. |
| P1 Table styles | **MODIFY** | Application d'un style existant (`tblStyle` + `tblLook`) en J05 ; création de styles de tableau différée (R-007). |
| P1 Audit et normalisation | **ACCEPT** (audit, J06) / **DEFER** (normalisation, R-007) | L'audit est en lecture seule sur le format effectif ; toute normalisation passe par les primitives existantes, pas par un outil magique. |
| Génération de la documentation | **ACCEPT** (J04) | `mcp.list_tools()` expose nom, description, schéma, annotations ; `TOOLS.md` et compteurs README générés, test de dérive. |
| Compatibilité | **ACCEPT** | Aucun renommage, aucun changement de format de retour des 120 outils ; nouveaux outils sous préfixe `doc_*` ; paramètres ajoutés optionnels seulement. |
| Ordre qualité (correctness > non destructif > testabilité > architecture > compat > perf > quantité) | **ACCEPT** | Le plan suit cet ordre : J01 mesure, J02 garantit, J03 corrige, J04–J06 ajoutent. |

Interactions dangereuses examinées : IDs × modifications concurrentes (résolu par l'absence d'IDs) ; transactions × undo natif (résolu par le rejet des transactions) ; remplacement cross-run × champs/signets/commentaires/révisions (résolu par les refus explicites et la conservation des marqueurs) ; styles × numérotation (le lien `numPr` d'un style est lu, jamais écrit) ; thème × formatage direct (règles ci-dessus) ; abstraction backend × comportements incompatibles (résolu par le rejet de l'abstraction).

## 3. Principaux risques techniques

1. Le moteur de plages (J02-P3) : un bug y corrompt tout ; d'où les tests de propriété sur plages aléatoires, l'instantané et le validateur avant toute ligne du moteur.
2. La migration J03 change des comportements observables en mieux (correspondances cross-run trouvées, remplacements dans les en-têtes) : documenté dans le CHANGELOG, jamais de rupture de signature.
3. Sérialisation python-docx : parties analysées réécrites canoniquement, `[Content_Types].xml` régénéré ; l'invariant est donc « canoniquement identique », pas « octet à octet », sauf parties binaires.
4. Espaces d'index : l'index V2 (corps + sdt de bloc, cellules exclues) coïncide avec `doc.paragraphs` pour les documents sans sdt de bloc ; les ancres `expect_text` rattrapent le reste.
5. LibreOffice n'est pas Word : un document qui s'ouvre dans LibreOffice peut encore déclencher une réparation Word ; le validateur de paquet et le respect des ordres de schéma limitent ce risque, une validation Word réelle reste une recommandation (R-001).
6. `merge_documents` par import d'éléments refuse notes et commentaires en v1 plutôt que de les perdre.

## 4. Architecture cible

```text
MCP tools (main.py : 120 wrappers inchangés + tools/v2/* découverts)
    ↓
tools/* (implémentations existantes, formats de retour conservés)   tools/v2/* (dict structuré, dry_run)
    ↓                                                                 ↓
word_document_server/engine/  — pur Python/lxml, synchrone, sans MCP ni Word
    package · ids · textmodel · ranges · format · find · revisions · locators · inspect
    styles · theme · effective · numbering · table_styles · merge · audit · compare
    ↓
python-docx OPC (parties, relations, types de contenu) → .docx (sauvegarde atomique)

tools/live_* (COM, JXA) — famille distincte, inchangée hors bugs prouvés ; reliée par tools/platforms.py
```

Alternatives comparées :

| Option | Verdict |
|---|---|
| Hiérarchie de backends (`OOXMLBackend`, `WordComBackend`, `WordJxaBackend`) | Rejetée : interface géante pour peu de comportements communs, dégradation silencieuse garantie là où les backends divergent. |
| Capabilities/services | Retenue en version minimale : `tools/platforms.py` + `doc_capabilities` décrivent ce qui existe où ; pas de dispatch dynamique. |
| Ports/adapters | Réduit à un seul port réel : `convert_to_pdf` (Word ou LibreOffice). |
| Domaine partagé + opérations natives | Retenue : le moteur est le domaine ; le live garde ses opérations natives (pagination, champs, capture, undo, OMath). |
| Couche OPC maison (zipfile + lxml) | Rejetée : python-docx préserve déjà toutes les parties (vérifié) ; une seconde couche paquet reproduirait les bugs de types de contenu et de `rId` de l'existant. |

## 5. Invariants

Vérifiés par `tests/support` (J01) puis par les tests de propriété du moteur (J02) et la caractérisation (J01-P5, J03) :

1. Parties non ciblées canoniquement identiques (c14n) ; parties binaires identiques octet à octet ; aucune partie perdue ni dupliquée ; chaque partie a un type de contenu ; chaque relation a une cible.
2. Texte visible et `rPr` de chaque run hors plage inchangés.
3. Signets, plages de commentaire, `permStart/End` jamais supprimés ; marqueurs d'une plage supprimée conservés à sa position de début.
4. Champs atomiques : aucune opération ne chevauche partiellement un champ ; `instrText` jamais visible.
5. `w:del` jamais apparié ni modifié ; `w:ins` visible et modifiable.
6. Ids d'annotation uniques sur toutes les stories ; ids de commentaires et de notes cohérents avec leurs parties ; `numId` résolus.
7. Enfants de `pPr`, `rPr`, `tblPr`, `tcPr` en ordre de schéma ; `bookmarkStart` jamais avant `pPr`.
8. Sauvegarde atomique : le fichier d'origine est intact tant que la nouvelle version n'est pas complète ; aucun temporaire résiduel.
9. Le document s'ouvre dans LibreOffice après chaque scénario de bout en bout.
10. Aucun outil ne répond succès après une exception avalée ou un repli différent de la demande.

## 6. Milestones

| Jalon | Contenu | Livrable vérifiable |
|---|---|---|
| J01 Harnais | Outillage, fixtures générées (deux producteurs), instantané, validateur, caractérisation | `tests/characterization` vert avec xfail stricts documentés dans `docs/audit/destructive-ops.md` |
| J02 Cœur | Paquet atomique, flux de texte, plages, formatage, recherche, révisions | Tests de propriété verts ; import du moteur sans fastmcp ni zipfile |
| J03 Migration | Outils texte existants sur le moteur, TOC/merge/en-têtes non destructifs, hook retiré, écritures atomiques | Plus aucun xfail de caractérisation pour ces outils |
| J04 Surface V2 | Locators, `doc_inspect`, `doc_edit_text`, `doc_format_range`, `doc_apply_edits`, `doc_capabilities`, TOOLS.md généré, registre corrigé, 3 plantages macOS | `gen_tools_md.py --check` vert ; ≥124 outils enregistrés |
| J05 Styles | Thème, styles en lecture, format effectif, styles paragraphe/caractère en écriture, numérotation, styles de tableau | `create_custom_style` et listes réels |
| J06 Validation | Audit, comparaison, scénarios agent, documentation | `tests/e2e` vert, README et CHANGELOG alignés |

## 7. Tests

| Jalon | Unitaires | OOXML | Round-trip | Word natif |
|---|---|---|---|---|
| J01 | constructeurs, validateur, instantané | validateur sur fixtures des deux producteurs | no-op python-docx sur toutes les fixtures | — |
| J02 | chaque module du moteur | ordre de schéma, ids, parties | propriété : plages aléatoires → invariants 1–8 | — |
| J03 | chaque outil migré | validateur après chaque outil | caractérisation (instantané avant/après) | — |
| J04 | locators, inspect, outils V2, registre | — | lot atomique : échec → fichier identique | JXA bouchonné (script généré) |
| J05 | thème, styles, effectif, numérotation | `styles.xml`, `numbering.xml` valides | fixtures de styles et LibreOffice | — |
| J06 | audit, compare | validateur | scénarios complets + ouverture LibreOffice | — |

## 8. Compatibilité

- Aucun outil renommé ni supprimé ; signatures inchangées ; formats de retour des outils existants inchangés (chaînes et JSON).
- Changements de comportement volontaires, tous documentés dans le CHANGELOG : correspondances cross-run et stories supplémentaires trouvées par `search_and_replace` ; `add_table_of_contents` insère un vrai champ ; `merge_documents` conserve le formatage et refuse notes/commentaires ; `add_header_footer` refuse d'effacer champs et images sans `replace_content=True` ; `create_custom_style` crée réellement le style.
- Dépendances : `pytest` sort des dépendances runtime ; `lxml` et `python-dotenv` déclarés ; rien de retiré. Pas de bump de version dans ce plan (R-005, R-009).
- Migration progressive : les anciens outils utilisent le moteur en interne ; les nouveaux outils `doc_*` sont additifs ; aucune dépréciation formelle avant retour d'usage.

## 9. Linux vs Word natif

| Catégorie | Éléments |
|---|---|
| Fully testable on Linux | tout `engine/`, les 75 outils cross-platform, les outils `doc_*`, la génération de docs, le registre, le validateur, LibreOffice (optionnel) |
| Partially testable on Linux | branches macOS des outils live par bouchonnage de `_run_jxa` (script généré vérifié, effet réel non vérifié) ; `convert_to_pdf` via LibreOffice seulement |
| Requires Windows Word | 45 outils `word_live_*` en COM : pagination, texte par page, diagnostics de mise en page, champs, capture, révisions natives, OMath, undo, `Selection` |
| Requires macOS Word | les 30 branches JXA, `mac_add_comment` (AppleScript), `mac_apply_list` (VBA), `mac_save_as_pdf` |

## 10. Décisions

Tranchées par le user le 2026-09-09 (D-001) : découpage en six jalons ; périmètre live limité aux trois plantages prouvés ; outils destructeurs avertis dès J01 puis réécrits en J03.

Tranchées par le plan sans arbitrage humain : python-docx OPC comme couche paquet ; pas d'IDs de session ; pas de transactions ; préfixe `doc_*` pour la surface V2 ; refus explicite plutôt qu'édition ambiguë ; LibreOffice comme producteur et validateur optionnel ; fixtures générées, jamais commitées en binaire ; code et docs d'outils en anglais.

Reportées : R-001 à R-011 dans `LONG_TERM_RECO.md` (raccordement live, notes de bas de page, protection, nom PyPI, fusion avec notes, styles de tableau en création, wrappers auto-générés, Docker, python-docx 1.2, divergence COM/JXA des listes).
