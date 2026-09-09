# Audit des outils existants — comportements destructeurs ou cassés

Ce tableau est rédigé à la main à partir des raisons `xfail` posées dans
`tests/characterization/test_existing_tools.py`. Il documente le comportement
**actuel** des outils cross-platform (`word_document_server/`), constaté en
exécutant chaque outil sur la fixture `combined` (ou une fixture minimale
dédiée) et en comparant l'instantané avant/après avec
`tests/support/snapshot.py`. Conformément à la décision D-001 du plan, ce
document ne corrige rien : la réécriture de ces outils est prévue pour le
jalon J03.

| Outil | Symptôme constaté | Test témoin |
|---|---|---|
| `format_text` (`word_document_server/tools/format_tools.py`) | Efface tous les runs du paragraphe ciblé (`run.clear()` sur `paragraph.runs`) puis les reconstruit en trois morceaux (avant / cible / après) : la mise en forme (gras, italique, style de caractère, langue…) de tout texte du paragraphe situé **hors** de la plage `start_pos:end_pos` est perdue, y compris quand elle n'a rien à voir avec l'appel. Déjà signalé `DESTRUCTIVE` dans la docstring (J01-P1). | `test_format_text_destroys_untouched_run_formatting` |
| `add_table_of_contents` (`word_document_server/tools/content_tools.py`) | Reconstruit le document entier dans un `Document()` vierge : ne copie que `paragraph.text` / `cell.text` (donc dépendant de `paragraph.text`, qui ignore les runs sous `w:ins`/`w:del`) et le nom de style des paragraphes. `comments.xml`, `commentsExtended.xml`, `footnotes.xml`, `endnotes.xml`, les en-têtes/pieds de page, les images, les hyperliens, les signets, les champs, les révisions suivies et les fusions de cellules de tableau disparaissent tous. Déjà signalé `DESTRUCTIVE` (J01-P1). | `test_add_table_of_contents_drops_ancillary_parts` |
| `merge_documents` (`word_document_server/tools/document_tools.py`) | Reconstruit la cible dans un `Document()` vierge à partir de chaque source : ne recopie que le texte des paragraphes, le nom de style s'il existe dans la cible, et un sous-ensemble du formatage direct des runs (gras/italique/souligné/taille). Mêmes pertes que `add_table_of_contents` (commentaires, révisions, notes, en-têtes/pieds de page, hyperliens, signets, champs, images). Déjà signalé `DESTRUCTIVE` (J01-P1). | `test_merge_documents_drops_ancillary_parts` |
| `add_header_footer` (`word_document_server/tools/layout_tools.py`) | Avant d'écrire le nouveau texte, vide tous les paragraphes existants de l'en-tête ou du pied de page (`for p in header.paragraphs: p.clear()`) : toute image, champ ou run déjà présent dans l'en-tête/pied de page ciblé est perdu, même s'il n'a aucun rapport avec le nouveau texte. | `test_add_header_footer_clears_existing_header_content` |
| `create_custom_style` (`word_document_server/tools/format_tools.py`, via `word_document_server/core/styles.py:create_style`) | Ne crée jamais le style demandé : `create_style` teste l'existence du style avec `doc.styles.get_by_id(style_name, WD_STYLE_TYPE.PARAGRAPH)`, or `get_by_id` ne lève jamais — il renvoie le style par défaut du type (`Normal`) quand l'id est introuvable. La branche `except:` censée créer le style n'est donc jamais atteinte : le style par défaut est renvoyé et rien n'est écrit dans `styles.xml`. L'outil répond pourtant toujours « Style '...' created successfully. ». | `test_create_custom_style_never_creates_the_style` |
| `replace_block_between_manual_anchors` (`word_document_server/utils/document_utils.py`, outil `replace_block_between_manual_anchors_tool`) | Ne trouve jamais l'ancre de départ, quel que soit le texte fourni : la comparaison `el.tag == CT_P.tag` (et `CT_Tbl.tag`) compare une chaîne à l'attribut de classe non lié `tag` de `lxml.etree._Element` (`CT_P` et `CT_Tbl` sont des classes d'éléments enregistrées via `element_class_lookup`, pas des instances), qui n'est jamais égal à une chaîne. `start_idx` reste `None` et l'outil renvoie systématiquement « Start anchor '...' not found. », même quand le texte existe mot pour mot dans le document. | `test_replace_block_between_manual_anchors_never_finds_the_anchor` |
| `track_replace` (`word_document_server/core/tracked_changes.py:track_replace_in_doc`) | Boucle infinie quand `new_text` contient `old_text` comme sous-chaîne : la boucle `while True` recherche `old_text` dans **tous** les runs du paragraphe, y compris ceux qui viennent d'être insérés dans le `<w:ins>` du remplacement précédent. Avec `track_replace("Risk", "Risk Risk")`, le texte inséré recontient "Risk", qui est retrouvé et remplacé indéfiniment (empaquetage `<w:ins>` de plus en plus profond) ; l'appel ne se termine jamais. | `test_track_replace_infinite_loop_when_replacement_contains_original` (sous `@pytest.mark.timeout(10)`) |
| `accept_tracked_changes` / `reject_tracked_changes` (`word_document_server/core/tracked_changes.py`) sur une **marque de paragraphe** suivie/insérée | Les deux fonctions traitent `w:rPr/w:del` et `w:rPr/w:ins` d'une marque de paragraphe exactement comme un `w:del`/`w:ins` de run : la balise est retirée (`parent.remove(...)`), rien de plus. Or accepter une marque supprimée devrait fusionner le paragraphe avec le suivant (l'effet réel de la suppression du saut de paragraphe), et rejeter une marque insérée devrait défaire la scission de la même façon. Constaté à l'exécution sur la fixture `combined` : le nombre de paragraphes du document ne change pas après `accept_tracked_changes()` ni après `reject_tracked_changes()`, alors qu'il devrait diminuer de un dans chaque cas. La révision est donc « acceptée »/« rejetée » (marqueur disparu, compteur `revisions` en baisse) sans que son effet ne soit jamais appliqué. | `test_accept_tracked_changes_does_not_merge_a_deleted_paragraph_mark`, `test_reject_tracked_changes_does_not_merge_an_inserted_paragraph_mark` |

## Notes complémentaires

- `accept_tracked_changes_in_doc` / `reject_tracked_changes_in_doc` ne
  suppriment jamais les marqueurs `w:rPrChange` / `w:pPrChange` : après
  acceptation ou rejet, une révision de mise en forme ou de propriétés de
  paragraphe laisse son historique `*Change` en place indéfiniment. Ce n'est
  pas une perte de contenu (le paragraphe reste correct) mais le compteur
  `revisions` ne retombe jamais à zéro même quand toutes les révisions
  visibles ont été traitées. Non couvert par un test dédié ici (hors du
  périmètre de ce jalon), à garder en tête pour J03.
- Le risque « `numId` 1/2 en dur » de `insert_numbered_list_near_text`
  (`word_document_server/utils/document_utils.py:add_bullet_numbering`) ne se
  reproduit pas avec les fixtures de ce dépôt : le gabarit par défaut de
  python-docx (`Document()`) fournit déjà des définitions `w:num` pour les
  `numId` 1 à 9, donc `validate_package` ne signale aucun `NUMID-UNRESOLVED`
  dans `tests/characterization/test_existing_tools.py`. Ce n'est constaté que
  pour les fixtures construites par ce dépôt ; un document dont le gabarit ne
  fournit pas ces `numId` resterait à risque.
- `protect_document` / `unprotect_document` (`word_document_server/tools/protection_tools.py`)
  ne sont pas destructeurs : le fichier chiffré n'est plus un paquet OPC valide
  tant qu'il est protégé (attendu, `msoffcrypto` produit un conteneur OLE), et
  `unprotect_document` restitue des octets strictement identiques à l'original.
  Voir `tests/characterization/test_existing_tools.py::test_protect_then_unprotect_roundtrips_bytes`.
