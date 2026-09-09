<!-- one decision, at most 100 lines; keep exactly one of the two states below.
     The queue is drained by the closing gate of each session; what is left here is what
     the next session's startup gate has to ask. -->

# DECISION — v2-semantic-engine

## Contention LibreOffice : `tests/support/libreoffice.py` non isolé

Question : les tests qui appellent LibreOffice échouent de façon non déterministe dès que deux
suites tournent en parallèle (une vague de 4 workers, ou une session concurrente). Faut-il isoler
le profil utilisateur de `soffice` ?

Preuve mesurée le 2026-09-09 en s4, au merge de J03-P1 : trois `soffice --headless --convert-to`
lancés en parallèle sur le même fichier — un des trois sort en **code 0 sans rien écrire** (aucune
ligne « convert … »), parce que `soffice` s'attache à l'instance déjà lancée qui détient le profil
`~/.config/libreoffice`. Côté test, `convert()` (`tests/support/libreoffice.py:44-80`) lève alors
`RuntimeError: soffice reported success but did not produce <out_path>`. Trois exécutions
successives de l'acceptation de J03-P1 sur la base ont donné 1 échec, puis 2 échecs sur d'autres
tests, puis 72 passed — victimes différentes à chaque fois, aucune régression de code.

Portée : `tests/characterization` figure dans les acceptations de J03-P1, J03-P5, J03-P7 et J03-P8,
et J06-P3 s'appuie aussi sur LibreOffice. Chaque vague parallèle de J03→J06 rejouera ce faux rouge,
et un vrai rouge y sera indistinguable d'un faux.

Correctif : passer `-env:UserInstallation=file://<profil temporaire unique>` à chaque invocation de
`soffice` dans `convert()`. `tests/support/libreoffice.py` n'est dans les `files` d'aucune part :
le correctif demande une part corrective ou l'élargissement d'une part existante.

Options :
1. Part corrective (T5, S) sur `tests/support/libreoffice.py` seul, dispatchée avant la suite de J03 — **recommandée** : l'instrument porte J03→J06, et un harnais qui rougit au hasard sous parallélisme rend illisible chaque merge de vague.
2. Élargir les `files` de J03-P7 (déjà porteuse du harnais d'écriture atomique) pour y ajouter le correctif — moins de parts, mais le faux rouge continue jusqu'à J03-P7, soit toute la vague J03-P5/P6.
3. Sérialiser à la main : ne jamais lancer deux acceptations LibreOffice en parallèle — aucun code changé, mais la contrainte n'est écrite nulle part et retombe sur chaque orchestrateur.

Bloque : rien (aucune part n'est empêchée ; c'est la fiabilité des acceptations qui est en jeu).
Source : orchestrateur s4, merge de J03-P1.

## Queue

(vide)
