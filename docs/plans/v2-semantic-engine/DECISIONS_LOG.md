# DECISIONS_LOG — v2-semantic-engine

Append-only, one line per decision. Never edit past lines.

- 2026-09-09 D-001 Validation du plan: 6 jalons J01→J06 approuvés ; périmètre live limité aux 3 plantages macOS prouvés par lecture (J04-P4) ; outils destructeurs conservés avec avertissement DESTRUCTIVE dès J01-P1 puis réécrits en J03 (reco: idem sur les trois points); long-term → R-001, R-002
- 2026-09-09 D-002 Révision du plan avant exécution: acceptations insatisfaisables ou non déterministes corrigées — J02-P1/P7 `'zipfile' not in sys.modules` (python-docx importe zipfile) → contrôle statique sur `engine/` ; J04-P4/P6 `-W error` sur import (masqué par le `.pyc`) → compilation depuis la source ; J03-P7 dépend de J03-P1..P6 ; `uv.lock` régénéré en commit séparé (périmé depuis v1.1.3, réécrit à chaque `uv run` ; `uv lock --check` stable ensuite) (reco: idem); plan updated: J02-P1, J02-P7, J03-P7, J04-P4, J04-P6
