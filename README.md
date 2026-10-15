# Résumé quotidien des sessions Claude

Chaque jour à 16h, résume les sessions Claude Code du jour et alimente le suivi des tâches :

- **version détaillée** (un sujet par ligne, avec ses points) → `Suivi taches details.ods` sur le Bureau Windows ;
- **version courte** (une ligne par sujet) → `Suivi tâches/Suivi taches.xlsx` sur kDrive, envoyé comme nouvelle version via WebDAV.

Le lendemain, ce qui a été fait après le résumé (donc après 16h) est ajouté à la fin du bloc de la veille.

## Fonctionnement

1. Lecture des sessions Claude Code du jour (`~/.claude/projects/*/*.jsonl`) : messages de l'utilisateur et dernière réponse de chaque session.
2. Résumé par `claude -p` (sans outils, sans session enregistrée), en JSON : `details` et `short`.
3. Écriture : en-tête « Jour JJ/MM/AAAA » au style des autres en-têtes, puis une ligne par sujet. L'onglet du mois (« Novembre 2026 ») est créé s'il n'existe pas. Les fichiers sont modifiés directement dans leur XML (ods et xlsx) pour ne rien perdre de la mise en forme.
4. kDrive : téléchargement de la version en ligne, ajout, renvoi avec `If-Match`, puis relecture pour vérifier.

Chaque passage est idempotent : le résumé d'un jour est généré une fois (`state/AAAA-MM-JJ.json`) et un jour déjà présent n'est jamais réécrit. Un fichier ouvert (verrou local, ou `423 Locked` sur kDrive quand il est ouvert dans l'éditeur en ligne) est retenté au passage suivant. Les jours manqués (PC éteint) sont rattrapés, jusqu'à 7 jours. Une copie de chaque fichier est gardée dans `backups/` avant écriture.

## Installation

- WSL (Ubuntu) avec `python3`, `lxml`, `requests`, et Claude Code dans `~/.local/bin/claude`.
- `cp kdrive.env.example kdrive.env && chmod 600 kdrive.env`, puis renseigner l'identifiant et un mot de passe d'application Infomaniak.
- Tâche planifiée Windows : exécuter `register-task.ps1` dans PowerShell.

## Utilisation à la main

```bash
python3 daily_summary.py --dry-run            # affiche le résumé du jour sans rien écrire (après 16h)
python3 daily_summary.py --date 2026-10-14    # traite un jour précis
python3 daily_summary.py --regenerate --date 2026-10-14   # ignore le résumé déjà généré
```

Journal : `daily_summary.log`.
