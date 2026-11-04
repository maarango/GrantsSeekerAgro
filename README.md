# GrantsSeekerAgro

An autonomous agent that explores **national (Colombia) and international grant calls** for agricultural research **every day**, syncs them to your cloud folder (OneDrive, Google Drive, Dropbox, iCloud or a local folder), and sends you a **weekly email digest**.

## What it does

| When | What |
|---|---|
| **Daily, 07:00** | 1. Queries official APIs: **Grants.gov** (USA) and **EU Funding & Tenders** (Horizon Europe and others), using your topics.<br>2. Runs **Claude Code (headless) with web search** over national sources (Minciencias, SGR, MADR, SENA, iNNpulsa, APC…) and international ones (FONTAGRO, IICA, FAO, IFAD, IDB, CGIAR, GCF/GEF, IDRC, Gates, UKRI, GIZ, AECID, JICA, CYTED…), prioritizing your topics.<br>3. Removes duplicates, updates the local database and writes the outputs to your sync folder. |
| **Weekly, Monday 09:00** | Sends an HTML email (through Outlook) with the week's new national and international calls and the calls closing within 30 days, with the Excel file attached. |

## Requirements
- Windows 10/11 with **Python 3.10+**
- **Outlook desktop** (classic) signed in, used to send the weekly email. Without Outlook, the digest is still saved as HTML in the sync folder.
- **[Claude Code](https://claude.com/claude-code)** installed and signed in (`claude`), used for the AI web research. It is optional: set `claude.enabled = false` to use only the APIs.

## Installation
```powershell
git clone https://github.com/maarango/GrantsSeekerAgro.git
cd GrantsSeekerAgro
pip install -r requirements.txt
python grants_seeker.py setup                                # wizard (see below)
python grants_seeker.py daily                                # first search
powershell -ExecutionPolicy Bypass -File install_tasks.ps1   # schedule daily and weekly runs
```

### Setup wizard
`python grants_seeker.py setup` asks for:
1. **Email(s)** to receive the weekly digest.
2. **Where to sync the calls.** It detects OneDrive (personal and work), Google Drive, Dropbox, iCloud Drive and Documents, or you can give any path.
3. **Topics to explore**, for example `cacao, soil carbon, livestock methane, agroforestry`. English works best for the international APIs. Leave it empty for a broad agricultural focus.
4. **Institution profile** (optional), which guides the AI search.

Your answers are saved in `config.json`, which **stays on your machine and is never committed** (it is in `.gitignore`). Run `setup` again at any time to change them. Defaults for everything else live in `config.example.json`.

## Outputs (in your sync folder)
- `Convocatorias_GrantsSeekerAgro.xlsx`: master list (sheets *Vigentes*, *Cerradas*, *Info*) with status, relevance, deadline, link and so on.
- `Reportes_diarios\YYYY\YYYY-MM-DD_convocatorias.md`: the calls detected each day.
- `Reportes_semanales\YYYY-MM-DD_resumen_semanal.html`: a copy of each weekly email.

## Commands
```powershell
python grants_seeker.py setup                                   # configure email, folder, topics
python grants_seeker.py daily                                   # full search
python grants_seeker.py daily --no-claude                       # APIs only (fast)
python grants_seeker.py daily --topics "coffee rust, irrigation" # topics for this run only
python grants_seeker.py weekly --preview                        # open the email in Outlook without sending
python grants_seeker.py weekly                                  # send the email
python grants_seeker.py export                                  # regenerate the Excel file only
```

## Scheduling
```powershell
powershell -ExecutionPolicy Bypass -File install_tasks.ps1                       # install (runs setup if needed)
powershell -ExecutionPolicy Bypass -File install_tasks.ps1 -DailyTime 06:30 -WeeklyDay Friday
Get-ScheduledTask -TaskName "GrantsSeekerAgro*"                                  # check
Unregister-ScheduledTask -TaskName "GrantsSeekerAgro*" -Confirm:$false           # uninstall
```
The tasks run while you are logged on. If the PC is off at the scheduled time, the task runs when it is turned back on.

## Files
| File | Purpose |
|---|---|
| `grants_seeker.py` | The agent |
| `config.example.json` | Shared defaults: keywords, API queries, web-search source groups, Claude parameters |
| `config.json` | *Your* settings (created by `setup`, not committed) |
| `agent_prompt.md` | Instructions for the AI web-research step |
| `install_tasks.ps1` | Creates the Windows scheduled tasks |
| `data\`, `logs\` | Local database and logs (not committed) |

## Customizing
- **Sources searched by AI:** `claude.search_groups` in `config.example.json`. You can override them in your `config.json`.
- **Relevance:** `core_keywords` (required for API results; your topics are added automatically) and `keywords` (raise relevance).
- **Disable the AI search:** set `"claude": {"enabled": false}` in `config.json`.

> The information is collected automatically. Always check requirements and deadlines on the official page of each call.
