"""
GrantsSeekerAgro - autonomous agent that explores national (Colombia) and
international grant calls for agricultural research.

Usage:
    python grants_seeker.py setup              # first-time wizard: email, sync folder, topics
    python grants_seeker.py daily --topics "cacao, soil carbon"   # one-off topic override
    python grants_seeker.py daily              # search sources, update DB, Excel and daily report
    python grants_seeker.py daily --no-claude  # API sources only (fast, no AI web research)
    python grants_seeker.py weekly             # send the weekly email digest via Outlook
    python grants_seeker.py weekly --preview   # open the email in Outlook without sending
    python grants_seeker.py export             # only regenerate the Excel file
"""
from __future__ import annotations

import argparse
import os
import datetime as dt
import hashlib
import html
import json
import logging
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

import requests

import ai_providers as ai

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
LOGS = BASE / "logs"
DB_PATH = DATA / "grants_db.json"
TODAY = dt.date.today()

for _stream in (sys.stdout, sys.stderr):
    if _stream is None:  # pythonw (scheduled task) has no console
        continue
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

DATA.mkdir(exist_ok=True)
LOGS.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler(LOGS / f"{TODAY:%Y-%m}.log", encoding="utf-8"),
    ] + ([logging.StreamHandler(sys.stdout)] if sys.stdout else []),
)
log = logging.getLogger("GrantsSeekerAgro")

FIELDS = [
    "id", "titulo", "entidad", "fuente", "ambito", "pais_entidad", "url",
    "fecha_apertura", "fecha_cierre", "monto", "temas", "elegibilidad_colombia",
    "relevancia", "resumen", "primera_deteccion", "ultima_actualizacion",
]


CONFIG_PATH = BASE / "config.json"            # user settings (local, not in git)
EXAMPLE_PATH = BASE / "config.example.json"   # shared defaults (in git)


# --------------------------------------------------------------------------- config
def deep_merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in over.items():
        out[k] = deep_merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def split_list(text: str) -> list[str]:
    return [t.strip() for t in re.split(r"[,;\n]", text or "") if t.strip()]


def load_config(topics_override: list[str] | None = None) -> dict:
    """Defaults from config.example.json overridden by the user's config.json."""
    cfg = json.loads(EXAMPLE_PATH.read_text(encoding="utf-8"))
    if not CONFIG_PATH.exists():
        if sys.stdin and sys.stdin.isatty():
            print("No configuration found - starting first-time setup.\n")
            run_setup()
        else:
            log.error("config.json not found. Run:  python grants_seeker.py setup")
            sys.exit(2)
    cfg = deep_merge(cfg, json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    if topics_override:
        cfg["topics"] = topics_override
    return apply_topics(cfg)


def apply_topics(cfg: dict) -> dict:
    """User topics become API queries, relevance keywords and priorities for the web research."""
    topics = [t for t in cfg.get("topics") or [] if t]
    if topics:
        cfg["grants_gov"]["queries"] = topics
        cfg["eu_portal"]["queries"] = topics
        cfg["core_keywords"] = list(dict.fromkeys(cfg["core_keywords"] + [t.lower() for t in topics]))
    return cfg


def detect_sync_folders() -> list[tuple[str, Path]]:
    """Find cloud-sync folders available on this Windows machine."""
    import os
    home = Path.home()
    found: list[tuple[str, Path]] = []
    for var, label in (("OneDriveCommercial", "OneDrive (work/school)"),
                       ("OneDriveConsumer", "OneDrive (personal)"), ("OneDrive", "OneDrive")):
        p = os.environ.get(var)
        if p and Path(p).is_dir() and all(Path(p) != f for _, f in found):
            found.append((label, Path(p)))
    for p in home.glob("OneDrive - *"):
        if p.is_dir() and all(p != f for _, f in found):
            found.append((f"OneDrive ({p.name[11:]})", p))
    gd = [Path(f"{d}:/{n}") for d in "DEFGHIJKLMNOPQRSTUVWXYZ" for n in ("My Drive", "Mi unidad")]
    gd += [home / "Google Drive" / "My Drive", home / "Google Drive", home / "My Drive"]
    for p in gd:
        if p.is_dir():
            found.append(("Google Drive", p))
            break
    dbx_info = Path(os.environ.get("LOCALAPPDATA", "")) / "Dropbox" / "info.json"
    dbx = home / "Dropbox"
    try:
        dbx = Path(next(iter(json.loads(dbx_info.read_text()).values()))["path"])
    except Exception:  # noqa: BLE001
        pass
    if dbx.is_dir():
        found.append(("Dropbox", dbx))
    if (home / "iCloudDrive").is_dir():
        found.append(("iCloud Drive", home / "iCloudDrive"))
    found.append(("Local folder (Documents)", home / "Documents"))
    return found


def setup_ai(cur: dict, defaults: dict) -> dict:
    """Wizard step: choose the AI service used for web research and sign in to it."""
    import getpass

    claude_cfg = deep_merge(defaults["claude"], cur.get("claude", {}))
    current = (cur.get("ai") or {}).get("provider") or ("claude_code" if CONFIG_PATH.exists() else "")
    keys = list(ai.PROVIDERS)
    print("\n5) AI service used to search the web for calls (national sources, foundations, etc.).")
    for i, k in enumerate(keys, 1):
        if k == "claude_code":
            state = "detected" if ai.available(k, claude_cfg) else "not installed"
        elif k == "none":
            state = ""
        else:
            state = "key stored" if ai.get_key(k) else ""
        print(f"   {i}. {ai.PROVIDERS[k]['label']:<52} {state}")
    if current in keys:
        default_idx = str(keys.index(current) + 1)
    elif ai.available("claude_code", claude_cfg):
        default_idx = str(keys.index("claude_code") + 1)
    else:
        default_idx = str(len(keys))
    while True:
        choice = input(f"   Choose an option [{default_idx}]: ").strip() or default_idx
        if choice.isdigit() and 1 <= int(choice) <= len(keys):
            provider = keys[int(choice) - 1]
            break
        print("   Invalid option.")
    info = ai.PROVIDERS[provider]
    model = ""
    if provider == "claude_code" and not ai.available(provider, claude_cfg):
        print("   Claude Code is not installed. Install it from https://claude.com/claude-code,")
        print("   then run `claude` once to sign in. The agent will use it automatically.")
    if info.get("env"):
        have = ai.get_key(provider)
        print(f"   Get an API key at: {info['key_url']}")
        if os.environ.get(info["env"]):
            print(f"   Using the key from environment variable {info['env']}.")
        else:
            hint = ", Enter = keep stored key" if have else ""
            while True:
                key = getpass.getpass(f"   Paste API key (hidden{hint}): ").strip()
                if key:
                    ai.set_key(provider, key)
                    print(f"   Key saved in Windows Credential Manager (GrantsSeekerAgro/{provider}).")
                    break
                if have:
                    break
                if provider == "anthropic":
                    print("   No key: the Anthropic SDK will use ANTHROPIC_API_KEY or an `ant auth login` profile.")
                    break
                print("   An API key is required for this service.")
        prev = cur.get("ai") or {}
        old_model = prev.get("model") if prev.get("provider") == provider else ""
        default_model = old_model or info["default_model"]
        model = input(f"   Model [{default_model}]: ").strip() or default_model
    if provider != "none":
        if input("   Test the connection now? [Y/n]: ").strip().lower() in ("", "y", "s", "yes", "si"):
            ok, msg = ai.test(provider, model or info["default_model"], claude_cfg, DATA)
            print("   Connection OK." if ok else f"   Connection FAILED: {msg}\n   You can re-run setup later.")
    return {"provider": provider, "model": model}


def run_setup() -> None:
    """Interactive first-time / re-configuration wizard. Writes config.json."""
    cur = json.loads(CONFIG_PATH.read_text(encoding="utf-8")) if CONFIG_PATH.exists() else {}
    defaults = json.loads(EXAMPLE_PATH.read_text(encoding="utf-8"))

    def ask(q: str, default: str = "") -> str:
        ans = input(f"{q}" + (f" [{default}]" if default else "") + ": ").strip()
        return ans or default

    print("=" * 64 + "\n GrantsSeekerAgro - setup\n" + "=" * 64)

    # 1) Email
    print("\n1) Weekly digest email (sent through your Outlook desktop app).")
    while True:
        emails = split_list(ask("   Recipient email(s), comma-separated", ", ".join(cur.get("email", {}).get("to", []))))
        if emails and all(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", e) for e in emails):
            break
        print("   Please enter at least one valid email address.")

    # 2) Sync folder
    print("\n2) Where should the calls be synced? (Excel + daily/weekly reports)")
    opts = detect_sync_folders()
    for i, (label, p) in enumerate(opts, 1):
        print(f"   {i}. {label:<26} {p}")
    print(f"   {len(opts) + 1}. Custom path")
    if cur.get("output_folder"):
        print(f"   Enter = keep current: {cur['output_folder']}")
    while True:
        choice = ask("   Choose an option").strip()
        if not choice and cur.get("output_folder"):
            folder = Path(cur["output_folder"])
            break
        if choice.isdigit() and 1 <= int(choice) <= len(opts):
            sub = ask("   Subfolder name", "Convocatorias_GrantsSeekerAgro")
            folder = opts[int(choice) - 1][1] / sub
            break
        if choice.isdigit() and int(choice) == len(opts) + 1:
            folder = Path(ask("   Full folder path").strip('"'))
            if str(folder):
                break
        print("   Invalid option.")
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        print(f"   Warning: could not create folder now ({e}); it will be retried at run time.")

    # 3) Topics
    print("\n3) Topics to explore, comma-separated (English works best for the international APIs).")
    print("   Example: cacao, soil carbon, livestock methane, plant genetic resources, agroforestry")
    print("   Leave empty to use the default broad agricultural focus.")
    topics = split_list(ask("   Topics", ", ".join(cur.get("topics", []))))

    # 4) Profile
    print("\n4) Short description of your institution / research profile (guides the AI web search).")
    focus = ask("   Profile (Enter = keep current/default)", "")
    focus = focus or cur.get("research_focus") or defaults["research_focus"]

    # 5) AI service
    ai_cfg = setup_ai(cur, defaults)

    new = dict(cur)
    new.update({"email": {**cur.get("email", {}), "to": emails}, "output_folder": str(folder),
                "topics": topics, "research_focus": focus, "ai": ai_cfg})
    CONFIG_PATH.write_text(json.dumps(new, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nSaved {CONFIG_PATH.name}.  Email: {', '.join(emails)}\n  Folder: {folder}\n"
          f"  Topics: {', '.join(topics) or '(default agricultural focus)'}\n"
          f"  AI service: {ai_cfg['provider']} {ai_cfg.get('model', '')}")
    print("Next: run  python grants_seeker.py daily   and   install_tasks.ps1  to schedule it.\n")


def load_db() -> dict:
    if DB_PATH.exists():
        return json.loads(DB_PATH.read_text(encoding="utf-8"))
    return {}


def save_db(db: dict) -> None:
    tmp = DB_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(db, ensure_ascii=False, indent=1), encoding="utf-8")
    if DB_PATH.exists():
        shutil.copy2(DB_PATH, DATA / "grants_db.bak.json")
    tmp.replace(DB_PATH)


def norm_url(url: str) -> str:
    url = (url or "").strip().lower()
    url = re.sub(r"^https?://(www\.)?", "", url)
    return url.rstrip("/").split("#")[0]


def norm_title(t: str) -> str:
    return re.sub(r"[^a-z0-9áéíóúñü]+", " ", (t or "").lower()).strip()


def make_id(source: str, key: str) -> str:
    return f"{source}:{hashlib.sha1(key.encode()).hexdigest()[:12]}"


def parse_date(s: str) -> str:
    """Return YYYY-MM-DD or ''."""
    s = (s or "").strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    m = re.match(r"(\d{4}-\d{2}-\d{2})", s)
    return m.group(1) if m else ""


def status_of(rec: dict) -> str:
    close = rec.get("fecha_cierre") or ""
    opening = rec.get("fecha_apertura") or ""
    if close and close < TODAY.isoformat():
        return "Cerrada"
    if opening and opening > TODAY.isoformat():
        return "Próxima"
    if close:
        return "Abierta"
    return "Sin fecha"


def keyword_relevance(text: str, keywords: list[str]) -> int:
    """Number of distinct keywords found at a word start (e.g. 'agri' matches 'agriculture')."""
    text = (text or "").lower()
    return sum(1 for k in keywords if re.search("(?<![a-z0-9])" + re.escape(k.lower()), text))


def api_relevance(cfg: dict, title: str, desc: str = "") -> str | None:
    """Relevance for API results; None means 'not relevant, discard'."""
    core_title = keyword_relevance(title, cfg["core_keywords"])
    core_all = keyword_relevance(f"{title} {desc}", cfg["core_keywords"])
    extra = keyword_relevance(f"{title} {desc}", cfg["keywords"])
    if core_title == 0 and core_all < 2:
        return None
    if core_title and core_all + extra >= 4:
        return "Alta"
    return "Media" if core_title else "Baja"


# --------------------------------------------------------------------------- sources
def source_grants_gov(cfg: dict) -> list[dict]:
    c = cfg["grants_gov"]
    out: dict[str, dict] = {}
    for q in c["queries"]:
        try:
            r = requests.post(
                "https://api.grants.gov/v1/api/search2",
                json={"keyword": q, "oppStatuses": "forecasted|posted", "rows": c["rows_per_query"]},
                timeout=60,
            )
            r.raise_for_status()
            hits = r.json().get("data", {}).get("oppHits", [])
        except Exception as e:  # noqa: BLE001
            log.warning("Grants.gov query '%s' failed: %s", q, e)
            continue
        for h in hits:
            title = h.get("title", "")
            rel = api_relevance(cfg, title)
            if rel is None:
                continue
            if not re.search(r"international|global|foreign|latin|americas|colombia|borlaug|cochran|developing", title, re.I):
                rel = "Baja"  # most Grants.gov programs are US-domestic
            url = f"https://www.grants.gov/search-results-detail/{h['id']}"
            out[h["id"]] = {
                "id": make_id("grantsgov", str(h["id"])),
                "titulo": title,
                "entidad": h.get("agency", ""),
                "fuente": "Grants.gov (API)",
                "ambito": "Internacional",
                "pais_entidad": "Estados Unidos",
                "url": url,
                "fecha_apertura": parse_date(h.get("openDate", "")),
                "fecha_cierre": parse_date(h.get("closeDate", "")),
                "monto": "",
                "temas": q,
                "elegibilidad_colombia": "Verificar",
                "relevancia": rel,
                "resumen": f"Oportunidad {h.get('number', '')} ({h.get('oppStatus', '')}). Verificar elegibilidad de entidades extranjeras.",
            }
        time.sleep(1)
    log.info("Grants.gov: %d candidates", len(out))
    return list(out.values())


def source_eu_portal(cfg: dict) -> list[dict]:
    c = cfg["eu_portal"]
    out: dict[str, dict] = {}
    query = {"bool": {"must": [
        {"terms": {"type": ["1", "2", "8"]}},
        {"terms": {"status": ["31094501", "31094502"]}},  # forthcoming, open
    ]}}
    files = {
        "query": ("blob", json.dumps(query), "application/json"),
        "languages": ("blob", json.dumps(["en"]), "application/json"),
        "sort": ("blob", json.dumps({"field": "sortStatus", "order": "ASC"}), "application/json"),
    }
    for q in c["queries"]:
        try:
            r = requests.post(
                "https://api.tech.ec.europa.eu/search-api/prod/rest/search",
                params={"apiKey": "SEDIA", "text": q, "pageSize": c["page_size"], "pageNumber": 1},
                files=files, timeout=60,
            )
            r.raise_for_status()
            results = r.json().get("results", [])
        except Exception as e:  # noqa: BLE001
            log.warning("EU portal query '%s' failed: %s", q, e)
            continue
        for x in results:
            m = x.get("metadata", {})
            first = lambda k: (m.get(k) or [""])[0]  # noqa: E731
            ident = first("identifier") or x.get("reference", "")
            title = first("title") or first("callTitle") or x.get("summary") or x.get("content") or ""
            desc = re.sub(r"<[^>]+>", " ", " ".join(m.get("descriptionByte", []) or []))[:6000]
            rel = api_relevance(cfg, title, desc)
            if rel is None or not title.strip():
                continue
            deadlines = sorted(parse_date(d) for d in (m.get("deadlineDate") or []) if d)
            future = [d for d in deadlines if d >= TODAY.isoformat()]
            close = future[0] if future else (deadlines[-1] if deadlines else "")
            status = "Abierta" if first("status") == "31094502" else "Próxima"
            if ident in out:
                continue
            out[ident] = {
                "id": make_id("eu", ident),
                "titulo": f"{title} ({ident})" if ident else title,
                "entidad": "Comisión Europea – " + (first("frameworkProgramme") and "Horizon Europe / EU" or "EU"),
                "fuente": "EU Funding & Tenders (API)",
                "ambito": "Internacional",
                "pais_entidad": "Unión Europea",
                "url": x.get("url", ""),
                "fecha_apertura": parse_date(first("startDate")),
                "fecha_cierre": close,
                "monto": "",
                "temas": q,
                "elegibilidad_colombia": "Verificar",
                "relevancia": rel,
                "resumen": f"Convocatoria UE ({status}). Verificar participación/financiación de entidades colombianas.",
            }
        time.sleep(1)
    log.info("EU portal: %d candidates", len(out))
    return list(out.values())


SOURCE_LABEL = {"claude_code": "Claude", "anthropic": "Claude API", "openai": "OpenAI", "gemini": "Gemini"}


def normalize_scope(rec: dict) -> str:
    """AI services sometimes answer 'Nacional / Internacional'; a Colombian funder means national."""
    amb = (rec.get("ambito") or "").lower()
    if "colombia" in (rec.get("pais_entidad") or "").lower() or amb.strip().startswith("nacional"):
        return "Nacional"
    return "Internacional"


def ai_provider(cfg: dict) -> tuple[str, str]:
    """Resolve (provider, model) from config; legacy configs fall back to Claude Code."""
    a = cfg.get("ai") or {}
    provider = a.get("provider") or ("claude_code" if cfg["claude"].get("enabled", True) else "none")
    model = a.get("model") or ai.PROVIDERS.get(provider, {}).get("default_model", "")
    return provider, model


def source_ai(cfg: dict, db: dict) -> list[dict]:
    """Run one AI web-research pass per source group, merging after each so later passes skip known calls."""
    provider, model = ai_provider(cfg)
    if provider == "none":
        return []
    if not ai.available(provider, cfg["claude"]):
        log.error("AI provider '%s' is not available (not signed in / no key). Run: python grants_seeker.py setup",
                  provider)
        return []
    log.info("AI web research with %s %s", provider, model)
    out: list[dict] = []
    for group in cfg["claude"].get("search_groups") or [{"name": "General", "sources": ""}]:
        items = ai_pass(cfg, db, group, provider, model)
        merge(db, items)
        out += items
    return out


def ai_pass(cfg: dict, db: dict, group: dict, provider: str, model: str) -> list[dict]:
    c = cfg["claude"]
    # Give the agent the known, still-relevant calls so it focuses on new ones.
    known = [r for r in db.values() if status_of(r) != "Cerrada" and not r["fuente"].endswith("(API)")]
    known_txt = "\n".join(f"- {r['titulo']} | {r['url']}" for r in known[-250:]) or "(none yet)"
    prompt = (BASE / "agent_prompt.md").read_text(encoding="utf-8")
    prompt = (prompt.replace("{TODAY}", TODAY.isoformat())
                    .replace("{FOCUS}", cfg["research_focus"])
                    .replace("{TOPICS}", ", ".join(cfg.get("topics") or []) or
                             "(no specific topics - use the institutional focus above)")
                    .replace("{MAX_RESULTS}", str(c["max_results"]))
                    .replace("{GROUP}", group["name"])
                    .replace("{SOURCES}", group["sources"])
                    .replace("{KNOWN}", known_txt))
    log.info("Web research [%s] (timeout %s min)...", group["name"], c["timeout_minutes"])
    try:
        raw = ai.run(provider, prompt, model, c, DATA) or ""
    except ai.ProviderError as e:
        log.error("Web research [%s] failed: %s", group["name"], e)
        return []
    (DATA / f"last_ai_output_{re.sub(r'[^A-Za-z]+', '_', group['name'])[:40]}.txt").write_text(raw, encoding="utf-8")
    m = re.search(r"<<<JSON\s*(\[.*\])\s*JSON>>>", raw, re.S) or re.search(r"(\[\s*\{.*\}\s*\])", raw, re.S)
    if not m:
        log.error("AI output had no JSON. See data/last_ai_output_*.txt")
        return []
    try:
        items = json.loads(m.group(1))
    except json.JSONDecodeError as e:
        log.error("Could not parse AI JSON: %s", e)
        return []
    out = []
    for it in items:
        if not isinstance(it, dict) or not it.get("titulo"):
            continue
        rec = {k: str(it.get(k, "") or "").strip() for k in FIELDS}
        rec["fecha_apertura"] = parse_date(rec["fecha_apertura"])
        rec["fecha_cierre"] = parse_date(rec["fecha_cierre"])
        rec["fuente"] = f"Búsqueda web ({SOURCE_LABEL.get(provider, provider)})"
        rec["ambito"] = normalize_scope(rec)
        rec["id"] = make_id("web", norm_url(rec["url"]) or norm_title(rec["titulo"]))
        out.append(rec)
    log.info("Web research [%s]: %d candidates", group["name"], len(out))
    return out


# --------------------------------------------------------------------------- merge
def find_similar(db: dict, it: dict) -> str | None:
    """Web results describe the same call with different wording/URLs: match on numbered
    calls ("Convocatoria 54") from the same funder, or on very similar titles."""
    import difflib
    t = norm_title(it["titulo"])
    num = re.search(r"(?:convocatoria|call|conv)\s*n?\D{0,3}(\d{2,4})", t)
    funder = norm_title(it.get("entidad", "")).split(" ")[0:1]
    for k, r in db.items():
        if r["fuente"].endswith("(API)") or it["fuente"].endswith("(API)"):
            continue
        rt = norm_title(r["titulo"])
        if num:
            rnum = re.search(r"(?:convocatoria|call|conv)\s*n?\D{0,3}(\d{2,4})", rt)
            if rnum and rnum.group(1) == num.group(1) and (
                    not funder or funder[0] in norm_title(r.get("entidad", "")) or
                    (r.get("fecha_cierre") and r.get("fecha_cierre") == it.get("fecha_cierre"))):
                return k
            continue
        if difflib.SequenceMatcher(None, t, rt).ratio() >= 0.85:
            return k
    return None


def merge(db: dict, new_items: list[dict]) -> list[dict]:
    """Insert/update items in db. Returns the list of newly discovered records."""
    by_url = {norm_url(r["url"]): k for k, r in db.items() if r.get("url")}
    by_title = {norm_title(r["titulo"]): k for k, r in db.items()}
    added = []
    now = TODAY.isoformat()
    for it in new_items:
        key = it["id"] if it["id"] in db else by_url.get(norm_url(it["url"])) or by_title.get(norm_title(it["titulo"]))
        if not key:
            key = find_similar(db, it)
        if key and key in db:
            old = db[key]
            for f in FIELDS:
                if f in ("id", "primera_deteccion"):
                    continue
                if it.get(f) and it[f] != old.get(f):
                    old[f] = it[f]
            old["ultima_actualizacion"] = now
            continue
        if it.get("fecha_cierre") and it["fecha_cierre"] < now:
            continue  # already closed, don't add
        it["primera_deteccion"] = now
        it["ultima_actualizacion"] = now
        db[it["id"]] = it
        by_url[norm_url(it["url"])] = it["id"]
        by_title[norm_title(it["titulo"])] = it["id"]
        added.append(it)
    return added


# --------------------------------------------------------------------------- outputs
def sort_key(r: dict):
    order = {"Abierta": 0, "Próxima": 1, "Sin fecha": 2, "Cerrada": 3}
    rel = {"Alta": 0, "Media": 1, "Baja": 2}
    return (order.get(status_of(r), 9), r.get("fecha_cierre") or "9999", rel.get(r.get("relevancia"), 3))


def export_excel(cfg: dict, db: dict) -> Path | None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    cols = [("Estado", None), ("Relevancia", "relevancia"), ("Ámbito", "ambito"), ("Título", "titulo"),
            ("Entidad", "entidad"), ("País entidad", "pais_entidad"), ("Fecha apertura", "fecha_apertura"),
            ("Fecha cierre", "fecha_cierre"), ("Monto", "monto"), ("Temas", "temas"),
            ("Elegible Colombia", "elegibilidad_colombia"), ("Resumen", "resumen"), ("Enlace", "url"),
            ("Fuente", "fuente"), ("Detectada", "primera_deteccion"), ("Actualizada", "ultima_actualizacion")]
    widths = [11, 10, 13, 60, 30, 16, 13, 13, 18, 30, 12, 70, 45, 22, 12, 12]
    fills = {"Abierta": "C6EFCE", "Próxima": "FFEB9C", "Cerrada": "D9D9D9", "Sin fecha": "DDEBF7"}

    wb = Workbook()
    for sheet_name, rows in (
        ("Vigentes", [r for r in db.values() if status_of(r) != "Cerrada"]),
        ("Cerradas", [r for r in db.values() if status_of(r) == "Cerrada"]),
    ):
        ws = wb.active if sheet_name == "Vigentes" else wb.create_sheet(sheet_name)
        ws.title = sheet_name
        ws.append([c[0] for c in cols])
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F6E43")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        for r in sorted(rows, key=sort_key):
            st = status_of(r)
            ws.append([st] + [r.get(f, "") for _, f in cols[1:]])
            row = ws.max_row
            ws.cell(row, 1).fill = PatternFill("solid", fgColor=fills[st])
            link = ws.cell(row, 13)
            if r.get("url"):
                link.hyperlink = r["url"]
                link.font = Font(color="0563C1", underline="single")
            for c in (4, 12):
                ws.cell(row, c).alignment = Alignment(wrap_text=True, vertical="top")
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    ws = wb.create_sheet("Info")
    ws.append(["Generado por", cfg["agent_name"]])
    ws.append(["Última actualización", dt.datetime.now().strftime("%Y-%m-%d %H:%M")])
    ws.append(["Total registros", len(db)])
    ws.append(["Nota", "Verifique siempre requisitos y fechas en la página oficial de cada convocatoria."])

    out_dir = Path(cfg["output_folder"])
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / cfg["excel_filename"]
    try:
        wb.save(target)
    except PermissionError:
        target = out_dir / f"{target.stem}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"
        log.warning("Excel is open/locked; saved copy as %s", target.name)
        wb.save(target)
    log.info("Excel written: %s", target)
    return target


def write_daily_report(cfg: dict, added: list[dict], db: dict) -> Path:
    out_dir = Path(cfg["output_folder"]) / "Reportes_diarios" / f"{TODAY:%Y}"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"{TODAY.isoformat()}_convocatorias.md"
    lines = [f"# {cfg['agent_name']} – Reporte diario {TODAY.isoformat()}", "",
             f"Nuevas convocatorias detectadas hoy: **{len(added)}**  ",
             f"Total vigentes en la base: **{sum(1 for r in db.values() if status_of(r) != 'Cerrada')}**", ""]
    for scope in ("Nacional", "Internacional"):
        items = sorted([r for r in added if r.get("ambito") == scope], key=sort_key)
        if not items:
            continue
        lines += [f"## {scope}", ""]
        for r in items:
            lines += [
                f"### {r['titulo']}",
                f"- **Entidad:** {r.get('entidad', '')}",
                f"- **Cierre:** {r.get('fecha_cierre') or 'Sin fecha'} | **Estado:** {status_of(r)} | **Relevancia:** {r.get('relevancia', '')}",
                f"- **Monto:** {r.get('monto') or '—'} | **Elegible Colombia:** {r.get('elegibilidad_colombia') or '—'}",
                f"- **Temas:** {r.get('temas', '')}",
                f"- {r.get('resumen', '')}",
                f"- 🔗 {r.get('url', '')}",
                "",
            ]
    path.write_text("\n".join(lines), encoding="utf-8")
    log.info("Daily report written: %s", path)
    return path


# --------------------------------------------------------------------------- email
def build_email_html(cfg: dict, db: dict) -> tuple[str, int, int]:
    week_ago = (TODAY - dt.timedelta(days=7)).isoformat()
    soon = (TODAY + dt.timedelta(days=cfg["email"]["closing_soon_days"])).isoformat()
    new = sorted([r for r in db.values() if r.get("primera_deteccion", "") >= week_ago and status_of(r) != "Cerrada"],
                 key=sort_key)
    closing = sorted([r for r in db.values() if status_of(r) == "Abierta" and r["fecha_cierre"] <= soon
                      and r not in new], key=sort_key)
    e = html.escape

    def table(rows: list[dict]) -> str:
        if not rows:
            return "<p style='color:#666'><i>Sin registros.</i></p>"
        head = "".join(f"<th style='background:#1F6E43;color:#fff;padding:6px;text-align:left'>{h}</th>"
                       for h in ("Título", "Entidad", "Ámbito", "Cierre", "Relevancia", "Elegible CO"))
        body = ""
        for i, r in enumerate(rows):
            bg = "#F3F8F4" if i % 2 else "#FFFFFF"
            body += (f"<tr style='background:{bg};vertical-align:top'>"
                     f"<td style='padding:6px'><a href='{e(r.get('url', ''))}'><b>{e(r['titulo'])}</b></a>"
                     f"<br><span style='color:#555;font-size:12px'>{e(r.get('resumen', ''))}</span></td>"
                     f"<td style='padding:6px'>{e(r.get('entidad', ''))}</td>"
                     f"<td style='padding:6px'>{e(r.get('ambito', ''))}</td>"
                     f"<td style='padding:6px;white-space:nowrap'>{e(r.get('fecha_cierre') or 'Sin fecha')}</td>"
                     f"<td style='padding:6px'>{e(r.get('relevancia', ''))}</td>"
                     f"<td style='padding:6px'>{e(r.get('elegibilidad_colombia', ''))}</td></tr>")
        return f"<table style='border-collapse:collapse;width:100%;font-size:13px'><tr>{head}</tr>{body}</table>"

    nat = [r for r in new if r.get("ambito") == "Nacional"]
    intl = [r for r in new if r.get("ambito") != "Nacional"]
    folder = cfg["output_folder"]
    body = f"""
<div style="font-family:Segoe UI,Arial,sans-serif;color:#222">
<h2 style="color:#1F6E43;margin-bottom:4px">{e(cfg['agent_name'])} – Resumen semanal</h2>
<p style="margin-top:0;color:#555">Semana del {week_ago} al {TODAY.isoformat()}</p>
<p><b>{len(new)}</b> convocatorias nuevas esta semana ({len(nat)} nacionales, {len(intl)} internacionales) ·
<b>{len(closing)}</b> convocatorias abiertas que cierran en los próximos {cfg['email']['closing_soon_days']} días ·
<b>{sum(1 for r in db.values() if status_of(r) != 'Cerrada')}</b> vigentes en total.</p>
<h3 style="color:#1F6E43">🇨🇴 Nuevas – Nacionales</h3>{table(nat)}
<h3 style="color:#1F6E43">🌎 Nuevas – Internacionales</h3>{table(intl)}
<h3 style="color:#B45F06">⏰ Cierran pronto (ya reportadas)</h3>{table(closing)}
<p style="font-size:12px;color:#666;margin-top:20px">Base completa: <b>{e(cfg['excel_filename'])}</b> en<br>
{e(folder)}<br>Reportes diarios en la subcarpeta <i>Reportes_diarios</i>.<br>
Información recopilada automáticamente; verifique requisitos y fechas en la página oficial de cada convocatoria.</p>
</div>"""
    return body, len(new), len(closing)


def send_weekly(cfg: dict, db: dict, preview: bool = False, to: list[str] | None = None) -> None:
    import win32com.client  # pywin32

    body, n_new, n_closing = build_email_html(cfg, db)
    out_dir = Path(cfg["output_folder"]) / "Reportes_semanales"
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / f"{TODAY.isoformat()}_resumen_semanal.html").write_text(
        f"<html><head><meta charset='utf-8'></head><body>{body}</body></html>", encoding="utf-8")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:  # noqa: BLE001
        log.error("Outlook desktop is not available (%s). The digest was saved in %s", e, out_dir)
        return
    mail = outlook.CreateItem(0)
    mail.To = "; ".join(to or cfg["email"]["to"])
    if cfg["email"].get("cc"):
        mail.CC = "; ".join(cfg["email"]["cc"])
    mail.Subject = f"{cfg['email']['subject_prefix']} – {TODAY.isoformat()} ({n_new} nuevas, {n_closing} cierran pronto)"
    mail.HTMLBody = body
    if cfg["email"].get("attach_excel"):
        xlsx = Path(cfg["output_folder"]) / cfg["excel_filename"]
        if xlsx.exists():
            mail.Attachments.Add(str(xlsx))
    if preview:
        mail.Display()
        log.info("Weekly email opened for preview (not sent)")
    else:
        recipients = mail.To  # read before Send(): the item is moved to Sent Items afterwards
        mail.Send()
        log.info("Weekly email sent to %s", recipients)


# --------------------------------------------------------------------------- main
def run_daily(cfg: dict, use_ai: bool) -> None:
    db = load_db()
    items: list[dict] = []
    if cfg["grants_gov"]["enabled"]:
        items += source_grants_gov(cfg)
    if cfg["eu_portal"]["enabled"]:
        items += source_eu_portal(cfg)
    if use_ai:
        items += source_ai(cfg, db)
    merge(db, items)
    added = [r for r in db.values() if r.get("primera_deteccion") == TODAY.isoformat()]
    save_db(db)
    log.info("New calls detected today: %d | DB total: %d", len(added), len(db))
    export_excel(cfg, db)
    write_daily_report(cfg, added, db)
    (DATA / "last_daily_run.txt").write_text(dt.datetime.now().isoformat(), encoding="utf-8")


def main() -> None:
    ap = argparse.ArgumentParser(description="GrantsSeekerAgro")
    ap.add_argument("mode", choices=["setup", "daily", "weekly", "export"])
    ap.add_argument("--topics", help='comma-separated topics for this run only, e.g. "cacao, soil carbon"')
    ap.add_argument("--no-ai", "--no-claude", dest="no_ai", action="store_true", help="skip the AI web research step")
    ap.add_argument("--ai", choices=list(ai.PROVIDERS), help="AI service for this run only")
    ap.add_argument("--preview", action="store_true", help="weekly: display email instead of sending")
    ap.add_argument("--to", nargs="+", help="weekly: override recipients")
    a = ap.parse_args()
    if a.mode == "setup":
        run_setup()
        return
    cfg = load_config(split_list(a.topics) if a.topics else None)
    if not cfg.get("output_folder"):
        log.error("output_folder is empty. Run:  python grants_seeker.py setup")
        sys.exit(2)
    if a.ai:
        cfg["ai"] = {"provider": a.ai, "model": ""}
    if cfg.get("topics"):
        log.info("Topics: %s", ", ".join(cfg["topics"]))
    log.info("=== %s %s ===", cfg["agent_name"], a.mode)
    try:
        if a.mode == "daily":
            run_daily(cfg, use_ai=not a.no_ai)
        elif a.mode == "weekly":
            send_weekly(cfg, load_db(), preview=a.preview, to=a.to)
        else:
            export_excel(cfg, load_db())
    except Exception:
        log.exception("Run failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
