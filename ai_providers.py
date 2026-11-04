"""
AI providers for GrantsSeekerAgro's web-research step.

Each provider receives the research prompt, searches the web with its own
built-in search tool, and returns the model's final text (which must contain
the <<<JSON ... JSON>>> block requested by agent_prompt.md).

Supported:
  claude_code  - Claude Code CLI, using the account you are signed in with (no key)
  anthropic    - Anthropic API (Claude) with server-side web search / web fetch
  openai       - OpenAI Responses API with the web_search tool
  gemini       - Google Gemini API with Google Search grounding
  none         - no AI research (official APIs only)

API keys are stored in Windows Credential Manager (never in config files).
An environment variable (ANTHROPIC_API_KEY / OPENAI_API_KEY / GEMINI_API_KEY)
takes precedence when set.
"""
from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

import requests

PROVIDERS = {
    "claude_code": {"label": "Claude Code (signed-in CLI, no API key)", "env": None,
                    "default_model": ""},
    "anthropic": {"label": "Anthropic API - Claude (API key)", "env": "ANTHROPIC_API_KEY",
                  "default_model": "claude-opus-5-5",
                  "key_url": "https://platform.claude.com/settings/keys"},
    "openai": {"label": "OpenAI - ChatGPT models (API key)", "env": "OPENAI_API_KEY",
               "default_model": "gpt-5",
               "key_url": "https://platform.openai.com/api-keys"},
    "gemini": {"label": "Google Gemini (API key)", "env": "GEMINI_API_KEY",
               "default_model": "gemini-2.5-flash",
               "key_url": "https://aistudio.google.com/app/apikey"},
    "none": {"label": "None - official APIs only (Grants.gov, EU portal)", "env": None,
             "default_model": ""},
}

CRED_PREFIX = "GrantsSeekerAgro/"


class ProviderError(RuntimeError):
    pass


# --------------------------------------------------------------------------- credentials
def get_key(provider: str) -> str | None:
    env = PROVIDERS[provider].get("env")
    if env and os.environ.get(env):
        return os.environ[env]
    try:
        import win32cred
        cred = win32cred.CredRead(CRED_PREFIX + provider, win32cred.CRED_TYPE_GENERIC)
        blob = cred["CredentialBlob"]
        return blob.decode("utf-16-le") if isinstance(blob, bytes) else str(blob)
    except Exception:  # noqa: BLE001 - not stored / not on Windows
        return None


def set_key(provider: str, key: str) -> None:
    import win32cred
    win32cred.CredWrite({
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": CRED_PREFIX + provider,
        "UserName": provider,
        "CredentialBlob": key,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
        "Comment": "API key used by GrantsSeekerAgro",
    }, 0)


def delete_key(provider: str) -> None:
    try:
        import win32cred
        win32cred.CredDelete(CRED_PREFIX + provider, win32cred.CRED_TYPE_GENERIC, 0)
    except Exception:  # noqa: BLE001
        pass


def find_claude_cli(configured: str = "") -> str | None:
    if configured and Path(configured).is_file():
        return configured
    exe = shutil.which("claude")
    if not exe and (Path.home() / ".local" / "bin" / "claude.exe").is_file():
        exe = str(Path.home() / ".local" / "bin" / "claude.exe")
    return exe


def available(provider: str, cfg_claude: dict | None = None) -> bool:
    if provider == "claude_code":
        return bool(find_claude_cli((cfg_claude or {}).get("exe", "")))
    if provider == "none":
        return True
    return bool(get_key(provider))


# --------------------------------------------------------------------------- runners
def run(provider: str, prompt: str, model: str, cfg_claude: dict, cwd: Path) -> str:
    """Run the research prompt and return the model's final text."""
    timeout = int(cfg_claude.get("timeout_minutes", 25)) * 60
    if provider == "claude_code":
        return _run_claude_code(prompt, cfg_claude, cwd, timeout)
    if provider == "anthropic":
        return _run_anthropic(prompt, model, timeout)
    if provider == "openai":
        return _run_openai(prompt, model, timeout)
    if provider == "gemini":
        return _run_gemini(prompt, model, timeout)
    raise ProviderError(f"Unknown AI provider: {provider}")


def _run_claude_code(prompt: str, c: dict, cwd: Path, timeout: int) -> str:
    exe = find_claude_cli(c.get("exe", ""))
    if not exe:
        raise ProviderError("Claude Code CLI not found (install it and run `claude` once to sign in)")
    cmd = [exe, "-p", "--allowedTools", "WebSearch,WebFetch",
           "--max-turns", str(c.get("max_turns", 60)), "--output-format", "text"]
    try:
        proc = subprocess.run(cmd, input=prompt, capture_output=True, text=True, encoding="utf-8",
                              timeout=timeout, cwd=cwd,
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired as e:
        raise ProviderError("Claude Code timed out") from e
    if proc.returncode != 0 and not proc.stdout:
        raise ProviderError(f"Claude Code exited {proc.returncode}: {proc.stderr[-500:]}")
    return proc.stdout or ""


def _run_anthropic(prompt: str, model: str, timeout: int) -> str:
    try:
        import anthropic
    except ImportError as e:
        raise ProviderError("Python package 'anthropic' is missing: pip install anthropic") from e
    key = get_key("anthropic")
    # Without a stored key the SDK resolves ANTHROPIC_API_KEY or an `ant auth login` profile.
    client = anthropic.Anthropic(api_key=key, timeout=timeout) if key else anthropic.Anthropic(timeout=timeout)
    tools = [
        {"type": "web_search_20260209", "name": "web_search", "max_uses": 40},
        {"type": "web_fetch_20260209", "name": "web_fetch", "max_uses": 30},
    ]
    messages: list = [{"role": "user", "content": prompt}]
    response = None
    try:
        for _ in range(8):  # resume server-tool loops that pause (pause_turn)
            with client.beta.messages.stream(
                model=model,
                max_tokens=64000,
                output_config={"effort": "high"},
                tools=tools,
                messages=messages,
                betas=["server-side-fallback-2026-07-01"],
                extra_body={"fallbacks": "default"},  # retry on another model if a request is declined
            ) as stream:
                response = stream.get_final_message()
            if response.stop_reason != "pause_turn":
                break
            messages.append({"role": "assistant", "content": response.content})
    except anthropic.AuthenticationError as e:
        raise ProviderError("Anthropic API key is invalid - run `python grants_seeker.py setup`") from e
    except anthropic.RateLimitError as e:
        raise ProviderError("Anthropic API rate limit reached; will retry next run") from e
    except anthropic.APIStatusError as e:
        raise ProviderError(f"Anthropic API error {e.status_code}: {e.message}") from e
    except anthropic.APIConnectionError as e:
        raise ProviderError("Could not reach the Anthropic API (network)") from e
    if response is None:
        return ""
    if response.stop_reason == "refusal":
        raise ProviderError("Anthropic API declined the request")
    return "".join(b.text for b in response.content if b.type == "text")


def _run_openai(prompt: str, model: str, timeout: int) -> str:
    key = get_key("openai")
    if not key:
        raise ProviderError("No OpenAI API key - run `python grants_seeker.py setup`")
    r = requests.post(
        "https://api.openai.com/v1/responses",
        headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
        json={"model": model, "input": prompt, "tools": [{"type": "web_search"}]},
        timeout=timeout,
    )
    if r.status_code == 401:
        raise ProviderError("OpenAI API key is invalid - run `python grants_seeker.py setup`")
    if r.status_code >= 400:
        raise ProviderError(f"OpenAI API error {r.status_code}: {r.text[:300]}")
    data = r.json()
    texts = [c.get("text", "") for item in data.get("output", []) if item.get("type") == "message"
             for c in item.get("content", []) if c.get("type") == "output_text"]
    return "\n".join(texts) or data.get("output_text", "")


def _raise_for_gemini(r: requests.Response, model: str) -> None:
    if r.status_code < 400:
        return
    if r.status_code == 429:
        raise ProviderError(f"Gemini quota exceeded for model '{model}'. Free-tier keys usually only allow "
                            "Flash models (e.g. gemini-2.5-flash); change the model with `setup` or enable billing.")
    if r.status_code == 404:
        raise ProviderError(f"Gemini model '{model}' not found - choose another model with `setup`")
    if r.status_code in (401, 403) or "API_KEY_INVALID" in r.text:
        raise ProviderError("Gemini API key is invalid or lacks permission - run `python grants_seeker.py setup`")
    raise ProviderError(f"Gemini API error {r.status_code}: {r.text[:300]}")


def _run_gemini(prompt: str, model: str, timeout: int) -> str:
    key = get_key("gemini")
    if not key:
        raise ProviderError("No Gemini API key - run `python grants_seeker.py setup`")
    r = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        headers={"x-goog-api-key": key, "Content-Type": "application/json"},
        json={"contents": [{"role": "user", "parts": [{"text": prompt}]}],
              "tools": [{"google_search": {}}]},
        timeout=timeout,
    )
    _raise_for_gemini(r, model)
    cands = r.json().get("candidates", [])
    if not cands:
        return ""
    return "".join(p.get("text", "") for p in cands[0].get("content", {}).get("parts", []))


# --------------------------------------------------------------------------- connection test
def test(provider: str, model: str, cfg_claude: dict, cwd: Path) -> tuple[bool, str]:
    """Cheap request without web tools, to check the sign-in / key works."""
    prompt = "Reply with exactly: OK"
    try:
        if provider == "claude_code":
            exe = find_claude_cli(cfg_claude.get("exe", ""))
            if not exe:
                return False, "Claude Code CLI not found"
            p = subprocess.run([exe, "-p", prompt], capture_output=True, text=True, encoding="utf-8",
                               timeout=180, cwd=cwd, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            out = (p.stdout or p.stderr).strip()
        elif provider == "anthropic":
            import anthropic
            key = get_key("anthropic")
            client = anthropic.Anthropic(api_key=key) if key else anthropic.Anthropic()
            msg = client.messages.create(model=model, max_tokens=1024, output_config={"effort": "low"},
                                         messages=[{"role": "user", "content": prompt}])
            out = "".join(b.text for b in msg.content if b.type == "text")
        elif provider == "openai":
            r = requests.post("https://api.openai.com/v1/responses", timeout=120,
                              headers={"Authorization": f"Bearer {get_key('openai')}"},
                              json={"model": model, "input": prompt})
            try:
                _raise_for_gemini(r, model)
            except ProviderError as e:
                return False, str(e)
            out = "ok" if r.json().get("output") else r.text[:200]
        elif provider == "gemini":
            r = requests.post(
                f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                headers={"x-goog-api-key": get_key("gemini") or ""}, timeout=120,
                json={"contents": [{"role": "user", "parts": [{"text": prompt}]}]})
            try:
                _raise_for_gemini(r, model)
            except ProviderError as e:
                return False, str(e)
            out = "ok"
        else:
            return True, "no AI provider"
    except Exception as e:  # noqa: BLE001
        return False, str(e)[:300]
    return ("ok" in out.lower()), out[:200]
