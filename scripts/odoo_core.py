#!/usr/bin/env python3
"""odoo_core.py — shared library for all odoo-cli commands.

Extracted from odoo_handshake.py. Provides:
- Config loading & profile resolution
- XML-RPC connection (dict-based)
- Generic RPC execution
- Hybrid knowledge retrieval (memory fast-lookup + runtime query)
"""

from __future__ import annotations
import json
import os
import ssl
import sys
import urllib.parse
import xmlrpc.client
from pathlib import Path


# ── .env auto-load ──────────────────────────────────────────────────────────
# Any script that imports this module gets prod creds without manual
# `set -a; . ./.env; set +a`. Path resolved relative to this file so
# cwd-independent. setdefault so an already-exported env wins over .env file.

_ENV_FILE = Path(__file__).resolve().parent.parent / ".env"
if _ENV_FILE.exists() and not os.environ.get("_ODOO_ENV_LOADED"):
    for _line in _ENV_FILE.read_text(encoding="utf-8").splitlines():
        _line = _line.strip()
        if not _line or _line.startswith("#") or "=" not in _line:
            continue
        _k, _v = _line.split("=", 1)
        os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))
    os.environ["_ODOO_ENV_LOADED"] = "1"


# ── Config ──────────────────────────────────────────────────────────────────

DEFAULT_CONFIG_CANDIDATES = [
    Path("config/odoo-environments.json"),
    Path("config/odoo-environments.example.json"),
]

MEMORY_DIR = Path("memory")


def load_config(config_path_arg: str | None = None) -> tuple[dict, Path]:
    candidates = [Path(config_path_arg)] if config_path_arg else DEFAULT_CONFIG_CANDIDATES
    for candidate in candidates:
        if candidate.exists():
            with candidate.open("r", encoding="utf-8") as handle:
                return json.load(handle), candidate
    raise FileNotFoundError("No config file found")


def merge_profile(config: dict, profile_name: str) -> dict:
    profiles = config.get("profiles", {})
    if profile_name not in profiles:
        raise KeyError(f"Unknown profile: {profile_name}")
    profile = dict(profiles[profile_name])
    inherit_from = profile.get("inheritAuthFrom")
    if inherit_from:
        if inherit_from not in profiles:
            raise KeyError(f"Profile {profile_name} inherits from unknown profile: {inherit_from}")
        parent = dict(profiles[inherit_from])
        parent.update(profile)
        profile = parent
    profile["name"] = profile_name
    return profile


def resolve_profile(config: dict, profile_name: str, base_url_override: str | None = None, db_override: str | None = None) -> dict:
    profile = merge_profile(config, profile_name)
    if base_url_override:
        profile["baseUrl"] = base_url_override
    if db_override:
        profile["db"] = db_override
    return profile


# ── URL helpers ──────────────────────────────────────────────────────────────

def normalize_base_url(base_url: str) -> str:
    parsed = urllib.parse.urlparse(base_url.rstrip("/"))
    path = parsed.path.rstrip("/")
    if path == "/odoo":
        path = ""
    return urllib.parse.urlunparse((parsed.scheme, parsed.netloc, path, "", "", "")).rstrip("/")


def get_endpoint(base_url: str, service: str) -> str:
    return f"{normalize_base_url(base_url)}/xmlrpc/2/{service}"


def read_secret(env_name: str | None, label: str) -> str:
    if not env_name:
        raise ValueError(f"Missing env binding for {label}")
    value = os.environ.get(env_name)
    if not value:
        raise ValueError(f"Environment variable {env_name} is not set")
    return value


def derive_db_name(base_url: str) -> str:
    parsed = urllib.parse.urlparse(base_url)
    hostname = parsed.hostname or ""
    return hostname.split(".")[0]


def resolve_db_name(profile: dict) -> str:
    db_name = profile.get("db")
    if db_name:
        return db_name
    return derive_db_name(profile["baseUrl"])


# ── Connection ───────────────────────────────────────────────────────────────

def connect_rpc(profile_name: str, config_path: str | None = None) -> dict:
    """Connect to Odoo XML-RPC. Returns dict with uid, db, models, api_key, profile."""
    config, _ = load_config(config_path)
    profile = resolve_profile(config, profile_name)
    username = read_secret(profile.get("usernameEnv"), "username")
    api_key = read_secret(profile.get("apiKeyEnv"), "api key")

    base_url = normalize_base_url(profile["baseUrl"])
    db_name = resolve_db_name(profile)

    ctx = ssl._create_unverified_context()
    common = xmlrpc.client.ServerProxy(get_endpoint(base_url, "common"), context=ctx)
    models = xmlrpc.client.ServerProxy(get_endpoint(base_url, "object"), context=ctx)
    uid = common.authenticate(db_name, username, api_key, {})
    if not uid:
        raise RuntimeError("Authentication failed")

    return {
        "profile": profile,
        "profileName": profile_name,
        "baseUrl": base_url,
        "db": db_name,
        "uid": uid,
        "username": username,
        "api_key": api_key,
        "common": common,
        "models": models,
    }


# ── RPC Execution ────────────────────────────────────────────────────────────

def rpc_execute(conn: dict, model: str, method: str, *args, **kwargs) -> any:
    """Generic execute_kw wrapper."""
    return conn["models"].execute_kw(
        conn["db"], conn["uid"], conn["api_key"],
        model, method, list(args), dict(kwargs),
    )


def search_read(conn: dict, model: str, domain: list, fields: list | None = None, limit: int | None = None, offset: int | None = None) -> list:
    kwargs = {}
    if fields:
        kwargs["fields"] = fields
    if limit:
        kwargs["limit"] = limit
    if offset:
        kwargs["offset"] = offset
    return rpc_execute(conn, model, "search_read", domain, **kwargs)


def read_record(conn: dict, model: str, ids: list, fields: list | None = None) -> list:
    kwargs = {}
    if fields:
        kwargs["fields"] = fields
    return rpc_execute(conn, model, "read", ids, **kwargs)


def search_count(conn: dict, model: str, domain: list) -> int:
    return rpc_execute(conn, model, "search_count", domain)


def fields_get(conn: dict, model: str, attributes: list | None = None) -> dict:
    kwargs = {}
    if attributes:
        kwargs["attributes"] = attributes
    return rpc_execute(conn, model, "fields_get", [], **kwargs)


# ── Hybrid Knowledge ─────────────────────────────────────────────────────────

def _load_yaml(path: Path) -> dict | None:
    """Load YAML file. Returns None if file doesn't exist."""
    if not path.exists():
        return None
    try:
        import yaml
        with path.open("r", encoding="utf-8") as handle:
            return yaml.safe_load(handle)
    except ImportError:
        # Minimal fallback — only handles simple key-value + nested sections
        # Full parsing requires PyYAML (pip install pyyaml)
        result = {}
        current_section = None
        current_model = None
        with path.open("r", encoding="utf-8") as handle:
            for line in handle:
                stripped = line.rstrip()
                if stripped.startswith("#") or not stripped.strip():
                    continue
                indent = len(stripped) - len(stripped.lstrip())
                content = stripped.strip()
                if indent == 0 and ":" in content:
                    key, val = content.split(":", 1)
                    key = key.strip()
                    val = val.strip()
                    if key == "models":
                        current_section = "models"
                        current_model = None
                    elif key == "verified_facts":
                        current_section = "verified_facts"
                        current_model = None
                    elif val:
                        result[key] = val if not val.isdigit() else int(val)
                    else:
                        current_section = key
                        current_model = None
                elif indent == 2 and current_section in ("models", "verified_facts"):
                    if ":" in content:
                        model_name = content.split(":")[0].strip()
                        current_model = model_name
                        result.setdefault(current_section, {})[model_name] = {}
                elif indent >= 4 and current_model:
                    if ":" in content:
                        k, v = content.split(":", 1)
                        k = k.strip()
                        v = v.strip()
                        if v.startswith("["):
                            v = [x.strip().strip("'\"") for x in v[1:-1].split(",")]
                        elif v.isdigit():
                            v = int(v)
                        elif v.startswith('"') or v.startswith("'"):
                            v = v.strip("'\"")
                        result.setdefault(current_section, {}).setdefault(current_model, {})[k] = v
        return result


def memory_lookup(model_name: str) -> dict:
    """Fast lookup from memory YAML files. Returns stored knowledge about a model."""
    result = {}

    # Top models (fast, always available)
    top = _load_yaml(MEMORY_DIR / "top-models.yaml")
    if top and model_name in top.get("models", {}):
        result["top"] = top["models"][model_name]

    # Verified facts (behavioral) — nested under verified_facts key
    facts_file = _load_yaml(MEMORY_DIR / "verified-facts.yaml")
    if facts_file:
        verified = facts_file.get("verified_facts", facts_file)
        if model_name in verified:
            result["facts"] = verified[model_name]

    # Dashboards referencing this model
    dashboards = _load_yaml(MEMORY_DIR / "dashboards.yaml")
    if dashboards:
        for dash_id, dash_info in dashboards.get("dashboards", {}).items():
            pivots = dash_info.get("pivots", [])
            for pivot in pivots:
                if pivot.get("model") == model_name:
                    result.setdefault("dashboards", []).append({
                        "id": dash_id,
                        "name": dash_info.get("name"),
                        "pivot": pivot,
                    })

    return result


def runtime_query(conn: dict, model_name: str) -> dict:
    """Query Odoo server directly for fresh model knowledge."""
    result = {}

    # Check model exists
    try:
        count = search_count(conn, "ir.model", [["model", "=", model_name]])
        result["exists"] = bool(count)
    except Exception as exc:
        result["exists"] = False
        result["error"] = str(exc)
        return result

    if not result["exists"]:
        return result

    # Record count
    try:
        result["record_count"] = search_count(conn, model_name, [])
    except Exception:
        result["record_count"] = "error"

    # Fields
    try:
        all_fields = fields_get(conn, model_name, ["string", "type", "relation", "required"])
        result["fields_total"] = len(all_fields)
        result["fields"] = all_fields
    except Exception as exc:
        result["fields_error"] = str(exc)

    # Custom fields (x_studio_*, x_sale_*, etc.)
    custom = {k: v for k, v in result.get("fields", {}).items() if k.startswith("x_")}
    result["custom_fields_count"] = len(custom)
    result["custom_fields"] = custom

    # Relations
    relations = {k: v for k, v in result.get("fields", {}).items()
                 if v.get("type") in ("many2one", "one2many", "many2many")}
    result["relations_count"] = len(relations)
    result["relations"] = relations

    return result


def get_model_knowledge(model_name: str, profile: str | None = None, fresh: bool = False) -> dict:
    """Hybrid model knowledge: memory first, runtime if stale or --fresh."""
    result = {"model": model_name}

    # Memory first (fast, 0 API calls)
    mem = memory_lookup(model_name)
    if mem:
        result["memory"] = mem

    # Runtime if fresh requested OR no memory data
    if fresh or not mem.get("top"):
        if not profile:
            config, _ = load_config()
            profile = config.get("defaultProfile", "prod")
        conn = connect_rpc(profile)
        rt = runtime_query(conn, model_name)
        result["runtime"] = rt

    return result


# ── CLI convenience ──────────────────────────────────────────────────────────

def get_connection(profile_name: str | None = None) -> dict:
    """Quick connect — uses default profile if none specified."""
    if not profile_name:
        config, _ = load_config()
        profile_name = config.get("defaultProfile", "prod")
    return connect_rpc(profile_name)