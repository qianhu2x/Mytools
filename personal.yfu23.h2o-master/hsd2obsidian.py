#!/usr/bin/env python
"""
Hsd2Obsidian  —  Convert Intel HSD tickets to Obsidian markdown notes.

CLI:  python hsd2obsidian.py HSD1 [HSD2 ...] <output_directory>
GUI:  python hsd2obsidian.py

HSD access: delegates to PowerShell / Windows auth via hsdes-plugin
  C:\\git\\personal.hcheng20.skills\\hsdes-plugin
"""

import sys
import re
import json
import shutil
import subprocess
import tempfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextEdit, QComboBox, QPushButton, QFileDialog,
    QSizePolicy, QLineEdit, QDialog, QListWidget, QListWidgetItem,
    QAbstractItemView, QMessageBox, QCheckBox, QStatusBar,
)
from PySide6.QtCore import QThread, Signal, QObject, QByteArray, Qt, QTimer, QElapsedTimer
from PySide6.QtGui import QFont

# ---------------------------------------------------------------------------
# Paths & constants
# ---------------------------------------------------------------------------

_APP_DIR       = Path(__file__).parent
_TEMPLATES_DIR = _APP_DIR / "Templates"
_CONFIG_FILE   = _APP_DIR / "hsd2obsidian_config.json"
_HSDES_API     = "https://hsdes-api.intel.com/rest"

TEMPLATE_MAP: dict[str, str] = {
    "sighting_central.sighting": "sighting_central.sighting.md",
    "server_platf_ae.bug":       "server_platf_ae.bug.md",
    "central_firmware.bug":      "central_firmware.bug.md",
    "central_firmware.feature":  "central_firmware.feature.md",
    "server_platf.test_case":    "server_platf.test_case.md",
}

_ARTICLE_FIELDS = (
    "id,tenant,subject,title,owner,submitted_by,submitted_date,"
    "co_owner,co_owners,sponsor,sponsor_org,submitter_org,"
    "priority,article_type,feature_type,feature_subtype,"
    "family,release,release_affected,program_milestone,"
    "component,component_affected,"
    "customer,customer_affected,customer_company,customer_project_name,"
    "forum,sub_forum,team_found,test_found,"
    "test_case.configuration,test_case.bios_rev,test_case.board_id,"
    "test_case.free_tag_1,test_case.free_tag_2,test_case.stepping,"
    "test_case.dimm_vendor,test_case.dimm_part_number,"
    "test_case.config_description,"
    "test_case.devices_or_dimms_per_channel,test_case.tested_speed,"
    "test_case.notes,test_case.actual_start,test_case.content_id,"
    "test_case.release_affected,test_case.component,"
    "test_case.component_affected,"
    "status,status_reason,reason_other,parent_id,"
    "updated_by,updated_date,nickname,"
    "notify,tag,description"
)

MAX_WORKERS = 10

# ---------------------------------------------------------------------------
# Config  (output-directory history + per-resolution window geometry)
# ---------------------------------------------------------------------------

def _load_config() -> dict:
    if _CONFIG_FILE.exists():
        try:
            return json.loads(_CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"output_dirs": [], "geometry": {}}


def _save_config(cfg: dict) -> None:
    try:
        _CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception:
        pass


def _add_output_dir(directory: str) -> list:
    cfg  = _load_config()
    dirs: list = cfg.get("output_dirs", [])
    if directory in dirs:
        dirs.remove(directory)
    dirs.insert(0, directory)
    cfg["output_dirs"] = dirs[:20]
    _save_config(cfg)
    return cfg["output_dirs"]


def _resolution_key() -> str:
    return "_".join(
        f"{s.size().width()}x{s.size().height()}"
        for s in QApplication.screens()
    )


# ---------------------------------------------------------------------------
# Obsidian vault tag scanner
# ---------------------------------------------------------------------------

def scan_vault_tags(vault_path: str) -> list:
    """Scan all .md files in an Obsidian vault and return sorted unique tags
    from YAML frontmatter (tags/tag fields)."""
    import yaml
    vault = Path(vault_path)
    if not vault.is_dir():
        return []
    tags = set()
    for md in vault.rglob("*.md"):
        try:
            text = md.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        if not text.startswith("---"):
            continue
        end = text.find("---", 3)
        if end == -1:
            continue
        try:
            parsed = yaml.safe_load(text[3:end])
            if not isinstance(parsed, dict):
                continue
            for key in ("tags", "tag"):
                val = parsed.get(key)
                if isinstance(val, list):
                    for t in val:
                        if t and str(t).strip():
                            tags.add(str(t).strip())
                elif isinstance(val, str) and val.strip():
                    for t in val.split(","):
                        if t.strip():
                            tags.add(t.strip())
        except Exception:
            continue
    return sorted(tags, key=str.lower)


# ---------------------------------------------------------------------------
# HSDes fetching via PowerShell Invoke-WebRequest
#   • Windows auth (UseDefaultCredentials) — no SSPI Python package needed
#   • NoProxy — bypasses corporate proxy
#   • Raw server JSON written to UTF-8 temp file — avoids console-codepage
#     corruption and PowerShell ConvertTo-Json re-serialisation issues
# ---------------------------------------------------------------------------

def _find_pwsh() -> str | None:
    for p in [
        Path.home() / "AppData/Local/Microsoft/WindowsApps/pwsh.exe",
        Path("C:/Program Files/PowerShell/7/pwsh.exe"),
    ]:
        if p.exists():
            return str(p)
    return shutil.which("pwsh")


def _ps_get(url: str) -> dict:
    """GET url with Windows auth; return parsed JSON."""
    pwsh = _find_pwsh()
    if not pwsh:
        raise RuntimeError(
            "PowerShell 7+ not found. Install: winget install Microsoft.PowerShell"
        )
    tmp = Path(tempfile.mktemp(suffix=".json"))
    try:
        ps = (
            "$ErrorActionPreference='Stop'; "
            f"$r = Invoke-WebRequest -Uri '{url}' -Method Get "
            "-UseDefaultCredentials -NoProxy "
            "-Headers @{Accept='application/json'}; "
            f"[System.IO.File]::WriteAllText('{tmp}', $r.Content, "
            "[System.Text.Encoding]::UTF8)"
        )
        r = subprocess.run(
            [pwsh, "-NoProfile", "-Command", ps],
            capture_output=True, timeout=120,
        )
        err = (r.stderr or b"").decode("utf-8", errors="replace").strip()
        if r.returncode != 0:
            raise RuntimeError(err or f"pwsh exited {r.returncode}")
        return json.loads(tmp.read_text(encoding="utf-8-sig"))
    finally:
        tmp.unlink(missing_ok=True)


def _strip_tenant_prefix(article: dict) -> dict:
    """Strip tenant.subject prefix from keys (e.g. 'server_platf_ae.bug.customer_company' → 'customer_company').
    Short/unprefixed keys are kept as-is; prefixed keys only added if the short key is absent."""
    tenant = article.get("tenant", "")
    subject = article.get("subject", "")
    prefix = f"{tenant}.{subject}." if tenant and subject else ""
    if not prefix:
        return article
    merged = dict(article)
    for k, v in article.items():
        if k.startswith(prefix):
            short = k[len(prefix):]
            if short not in merged or not merged[short]:
                merged[short] = v
    return merged


def fetch_article(hsd_id: str) -> dict:
    payload = _ps_get(f"{_HSDES_API}/article/{hsd_id}?fields={_ARTICLE_FIELDS}")
    data = payload.get("data", [])
    if not data:
        return {}
    article = data[0]

    # Some fields are tenant-specific and require a tenant.subject prefix
    # in the API query. Re-fetch with prefixed field names and merge.
    tenant = article.get("tenant", "")
    subject = article.get("subject", "")
    if tenant and subject:
        prefix = f"{tenant}.{subject}"
        all_fields = [f.strip() for f in _ARTICLE_FIELDS.split(",") if f.strip()]
        prefixed = ",".join(f"{prefix}.{f}" for f in all_fields)
        payload2 = _ps_get(f"{_HSDES_API}/article/{hsd_id}?fields={prefixed}")
        data2 = payload2.get("data", [])
        if data2:
            article.update(data2[0])

    return _strip_tenant_prefix(article)


# ---------------------------------------------------------------------------
# HSD query → article-ID expansion
# ---------------------------------------------------------------------------

_QUERY_PAGE_SIZE = 100

HSD_STATUSES = ("blocked", "complete", "future", "open", "rejected")


def _is_article_id(hsd_number: str) -> tuple:
    """Check if an HSD number is a valid article ID with a supported template.
    Returns (True, status) if valid, (False, None) otherwise."""
    try:
        payload = _ps_get(
            f"{_HSDES_API}/article/{hsd_number}?fields=id,tenant,subject,status"
        )
        data = payload.get("data")
        if not data:
            return False, None
        article = data[0]
        key = f"{article.get('tenant', '')}.{article.get('subject', '')}"
        return (key in TEMPLATE_MAP, article.get("status", ""))
    except Exception:
        return False, None


def fetch_query_ids(query_id: str, status_filter: set = None,
                    progress_cb=None) -> list:
    """Return all article IDs from an HSD query (handles pagination).
    If status_filter is provided, only include articles with matching status."""
    ids = []
    skipped = 0
    start = 1
    while True:
        payload = _ps_get(
            f"{_HSDES_API}/query/execution/{query_id}"
            f"?start_at={start}&max_results={_QUERY_PAGE_SIZE}"
        )
        data = payload.get("data", [])
        for item in data:
            aid = str(item.get("id", "")).strip()
            if not aid:
                continue
            if status_filter and item.get("status", "").lower() not in status_filter:
                skipped += 1
                continue
            ids.append(aid)
        total = payload.get("total", 0)
        if progress_cb:
            msg = f"Query {query_id}: fetched {len(ids)}/{total} IDs"
            if skipped:
                msg += f" (skipped {skipped} filtered)"
            progress_cb(msg)
        if start + len(data) > total or not data:
            break
        start += len(data)
    return ids


def resolve_hsd_numbers(numbers: list, status_filter: set = None,
                        progress_cb=None) -> list:
    """Expand a mixed list of article IDs and query IDs into article IDs.
    Tries article lookup first; falls back to query if not found.
    If status_filter is provided, only include articles with matching status."""
    resolved = []
    for num in numbers:
        num = num.strip()
        if not num:
            continue
        is_article, status = _is_article_id(num)
        if is_article:
            if status_filter and (status or "").lower() not in status_filter:
                if progress_cb:
                    progress_cb(f"⊘ HSD#{num}: status '{status}' filtered out")
                continue
            resolved.append(num)
        else:
            # Not an article — try as a query ID
            if progress_cb:
                progress_cb(f"{num} is not an article ID, trying as query…")
            try:
                qids = fetch_query_ids(num, status_filter, progress_cb)
                if qids:
                    resolved.extend(qids)
                else:
                    if progress_cb:
                        progress_cb(f"✗ {num}: query returned no articles after filtering")
            except Exception as exc:
                if progress_cb:
                    progress_cb(f"✗ {num}: not a valid article or query ID ({exc})")
    return resolved


def fetch_comments(hsd_id: str) -> list:
    payload = _ps_get(
        f"{_HSDES_API}/article/{hsd_id}/children"
        f"?child_subject=comment&fields=id,owner,description,submitted_date"
    )
    return payload.get("data", [])


# ---------------------------------------------------------------------------
# Template rendering  (str.format_map — templates use {key} syntax)
# ---------------------------------------------------------------------------

class _SafeNamespace:
    """Attribute-access object that returns '' for any missing attr.
    Needed for dotted template keys like {test_case.bios_rev}."""
    def __init__(self, **kwargs):
        self.__dict__.update(kwargs)
    def __getattr__(self, name):
        return ""
    def __format__(self, spec):
        return ""


class _SafeDict(defaultdict):
    """Returns '' for any key not in the context."""
    def __init__(self, data: dict):
        super().__init__(str, data)
    def __missing__(self, key):
        return ""


import re

_PLACEHOLDER_RE = re.compile(r"\{([^{}]+)\}")


_EMPTY_LIST_ITEM_RE = re.compile(
    r'^\s*-\s+'           # YAML list marker
    r'(?:'
    r'[,\s]*'             # only commas/spaces  (e.g. ", , ")
    r'|'
    r'"\[\[\]\]"'         # empty wiki-link     (e.g. "[[]]")
    r')\s*$'
)


def _clean_frontmatter_csv(text: str) -> str:
    """Within YAML frontmatter, clean up empty or artifact-only entries.

    Handles:
    - Inline CSV:  'notify: , , '        → 'notify: '
    - Mixed CSV:   'notify: jim, , bob'  → 'notify: jim, bob'
    - List CSV:    '- , , '              → removed
    - Empty links: '- "[[]]"'            → removed
    """
    if not text.startswith("---"):
        return text
    end = text.find("---", 3)
    if end == -1:
        return text
    fm = text[3:end]
    rest = text[end:]
    lines = []
    for line in fm.split("\n"):
        if ": " in line and "," in line:
            key_part, _, val_part = line.partition(": ")
            parts = [p.strip() for p in val_part.split(",")]
            non_empty = [p for p in parts if p]
            lines.append(f"{key_part}: {', '.join(non_empty)}")
        elif _EMPTY_LIST_ITEM_RE.match(line):
            # Drop empty / artifact-only YAML list items
            continue
        elif line.lstrip().startswith("- ") and "," in line:
            indent = line[:len(line) - len(line.lstrip())]
            val_part = line.lstrip()[2:]
            parts = [p.strip() for p in val_part.split(",")]
            non_empty = [p for p in parts if p]
            if non_empty:
                lines.append(f"{indent}- {', '.join(non_empty)}")
        else:
            lines.append(line)
    return "---" + "\n".join(lines) + rest


def _render(template_name: str, ctx: dict) -> str:
    text = (_TEMPLATES_DIR / template_name).read_text(encoding="utf-8")
    # Regex-based replacement so dotted keys like {server_platf_ae.bug.customer_company}
    # are looked up as flat keys first, falling back to str.format_map for nested access.
    def _replace(m):
        key = m.group(1)
        if key in ctx:
            val = ctx[key]
            return str(val) if val else ""
        # Fall back to format_map for nested attribute access (e.g. {test_case.bios_rev})
        try:
            return m.group(0).format_map(_SafeDict(ctx))
        except (KeyError, AttributeError, ValueError):
            return ""
    rendered = _PLACEHOLDER_RE.sub(_replace, text)
    return _clean_frontmatter_csv(rendered)


# ---------------------------------------------------------------------------
# Context builder
# ---------------------------------------------------------------------------

def _csv(value) -> list:
    if not value:
        return []
    return [v.strip() for v in str(value).split(",") if v.strip()]


# ---------------------------------------------------------------------------
# Title / filename sanitisation
# ---------------------------------------------------------------------------

_SPECIAL = re.compile(r'[\\\/\{\}\[\]\+\-\*\:\|\?]')
_MULTI   = re.compile(r'\s{2,}')


def _sanitize(title: str) -> str:
    title = _SPECIAL.sub(" ", title)
    title = title.replace(">", "higher than").replace("<", "lower than")
    return _MULTI.sub(" ", title).strip()


def _out_filename(hsd_id: str, title: str) -> str:
    return f"{hsd_id} {_sanitize(title)}.md"


# ---------------------------------------------------------------------------


def build_context(article: dict, comments: list, user_tags: str = "",
                  user_parent_ids: list = None) -> dict:
    def g(k):
        return article.get(k) or ""

    notify     = _csv(g("notify"))
    hsd_tags   = _csv(g("tag"))
    # Parent Nodes come exclusively from user input, not from HSD parent_id
    parent_ids = user_parent_ids or []
    utags      = [t.strip() for t in user_tags.split(",") if t.strip()]

    ctx = {
        "tenant":  g("tenant"),   "subject": g("subject"),
        "id":      g("id"),       "title":   _sanitize(g("title")),
        "submitted_by":   g("submitted_by"),
        "submitted_date": g("submitted_date"),
        "owner":          g("owner"),
        "co_owner":       g("co_owner"),
        "co_owners":      g("co_owners"),
        "sponsor":        g("sponsor"),
        "sponsor_org":    g("sponsor_org"),
        "submitter_org":  g("submitter_org"),
        "priority":        g("priority"),
        "article_type":    g("article_type"),
        "feature_type":    g("feature_type"),
        "feature_subtype": g("feature_subtype"),
        "family":             g("family"),
        "release":            g("release"),
        "release_affected":   g("release_affected"),
        "program_milestone":  g("program_milestone"),
        "component":          g("component"),
        "component_affected": g("component_affected"),
        "customer":              g("customer"),
        "customer_affected":     g("customer_affected"),
        "customer_company":      g("customer_company"),
        "customer_project_name": g("customer_project_name"),
        "forum":      g("forum"),
        "sub_forum":  g("sub_forum"),
        "team_found": g("team_found"),
        "test_found": g("test_found"),
        "test_case": _SafeNamespace(
            configuration=g("test_case.configuration"),
            release_affected=g("test_case.release_affected"),
            bios_rev=g("test_case.bios_rev"),
            board_id=g("test_case.board_id"),
            free_tag_1=g("test_case.free_tag_1"),
            stepping=g("test_case.stepping"),
            dimm_vendor=g("test_case.dimm_vendor"),
            dimm_part_number=g("test_case.dimm_part_number"),
            config_description=g("test_case.config_description"),
            devices_or_dimms_per_channel=g("test_case.devices_or_dimms_per_channel"),
            tested_speed=g("test_case.tested_speed"),
            component=g("test_case.component"),
            component_affected=g("test_case.component_affected"),
            notes=g("test_case.notes"),
            actual_start=g("test_case.actual_start"),
        ),
        "reason_other":               g("reason_other"),
        "status_reason":              g("status_reason"),
        "bios_rev_prefix":            g("test_case.bios_rev").rsplit("_", 1)[0] if g("test_case.bios_rev") else "",
        "bios_rev_suffix":            g("test_case.bios_rev").rsplit("_", 1)[-1] if "_" in g("test_case.bios_rev") else "",
        "updated_by":                 g("updated_by"),
        "updated_date":               g("updated_date"),
        "nickname":                   g("nickname"),
        "parent_id1": parent_ids[0] if len(parent_ids) > 0 else "",
        "parent_id2": parent_ids[1] if len(parent_ids) > 1 else "",
        "parent_id3": parent_ids[2] if len(parent_ids) > 2 else "",
        "parent_node1": parent_ids[0] if len(parent_ids) > 0 else "",
        "parent_node2": parent_ids[1] if len(parent_ids) > 1 else "",
        "parent_node3": parent_ids[2] if len(parent_ids) > 2 else "",
        "notify1": notify[0] if len(notify) > 0 else "",
        "notify2": notify[1] if len(notify) > 1 else "",
        "notify3": notify[2] if len(notify) > 2 else "",
        "notify4": notify[3] if len(notify) > 3 else "",
        "tag1": hsd_tags[0] if len(hsd_tags) > 0 else "",
        "tag2": hsd_tags[1] if len(hsd_tags) > 1 else "",
        "tag3": hsd_tags[2] if len(hsd_tags) > 2 else "",
        "other_tag1": utags[0] if len(utags) > 0 else "",
        "other_tag2": utags[1] if len(utags) > 1 else "",
        "Description": g("description"),
        "user_tags":   ", ".join(utags),
    }

    # Include raw article keys (with tenant prefixes) so templates
    # can use full HSD field names like {server_platf_ae.bug.customer_company}
    for k, v in article.items():
        if k not in ctx:
            ctx[k] = v if v else ""

    return ctx


# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------

def _check_existing_files(hsd_ids: list, output_dir: str,
                          progress_cb=None) -> list:
    """Return list of (hsd_id, filepath) for output files that already exist."""
    existing = []

    def _check_one(hsd_id):
        try:
            payload = _ps_get(
                f"{_HSDES_API}/article/{hsd_id}?fields=id,title")
            data = payload.get("data", [])
            if not data:
                return None
            a = data[0]
            out = Path(output_dir) / _out_filename(
                str(a.get("id", hsd_id)), str(a.get("title", hsd_id)))
            if out.exists():
                return (hsd_id, str(out.name))
        except Exception:
            pass
        return None

    with ThreadPoolExecutor(max_workers=min(len(hsd_ids), MAX_WORKERS)) as ex:
        futures = {ex.submit(_check_one, hid): hid for hid in hsd_ids}
        for f in as_completed(futures):
            result = f.result()
            if result:
                existing.append(result)
            if progress_cb:
                progress_cb(f"Checked HSD#{futures[f]}")
    return existing


def process_hsd(hsd_id: str, output_dir: str, tags: str = "",
                user_parent_ids: list = None) -> tuple:
    """Returns (hsd_id, success, message)."""
    try:
        with ThreadPoolExecutor(max_workers=2) as ex:
            fa = ex.submit(fetch_article,  hsd_id)
            fc = ex.submit(fetch_comments, hsd_id)
            article  = fa.result()
            comments = fc.result()

        if not article:
            return hsd_id, False, f"✗ HSD#{hsd_id}: API returned no data."

        key = f"{article.get('tenant','')}.{article.get('subject','')}"
        if key not in TEMPLATE_MAP:
            return hsd_id, False, (
                f"✗ HSD#{hsd_id}: No template for '{key}'. "
                f"Supported: {', '.join(TEMPLATE_MAP)}"
            )

        ctx      = build_context(article, comments, tags, user_parent_ids)
        rendered = _render(TEMPLATE_MAP[key], ctx)

        if comments:
            rendered += "\n\n---\n##### Comments:\n"
            for c in comments:
                author = c.get("owner") or c.get("submitted_by") or "unknown"
                rendered += (
                    f"\n**{author}** ({c.get('submitted_date','')}):\n"
                    f"{c.get('description','')}\n"
                )

        out = Path(output_dir) / _out_filename(
            str(article.get("id", hsd_id)),
            str(article.get("title", hsd_id)),
        )
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(rendered, encoding="utf-8")
        return hsd_id, True, f"✓ HSD#{hsd_id}: Saved → '{out.name}'"

    except Exception as exc:
        return hsd_id, False, f"✗ HSD#{hsd_id}: {type(exc).__name__}: {exc}"


def run_batch(hsd_ids: list, output_dir: str,
              tags: str = "", user_parent_ids: list = None,
              progress_cb=None) -> list:
    results = []
    with ThreadPoolExecutor(max_workers=min(len(hsd_ids), MAX_WORKERS)) as ex:
        futures = {ex.submit(process_hsd, hid, output_dir, tags, user_parent_ids): hid
                   for hid in hsd_ids}
        for f in as_completed(futures):
            res = f.result()
            results.append(res)
            if progress_cb:
                progress_cb(res[2])
    return results


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _cli() -> None:
    args = sys.argv[1:]
    if len(args) < 2:
        print("Usage: python hsd2obsidian.py HSD1 [HSD2 ...] <output_directory>")
        sys.exit(1)
    raw_ids, output_dir = args[:-1], args[-1]

    # Resolve query IDs into article IDs
    hsd_ids = resolve_hsd_numbers(raw_ids, progress_cb=print)
    if not hsd_ids:
        print("No HSD article IDs found.")
        sys.exit(1)

    # Check for existing output files
    existing = _check_existing_files(hsd_ids, output_dir)
    if existing:
        print("The following file(s) already exist:")
        for _, name in existing:
            print(f"  • {name}")
        answer = input("Overwrite? (y/n): ").strip().lower()
        if answer != "y":
            print("Aborted.")
            sys.exit(0)

    print(f"Processing {len(hsd_ids)} HSD(s) → {output_dir}\n")
    results = run_batch(hsd_ids, output_dir, progress_cb=print)
    ok = sum(1 for _, s, _ in results if s)
    print(f"\nDone: {ok} succeeded, {len(results)-ok} failed.")


# ---------------------------------------------------------------------------
# GUI — background worker
# ---------------------------------------------------------------------------

class _Worker(QObject):
    progress = Signal(str)
    finished = Signal()

    def __init__(self, hsd_ids: list, output_dir: str, tags: str = "",
                 user_parent_ids: list = None):
        super().__init__()
        self.hsd_ids         = hsd_ids
        self.output_dir      = output_dir
        self.tags            = tags
        self.user_parent_ids = user_parent_ids

    def run(self) -> None:
        run_batch(self.hsd_ids, self.output_dir, self.tags,
                  self.user_parent_ids,
                  progress_cb=lambda m: self.progress.emit(m))
        self.finished.emit()


# ---------------------------------------------------------------------------
# GUI — main window
class OptionsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Options")
        self.setMinimumSize(400, 150)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 12, 12, 12)
        lay.setSpacing(8)

        # Obsidian vault path
        lay.addWidget(QLabel("Obsidian Vault Path:"))
        row = QHBoxLayout()
        self._vault_edit = QLineEdit()
        self._vault_edit.setPlaceholderText("e.g. C:\\Users\\...\\MyVault")
        cfg = _load_config()
        self._vault_edit.setText(cfg.get("obsidian_vault", ""))
        row.addWidget(self._vault_edit)
        browse_btn = QPushButton("Browse…")
        browse_btn.clicked.connect(self._browse_vault)
        row.addWidget(browse_btn)
        lay.addLayout(row)

        lay.addStretch()

        # OK / Cancel
        btn_row = QHBoxLayout()
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self._accept)
        btn_row.addWidget(ok_btn)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        lay.addLayout(btn_row)

    def _browse_vault(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "Select Obsidian Vault")
        if d:
            self._vault_edit.setText(str(Path(d)))

    def _accept(self) -> None:
        cfg = _load_config()
        cfg["obsidian_vault"] = self._vault_edit.text().strip()
        _save_config(cfg)
        self.accept()

    def vault_path(self) -> str:
        return self._vault_edit.text().strip()


# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Hsd2Obsidian")
        self.setMinimumSize(720, 580)
        self._thread = self._worker = None
        self._vault_tags: list = []
        self._total_hsd = 0
        self._completed_hsd = 0
        self._build_ui()
        self._load_vault_tags()
        self._restore_geometry()

        # Elapsed-time timer
        self._elapsed_timer = QElapsedTimer()
        self._tick_timer = QTimer(self)
        self._tick_timer.setInterval(1000)
        self._tick_timer.timeout.connect(self._update_elapsed)

    # ---- geometry ----

    def _restore_geometry(self) -> None:
        geom = _load_config().get("geometry", {}).get(_resolution_key())
        if geom:
            self.restoreGeometry(QByteArray.fromBase64(geom.encode()))
        else:
            self.showNormal()

    def _save_geometry(self) -> None:
        cfg = _load_config()
        cfg.setdefault("geometry", {})[_resolution_key()] = (
            self.saveGeometry().toBase64().data().decode()
        )
        _save_config(cfg)

    def closeEvent(self, event):
        self._save_geometry()
        super().closeEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.close()
        else:
            super().keyPressEvent(event)

    # ---- UI ----

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        lay = QVBoxLayout(root)
        lay.setSpacing(8)
        lay.setContentsMargins(12, 12, 12, 12)

        # HSD numbers
        lay.addWidget(QLabel("HSD Numbers (one per line):"))
        self._hsd_edit = QTextEdit()
        self._hsd_edit.setAcceptRichText(False)
        self._hsd_edit.setPlaceholderText("123456789\n987654321\n...")
        self._hsd_edit.setFixedHeight(120)
        lay.addWidget(self._hsd_edit)

        # Output directory
        lay.addWidget(QLabel("Output Directory:"))
        row = QHBoxLayout()
        self._dir_combo = QComboBox()
        self._dir_combo.setEditable(True)
        self._dir_combo.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        for d in _load_config().get("output_dirs", []):
            self._dir_combo.addItem(d)
        row.addWidget(self._dir_combo)
        btn = QPushButton("Browse…")
        btn.clicked.connect(self._browse)
        row.addWidget(btn)
        lay.addLayout(row)

        # Tags
        lay.addWidget(QLabel("Tags (comma-separated):"))
        tag_row = QHBoxLayout()
        self._tags_edit = QLineEdit()
        self._tags_edit.setPlaceholderText("e.g. DDR5, memory, debug")
        tag_row.addWidget(self._tags_edit)
        self._pick_tags_btn = QPushButton("Pick Tags…")
        self._pick_tags_btn.clicked.connect(self._pick_tags)
        self._pick_tags_btn.setEnabled(False)
        tag_row.addWidget(self._pick_tags_btn)
        lay.addLayout(tag_row)

        # Parent Nodes
        lay.addWidget(QLabel("Parent Nodes (one per line):"))
        self._parent_id_edit = QTextEdit()
        self._parent_id_edit.setAcceptRichText(False)
        self._parent_id_edit.setPlaceholderText("e.g. 12345678\n87654321")
        self._parent_id_edit.setFixedHeight(80)
        lay.addWidget(self._parent_id_edit)

        # Status filter — checked means excluded
        lay.addWidget(QLabel("Filters (check to exclude):"))
        status_row = QHBoxLayout()
        self._status_cbs: dict[str, QCheckBox] = {}
        for s in HSD_STATUSES:
            cb = QCheckBox(s)
            self._status_cbs[s] = cb
            status_row.addWidget(cb)
        lay.addLayout(status_row)

        # Options / Start
        btn_row = QHBoxLayout()
        self._options_btn = QPushButton("Options")
        self._options_btn.clicked.connect(self._show_options)
        btn_row.addWidget(self._options_btn)
        self._start_btn = QPushButton("Start")
        self._start_btn.clicked.connect(self._start)
        btn_row.addWidget(self._start_btn)
        lay.addLayout(btn_row)

        # Log
        lay.addWidget(QLabel("Progress / Status:"))
        self._log = QTextEdit()
        self._log.setReadOnly(True)
        self._log.setFont(QFont("Consolas", 9))
        lay.addWidget(self._log)

        # Status bar
        sb = QStatusBar()
        self.setStatusBar(sb)
        self._sb_status = QLabel("Idle")
        self._sb_elapsed = QLabel("00:00")
        self._sb_progress = QLabel("0/0")
        sb.addWidget(self._sb_status, 1)
        sb.addWidget(self._sb_elapsed)
        sb.addWidget(self._sb_progress)

    # ---- slots ----

    def _load_vault_tags(self) -> None:
        vault = _load_config().get("obsidian_vault", "")
        if vault and Path(vault).is_dir():
            self._vault_tags = scan_vault_tags(vault)
        else:
            self._vault_tags = []
        self._pick_tags_btn.setEnabled(bool(self._vault_tags))

    def _pick_tags(self) -> None:
        """Open a dialog with a checkable list of vault tags for multi-select."""
        dlg = QDialog(self)
        dlg.setWindowTitle("Pick Tags")
        dlg.setMinimumSize(350, 400)
        lay = QVBoxLayout(dlg)

        # Search filter
        search = QLineEdit()
        search.setPlaceholderText("Filter tags…")
        lay.addWidget(search)

        tag_list = QListWidget()
        tag_list.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        # Pre-check tags already in the text field
        current = {t.strip() for t in self._tags_edit.text().split(",") if t.strip()}
        for tag in self._vault_tags:
            item = QListWidgetItem(tag)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if tag in current else Qt.CheckState.Unchecked)
            tag_list.addItem(item)
        lay.addWidget(tag_list)

        def _filter(text):
            text = text.lower()
            for i in range(tag_list.count()):
                item = tag_list.item(i)
                item.setHidden(text not in item.text().lower())

        search.textChanged.connect(_filter)

        btn_row = QHBoxLayout()
        ok_btn = QPushButton("OK")
        cancel_btn = QPushButton("Cancel")
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        lay.addLayout(btn_row)

        def _accept():
            picked = []
            for i in range(tag_list.count()):
                item = tag_list.item(i)
                if item.checkState() == Qt.CheckState.Checked:
                    picked.append(item.text())
            # Merge with any manually typed tags not in vault
            manual = [t.strip() for t in self._tags_edit.text().split(",")
                       if t.strip() and t.strip() not in self._vault_tags]
            all_tags = manual + picked
            self._tags_edit.setText(", ".join(all_tags))
            dlg.accept()

        ok_btn.clicked.connect(_accept)
        cancel_btn.clicked.connect(dlg.reject)
        dlg.exec()

    def _show_options(self) -> None:
        dlg = OptionsDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._load_vault_tags()

    def _browse(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if d:
            d = str(Path(d))
            if self._dir_combo.findText(d) == -1:
                self._dir_combo.insertItem(0, d)
            self._dir_combo.setCurrentText(d)

    def _start(self) -> None:
        raw = self._hsd_edit.toPlainText().strip()
        if not raw:
            self._append("ERROR: Please enter at least one HSD number.")
            return
        output_dir = self._dir_combo.currentText().strip()
        if not output_dir:
            self._append("ERROR: Please specify an output directory.")
            return

        hsd_ids = [ln.strip() for ln in raw.splitlines() if ln.strip()]
        dirs    = _add_output_dir(output_dir)
        cur     = self._dir_combo.currentText()
        self._dir_combo.clear()
        for d in dirs:
            self._dir_combo.addItem(d)
        self._dir_combo.setCurrentText(cur)

        # Resolve query IDs into article IDs
        # Checked statuses are excluded
        excluded = {s for s, cb in self._status_cbs.items() if cb.isChecked()}
        status_filter = {s for s in HSD_STATUSES if s not in excluded} if excluded else None
        self._log.clear()
        if excluded:
            self._append(f"Excluded statuses: {', '.join(sorted(excluded))}")
        self._append("Resolving HSD numbers (queries → article IDs)…")
        hsd_ids = resolve_hsd_numbers(hsd_ids, status_filter,
                                      progress_cb=self._append)
        if not hsd_ids:
            self._append("ERROR: No HSD article IDs found.")
            return

        # Parse user-supplied parent IDs
        raw_pids = self._parent_id_edit.toPlainText().strip()
        user_parent_ids = [ln.strip() for ln in raw_pids.splitlines() if ln.strip()] or None

        # Check for existing output files
        self._log.clear()
        self._append("Checking for existing files…")
        existing = _check_existing_files(hsd_ids, output_dir)
        if existing:
            existing_ids = {hid for hid, _ in existing}
            dlg = QDialog(self)
            dlg.setWindowTitle("Existing Files")
            dlg.setMinimumSize(500, 300)
            lay_dlg = QVBoxLayout(dlg)
            lay_dlg.addWidget(QLabel("The following file(s) already exist:"))
            file_list = QTextEdit()
            file_list.setReadOnly(True)
            file_list.setPlainText("\n".join(name for _, name in existing))
            lay_dlg.addWidget(file_list)
            btn_row = QHBoxLayout()
            proceed_btn = QPushButton("Override")
            skip_btn = QPushButton("Skip Existing Files")
            cancel_btn = QPushButton("Cancel")
            dlg._result_action = None
            def _proceed():
                dlg._result_action = "proceed"
                dlg.accept()
            def _skip():
                dlg._result_action = "skip"
                dlg.accept()
            proceed_btn.clicked.connect(_proceed)
            skip_btn.clicked.connect(_skip)
            cancel_btn.clicked.connect(dlg.reject)
            btn_row.addWidget(proceed_btn)
            btn_row.addWidget(cancel_btn)
            btn_row.addWidget(skip_btn)
            lay_dlg.addLayout(btn_row)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                self._append("Aborted.")
                return
            if dlg._result_action == "skip":
                hsd_ids = [h for h in hsd_ids if h not in existing_ids]
                if not hsd_ids:
                    self._append("All HSD IDs already have existing files — nothing to process.")
                    return
                self._append(f"Skipping {len(existing_ids)} existing file(s).")

        self._log.clear()
        self._append(f"Processing {len(hsd_ids)} HSD(s) → {output_dir}")
        self._start_btn.setEnabled(False)

        # Reset status bar
        self._total_hsd = len(hsd_ids)
        self._completed_hsd = 0
        self._sb_status.setText("Working")
        self._sb_elapsed.setText("00:00")
        self._sb_progress.setText(f"0/{self._total_hsd}")
        self._elapsed_timer.start()
        self._tick_timer.start()

        self._worker = _Worker(hsd_ids, output_dir, self._tags_edit.text().strip(),
                               user_parent_ids)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._done)
        self._thread.start()

    def _done(self) -> None:
        self._tick_timer.stop()
        self._update_elapsed()
        self._sb_status.setText("Done")
        self._append("Done.")
        self._start_btn.setEnabled(True)
        if self._thread:
            self._thread.quit()
            self._thread.wait()
        self._thread = self._worker = None

    def _update_elapsed(self) -> None:
        secs = self._elapsed_timer.elapsed() // 1000
        self._sb_elapsed.setText(f"{secs // 60:02d}:{secs % 60:02d}")

    def _on_progress(self, msg: str) -> None:
        self._append(msg)
        if msg.startswith("✓") or msg.startswith("✗"):
            self._completed_hsd += 1
            self._sb_progress.setText(f"{self._completed_hsd}/{self._total_hsd}")

    def _append(self, msg: str) -> None:
        self._log.append(msg)
        self._log.verticalScrollBar().setValue(
            self._log.verticalScrollBar().maximum()
        )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    if len(sys.argv) >= 3:
        _cli()
    else:
        app = QApplication(sys.argv)
        win = MainWindow()
        win.show()
        sys.exit(app.exec())


if __name__ == "__main__":
    main()
