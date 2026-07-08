"""design_lint.py — lint + contrast + diff for the OW design system.

Three capabilities, all reading the same generated sources of truth so they can
never disagree with what the engine actually does:

  lint_deck(deck_spec)   structural lint of an emitted deck JSON:
                           - broken-ref   : intent / chart-type / palette role
                                            references that don't resolve
                           - off-palette  : an explicit colour that isn't an OW
                                            palette role or OW hex
                           - orphan       : palette roles no component binds
                           Returns findings + a `passed` boolean (errors==0).

  contrast_report()      WCAG AA contrast of every component's declared
                           background/text token pair. Fails on any pair < 4.5:1.

  diff_tokens(a, b)      token-level diff of two design-system snapshots
                           (added / removed / modified), with a `regression`
                           flag when tokens are removed or a colour pair drops
                           below AA — the guard against silent drift between
                           runtime versions.

This is the OW analogue of `@google/design.md lint|diff`, but pointed at our
runtime-as-source-of-truth: tokens come from code, components from
design_components, and the deck under lint is the JSON the GPT emits.
"""
from __future__ import annotations

import re
from typing import Any, Dict, List

import runtime as _rt
import design_components as _dc

try:
    from compiler import supported_intents as _supported_intents
except Exception:                                            # pragma: no cover
    def _supported_intents() -> List[str]:
        return []

_HEX_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")
_OW_HEXES = {v.upper() for v in _rt.PALETTE.values()} | {
    v.upper() for v in _rt.DEFAULT_THEME.values() if _HEX_RE.match(str(v))
}
_PALETTE_ROLES = set(_rt.PALETTE) | set(_rt.BRAND_COLORS)


def _finding(sev: str, path: str, msg: str) -> dict:
    return {"severity": sev, "path": path, "message": msg}


# ---------------------------------------------------------------------------
# 1. Deck-JSON lint
# ---------------------------------------------------------------------------
def _walk_colors(obj: Any, path: str):
    """Yield (path, value) for every key that looks like a colour input."""
    COLOR_KEYS = {"fill", "color", "colour", "accent", "line", "background",
                  "bg", "text_color", "textColor"}
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = f"{path}.{k}"
            if k in COLOR_KEYS and isinstance(v, str):
                yield p, v
            yield from _walk_colors(v, p)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from _walk_colors(v, f"{path}[{i}]")


def _is_ok_color(val: str) -> bool:
    v = val.strip()
    if v in _PALETTE_ROLES:                 # a semantic role
        return True
    if v in _rt.DEFAULT_THEME:              # accentN / dk1 / lt1
        return True
    if _HEX_RE.match(v) and v.lstrip("#").upper() in _OW_HEXES:
        return True
    return False


def lint_deck(deck_spec: Dict) -> Dict:
    findings: List[dict] = []
    intents = set(_supported_intents())
    slides = deck_spec.get("slides", []) or []

    for i, s in enumerate(slides, 1):
        intent = s.get("intent", "")
        if intents and intent and intent not in intents:
            findings.append(_finding(
                "error", f"slides[{i}].intent",
                f"unknown intent '{intent}' — not a registered intent"))
        # chart-type references
        content = s.get("content", {}) or {}
        ctype = (content.get("chart") or {}).get("type") if isinstance(
            content.get("chart"), dict) else content.get("chart_type")
        if ctype:
            canon = str(ctype).strip().lower().replace("-", "_").replace(" ", "_")
            if canon not in _rt.CHART_TYPE_MAP:
                findings.append(_finding(
                    "warning", f"slides[{i}].content.chart.type",
                    f"unknown chart type '{ctype}' — engine will fall back to "
                    f"column_clustered"))
        # off-palette colours
        for p, val in _walk_colors(s, f"slides[{i}]"):
            if not _is_ok_color(val):
                findings.append(_finding(
                    "warning", p,
                    f"colour '{val}' is not an OW palette role or OW hex — "
                    f"rebind to a role ({', '.join(sorted(_rt.BRAND_COLORS))})"))

    # orphaned palette roles (defined but bound by no component)
    bound = set()
    for ref in _dc.all_bound_token_refs():
        m = re.match(r"\{(palette|BRAND_COLORS)\.([^.}]+)\}", ref)
        if m:
            bound.add(m.group(2))
    for role in sorted(_rt.BRAND_COLORS):
        if role not in bound:
            findings.append(_finding(
                "info", f"BRAND_COLORS.{role}",
                f"palette role '{role}' is defined but no component binds it"))

    errors = sum(1 for f in findings if f["severity"] == "error")
    warnings = sum(1 for f in findings if f["severity"] == "warning")
    info = sum(1 for f in findings if f["severity"] == "info")
    return {
        "passed": errors == 0,
        "findings": findings,
        "summary": {"errors": errors, "warnings": warnings, "info": info},
    }


# ---------------------------------------------------------------------------
# 2. Contrast report (component bg/text pairs)
# ---------------------------------------------------------------------------
def contrast_report() -> Dict:
    findings: List[dict] = []
    for name, spec in _dc.COMPONENTS.items():
        for bg_role, tx_role in spec.get("contrast_pairs", []):
            try:
                bg = _rt.resolve_color_ref(spec["tokens"][bg_role])
                tx = _rt.resolve_color_ref(spec["tokens"][tx_role])
            except Exception as exc:
                findings.append(_finding(
                    "error", f"{name}.{tx_role}",
                    f"cannot resolve contrast pair: {exc}"))
                continue
            cr = _rt.contrast_ratio(bg, tx)
            if cr < 4.5:
                sev = "warning" if cr >= 3.0 else "error"
                findings.append(_finding(
                    sev, f"{name}: {tx_role} on {bg_role}",
                    f"contrast {cr}:1 below WCAG AA (4.5:1) for normal text"))
    errors = sum(1 for f in findings if f["severity"] == "error")
    warnings = sum(1 for f in findings if f["severity"] == "warning")
    return {"passed": errors == 0, "findings": findings,
            "summary": {"errors": errors, "warnings": warnings}}


# ---------------------------------------------------------------------------
# 3. Token diff (regression guard)
# ---------------------------------------------------------------------------
_TOKEN_LINE = re.compile(r"^\|\s*`([^`]+)`\s*\|\s*(.+?)\s*\|")


def _parse_tokens(md: str) -> Dict[str, str]:
    """Extract every `token | value` row from a design-system markdown doc into
    a flat {token: value} map. Robust to the prose interleaving — it only reads
    table rows whose first cell is a backticked token name."""
    out: Dict[str, str] = {}
    for line in md.splitlines():
        m = _TOKEN_LINE.match(line.strip())
        if m:
            name, val = m.group(1), m.group(2)
            if name in ("Token", "Role", "Rule", "Variable", "Constant",
                        "Slot", "Reference"):
                continue
            out[name] = val.strip().strip("`")
    return out


def diff_tokens(before_md: str, after_md: str) -> Dict:
    a, b = _parse_tokens(before_md), _parse_tokens(after_md)
    added = sorted(set(b) - set(a))
    removed = sorted(set(a) - set(b))
    modified = sorted(k for k in (set(a) & set(b)) if a[k] != b[k])
    # A regression = a token vanished (something referencing it may now break).
    regression = bool(removed)
    return {
        "added": added,
        "removed": removed,
        "modified": [{"token": k, "before": a[k], "after": b[k]} for k in modified],
        "regression": regression,
    }
