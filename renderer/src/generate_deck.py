"""Compile a semantic deck JSON into a PPTX.

Usage:
    python generate_deck.py \\
        --template ow_default.pptx \\
        --deck sample-json/semantic_deck.json \\
        --output output/semantic_deck.pptx
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from runtime import TemplateRuntime
from compiler import (
    SemanticCompiler, supported_intents, supported_strategies,
    supported_content_kinds, describe_capabilities,
)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", default="ow_default.pptx")
    parser.add_argument("--deck",     default="sample-json/semantic_deck.json")
    parser.add_argument("--output",   default="output/semantic_deck.pptx")
    parser.add_argument("--list-intents", action="store_true",
                        help="Print the registered intents and exit.")
    parser.add_argument("--list-capabilities", action="store_true",
                        help="Print the layout capability registry and exit.")
    parser.add_argument("--version", action="store_true",
                        help="Print the runtime version and exit.")
    parser.add_argument("--list-tokens", action="store_true",
                        help="Print the design-system reference (DESIGN_SYSTEM.md) and exit.")
    parser.add_argument("--lint-deck", metavar="DECK",
                        help="Lint a deck JSON/YAML for broken refs + off-palette "
                             "colours; exit 1 on errors.")
    parser.add_argument("--check-contrast", action="store_true",
                        help="WCAG-AA contrast check of component colour pairs; exit 1 on fail.")
    parser.add_argument("--diff-tokens", nargs=2, metavar=("BEFORE.md", "AFTER.md"),
                        help="Diff two design-system snapshots; exit 1 on regression.")
    parser.add_argument("--list-voice", action="store_true",
                        help="Print the voice-system reference (60_voice_system.md) and exit.")
    parser.add_argument("--voice-spec", metavar="PROFILE",
                        help="Print the compact prompt-injectable spec for one voice profile.")
    parser.add_argument("--lint-text", metavar="FILE",
                        help="Voice-lint a draft (filler, clutter, hedging, profile rules).")
    parser.add_argument("--profile", default="consulting",
                        help="Voice profile for --lint-text (consulting | longform).")
    args = parser.parse_args()

    if args.version:
        from runtime import __version__
        print(f"deck runtime {__version__}")
        return 0

    if args.list_voice:
        import voice_profiles as _vp
        print(_vp.voice_markdown())
        return 0

    if args.voice_spec:
        import voice_profiles as _vp
        print(_vp.voice_spec(args.voice_spec))
        return 0

    if args.lint_text:
        import json as _json
        import voice_lint as _vl
        with open(args.lint_text) as f:
            text = f.read()
        rep = _vl.lint_text(text, args.profile)
        print(_json.dumps(rep, indent=2))
        return 0 if rep["passed"] else 1

    if args.list_tokens:
        from runtime import design_system_markdown
        print(design_system_markdown())
        return 0

    if args.check_contrast:
        import json as _json
        import design_lint as _lint
        rep = _lint.contrast_report()
        print(_json.dumps(rep, indent=2))
        return 0 if rep["passed"] else 1

    if args.diff_tokens:
        import json as _json
        import design_lint as _lint
        before = open(args.diff_tokens[0]).read()
        after = open(args.diff_tokens[1]).read()
        rep = _lint.diff_tokens(before, after)
        print(_json.dumps(rep, indent=2))
        return 1 if rep["regression"] else 0

    if args.lint_deck:
        import json as _json
        import design_lint as _lint
        with open(args.lint_deck) as f:
            if args.lint_deck.lower().endswith((".yaml", ".yml")):
                import yaml
                spec = yaml.safe_load(f)
            else:
                spec = _json.load(f)
        rep = _lint.lint_deck(spec)
        print(_json.dumps(rep, indent=2))
        return 0 if rep["passed"] else 1

    if args.list_intents:
        print("Registered intents:")
        for name in supported_intents():
            print(f"  - {name}")
        print("\nRegistered render strategies:")
        for name in supported_strategies():
            print(f"  - {name}")
        print(f"\nKnown content kinds: {', '.join(supported_content_kinds())}")
        return 0

    if args.list_capabilities:
        print("Layout capability registry:")
        print(describe_capabilities())
        return 0

    # Accept either JSON or YAML (.yaml/.yml). Both deserialize to the same
    # deck_spec dict — YAML is just a more author-friendly surface (mode b2).
    with open(args.deck) as f:
        if args.deck.lower().endswith((".yaml", ".yml")):
            try:
                import yaml
            except ImportError:
                print("PyYAML not installed: pip install pyyaml "
                      "(or pass a .json deck)")
                return 2
            deck_spec = yaml.safe_load(f)
        else:
            deck_spec = json.load(f)
    if not isinstance(deck_spec, dict) or "slides" not in deck_spec:
        print("deck file must be a mapping with a top-level 'slides' list")
        return 2

    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    rt = TemplateRuntime(args.template)

    # Image registry: deck-level sources can be reused across slides with
    # different crops (deck_spec["images"] = {id: "path" | {"path"|"prompt"}}).
    try:
        import images as _img
        reg = _img.ImageRegistry()
        for sid, spec in (deck_spec.get("images") or {}).items():
            if isinstance(spec, str):
                reg.register(sid, spec)
            elif isinstance(spec, dict):
                reg.register(sid, spec.get("path"), spec.get("prompt"))
        rt._image_registry = reg
        rt._allow_image_gen = bool(deck_spec.get("allow_image_generation", False))
    except Exception:
        pass

    compiler = SemanticCompiler(rt)
    compiler.build(deck_spec)
    rt.save(args.output)   # save() also normalizes charts (ticks off, value axis)
    summary = compiler.build_summary()
    if summary:
        print(summary)

    # Machine-readable build gate (overflow + icon coverage + action titles +
    # placeholders + sources). Writes gate_result.json next to the output; the
    # GPT must READ it before declaring done -- passed is a Python boolean.
    try:
        import gate as _gate
        res = _gate.run_gate(deck_spec, args.output, args.template)
        print(f"[gate] {res['verdict']}  (score {res['overall_score']}/100) "
              f"-> {res.get('_path', 'gate_result.json')}")
        for item in res["fail_items"]:
            print(f"  FAIL {item}")
        for item in res["warn_items"]:
            print(f"  warn {item}")
    except Exception as exc:
        print(f"[gate] check skipped: {exc}")

    print(args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
