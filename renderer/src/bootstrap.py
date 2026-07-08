#!/usr/bin/env python3
"""OW runtime bootstrap — locate, extract, and verify the deck runtime.

MUST be *executed* in Code Interpreter (not described). Running it is also what
makes ChatGPT stage the GPT's Knowledge files into /mnt/data — those files are
copied in lazily on the first Code Interpreter invocation, so a first miss is
retried after a short wait before concluding anything is missing.

    python bootstrap.py        # or:  exec(open('bootstrap.py').read())

On success prints exactly one gate line:
    RUNTIME OK v2.1.0 TYPE_KPI=26  src=/mnt/data/ow_runtime/src
On failure prints RUNTIME MISSING plus a listing of /mnt/data for diagnosis.
Never hand-build a deck unless this has actually run and printed MISSING.
"""
import os, sys, zipfile, glob, time

DEST = "/mnt/data/ow_runtime"
ROOTS = ("/mnt/data", "/mnt", os.getcwd())
ZIP_PATTERNS = ("deck_system_runtime*clean*.zip", "deck_system_runtime*.zip",
                "*runtime*clean*.zip", "*runtime*.zip")


def _find_src():
    for c in (os.path.join(DEST, "src"), DEST):
        if os.path.isfile(os.path.join(c, "runtime.py")):
            return c
    for r in ROOTS:
        h = glob.glob(os.path.join(r, "**", "runtime.py"), recursive=True)
        if h:
            return os.path.dirname(h[0])
    return None


def _find_zip():
    for r in ROOTS:
        for p in ZIP_PATTERNS:
            h = sorted(glob.glob(os.path.join(r, "**", p), recursive=True))
            if h:
                return h[0]
    return None


def _try_once():
    src = _find_src()
    if not src:
        z = _find_zip()
        if z:
            os.makedirs(DEST, exist_ok=True)
            zipfile.ZipFile(z).extractall(DEST)
            src = _find_src()
    return src


def load(retries: int = 2, wait: float = 2.0):
    src = _try_once()
    # Knowledge files copy into /mnt/data lazily on first CI run — retry.
    for _ in range(retries):
        if src:
            break
        time.sleep(wait)
        src = _try_once()
    if not src:
        here = sorted(os.listdir("/mnt/data")) if os.path.isdir("/mnt/data") else "N/A"
        print("RUNTIME MISSING. /mnt/data contains:", here)
        print("Attach deck_system_runtime_clean.zip to THIS chat (drag it into the "
              "message box) — it lands in /mnt/data. Do NOT hand-build the deck.")
        return None
    if src not in sys.path:
        sys.path.insert(0, src)
    import runtime  # noqa: E402
    print(f"RUNTIME OK v{runtime.__version__} TYPE_KPI={runtime.Tokens.TYPE_KPI}  src={src}")
    return src


if __name__ == "__main__":
    sys.exit(0 if load() else 1)
