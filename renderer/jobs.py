"""In-memory background jobs for /agent.

Why: a full deck generation (storyline + parallel slide fills + render +
repair pass) can take several minutes with the strong model. Render.com's
proxy kills long-lived HTTP requests well before that, which the pane sees as
an opaque 502/timeout. The fix is submit-and-poll: POST /agent with
{"background": true} returns a job id immediately, the work continues in a
daemon thread, and the pane polls GET /agent/jobs/{id} (each poll is a fast,
cheap request that also keeps the free-tier instance awake).

Single-process only (uvicorn default, one Render instance) — the store is a
plain dict. If the service restarts mid-job the job is gone; the poller gets
a 404 and tells the user to retry.

Pipeline code reports progress via progress("stage text"); it is a no-op
outside a job, so the same code serves the synchronous path unchanged.
"""
from __future__ import annotations

import threading
import time
import uuid
from typing import Any, Callable

_JOBS: dict[str, dict[str, Any]] = {}
_LOCK = threading.Lock()
_TTL_SECONDS = 30 * 60          # forget finished jobs after 30 min
_MAX_JOBS = 50                  # hard cap — this is a single-user test service

# The job dict for the CURRENT thread (set in the worker; inherited nowhere
# else — milestone progress() calls all happen on the worker thread itself).
_local = threading.local()


def progress(stage: str) -> None:
    """Record a human-readable stage on the current job. No-op outside a job."""
    job = getattr(_local, "job", None)
    if job is not None:
        job["stage"] = stage
        job["updated"] = time.time()


def submit(fn: Callable[[], Any]) -> str:
    """Run fn() on a daemon thread; return the job id immediately."""
    job_id = uuid.uuid4().hex
    job: dict[str, Any] = {
        "id": job_id, "status": "running", "stage": "queued",
        "created": time.time(), "updated": time.time(),
        "result": None, "error": None,
    }
    with _LOCK:
        _reap_locked()
        if len(_JOBS) >= _MAX_JOBS:
            raise RuntimeError("too many concurrent jobs — retry in a minute")
        _JOBS[job_id] = job

    def _run() -> None:
        _local.job = job
        try:
            job["result"] = fn()
            job["status"] = "done"
            job["stage"] = "done"
        except Exception as exc:                              # noqa: BLE001
            import traceback
            traceback.print_exc()
            job["error"] = f"{type(exc).__name__}: {exc}"
            job["status"] = "error"
        finally:
            job["updated"] = time.time()
            _local.job = None

    threading.Thread(target=_run, daemon=True, name=f"job-{job_id[:8]}").start()
    return job_id


def get(job_id: str) -> dict[str, Any] | None:
    with _LOCK:
        _reap_locked()
        return _JOBS.get(job_id)


def public_view(job: dict[str, Any]) -> dict[str, Any]:
    """What the poller sees. The (large) result is only attached when done."""
    out = {
        "job_id": job["id"], "status": job["status"], "stage": job["stage"],
        "elapsed_s": round(time.time() - job["created"], 1),
    }
    if job["status"] == "done":
        out["result"] = job["result"]
    elif job["status"] == "error":
        out["error"] = job["error"]
    return out


def _reap_locked() -> None:
    now = time.time()
    dead = [k for k, j in _JOBS.items()
            if j["status"] != "running" and now - j["updated"] > _TTL_SECONDS]
    for k in dead:
        del _JOBS[k]
