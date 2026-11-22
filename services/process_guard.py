import logging
import time
from typing import Optional

import psutil

logger = logging.getLogger(__name__)


def watch_process(pid: int) -> Optional[psutil.Process]:
    """Capture a handle to pid now so it can be killed later without PID-reuse risk.

    Returns None if the process is already gone.
    """
    try:
        return psutil.Process(pid)
    except (psutil.NoSuchProcess, ValueError):
        return None


def ensure_gone(proc: Optional[psutil.Process], label: str = "", grace: float = 5.0,
                poll: float = 0.5) -> bool:
    """Wait up to `grace` seconds for proc to exit on its own, then force-kill it.

    Returns True only if a kill was actually needed.
    """
    if proc is None:
        return False

    deadline = time.monotonic() + grace
    while time.monotonic() < deadline:
        try:
            if not proc.is_running() or proc.status() == psutil.STATUS_ZOMBIE:
                return False
        except psutil.NoSuchProcess:  # exited between is_running() and status()
            return False
        time.sleep(poll)

    try:
        proc.kill()
        proc.wait(timeout=grace)
        logger.warning("%s: force-killed leftover process PID %s", label, proc.pid)
        return True
    except psutil.NoSuchProcess:
        return False
    except Exception as exc:
        logger.error("%s: could not kill PID %s: %s", label, proc.pid, exc)
        return False
