"""
High-frequency memory logger for OOM debugging.
Call start_mem_logger() at app startup when ENABLE_MEM_LOG=1.
"""
import os
import time
import threading

try:
    import psutil
except ImportError:
    psutil = None


def start_mem_logger(interval_s: float = 0.5) -> None:
    """Start a daemon thread that logs RSS and peak memory every interval_s seconds."""
    if not psutil:
        print("[mem] psutil not installed, skipping memory logger")
        return

    p = psutil.Process(os.getpid())

    def loop():
        peak = 0
        while True:
            rss = p.memory_info().rss
            peak = max(peak, rss)
            print(
                f"[mem] rss_mb={rss/1024/1024:.1f} peak_mb={peak/1024/1024:.1f}",
                flush=True,
            )
            time.sleep(interval_s)

    t = threading.Thread(target=loop, daemon=True)
    t.start()
    print(f"[mem] Started memory logger (interval={interval_s}s)", flush=True)


def log_mem(tag: str) -> float:
    """Log current RSS for a single point. Returns RSS in MB."""
    if not psutil:
        return 0.0
    p = psutil.Process(os.getpid())
    rss_mb = p.memory_info().rss / (1024 * 1024)
    print(f"[mem] {tag} rss_mb={rss_mb:.1f}", flush=True)
    return rss_mb
