import sys

# Pastikan output UTF-8 aman di terminal Windows
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from cleaner.reader import baca_data, read_csv_with_fallback
from cleaner.engine import CleaningEngine

__version__ = "1.0.0"
__all__ = ["baca_data", "read_csv_with_fallback", "CleaningEngine"]

