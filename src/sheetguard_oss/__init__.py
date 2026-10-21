"""Read-only XLSX release QA."""
from .core import VERSION as __version__, Finding, Report, Sheet, inspect, compare
__all__ = ["__version__", "Finding", "Report", "Sheet", "inspect", "compare"]
