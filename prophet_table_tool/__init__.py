"""Prophet Table Change Consolidation & Integration Tool."""

from .changelog import generate_change_log
from .integrate import integrate_changes

__all__ = ["generate_change_log", "integrate_changes"]
__version__ = "1.0.0"
