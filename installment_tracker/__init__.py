"""Track credit-card purchases paid in installments, with a live Excel tracker."""

from .config import Config, load_config
from .model import Purchase, Schedule, build_schedule, read_purchases

__all__ = ["Config", "load_config", "Purchase", "Schedule", "build_schedule", "read_purchases"]
__version__ = "0.1.0"
