"""
Utility functions: logging, configuration, error handling, math helpers.
"""
import os
import math
import logging
import functools

# ---------------------------------------------------------------------------
#  Project root
# ---------------------------------------------------------------------------

def get_project_root():
    """Returns the project root directory.

    When running from a PyInstaller bundle, ``sys._MEIPASS`` points to the
    temp extraction folder that contains the bundled data files.
    """
    import sys
    if getattr(sys, '_MEIPASS', None):
        return sys._MEIPASS
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ---------------------------------------------------------------------------
#  Logging
# ---------------------------------------------------------------------------

def setup_logger(name):
    """Creates a console logger with timestamp + level + message."""
    logger = logging.getLogger(name)
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            "[%(asctime)s] %(levelname)-7s %(name)s — %(message)s",
            datefmt="%H:%M:%S",
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)
    return logger


# ---------------------------------------------------------------------------
#  COM error decorator
# ---------------------------------------------------------------------------

def com_safe(func):
    """Decorator that catches COM errors and logs them, returning None."""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger = setup_logger(func.__module__ or "com_safe")
            logger.error(f"{func.__name__} failed: {e}")
            return None
    return wrapper


# ---------------------------------------------------------------------------
#  Configuration
# ---------------------------------------------------------------------------

_DEFAULT_CONFIG = {
    "centerline": {
        "straightness_tolerance": 0.5,
        "dedup_tolerance": 5.0,
        "min_segment_length": 1.0,
    },
    "pipe": {
        "default_radius": 2.0,
        "default_wall_thickness": 1.0,
    },
    "connect": {
        "continuity": 1,
        "tension": 1.0,
    },
}


def load_config(config_path=None):
    """
    Load configuration from defaults.yaml.
    Falls back to hardcoded defaults if file is missing.
    """
    if config_path is None:
        config_path = os.path.join(get_project_root(), "config", "defaults.yaml")

    try:
        import yaml
        with open(config_path, "r") as f:
            cfg = yaml.safe_load(f)
        if cfg:
            # Merge with defaults so missing keys still have values
            merged = _DEFAULT_CONFIG.copy()
            for section in merged:
                if section in cfg:
                    merged[section].update(cfg[section])
            return merged
    except FileNotFoundError:
        pass
    except ImportError:
        pass
    except Exception:
        pass

    return _DEFAULT_CONFIG.copy()


# ---------------------------------------------------------------------------
#  Math helpers
# ---------------------------------------------------------------------------

def distance_3d(p1, p2):
    """Euclidean distance between two dicts with 'x', 'y', 'z' keys."""
    return math.sqrt(
        (p2["x"] - p1["x"]) ** 2 +
        (p2["y"] - p1["y"]) ** 2 +
        (p2["z"] - p1["z"]) ** 2
    )
