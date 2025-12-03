"""
Version management for DerivativeMill

This file maintains the current version of the application.
Version format: vMAJOR.MINOR (e.g., v0.6)
"""

__version__ = "v0.6"

def get_version():
    """Get the current version"""
    return __version__

def increment_patch():
    """
    Increment the patch version (e.g., v0.6 -> v0.7)
    Call this when making a new release
    """
    global __version__
    parts = __version__[1:].split('.')  # Remove 'v' and split
    major, minor = int(parts[0]), int(parts[1])
    minor += 1
    __version__ = f"v{major}.{minor}"
    return __version__
