"""
pycombiner - A tool for combining Python files
"""

import os
from pathlib import Path

def _get_version():
    """Get version from pyproject.toml"""
    try:
        # Get the directory containing this file
        package_dir = Path(__file__).parent.parent
        pyproject_path = package_dir / "pyproject.toml"
        
        if not pyproject_path.exists():
            # If not found in parent directory, try the current directory
            pyproject_path = Path("pyproject.toml")
            
        if not pyproject_path.exists():
            return "0.0.0"  # Fallback version
            
        with open(pyproject_path, "rb") as f:
            import tomli
            return tomli.load(f)["project"]["version"]
    except Exception:
        return "0.0.0"  # Fallback version

__version__ = _get_version()

from .combiner import PyCombiner

__all__ = ["PyCombiner"] 