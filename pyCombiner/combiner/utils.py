"""Utility functions for pycombiner."""

def is_python_file(filename: str) -> bool:
    """Check if a file is a Python file."""
    return filename.endswith('.py')

def get_module_name(filepath: str) -> str:
    """Get module name from filepath."""
    return filepath.replace('/', '.').replace('\\', '.')[:-3] 