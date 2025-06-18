"""
This module is the core of pycombiner.
"""

# 定义包版本
__version__ = "2.0.0"

# 导入核心函数，方便外部直接访问
from .file_handler import find_python_files, read_file
from .ast_parser import analyze_file, build_dependency_graph
from .merger import merge_files, topological_sort_files
from .combiner import PyCombiner
from .config import Config
from .utils import *

__all__ = [
    "find_python_files",
    "read_file",
    "analyze_file",
    "build_dependency_graph",
    "merge_files",
    "topological_sort_files",
    "PyCombiner",
    "Config",
    '__version__',
]
