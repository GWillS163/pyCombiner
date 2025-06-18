"""
AST Parser module for analyzing Python files and their dependencies
"""

import ast
import os
from typing import List, Tuple, Dict, Set, NamedTuple, Optional
from pathlib import Path
from dataclasses import dataclass

@dataclass
class ImportInfo:
    """Information about an import statement"""
    module: str
    name: str
    is_from_import: bool
    alias: Optional[str] = None
def analyze_file(content: str, filepath: str) -> Tuple[List[ImportInfo], Set[str]]:
    """Analyze a Python file and return its imports and defined names"""
    imports = []
    defined_names = set()
    
    try:
        tree = ast.parse(content)
        print(f"\n[DEBUG] Analyzing file: {filepath}")
        
        # First pass: collect all defined names
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef):
                defined_names.add(node.name)
                print(f"[DEBUG]{('[Function definition]'):>25} \tdef {node.name}()")
            elif isinstance(node, ast.ClassDef):
                defined_names.add(node.name)
                print(f"[DEBUG]{('[Class definition]'):>25} \tclass {node.name}")
            elif isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store):
                defined_names.add(node.id)
                print(f"[DEBUG]{('[Variable definition]'):>25} \t{node.id}")
        
        # Second pass: collect imports
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for name in node.names:
                    import_info = ImportInfo(
                        module=name.name,
                        name=name.name,
                        is_from_import=False,
                        alias=name.asname
                    )
                    imports.append(import_info)
                    print(f"[DEBUG]{'[Import statement]':>25} \timport {name.name}")
            elif isinstance(node, ast.ImportFrom):
                module = node.module if node.module else ''
                for name in node.names:
                    import_info = ImportInfo(
                        module=module,
                        name=name.name,
                        is_from_import=True,
                        alias=name.asname
                    )
                    imports.append(import_info)
                    print(f"[DEBUG]{('[From-import statement]'):>25} \tfrom {module} import {name.name}")
    except SyntaxError as e:
        print(f"Warning: Syntax error in {filepath}: {e}")
    return imports, defined_names

def _resolve_module_to_filepath(module_name: str, project_files: List[str], input_dir: str) -> str | None:
    """
    尝试将导入的模块名映射到项目中的文件路径。
    这是一个简化的实现，对于复杂的包结构或第三方库导入可能不准确。

    Args:
        module_name: 导入的模块名 (e.g., 'my_module.utils').
        project_files: 项目中所有待处理的文件路径列表。
        input_dir: 项目的根目录路径。

    Returns:
        对应的文件路径，如果找到则返回绝对路径，否则返回 None。
    """
    # 移除相对导入前缀 ('.', '..', etc.) for simplified mapping
    original_module_name = module_name
    while module_name.startswith('.'):
        module_name = module_name[1:]
    if not module_name: # If it was just '.' or '..', etc.
         # This case is tricky and needs context of the importing file's path
         # We'll skip resolving pure relative imports for now or handle them differently
         return None # Cannot resolve purely relative imports without more context

    # 尝试匹配 module_name 到 project_files
    # 模块名 'a.b.c' 可能对应文件 'path/to/a/b/c.py' 或目录 'path/to/a/b/c/__init__.py'
    # 构造可能的路径后缀
    module_path_suffix = module_name.replace('.', os.sep)

    # 检查作为文件存在
    possible_filepath = os.path.join(input_dir, module_path_suffix + ".py")
    if possible_filepath in project_files: # Check if this file is in our processed list
        return possible_filepath

    # 检查作为包存在 (__init__.py)
    possible_initpath = os.path.join(input_dir, module_path_suffix, "__init__.py")
    if possible_initpath in project_files: # Check if this init file is in our processed list
        return possible_initpath

    # print(f"DEBUG: Could not resolve module '{original_module_name}' to a file in processed list.") # Debugging
    return None # 未找到匹配的文件

def get_module_name(file_path: Path, source_dir: Path) -> str:
    """Get the module name for a file relative to the source directory"""
    # Convert to absolute paths
    file_path = Path(file_path).resolve()
    source_dir = Path(source_dir).resolve()
    
    try:
        relative_path = file_path.relative_to(source_dir)
    except ValueError:
        # If file_path is not in source_dir, use the file name without extension
        return file_path.stem
        
    return str(relative_path.with_suffix('')).replace('\\', '.').replace('/', '.')

def build_dependency_graph(imports_by_file: Dict[str, List[ImportInfo]], source_dir: str) -> Dict[str, Set[str]]:
    """Build a dependency graph from imports"""
    graph = {}
    source_path = Path(source_dir).resolve()
    
    for file_path, imports in imports_by_file.items():
        file_path = Path(file_path).resolve()
        module_name = get_module_name(file_path, source_path)
        graph[module_name] = set()
        
        for imp in imports:
            if imp.is_from_import:
                # Handle relative imports
                if imp.module.startswith('.'):
                    # Convert relative import to absolute
                    parts = imp.module.split('.')
                    current_dir = file_path.parent
                    for _ in range(len(parts) - 1):
                        current_dir = current_dir.parent
                    target_module = get_module_name(current_dir / f"{imp.name}.py", source_path)
                else:
                    target_module = imp.module
            else:
                target_module = imp.module
                
            # Only add local module dependencies
            if target_module.split('.')[0] in ['utils', 'models', 'services']:
                graph[module_name].add(target_module)
                
    return graph
