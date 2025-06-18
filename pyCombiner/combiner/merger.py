"""
Merger module for combining Python files
"""

import os
from typing import List, Dict, Tuple, Set
from collections import defaultdict, deque
from pathlib import Path
from .file_handler import read_file
from .ast_parser import analyze_file, get_module_name, ImportInfo

def topological_sort_files(dependency_graph: Dict[str, Set[str]]) -> List[Path]:
    """Sort files based on their dependencies"""
    visited = set()
    temp = set()
    order = []
    
    def visit(node: str):
        if node in temp:
            raise ValueError(f"Circular dependency detected: {node}")
        if node in visited:
            return
            
        temp.add(node)
        for neighbor in dependency_graph.get(node, set()):
            visit(neighbor)
        temp.remove(node)
        visited.add(node)
        order.append(Path(node))
    
    for node in dependency_graph:
        if node not in visited:
            visit(node)
            
    return order[::-1]  # Reverse to get correct order

def deduplicate_imports(all_imports: List[ImportInfo]) -> List[str]:
    """
    从所有文件中收集到的导入信息中去重，并格式化为字符串列表。

    Args:
        all_imports: 所有文件提取的原始导入信息列表。

    Returns:
        去重并格式化后的导入语句字符串列表。
    """
    from_imports: Dict[str, Set[str]] = defaultdict(set)
    simple_imports: Set[str] = set()
    
    for module_name, imported_name in all_imports:
         if imported_name == module_name or imported_name is None:
              simple_imports.add(f"import {module_name}")
         elif imported_name == '*':
              simple_imports.add(f"from {module_name} import *")
         else:
              from_imports[module_name].add(imported_name)

    formatted_imports: List[str] = sorted(list(simple_imports))
    for module in sorted(from_imports.keys()):
        names = sorted(list(from_imports[module]))
        if names:
             formatted_imports.append(f"from {module} import {', '.join(names)}")

    return formatted_imports

def merge_files(
    files: List[Path],
    dependency_graph: Dict[str, Set[str]],
    source_dir: Path,
    output_file: Path,
    entry_file: Path = None
) -> None:
    """Merge Python files into a single file"""
    print("\n[DEBUG] Starting file merge process:")
    print("----------------------------------------")
    
    # Use provided entry file or find main.py
    if not entry_file:
        for file in files:
            if file.name == 'main.py':
                entry_file = file
                break
    
    if not entry_file:
        raise ValueError("No entry file specified and no main.py found in the source directory")
    
    print(f"\n[DEBUG] Entry point: {entry_file}")
    
    # Get all files that are referenced from the entry point
    referenced_files = set()
    to_process = {entry_file}
    processed = set()
    
    while to_process:
        current_file = to_process.pop()
        if current_file in processed:
            continue
            
        processed.add(current_file)
        referenced_files.add(current_file)
        
        # Get imports from current file
        abs_path = source_dir / current_file if not current_file.is_absolute() else current_file
        content = read_file(abs_path)
        if not content:
            continue
            
        imports, _ = analyze_file(content, str(abs_path))
        
        # Find referenced files
        for imp in imports:
            if imp.is_from_import:
                module_parts = imp.module.split('.')
                if module_parts[0] in ['utils', 'models', 'services']:  # Add other local module prefixes
                    # Convert module path to file path
                    module_path = Path(*module_parts)
                    py_file = source_dir / f"{module_path}.py"
                    if py_file.exists():
                        to_process.add(py_file)
            elif imp.module.split('.')[0] in ['utils', 'models', 'services']:  # Add other local module prefixes
                # Convert module path to file path
                module_path = Path(imp.module)
                py_file = source_dir / f"{module_path}.py"
                if py_file.exists():
                    to_process.add(py_file)
    
    print("\n[DEBUG] Referenced files:")
    for file in sorted(referenced_files):
        print(f"  - {file}")
    
    # Read and merge file contents
    merged_content = []
    processed_imports = set()
    main_content = []
    module_definitions = {}  # Maps module names to their defined names
    all_imports = []  # Store all imports for later processing
    main_py_content = []  # Store main.py's content
    local_modules = set()  # Store all local module names
    
    print("\n[DEBUG] Collecting module definitions and imports:")
    print("----------------------------------------")
    
    # First pass: collect all module definitions and imports
    for file_path in referenced_files:
        abs_path = source_dir / file_path if not file_path.is_absolute() else file_path
        content = read_file(abs_path)
        if not content:
            continue
            
        module_name = get_module_name(abs_path, source_dir)
        local_modules.add(module_name)
        # Also add parent modules
        parts = module_name.split('.')
        for i in range(1, len(parts)):
            local_modules.add('.'.join(parts[:i]))
            
        imports, defined_names = analyze_file(content, str(abs_path))
        module_definitions[module_name] = defined_names
        all_imports.extend(imports)
        
        # If this is the entry file, collect its content
        if file_path == entry_file:
            lines = content.split('\n')
            in_main_function = False
            for line in lines:
                if line.strip().startswith('def main()'):
                    in_main_function = True
                    main_py_content.append(line)
                elif in_main_function and line.strip() and not line.strip().startswith('if __name__'):
                    main_py_content.append(line)
                elif line.strip().startswith('if __name__'):
                    main_py_content.append(line)
                    main_py_content.append('    main()')
                    break
            print(f"  - Found entry file content: {len(main_py_content)} lines")
        
        print(f"\n[DEBUG] Module {module_name} defines:")
        for name in defined_names:
            print(f"  - {name}")
    
    print("\n[DEBUG] Local modules:")
    for module in sorted(local_modules):
        print(f"  - {module}")
    
    print("\n[DEBUG] Processing imports:")
    print("----------------------------------------")
    
    # Process all imports
    for imp in all_imports:
        # Skip if the imported name is defined in our code
        if imp.is_from_import:
            module_parts = imp.module.split('.')
            if module_parts[0] in local_modules:
                print(f"  - Skipping {imp} (local module)")
                continue
        elif imp.module.split('.')[0] in local_modules:
            print(f"  - Skipping {imp} (local module)")
            continue
            
        # Format the import statement
        if imp.is_from_import:
            if imp.name == '*':
                import_stmt = f"from {imp.module} import *"
            else:
                import_stmt = f"from {imp.module} import {imp.name}"
        else:
            import_stmt = f"import {imp.module}"
            
        if import_stmt not in processed_imports:
            processed_imports.add(import_stmt)
            merged_content.append(import_stmt)
    
    # Add a blank line after imports
    if merged_content:
        merged_content.append('')
    
    # Second pass: write file contents
    for file_path in referenced_files:
        abs_path = source_dir / file_path if not file_path.is_absolute() else file_path
        content = read_file(abs_path)
        if not content:
            continue
            
        # Write file header
        merged_content.append(f"\n#{'='*80}")
        merged_content.append(f"# File: {file_path}")
        merged_content.append(f"#{'='*80}\n")
        
        # Write non-import statements
        lines = content.split('\n')
        in_import_section = True
        for line in lines:
            stripped = line.strip()
            if in_import_section and not (stripped.startswith('import ') or stripped.startswith('from ')):
                in_import_section = False
            if not in_import_section:
                merged_content.append(line)
    
    # Write the merged content to the output file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(merged_content))
    print(f"  - Output written to {output_file}")
