"""
Main module for pycombiner
"""
from pathlib import Path
from typing import Dict, List, Set, Tuple
import ast
from .output import MergeReport, print_merge_report

class PyCombiner:
    def __init__(self, entry_file: Path, source_dir: Path, output_file: Path, debug: bool = False, show_details: bool = False):
        self.entry_file = entry_file
        self.source_dir = source_dir
        self.output_file = output_file
        self.debug = debug
        self.show_details = show_details
        self.report = MergeReport(entry_file, source_dir, output_file, debug, show_details)
        self.imports_by_file: Dict[str, List[str]] = {}
        self.dependency_graph: Dict[str, List[str]] = {}
        self.merge_order: List[Path] = []
        self.stats = {
            'total_imports': 0,
            'duplicate_imports': 0,
            'redundant_imports': 0,
            'total_lines': 0,
            'functions': 0,
            'classes': 0
        }

    def debug_print(self, message: str):
        """Print debug message if debug mode is enabled"""
        if self.debug:
            print(f"\033[33m[DEBUG]\033[0m {message}")

    def _is_relative_import(self, import_path: str) -> bool:
        """Check if an import is relative to the source directory"""
        try:
            # Convert import path to possible file paths
            possible_paths = []
            
            # Check for direct file
            file_path = self.source_dir / import_path.replace('.', '/')
            if file_path.suffix != '.py':
                file_path = file_path.with_suffix('.py')
            possible_paths.append(file_path)
            
            # Check for package with __init__.py
            init_path = self.source_dir / import_path.replace('.', '/') / '__init__.py'
            possible_paths.append(init_path)
            
            # Check if any of the possible paths exist in our project
            for path in possible_paths:
                if path.exists() and path.is_file():
                    # Verify the file is within our source directory
                    try:
                        path.resolve().relative_to(self.source_dir)
                        return True
                    except ValueError:
                        continue
            
            return False
        except Exception:
            return False

    def _get_import_path(self, import_path: str) -> Path:
        """Convert import path to file path"""
        import_path = import_path.replace('.', '/')
        if not import_path.endswith('.py'):
            import_path += '.py'
        return self.source_dir / import_path

    def _parse_imports(self, file_path: Path) -> Tuple[List[str], Set[str]]:
        """Parse imports from a Python file and return ordered imports and unhandled imports"""
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        try:
            tree = ast.parse(content)
        except SyntaxError as e:
            self.debug_print(f"Syntax error in {file_path}: {e}")
            return [], set()

        ordered_imports = []  # Keep track of import order
        unhandled_imports = set()

        for node in ast.walk(tree):
            if isinstance(node, (ast.Import, ast.ImportFrom)):
                if isinstance(node, ast.Import):
                    for name in node.names:
                        import_path = name.name
                        if self._is_relative_import(import_path):
                            ordered_imports.append(import_path)
                        else:
                            unhandled_imports.add(import_path)
                else:  # ImportFrom
                    if node.module:
                        if self._is_relative_import(node.module):
                            ordered_imports.append(node.module)
                        else:
                            unhandled_imports.add(node.module)

        return ordered_imports, unhandled_imports

    def _build_dependency_graph(self):
        """Build dependency graph between files based on import order"""
        for file_path in self.source_dir.rglob('*.py'):
            ordered_imports, _ = self._parse_imports(file_path)
            self.dependency_graph[str(file_path)] = []
            for imp in ordered_imports:
                imp_path = self._get_import_path(imp)
                if imp_path.exists():
                    self.dependency_graph[str(file_path)].append(str(imp_path))

    def _get_merge_order(self) -> List[Path]:
        """Get the order to merge files based on dependencies"""
        visited = set()
        order = []

        def visit(file_path: str):
            if file_path in visited:
                return
            visited.add(file_path)
            # Process dependencies in their original order from the file
            for dep in self.dependency_graph.get(file_path, []):
                visit(dep)
            order.append(Path(file_path))

        # Start with entry file
        visit(str(self.entry_file))

        # Add any remaining files in their original order
        for file_path in self.source_dir.rglob('*.py'):
            if str(file_path) not in visited:
                visit(str(file_path))

        return order

    def _merge_files(self):
        """Merge all Python files in the correct order"""
        with open(self.output_file, 'w', encoding='utf-8') as out:
            # Write header
            out.write(f"# generated by pycombiner\n")
            out.write(f"# Entry file: {self.entry_file}\n")
            out.write(f"# Source directory: {self.source_dir}\n\n")

            # Track imports to avoid duplicates
            unhandled_imports = set()  # Only track imports that can't be resolved
            
            # First pass: collect all unhandled imports and update stats
            for file_path in self.merge_order:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                try:
                    tree = ast.parse(content)
                except SyntaxError as e:
                    self.debug_print(f"Syntax error in {file_path}: {e}")
                    continue

                # Track imports for this file
                file_unhandled_imports = set()
                file_handled_imports = set()
                info = {}

                # Process imports using AST
                for node in ast.iter_child_nodes(tree):
                    if isinstance(node, (ast.Import, ast.ImportFrom)):
                        self.stats['total_imports'] += 1  # 增加导入语句计数
                        if isinstance(node, ast.Import):
                            for name in node.names:
                                import_path = name.name
                                if not self._is_relative_import(import_path):
                                    import_stmt = f"import {import_path}"
                                    if import_stmt in unhandled_imports:
                                        self.stats['duplicate_imports'] += 1  # 增加重复导入计数
                                    unhandled_imports.add(import_stmt)
                                    file_unhandled_imports.add(import_stmt)
                                    # 保存原始导入语句
                                    if 'unhandled_import_statements' not in info:
                                        info['unhandled_import_statements'] = []
                                    info['unhandled_import_statements'].append(import_stmt)
                                else:
                                    if import_path in file_handled_imports:
                                        self.stats['redundant_imports'] += 1  # 增加冗余导入计数
                                    file_handled_imports.add(import_path)
                                    # 保存原始导入语句
                                    if 'import_statements' not in info:
                                        info['import_statements'] = []
                                    info['import_statements'].append(f"import {import_path}")
                        else:  # ImportFrom
                            if node.module:
                                if not self._is_relative_import(node.module):
                                    import_stmt = f"from {node.module} import {', '.join(n.name for n in node.names)}"
                                    if import_stmt in unhandled_imports:
                                        self.stats['duplicate_imports'] += 1  # 增加重复导入计数
                                    unhandled_imports.add(import_stmt)
                                    file_unhandled_imports.add(import_stmt)
                                    # 保存原始导入语句
                                    if 'unhandled_import_statements' not in info:
                                        info['unhandled_import_statements'] = []
                                    info['unhandled_import_statements'].append(import_stmt)
                                else:
                                    if node.module in file_handled_imports:
                                        self.stats['redundant_imports'] += 1  # 增加冗余导入计数
                                    file_handled_imports.add(node.module)
                                    # 保存原始导入语句
                                    if 'import_statements' not in info:
                                        info['import_statements'] = []
                                    info['import_statements'].append(f"from {node.module} import {', '.join(n.name for n in node.names)}")

                # Update file info in report
                self.report.add_file_info(
                    file_path,
                    len(content.split('\n')),
                    file_handled_imports,
                    file_unhandled_imports,
                    info
                )

            # Write all unhandled imports at the beginning
            for imp in sorted(unhandled_imports):
                out.write(imp + '\n')
            out.write('\n')

            # Second pass: write file contents
            for idx, file_path in enumerate(self.merge_order, 1):
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                try:
                    tree = ast.parse(content)
                except SyntaxError as e:
                    self.debug_print(f"Syntax error in {file_path}: {e}")
                    continue

                # Write file header
                out.write(f"\n#{'='*80}\n")
                out.write(f"# [{idx}] {file_path.name} : {file_path}\n")
                out.write(f"#{'='*80}\n\n")

                # Get the line numbers of handled imports to skip
                handled_import_lines = set()
                for node in ast.iter_child_nodes(tree):
                    if isinstance(node, (ast.Import, ast.ImportFrom)):
                        if isinstance(node, ast.Import):
                            for name in node.names:
                                if self._is_relative_import(name.name):
                                    handled_import_lines.add(node.lineno)
                        else:  # ImportFrom
                            if node.module and self._is_relative_import(node.module):
                                handled_import_lines.add(node.lineno)
                                # Also skip the 'from' line
                                handled_import_lines.add(node.lineno - 1)

                # Write content, skipping only handled import lines
                lines = content.split('\n')
                for i, line in enumerate(lines, 1):
                    if i not in handled_import_lines:
                        out.write(line + '\n')

                # Update stats
                for node in ast.iter_child_nodes(tree):
                    if isinstance(node, ast.FunctionDef):
                        self.stats['functions'] += 1
                    elif isinstance(node, ast.ClassDef):
                        self.stats['classes'] += 1

    def combine(self):
        """Combine all Python files into a single file"""
        self.debug_print("Starting file combination process...")
        
        # Build dependency graph
        self.debug_print("Building dependency graph...")
        self._build_dependency_graph()
        self.report.set_dependency_graph(self.dependency_graph)

        # Get merge order
        self.debug_print("Determining merge order...")
        self.merge_order = self._get_merge_order()
        self.report.set_merge_order(self.merge_order)

        # Process each file for report
        self.debug_print("Processing files...")
        for file_path in self.merge_order:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = len(f.readlines())
            self.report.add_file_info(file_path, lines, set(), set())  # Empty sets as imports are handled in _merge_files

        # Merge files
        self.debug_print("Merging files...")
        self._merge_files()

        # Update report
        self.debug_print("Updating report...")
        self.report.update_stats(self.stats)

        # Print report
        self.debug_print("Printing report...")
        print_merge_report(self.report) 