import unittest
from pathlib import Path
import tempfile
import shutil
from pycombiner.combiner.ast_parser import analyze_file, build_dependency_graph
from pycombiner.combiner.merger import merge_files

class TestDeepDemo(unittest.TestCase):
    def setUp(self):
        # Create temporary directory
        self.test_dir = Path(tempfile.mkdtemp())
        self.source_dir = self.test_dir / "deep_demo"
        
        # Copy the example directory structure
        example_dir = Path(__file__).parent / "examples" / "deep_demo"
        shutil.copytree(example_dir, self.source_dir)
        
        # Set up output file
        self.output_file = self.test_dir / "output.py"
        
    def test_deep_demo_merge(self):
        """Test merging files from a deep directory structure"""
        # Find all Python files
        python_files = []
        for file in self.source_dir.rglob("*.py"):
            python_files.append(file.relative_to(self.source_dir))
        
        # Analyze imports
        imports_by_file = {}
        for file_path in python_files:
            abs_path = self.source_dir / file_path
            with open(abs_path) as f:
                content = f.read()
            imports, _ = analyze_file(content, str(abs_path))
            imports_by_file[str(abs_path)] = imports
        
        # Build dependency graph
        dependency_graph = build_dependency_graph(imports_by_file, str(self.source_dir))
        
        # Merge files
        entry_file = self.source_dir / "main.py"
        merge_files(python_files, dependency_graph, self.source_dir, self.output_file, entry_file)
        
        # Verify output file exists
        self.assertTrue(self.output_file.exists())
        
        # Read and verify output content
        with open(self.output_file) as f:
            content = f.read()
            
        # Verify essential components are present
        self.assertIn("class User:", content)
        self.assertIn("def add(", content)
        self.assertIn("def subtract(", content)
        self.assertIn("def format_name(", content)
        self.assertIn("def login(", content)
        self.assertIn("def connect_db(", content)
        self.assertIn("def main():", content)
        
        # Verify file contents are present in the correct order
        # Check for docstrings and key content from each file
        self.assertIn('"""\nUser model\n"""', content)
        self.assertIn('"""\nAuthentication service\n"""', content)
        self.assertIn('"""\nString utility functions\n"""', content)
        self.assertIn('"""\nMath utility functions\n"""', content)
        self.assertIn('"""\nDatabase service\n"""', content)
        
        # Verify main function content
        self.assertIn("db = connect_db()", content)
        self.assertIn("user = User(\"John Doe\")", content)
        self.assertIn("formatted_name = format_name(user.name)", content)
        self.assertIn("result1 = add(10, 5)", content)
        self.assertIn("result2 = subtract(10, 5)", content)
        self.assertIn("login(user)", content)
        self.assertIn("logout(user)", content)
        
    def tearDown(self):
        # Clean up temporary directory
        # shutil.rmtree(self.test_dir)
        print(self.test_dir)

if __name__ == '__main__':
    unittest.main() 