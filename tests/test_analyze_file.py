import unittest
from pathlib import Path
import tempfile
from pycombiner.combiner.ast_parser import analyze_file

class TestAnalyzeFile(unittest.TestCase):
    def setUp(self):
        # Create temporary directory
        self.test_dir = tempfile.mkdtemp()
        self.source_dir = Path(self.test_dir) / "src"
        self.source_dir.mkdir()
        
        # Create test file
        self.create_test_file()
        
    def create_test_file(self):
        # Create a test file with various imports and definitions
        test_content = """
# Test various imports
from utils.helper import helper_function
from models.user import User
import services.auth
import os.path as path
from datetime import datetime, timedelta
from typing import List, Dict

# Test various definitions
def test_function():
    pass

class TestClass:
    def __init__(self):
        self.name = "Test"

variable = 42
print("233")

# Test nested definitions
def outer_function():
    def inner_function():
        pass
    class InnerClass:
        pass
"""
        test_file = self.source_dir / "test_file.py"
        with open(test_file, "w", encoding='utf-8') as f:
            f.write(test_content)
        self.test_file = test_file

    def test_analyze_file_imports(self):
        """Test file import analysis"""
        with open(self.test_file) as f:
            content = f.read()
        
        imports, _ = analyze_file(content, str(self.test_file))
        
        # Verify import count
        self.assertEqual(len(imports), 8)
        
        # Verify import types
        import_modules = {imp.module for imp in imports}
        self.assertIn("utils.helper", import_modules)
        self.assertIn("models.user", import_modules)
        self.assertIn("services.auth", import_modules)
        self.assertIn("os.path", import_modules)
        self.assertIn("datetime", import_modules)
        self.assertIn("typing", import_modules)

    def test_analyze_file_definitions(self):
        """Test file definition analysis"""
        with open(self.test_file) as f:
            content = f.read()
        
        _, defined_names = analyze_file(content, str(self.test_file))
        
        # Verify defined names
        self.assertIn("test_function", defined_names)
        self.assertIn("TestClass", defined_names)
        self.assertIn("variable", defined_names)
        self.assertIn("outer_function", defined_names)
        self.assertIn("inner_function", defined_names)
        self.assertIn("InnerClass", defined_names)

    def test_analyze_file_empty(self):
        """Test empty file analysis"""
        empty_file = self.source_dir / "empty.py"
        with open(empty_file, "w", encoding='utf-8') as f:
            f.write("")
        
        imports, defined_names = analyze_file("", str(empty_file))
        self.assertEqual(len(imports), 0)
        self.assertEqual(len(defined_names), 0)

    def test_analyze_file_syntax_error(self):
        """Test syntax error handling"""
        invalid_content = "def test_function()\n    pass"  # Missing colon
        invalid_file = self.source_dir / "invalid.py"
        with open(invalid_file, "w", encoding='utf-8') as f:
            f.write(invalid_content)
        
        # Syntax error files should return empty results instead of raising exceptions
        imports, defined_names = analyze_file(invalid_content, str(invalid_file))
        self.assertEqual(len(imports), 0)
        self.assertEqual(len(defined_names), 0)

    def tearDown(self):
        # Clean up temporary directory
        import shutil
        shutil.rmtree(self.test_dir)

if __name__ == '__main__':
    unittest.main() 