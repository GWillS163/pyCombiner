import unittest
from pathlib import Path
import tempfile
import shutil
from pycombiner.combiner.ast_parser import analyze_file, _resolve_module_to_filepath, get_module_name, build_dependency_graph

class TestASTParser(unittest.TestCase):
    def setUp(self):
        # Create temporary directory
        self.test_dir = Path(tempfile.mkdtemp())
        self.source_dir = self.test_dir / "test_project"
        
        # Create test files
        self.create_test_files()
        
    def create_test_files(self):
        """Create test files with various import patterns"""
        # Create main.py
        main_content = '''"""
Main entry point
"""
from utils.helpers.math_utils import add, subtract
from utils.helpers.string_utils import format_name
from models.user import User
from services.auth import login, logout
from services.database import connect_db

def main():
    db = connect_db()
    user = User("John Doe")
    formatted_name = format_name(user.name)
    result1 = add(10, 5)
    result2 = subtract(10, 5)
    login(user)
    logout(user)
    print(f"Results: {result1}, {result2}")
    print(f"Formatted name: {formatted_name}")

if __name__ == "__main__":
    main()
'''
        # Create utils/helpers/math_utils.py
        math_utils_content = '''"""
Math utility functions
"""
def add(a: int, b: int) -> int:
    """Add two numbers"""
    return a + b

def subtract(a: int, b: int) -> int:
    """Subtract b from a"""
    return a - b
'''
        # Create utils/helpers/string_utils.py
        string_utils_content = '''"""
String utility functions
"""
def format_name(name: str) -> str:
    """Format a name with proper capitalization"""
    return name.title()
'''
        # Create models/user.py
        user_content = '''"""
User model
"""
class User:
    def __init__(self, name: str):
        self.name = name
        self.is_logged_in = False
'''
        # Create services/auth.py
        auth_content = '''"""
Authentication service
"""
from models.user import User

def login(user: User) -> None:
    """Log in a user"""
    user.is_logged_in = True

def logout(user: User) -> None:
    """Log out a user"""
    user.is_logged_in = False
'''
        # Create services/database.py
        database_content = '''"""
Database service
"""
class Database:
    def __init__(self):
        self.connected = False

def connect_db() -> Database:
    """Create and connect to a database"""
    return Database()
'''
        # Create the directory structure and write files
        (self.source_dir / "utils" / "helpers").mkdir(parents=True)
        (self.source_dir / "models").mkdir()
        (self.source_dir / "services").mkdir()
        
        (self.source_dir / "main.py").write_text(main_content)
        (self.source_dir / "utils" / "helpers" / "math_utils.py").write_text(math_utils_content)
        (self.source_dir / "utils" / "helpers" / "string_utils.py").write_text(string_utils_content)
        (self.source_dir / "models" / "user.py").write_text(user_content)
        (self.source_dir / "services" / "auth.py").write_text(auth_content)
        (self.source_dir / "services" / "database.py").write_text(database_content)
        
    def test_analyze_file(self):
        """Test analyzing a file for imports and definitions"""
        file_path = self.source_dir / "main.py"
        imports, defined_names = analyze_file(file_path.read_text(), str(file_path))
        
        # Check imports
        self.assertEqual(len(imports), 7)
        import_modules = {imp.module for imp in imports}
        self.assertIn("utils.helpers.math_utils", import_modules)
        self.assertIn("utils.helpers.string_utils", import_modules)
        self.assertIn("models.user", import_modules)
        self.assertIn("services.auth", import_modules)
        self.assertIn("services.database", import_modules)
        
        # Check defined names
        # self.assertIn("main", defined_names)
        
    def test_resolve_module_to_filepath(self):
        """Test resolving module names to file paths"""
        # Test with a simple module
        filepath = _resolve_module_to_filepath("utils.helpers.math_utils", self.source_dir)
        self.assertEqual(filepath, self.source_dir / "utils" / "helpers" / "math_utils.py")
        
        # Test with a non-existent module
        filepath = _resolve_module_to_filepath("nonexistent.module", self.source_dir)
        self.assertIsNone(filepath)
        
    def test_get_module_name(self):
        """Test getting module name from file path"""
        file_path = self.source_dir / "utils" / "helpers" / "math_utils.py"
        module_name = get_module_name(file_path, self.source_dir)
        self.assertEqual(module_name, "utils.helpers.math_utils")
        
    def test_build_dependency_graph(self):
        """Test building dependency graph"""
        # Analyze all files
        imports_by_file = {}
        for file in self.source_dir.rglob("*.py"):
            with open(file) as f:
                content = f.read()
            imports, _ = analyze_file(content, str(file))
            imports_by_file[str(file)] = imports
            
        # Build dependency graph
        graph = build_dependency_graph(imports_by_file, str(self.source_dir))
        
        # Check dependencies
        main_file = str(self.source_dir / "main.py")
        self.assertIn(main_file, graph)
        self.assertGreater(len(graph[main_file]), 0)
        
        # Check that auth.py depends on user.py
        auth_file = str(self.source_dir / "services" / "auth.py")
        user_file = str(self.source_dir / "models" / "user.py")
        self.assertIn(user_file, graph[auth_file])
        
    def tearDown(self):
        # Clean up temporary directory
        shutil.rmtree(self.test_dir)

if __name__ == '__main__':
    unittest.main() 