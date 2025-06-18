import unittest
from pathlib import Path
import tempfile
import shutil
from combiner.ast_parser import analyze_file, build_dependency_graph
from combiner.merger import merge_files

class TestDeepDemo(unittest.TestCase):
    def setUp(self):
        # Create temporary directory
        self.test_dir = Path(tempfile.mkdtemp())
        self.source_dir = self.test_dir / "deep_demo"
        
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
        output_file = self.test_dir / "output.py"
        merge_files(python_files, dependency_graph, self.source_dir, output_file, entry_file)
        
        # Verify output file exists
        self.assertTrue(output_file.exists())
        
        # Read and verify output content
        with open(output_file) as f:
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
        self.assertIn('"""\nMath utility functions\n"""', content)
        self.assertIn('"""\nString utility functions\n"""', content)
        self.assertIn('"""\nUser model\n"""', content)
        self.assertIn('"""\nDatabase service\n"""', content)
        self.assertIn('"""\nAuthentication service\n"""', content)
        
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
    