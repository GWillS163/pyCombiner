import unittest
from pathlib import Path
from pycombiner.combiner.output import MergeReport, print_merge_report

class TestOutput(unittest.TestCase):
    def test_print_merge_report(self):
        # Create a sample MergeReport
        source_dir = Path('pycombiner/tests/examples/deep_demo')
        entry_file = Path('main.py')
        output_file = Path('demo-result.py')
        report = MergeReport(entry_file, source_dir, output_file)

        # Add sample file info
        report.add_file_info(
            Path('main.py'),
            20,
            {'utils.helpers.math_utils', 'utils.helpers.string_utils', 'models.user', 'services.auth', 'services.database'},
            {'os', 'sys'}
        )
        report.add_file_info(
            Path('models/user.py'),
            8,
            set(),
            {'typing.Optional', 'json'}
        )
        report.add_file_info(
            Path('services/auth.py'),
            12,
            {'models.user'},
            {'hashlib'}
        )
        report.add_file_info(
            Path('services/database.py'),
            10,
            set(),
            {'os', 'logging'}
        )
        report.add_file_info(
            Path('utils/helpers/math_utils.py'),
            15,
            set(),
            {'math'}
        )
        report.add_file_info(
            Path('utils/helpers/string_utils.py'),
            11,
            set(),
            {'re'}
        )

        # Set sample dependency graph
        report.set_dependency_graph({
            'main': {'utils.helpers.math_utils', 'utils.helpers.string_utils', 'models.user', 'services.auth', 'services.database'},
            'models.user': set(),
            'services.auth': {'models.user'},
            'services.database': set(),
            'utils.helpers.math_utils': set(),
            'utils.helpers.string_utils': set()
        })

        # Set sample merge order
        report.set_merge_order([
            Path('models/user.py'),
            Path('services/auth.py'),
            Path('services/database.py'),
            Path('utils/helpers/math_utils.py'),
            Path('utils/helpers/string_utils.py'),
            Path('main.py')
        ])

        # Update sample stats
        report.update_stats({
            'total_imports': 10,
            'duplicate_imports': 2,
            'redundant_imports': 1,
            'functions': 12,
            'classes': 2,
            'total_time': 0.74,
            'total_lines': 120
        })

        # Set entry file imports for detailed display
        report.imports_by_file = {
            'main.py': [
                'from utils.helpers.math_utils   import add, subtract',
                'from utils.helpers.string_utils import format_name',
                'from models.user                import User',
                'from services.auth              import login, logout',
                'from services.database          import connect_db'
            ]
        }

        # Get the formatted report
        formatted_report = report.format_report()

        # Print the actual report
        print("\nActual Output:")
        print("=" * 100)
        print(formatted_report)
        print("=" * 100)

        # Verify the report contains expected sections
        self.assertIn("Import Handling Summary", formatted_report)
        self.assertIn("Import Handling Details", formatted_report)
        self.assertIn("Summary", formatted_report)

        # Verify directory tree structure
        self.assertIn("models/", formatted_report)
        self.assertIn("services/", formatted_report)
        self.assertIn("utils/", formatted_report)
        self.assertIn("helpers/", formatted_report)


        # Verify file information with order numbers
        self.assertIn("[1] user.py", formatted_report)
        self.assertIn("[2] auth.py", formatted_report)
        self.assertIn("[3] database.py", formatted_report)
        self.assertIn("[4] math_utils.py", formatted_report)
        self.assertIn("[5] string_utils.py", formatted_report)
        self.assertIn("main.py 🚩 entry file", formatted_report)

        # Verify import information
        self.assertIn("Handled Imports", formatted_report)
        self.assertIn("Unhandled Imports", formatted_report)
        self.assertIn("from utils.helpers.math_utils", formatted_report)
        self.assertIn("from models.user", formatted_report)

        # Verify unhandled imports (check for individual imports)
        self.assertIn("json", formatted_report)
        self.assertIn("typing.Optional", formatted_report)
        self.assertIn("hashlib", formatted_report)
        self.assertIn("logging", formatted_report)
        self.assertIn("os", formatted_report)
        self.assertIn("math", formatted_report)
        self.assertIn("re", formatted_report)

        # Verify statistics
        self.assertIn("Total import statements analyzed", formatted_report)
        self.assertIn("Dependency graph built", formatted_report)
        self.assertIn("Duplicate local imports skipped", formatted_report)
        self.assertIn("Lines written to merged output", formatted_report)
        self.assertIn("Redundant imports removed", formatted_report)
        self.assertIn("Total time elapsed", formatted_report)

if __name__ == '__main__':
    unittest.main() 