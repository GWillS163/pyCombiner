"""
Command-line interface for pycombiner
"""
import argparse
from pathlib import Path
import sys

# Add the project root directory to Python path
project_root = str(Path(__file__).parent.parent)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from pycombiner.combiner.combiner import PyCombiner

def run():
    parser = argparse.ArgumentParser(description='Combine Python files into a single file')
    parser.add_argument('source_path', type=str, help='Source directory or entry point Python file')
    parser.add_argument('output_file', type=str, help='Output file path')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    parser.add_argument('--show-details', action='store_true', help='Show detailed import information')

    args = parser.parse_args()

    source_path = Path(args.source_path).resolve()
    output_file = Path(args.output_file).resolve()

    if not source_path.exists():
        print(f"Error: '{source_path}' does not exist")
        return

    # If source path is a file, use its directory as source_dir
    if source_path.is_file():
        source_dir = source_path.parent
        entry_file = source_path
    else:
        source_dir = source_path
        entry_file = source_dir / 'main.py'
        if not entry_file.exists():
            print(f"Error: No main.py found in '{source_dir}'")
            return

    # Use new implementation with debug and detail options
    combiner = PyCombiner(entry_file, source_dir, output_file, args.debug, args.show_details)
    combiner.combine()

if __name__ == '__main__':
    run() 