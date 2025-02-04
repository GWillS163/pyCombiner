import os
import unittest

from pyCombiner.echos.print_tree import parse_file
from pyCombiner.main import main

output_dir = r"C:\Users\PFS\Documents\Codes\pyCombiner\output_temp"

class MyTestCase(unittest.TestCase):

    def test_deep_demo(self):
        project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo"
        entry_file = os.path.join(project_root, "main.py")  # 修改为你的入口文件路径
        self.assertEqual(None, main(project_root, entry_file, output_dir))

    def test_deep_complicate(self):
        project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_complicate"
        entry_file = os.path.join(project_root, "runMeCom.py")  # 修改为你的入口文件路径
        self.assertEqual(None, main(project_root, entry_file, output_dir))

    def test_simple(self):
        project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\demo_simple"
        entry_file = os.path.join(project_root, "runMeSim.py")
        self.assertEqual(None, main(project_root, entry_file, output_dir))

    def test_project2(self):
        project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\project2"
        entry_file = os.path.join(project_root, "batches\inspect_all_table.py")

        output_reference = r"tests/examples_refer_result/project2/combined.py"
        save_path = main(project_root, entry_file, output_dir)

        self.assertEqual(True, True)

    def test_realProject1(self):
        project_root = r"tests\examples\realProject1"
        entry_file = os.path.join(project_root, "realTest.py")

        output_reference = r"tests\examples_refer_result\realProject1\realTest20250203_143439.py"
        save_path = main(project_root, entry_file, output_dir)

        self.assertEqual(True, True)


if __name__ == '__main__':
    unittest.main()

