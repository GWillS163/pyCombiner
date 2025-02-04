import os
import unittest

from pyCombiner.locate_module import locate_module, find_python_packages


class MyTestCase(unittest.TestCase):
    def test_something(self):
        # 测试用代码：仅供调试，内联工具中不会执行
        # os.chdir(r"../tests/examples/deep_demo/main.py")
        project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo"
        current_file = os.path.join(project_root, "main.py")

        self.assertEqual(['C:\\Users\\PFS\\Documents\\Codes\\pyCombiner\\tests\\examples\\deep_demo\\module1.py'],
                         locate_module("module1", current_file, project_root))  # add assertion here

    def test_something2(self):
        project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo"
        current_file = os.path.join(project_root, "module1.py")

        self.assertEqual([
            'C:\\Users\\PFS\\Documents\\Codes\\pyCombiner\\tests\\examples\\deep_demo\\subdir\\deeper\\module3.py',
            'C:\\Users\\PFS\\Documents\\Codes\\pyCombiner\\tests\\examples\\deep_demo\\subdir\\deeper\\deepest\\module4.py'
        ],
                         locate_module("subdir.deeper", current_file, project_root))

    def test_something3(self):
        project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo"
        current_file = os.path.join(project_root, "module1.py")

        self.assertEqual(['C:\\Users\\PFS\\Documents\\Codes\\pyCombiner\\tests\\examples\\deep_demo\\subdir\\module2.py'],
                         locate_module("subdir.module2", current_file, project_root))

    def test_findup(self):
        project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\project2"
        current_file = os.path.join(project_root, "batches\inspect_all_table.py")

        self.assertEqual([
                                r'C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\project2\tests\util\operation.py'
                            ],
                         locate_module("tests.util.operation", current_file, project_root))


    def test_deep_finder(self):
        self.assertEqual(
            ['deepest', 'module3'],
            find_python_packages(
                'C:\\Users\\PFS\\Documents\\Codes\\pyCombiner\\tests\\examples\\deep_demo\\subdir\\deeper'))

    def test_import_py_file(self):
        self.assertEqual(True,
            find_python_packages( False)
                         )
    # import module
    # import module as mod
    # import module1, module2
    # import module1 as m1, module2 as m2
    # from module import name1 as n1, name2 as n2
    # from module import *
    # from module import (
    #     name1,
    #     name2,
    #     name3,
    # )
    # from . import submodule
    # from ..subpackage import submodule
    # # 动态导入
    # TODO: as xx 这种语法解析不了
    # mod = __import__("module")

if __name__ == '__main__':
    unittest.main()
