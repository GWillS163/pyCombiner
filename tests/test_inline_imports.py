import os
import unittest

from pyCombiner.inline_imports import inline_imports


class MyTestCase(unittest.TestCase):
    def test_something(self):

        # 全局字典，保存文件依赖关系（导入关系）
        file_tree = {}
        project_dir = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo"
        entry_file = os.path.join(project_dir, "main.py")

        combined_code = inline_imports(entry_file, project_dir, project_dir, file_tree, entry_file=entry_file)
        print("内联后的代码内容：\n")
        print(combined_code)

        self.assertEqual(True, False)  # add assertion here


if __name__ == '__main__':
    unittest.main()
