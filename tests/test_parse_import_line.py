import unittest
from pyCombiner.parse_import_line import parse_import_line


class MyTestCase(unittest.TestCase):
    def test_1(self):
        self.assertEqual({'modules': [('module1', None)], 'type': 'import'},
                         parse_import_line("import module1"))

    def test_2(self):
        self.assertEqual({'modules': [('module1', None), ('module2', 'm2')], 'type': 'import'},
                         parse_import_line("import module1, module2 as m2"))

    def test_3(self):
        self.assertEqual({'module': 'subdir', 'names': [('module2', None)], 'type': 'from'},
                         parse_import_line("from subdir import module2"))

    def test_4(self):
        self.assertEqual({'module': 'subdir.deeper',
                          'names': [('module3', 'mod3'), ('module4', None)],
                          'type': 'from'},
                         parse_import_line("from subdir.deeper import module3 as mod3, module4"))

    def test_5(self):
        self.assertEqual({'module': 'module1', 'names': [('*', None)], 'type': 'from'},
                         parse_import_line("from module1 import *"))

    # def test_6(self):
    #     self.assertEqual({'modules': [('module1', None)], 'type': 'import'},
    #                      parse_import_line("import module1"))
    #
    # def test_7(self):
    #     self.assertEqual({'modules': [('module1', None)], 'type': 'import'},
    #                      parse_import_line("import module1"))


if __name__ == '__main__':
    unittest.main()
