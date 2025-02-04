import re

def parse_import_line(line):
    """
    解析一行代码，检测是否为 import 语句。

    返回：
      - 如果匹配到 import 语句，则返回字典。例如：
            {'type': 'import', 'modules': [('module1', None), ('module2', None)]}
        或
            {'type': 'from', 'module': 'subdir.module', 'names': [('name1', None), ('name2', 'alias2')]}
      - 如果不是 import 语句，返回 None。
    """
    line = line.strip()
    # 处理多行 import 时可能已去除括号，默认处理单行情况
    pattern_main = r"^if __name__ == [\'\"]__main__[\'\"]:"

    m = re.match(r"^import\s+(.+)$", line)
    if m:
        modules_part = m.group(1)
        modules = []
        for part in modules_part.split(","):
            part = part.strip()
            if " as " in part:
                mod, alias = part.split(" as ")
                modules.append((mod.strip(), alias.strip()))
            else:
                modules.append((part, None))
        return {'type': 'import', 'modules': modules}

    m = re.match(r"^from\s+([\w\.]+)\s+import\s+(.+)$", line)
    if m:
        module_base = m.group(1).strip()
        names_part = m.group(2).replace("(", "").replace(")", "")
        names = []
        for part in names_part.split(","):
            part = part.strip()
            if not part:
                continue
            if " as " in part:
                name, alias = part.split(" as ")
                names.append((name.strip(), alias.strip()))
            else:
                names.append((part, None))
        return {'type': 'from', 'module': module_base, 'names': names}

    # m = re.match(pattern_main, line)
    # if m:
    #     return {'type': '__main__'}

    return None

if __name__ == '__main__':
    test_lines = [
        "import module1",
        "import module1, module2 as m2",
        "from subdir import module2",
        "from subdir.deeper import module3 as mod3, module4",
        "from module import *"
    ]
    for line in test_lines:
        result = parse_import_line(line)
        print(f"Line: {line}\nParsed: {result}\n")
