
import re
from typing import Any


def locate_module(module_name, current_file, project_root, target_names=None, import_list=None) -> list[Any] | Any:
    """
    根据模块名、当前文件路径和项目根目录定位目标模块文件。 最小单位是文件，而不是函数。

    参数：
      - module_name: 字符串，例如 "module1"、"subdir.module2" 或相对导入 ".getScoreData"
      - current_file: 当前处理文件的绝对路径（用于相对导入时参考）
      - project_root: 项目的根目录绝对路径

    返回：
      - 模块文件的绝对路径（如果找到），否则返回 None
    """
    possibility = []
    if not import_list:
        import_list = []

    # 判断是否为相对导入：开头为点号
    # 当前文件所在目录
    current_dir = os.path.dirname(current_file)
    if module_name.startswith('.'):
        # 计算相对层级：点的个数减1表示上移层级数
        m = re.match(r'^(\.+)(.*)', module_name)
        if not m:
            return import_list
        dots, remainder = m.groups()
        level = len(dots) - 1
        # 上移层级
        for _ in range(level):
            current_dir = os.path.dirname(current_dir)
        rel_path = remainder.replace(".", os.sep)
        possibility.append(os.path.join(current_dir, rel_path))
    else:
        # 绝对导入：将模块名转换为相对路径（以 .py 结尾）
        rel_path = module_name.replace(".", os.sep)
        candidate_folder = os.path.join(current_dir, rel_path)
        possibility.append(candidate_folder)

    # 如果没有找到文件，尝试查找包目录下的 __init__.py
    if module_name.startswith('.'):
        possibility.append(os.path.dirname(module_name))
    else:
        possibility.append(os.path.join(project_root, module_name.replace(".", os.sep)))

    for candidate_folder in set(possibility):
        if_file = candidate_folder  + ".py"
        if_pack = os.path.join(candidate_folder, "__init__.py")
        if os.path.exists(if_file):
            import_list.append(if_file)
            return import_list

        if os.path.exists(if_pack):
            if target_names is None:  # import subdir.deeper.deepest
                target_names = find_python_packages(candidate_folder)
                for target_name in target_names:
                    import_list += locate_module(target_name, if_pack, project_root, import_list=import_list)
                return list(set(import_list))

            for target_name_alias in target_names:
                assert type(target_name_alias) == tuple
                if target_name_alias[0]: # Module has Alias
                    # import_list.append() TODO:  target_name_alias[1] 如何处理？
                    import_list += locate_module(target_name_alias[0], if_pack, project_root, import_list=import_list)
            return list(set(import_list))

    return list(set(import_list))

def locate_module_file():
    pass

import os

def find_python_packages(directory):
    """
    递归查找目录中的所有 .py 文件，如果是包则继续查找。

    :param directory: 要扫描的根目录
    :return: 所有 .py 文件的绝对路径列表
    """
    python_files = []
    for file in os.listdir(directory):
        # 查找 .py 文件
        if "__init__.py" in file:
            continue
        python_files.append(file.strip(".py"))
    return python_files

