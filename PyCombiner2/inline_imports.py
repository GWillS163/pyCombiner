import os
from locate_module import locate_module
from PyCombiner2.parse_import_line import parse_import_line


def remove_main_guard(code):
    """
    将代码中 if __name__ == '__main__': 块及其缩进部分全部注释掉。
    简单实现：对每行代码，如果匹配到 if __name__ == '__main__': 则在前面添加 '#'，
    后续缩进的行也加 '#'，直到遇到非空且缩进相同或更少的行。
    此处采用简单逐行处理的策略，适合大部分情况。
    """
    lines = code.splitlines()
    new_lines = []
    in_main = False
    main_indent = None
    for line in lines:
        stripped = line.lstrip()
        indent = len(line) - len(stripped)
        if not in_main and stripped.startswith("if __name__") and "__main__" in stripped:
            in_main = True
            main_indent = indent
            new_lines.append("# " + line)
            continue
        if in_main:
            # 如果本行缩进大于 main_indent，认为属于 if 块
            if line.strip() == "":
                new_lines.append(line)
            elif indent > main_indent:
                new_lines.append("# " + line)
            else:
                in_main = False
                new_lines.append(line)
        else:
            new_lines.append(line)
    return "\n".join(new_lines) + "\n"


def inline_imports(file_path, project_root, current_dir, file_tree, visited=None, entry_file=None, settings=None):
    """
    递归内联文件中所有 import 语句。

    参数：
      - file_path: 当前文件的绝对路径
      - project_root: 项目的根目录
      - visited: 已内联的文件集合（防止重复和循环依赖）
      - entry_file: 主入口文件路径（该文件保留 if __name__ == '__main__': 块，其它文件内的此块全部注释掉）

    返回：一个字符串，代表内联后的代码内容。
    """
    if visited is None:
        visited = set()

    # 初始化当前文件在字典中的依赖列表
    if file_path not in file_tree:
        file_tree[file_path] = []

    if file_path in visited:
        if settings.debug_log:
            return f"# [Already inlined: {file_path}]\n"
        return ""
    visited.add(file_path)

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except Exception as e:
        return f"# 错误：无法读取文件 {file_path}，错误信息: {e}\n"

    new_lines = []
    for i, line in enumerate(lines):
        parsed = parse_import_line(line)
        if not parsed:
            new_lines.append(line)
            continue

        if parsed['type'] == 'import':
            for mod, alias in parsed['modules']:
                target_file_abs_paths = locate_module(mod, file_path, project_root)
                if len(target_file_abs_paths) == 0:
                    new_lines.append(line)
                    continue
                for target_file_abs_path in target_file_abs_paths:
                    if not target_file_abs_path or not os.path.exists(target_file_abs_path):
                        new_lines.append(line)
                        continue
                    inlined_code = inline_imports(target_file_abs_path, project_root, current_dir, file_tree, visited=visited, entry_file=entry_file, settings=settings)
                    new_lines, file_tree = print_imported_file(mod, inlined_code, new_lines, file_path, target_file_abs_path, file_tree, settings=settings)

        elif parsed['type'] == 'from':
            target_mod = parsed['module']
            target_names = parsed['names']

            target_file_abs_paths = locate_module(target_mod, file_path, project_root, target_names)
            if len(target_file_abs_paths) == 0:
                new_lines.append(line)
                continue

            for target_file_abs_path in target_file_abs_paths:
                if not target_file_abs_path or not os.path.exists(target_file_abs_path):
                    new_lines.append(line)
                    continue
                inlined_code = inline_imports(target_file_abs_path, project_root, current_dir, file_tree, visited=visited, entry_file=entry_file, settings=settings)
                new_lines, file_tree = print_imported_file(target_mod, inlined_code, new_lines, file_path, target_file_abs_path, file_tree, settings=settings)


    code = "".join(new_lines)
    # 对非入口文件，注释掉 if __name__ == '__main__': 块
    if entry_file is None or os.path.abspath(file_path) != os.path.abspath(entry_file):
        code = remove_main_guard(code)
    return code

def print_imported_file(module_file_name, file_codes, all_lines, file_path, target_file_abs_path, file_tree, settings=None):
    if settings.debug_log:
        all_lines.append(f"# ========================  BEGIN inlined module: {module_file_name}\n")
    all_lines.append(file_codes)
    if settings.debug_log:
        all_lines.append(f"# ========================  END inlined module: {module_file_name}\n")
    file_tree[file_path].append(target_file_abs_path)
    return all_lines, file_tree

if __name__ == '__main__':
    # 测试 inline_imports 模块，注意：此处仅用于调试
    # os.chdir(r"../tests/examples/deep_demo")
    # 全局字典，保存文件依赖关系（导入关系）
    file_tree = {}
    project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo"
    entry_file = os.path.join(project_root, "main.py")
    combined_code = inline_imports(entry_file, project_root, project_root, file_tree, entry_file=entry_file)
    print("内联后的代码内容：\n")
    print(combined_code)
