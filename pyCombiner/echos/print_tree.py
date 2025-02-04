import os

# # 全局字典，保存文件依赖关系（导入关系）
# file_tree = {}


def parse_file(file_path):
    """
    读取 Python 文件，并返回其中所有 import/from 语句的列表
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except Exception as e:
        print(f"读取 {file_path} 时发生错误: {e}")
        return []

    imports = []
    for line in lines:
        line = line.strip()
        if line.startswith("import ") or line.startswith("from "):
            imports.append(line)
    return imports


def find_module_path(import_line, current_dir):
    """
    解析 import 语句，返回对应模块的文件路径（如果存在）

    示例：
      - import module1
      - from subdir import module2
      - from subdir.deeper import module3
    """
    module_name = None
    if import_line.startswith("import "):
        # 例如 "import module1" 或 "import module1, moduleX"
        modules = import_line.split()[1].split(',')
        module_name = modules[0].strip()
    elif import_line.startswith("from "):
        # 例如 "from subdir import module2" 或 "from subdir.deeper import module3"
        parts = import_line.split()
        if len(parts) >= 4:
            from_module = parts[1].strip()
            # parts[3] 可能包含多个逗号分隔的项，只取第一个示例
            imported = parts[3].split(',')[0].strip()
            module_name = from_module + "." + imported
        else:
            return None
    else:
        return None

    # 将模块名转换为文件路径（假设模块文件名以 .py 结尾）
    module_path = module_name.replace(".", os.sep) + ".py"
    possible_path = os.path.join(current_dir, module_path)  # pyCombiner/echos/print_tree.py:55
    if os.path.exists(possible_path):
        return os.path.normpath(possible_path)
    return None


def print_dependency_tree(node, file_tree, prefix="", is_last=True):
    """
    递归打印文件依赖树，使用树形图格式
    """
    connector = "└── " if is_last else "├── "
    print(prefix + connector + os.path.basename(node))
    children = file_tree.get(node, [])
    new_prefix = prefix + ("    " if is_last else "│   ")
    for i, child in enumerate(children):
        is_last_child = (i == len(children) - 1)
        print_dependency_tree(child, file_tree, new_prefix, is_last_child)


def print_directory_tree(root_dir, prefix="", entry_file=None):
    """
    递归打印整体物理目录结构，从 root_dir 开始。
    只有根目录下的文件（非目录）前使用黑色方块“■”标记，
    子目录和其内容使用常规标记。
    """
    entries = os.listdir(root_dir)
    entries = [e for e in entries if not e.startswith('.')]
    entries.sort()
    for i, entry in enumerate(entries):
        full_path = os.path.join(root_dir, entry)
        is_last = (i == len(entries) - 1)
        connector = "└── " if is_last else "├── "
        # 仅在根目录层，且当前文件正是入口文件时使用黑色方块标记
        if prefix == "" and entry_file is not None and os.path.abspath(full_path) == os.path.abspath(entry_file):
            print(prefix + connector + "■ " + entry)
        else:
            print(prefix + connector + entry)
        if os.path.isdir(full_path):
            extension = "    " if is_last else "│   "
            print_directory_tree(full_path, prefix + extension, entry_file)



def build_import_dir_tree(project_root):
    """
    根据通过 import 语句解析到的所有文件（file_tree 中的键和值），
    构造一个目录树（字典形式），便于后续打印出整体的文件目录结构。
    """
    all_files = set(file_tree.keys())
    for children in file_tree.values():
        all_files.update(children)

    tree = {}  # 目录树，形如 { "subdir": { "module2.py": None, "deeper": { ... } }, ... }

    for file_path in all_files:
        # 得到相对于 project_root 的相对路径
        rel_path = os.path.relpath(file_path, project_root)
        parts = rel_path.split(os.sep)
        current = tree
        for part in parts[:-1]:
            current = current.setdefault(part, {})
        # 文件作为叶子，用 None 表示
        current[parts[-1]] = None
    return tree


def print_tree_from_dict(tree, prefix="", entry_file=None):
    """
    递归打印目录树（字典形式）。
    节点为 None 表示叶子（文件），否则为子目录字典。
    """
    entries = list(tree.items())
    entries.sort(key=lambda x: x[0])
    for i, (name, subtree) in enumerate(entries):
        is_last = (i == len(entries) - 1)
        connector = "└── " if is_last else "├── "
        # 如果在根层（prefix为空）并且该名称就是入口文件的文件名，则加上黑色方块
        if prefix == "" and entry_file is not None and name == os.path.basename(entry_file):
            print(prefix + connector + "■ " + name)
        else:
            print(prefix + connector + name)
        if subtree is not None:
            extension = "    " if is_last else "│   "
            print_tree_from_dict(subtree, prefix + extension, entry_file)


if __name__ == "__main__":
    # 项目根目录（当前文件所在目录）
    project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo"
    # 入口文件，例如 main.py
    entry_file = os.path.join(project_root, "main.py")

    # 1. 解析依赖关系，构造 file_tree（导入依赖树）
    file_tree = {}

    print("文件引用结构树：")
    print_dependency_tree(entry_file, prefix="", is_last=True)

    print("\n整体物理文件目录结构：")
    print(f"根目录路径：{project_root}\n")
    # 打印整体目录结构，只有根目录的文件前有黑色方块标记
    print_directory_tree(project_root, entry_file=entry_file)

    # 2. 根据 import 语句推测的整体文件目录结构
    print("\n通过 import 语句推测的文件目录结构：")
    import_dir_tree = build_import_dir_tree(project_root)
    # 如果你希望标记根目录下的文件，这里你可以在最外层处理（示例中直接打印树）
    print_tree_from_dict(import_dir_tree, entry_file=entry_file)
