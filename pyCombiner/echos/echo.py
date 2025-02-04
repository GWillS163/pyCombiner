# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 23:15

def titleBanner(prefix):
    return f"\t------{prefix}------"


def decorate(func):
    def wrapper(prefix, data):
        print(titleBanner(prefix))
        for line in data:
            func(data, line)
        print(titleBanner(prefix))
    return wrapper


@decorate
def showImport(importData, line):
    print(importData[line])


@decorate
def showFrom(fromData, line):
    for importLine in fromData[line].values():
        print(importLine)


@decorate
def showCodes(codes, line):
    print(line)


def showRests(path, imports, froms, codes):
    print("\n\n", path, "-  " * 10)
    showImport("imports", imports)
    showFrom("froms", froms)
    showCodes("codes", codes)


def showMainRests(path, imports, codes):
    print("\n\n", path, "-  " * 10)

    print(titleBanner("restImports"))
    for line in imports:
        print(line)
    print(titleBanner("restImports"))
    # print(titleBanner("code"))
    # for line in codes:
    #     print(line, end="")
    print(f"number of import line: {len(imports)}")
    print(f"number of code lines: {len(codes)}")
    # print(titleBanner("code"))

def print_tree(node, file_tree, prefix=""):
    """
    递归打印文件引用树
    :param node: 当前节点（文件路径）
    :param file_tree: 文件引用字典
    :param prefix: 树形前缀
    """
    print(f"{prefix}└── {node}")
    children = file_tree.get(node, [])
    for i, child in enumerate(children):
        new_prefix = prefix + ("    " if i == len(children) - 1 else "│   ")
        print_tree(child, file_tree, new_prefix)

def print_imported_file(module_file_name, file_codes, all_lines, file_path, target_file_abs_path, file_tree, settings=None):
    if settings.debug_log:
        all_lines.append(f"# ========================  BEGIN inlined module: {module_file_name}\n")
    all_lines.append(file_codes)
    if settings.debug_log:
        all_lines.append(f"# ========================  END inlined module: {module_file_name}\n")
    file_tree[file_path].append(target_file_abs_path)
    return all_lines, file_tree


if __name__ == '__main__':
    importModules = {'update': 'import update as up'}
    fromModules = {'bin': {'*': 'from bin import *'}}
    otherLines = ['', '# 4', 'up.showBanner()', 'main()']
    showRests("Mock Import data", importModules, fromModules, otherLines)
