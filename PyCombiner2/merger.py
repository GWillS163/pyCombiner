import os

from PyCombiner2.inline_imports import inline_imports
from pyCombiner.echos.print_tree import print_dependency_tree

class settings:
    debug_log = True


def main(project_root, entry_file):
    print(f"项目根目录：{project_root}")
    print(f"入口文件：{entry_file}")

    # 全局字典，保存文件依赖关系（导入关系）
    file_tree = {}

    # Main
    combined_code = inline_imports(entry_file, project_root, project_root, file_tree, settings=settings)

    output_file = os.path.join(project_root, "combined.py")
    with open(output_file, 'w', encoding='utf-8') as out_f:
        out_f.write(combined_code)

    print(f"合并后的文件已生成：{output_file}")
    print("文件引用结构树：")
    print_dependency_tree(entry_file, file_tree, prefix="", is_last=True)

if __name__ == "__main__":
    # project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\project2"
    # entry_file = os.path.join(project_root, "batches\inspect_all_table.py")  # 修改为你的入口文件路径

    project_root = r"C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo"
    entry_file = os.path.join(project_root, "main.py")  # 修改为你的入口文件路径

    main(project_root, entry_file)