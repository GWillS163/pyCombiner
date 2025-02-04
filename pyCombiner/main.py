import os
import time

from pyCombiner.inline_imports import inline_imports
from pyCombiner.echos.print_tree import print_dependency_tree

class settings:
    debug_log = False

_output_dir = "CombineOutput"

def get_save_path(starting_file, output_dir=None):
    """
    获取保存路径
    保存パスを取得
    Get the save path
    """
    currentTime = time.strftime("%Y%m%d_%H%M%S_combined", time.localtime())
    filename = os.path.basename(starting_file).replace(".py", f"{currentTime}.py")

    # save_folder = os.path.join(os.path.dirname(starting_file), _output_dir)
    # os.makedirs(save_folder, exist_ok=True)
    output_dir2 = output_dir if output_dir else os.path.dirname(starting_file)
    return os.path.join(output_dir2, filename)


def main(project_root, entry_file, output_dir):

    # 全局字典，保存文件依赖关系（导入关系）
    file_tree = {}
    # Main
    combined_code = inline_imports(entry_file, project_root, project_root, file_tree, settings=settings)

    save_path = get_save_path(entry_file, output_dir)
    with open(save_path, 'w', encoding='utf-8') as out_f:
        out_f.write(combined_code)

    print("File Importing Tree：")
    print_dependency_tree(entry_file, file_tree, prefix="", is_last=True)
    return save_path
