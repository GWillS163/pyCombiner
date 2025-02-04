# Github: GWillS163
# User: 駿清清 
# Date: 09/11/2022 
# Time: 15:43
import argparse
import sys
import os.path
from pyCombiner.echos.colorfulPrint import *
from pyCombiner.main import main
from pyCombiner import description

import locale

# 定义多语言支持
LANGUAGES = {
    "zh": {
        "start_file_error": "错误: 指定的起始文件路径不存在: {path}",
        "project_dir_error": "错误: 项目文件目录无效: {path}",
        "save_dir_error": "错误: 保存文件夹无效: {path}",
        "runtime_error": "运行错误: {error}",
        "result_info": "起始文件路径: {start_file}\n项目目录: {project_dir}\n保存目录: {save_dir}",
        "start_file_intro": "起始文件路径，支持相对或绝对路径",
        "project_dir_intro": "项目文件目录，如果为空则使用起始文件所在目录",
        "save_dir_intro": "保存文件夹，如果为空则使用起始文件所在目录"
    },
    "en": {
        "start_file_error": "Error: The specified start file path does not exist: {path}",
        "project_dir_error": "Error: The specified project directory is invalid: {path}",
        "save_dir_error": "Error: The specified save directory is invalid: {path}",
        "runtime_error": "Runtime Error: {error}",
        "result_info": "Start file path: {start_file}\nProject directory: {project_dir}\nSave directory: {save_dir}",
        "start_file_intro": "Start file path, supports relative or absolute path",
        "project_dir_intro": "Project directory; if not provided, the directory of the start file will be used",
        "save_dir_intro": "Save directory; if not provided, the directory of the start file will be used"
    },
    "jp": {
        "start_file_error": "エラー: 指定された開始ファイルパスが存在しません: {path}",
        "project_dir_error": "エラー: 指定されたプロジェクトディレクトリが無効です: {path}",
        "save_dir_error": "エラー: 指定された保存ディレクトリが無効です: {path}",
        "runtime_error": "実行エラー: {error}",
        "result_info": "開始ファイルパス: {start_file}\nプロジェクトディレクトリ: {project_dir}\n保存ディレクトリ: {save_dir}",
        "start_file_intro": "開始ファイルのパス（相対パスまたは絶対パスが使用可能）",
        "project_dir_intro": "プロジェクトディレクトリ（指定がない場合は開始ファイルのディレクトリが使用されます）",
        "save_dir_intro": "保存ディレクトリ（指定がない場合は開始ファイルのディレクトリが使用されます）"
    }
}


def detect_language():
    system_lang = locale.getdefaultlocale()[0]
    if system_lang.startswith("zh"):
        return "zh"
    elif system_lang.startswith("ja"):
        return "jp"
    else:
        return "en"

LANG = detect_language()  # 自动设置语言


def parse_arguments():
    """
    解析命令行参数。
    """
    parser = argparse.ArgumentParser(
        description="PyCombiner Introduce"
    )

    # 必选参数: 起始文件路径
    parser.add_argument('start_file', type=str, help=LANGUAGES[LANG]["start_file_intro"])

    # 可选参数: 项目文件目录
    parser.add_argument('--project-dir', '-p', type=str, default=None,
                        help=LANGUAGES[LANG]["project_dir_intro"])

    # 可选参数: 保存文件夹
    parser.add_argument('--save-dir', '-s', type=str, default=None,
                        help=LANGUAGES[LANG]["save_dir_intro"])

    return parser.parse_args(), parser


def validate_paths(args, parser):
    """
    校验路径参数。
    """
    # 检查起始文件路径是否存在
    if not os.path.isfile(args.start_file):
        sys.exit(LANGUAGES[LANG]["start_file_error"].format(path=args.start_file))

    # 获取起始文件所在目录
    start_file_dir = os.path.dirname(os.path.abspath(args.start_file))

    # 设置项目目录和保存目录
    project_dir = args.project_dir or start_file_dir
    save_dir = args.save_dir or start_file_dir

    # 检查路径有效性
    if not os.path.isdir(project_dir):
        sys.exit(LANGUAGES[LANG]["project_dir_error"].format(path=project_dir))

    if not os.path.isdir(save_dir):
        sys.exit(LANGUAGES[LANG]["save_dir_error"].format(path=save_dir))

    return project_dir, save_dir


def run():
    args, parser = parse_arguments()
    project_dir, save_dir = validate_paths(args, parser)
    # 打印参数结果
    print(LANGUAGES[LANG]["result_info"].format(
        start_file=args.start_file, project_dir=project_dir, save_dir=save_dir
    ))

    project_root = os.path.abspath(project_dir)
    entrance_file = os.path.abspath(args.start_file)


    outputPath = main(project_root, entrance_file, save_dir)
    print("\nThe output file is: ", end="")
    colorPrint(outputPath, color="green")


if __name__ == '__main__':
    run()
