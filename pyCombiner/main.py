# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 20:20
import os.path
import time
from os import mkdir

from pyCombiner.echos.echo import showMainRests
from pyCombiner.lineParse import *
from pyCombiner.osFileOpr import *
from pyCombiner.parseCore import *
from pyCombiner.depend.findOri import *

_output_dir = "CombineOutput"

def recursiveParser(entranceFilePath, funcList=None, visited=None) -> list:
    """
    单个文件的递归解析
    単一ファイルの再帰的解析
    Recursive parsing of a single file

    :param visited: 已访问的文件列表
                   訪問済みのファイルリスト
                   List of visited files
    :param funcList: 需要导入的函数
                    インポートする必要のある関数
                    Functions to be imported
    :param entranceFilePath: 入口文件路径
                            エントリファイルのパス
                            Path to the entrance file
    :return: 返回解析结果
             解析結果を返す
             Returns the parsing result
    """

    def enterDeeper(sumImportModules, sumOtherLines, deeperFilePath):
        """
        进入下一层
        次の層に入る
        Enter deeper
        """
        subImportModule, subOtherLines = recursiveParser(deeperFilePath, visited=visited)
        sumImportModules += subImportModule
        sumOtherLines += subOtherLines
        return sumImportModules, sumOtherLines

    if visited is None:
        visited = []
    lines = parseFile(entranceFilePath)
    workFolder, entranceFile = getWorkFolderWithFile(entranceFilePath)
    if entranceFilePath in visited:
        return [[], []]
    importModules, fromModules, otherLines = parseSinglePy(lines)
    visited.append(entranceFilePath)
    restImportModules = []
    restOtherLines = []

    # 处理import语句的模块
    # import文のモジュールを処理
    # Process import statements
    for module in importModules:
        deeperFilePath = findFromOrigin(module, workFolder)
        if not deeperFilePath:
            restImportModules.append(importModules[module])
        else:
            restImportModules, restOtherLines = enterDeeper(restImportModules, restOtherLines, deeperFilePath)

    # 处理from import语句的模块
    # from import文のモジュールを処理
    # Process from import statements
    for module in fromModules:
        deeperFilePath = findFromOrigin(module, workFolder)
        if not deeperFilePath:
            restImportModules += fromModules[module].values()
        else:
            restImportModules, childOtherLines = enterDeeper(restImportModules, restOtherLines, deeperFilePath)
            if funcList:
                importedPart = getPartImportedFuncs(childOtherLines, fromModules[module].keys())
                restOtherLines = importedPart + restOtherLines

    return [restImportModules, restOtherLines + otherLines]


def combineImportLines(restImportModules):
    """
    将import语句合并
    import文を結合
    Combine import statements

    :param restImportModules: 未合并的import语句
                              未結合のimport文
                              Uncombined import statements
    :return: 合并后的import语句
             結合後のimport文
             Combined import statements
    """
    combineRestImportLines = list(set(restImportModules))
    combineRestImportLines.sort(key=lambda x: len(x))
    return combineRestImportLines


def getSavePath(entranceFile):
    """
    获取保存路径
    保存パスを取得
    Get the save path

    :param outFolder: 输出文件夹
                      出力フォルダ
                      Output folder
    :param entranceFile: 入口文件
                        エントリファイル
                        Entrance file
    :return: 保存路径
             保存パス
             Save path
    """
    currentTime = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    filename = os.path.basename(entranceFile).replace(".py", f"{currentTime}.py")

    save_folder = os.path.join(os.path.dirname(entranceFile), _output_dir)
    os.makedirs(save_folder, exist_ok=True)

    return os.path.join(save_folder, filename)

def main(entranceFilePath):
    """
    主函数
    メイン関数
    Main function

    :param entranceFilePath: 入口文件路径
                            エントリファイルのパス
                            Path to the entrance file
    :return: 保存路径
             保存パス
             Save path
    """
    # workFolder, entranceFile = getWorkFolderWithFile(entranceFilePath)
    # os.chdir(workFolder)
    entranceFilePath = os.path.abspath(entranceFilePath)
    savePath = getSavePath(entranceFilePath)

    # 分離導入とコード
    # インポートとコードを分離
    # Separate import and code statements
    restImportModules, otherLines = recursiveParser(entranceFilePath, visited=[])

    # 合并import语句
    # import文を結合
    # Combine import statements
    combineRestImportLines = combineImportLines(restImportModules)

    # 输出结果文件
    # 結果ファイルを出力
    # Output result file
    showMainRests(entranceFilePath, combineRestImportLines, otherLines)
    outputPyFile(combineRestImportLines, otherLines, savePath)

    return savePath


if __name__ == '__main__':
    # 测试
    # テスト
    # Test
    # print(main(r"..\\examples\\realProject1\\realTest.py"))
    print(main(r"..\\examples/Project2/batches/inspect_all_table.py"))