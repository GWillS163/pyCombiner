# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 20:20
import os.path

from pyCombiner.echos.echo import showMainRests
from pyCombiner.lineParse import *
from pyCombiner.osFileOpr import *
from pyCombiner.parseCore import *
from pyCombiner.depend.findOri import *


def recursiveParser(entranceFilePath, funcList=None, visited=None) -> list:
    """
    单个文件的递归解析
    :param visited:
    :param funcList: 需要导入的函数
    :param entranceFilePath:
    :return:
    """

    def enterDeeper(sumImportModules, sumOtherLines):
        """进入下一层; enter deeper"""
        subImportModule, subOtherLines = recursiveParser(deeperFilePath, visited=visited)
        sumImportModules += subImportModule
        sumOtherLines += subOtherLines
        return subOtherLines

    # 得到imports 和 fromImports
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
    # 处理import 语句的模块
    for module in importModules:
        # 保存未找到的模块的原文
        deeperFilePath = findFromOrigin(module, workFolder)
        if not deeperFilePath:  # 若未找到
            # print("can't find module file: ", module)
            restImportModules.append(importModules[module])
        else:
            enterDeeper(restImportModules, restOtherLines)

    # 处理 from import 语句的模块
    for module in fromModules:
        # 尝试找到模块的文件

        deeperFilePath = findFromOrigin(module, workFolder)
        if not deeperFilePath:
            # print("can't find module file: ", module)
            restImportModules += fromModules[module].values()
            continue
        else:
            # 进入模块文件
            childOtherLines = enterDeeper(restImportModules, restOtherLines)
            if funcList:
                importedPart = getPartImportedFuncs(childOtherLines, fromModules[module].keys())
                restOtherLines = importedPart + restOtherLines
    return [restImportModules, restOtherLines + otherLines]


def combineImportLines(restImportModules):
    """
    将import语句合并
    :param restImportModules:
    :return:
    """
    combineRestImportLines = list(set(restImportModules))
    # sort by length
    combineRestImportLines.sort(key=lambda x: len(x))
    # TODO: 优化合并算法
    # afterCombined = []
    # combine some import lines of the same start words
    # from typing import Optional" to "from typing import Dict, Optional"
    # for importLine in combineRestImportLines:
    #     if not importLine.startswith("from"):
    #         continue
    #     if importLine.split(" ")[1] in
    return combineRestImportLines


def getSavePath(outFolder, entranceFile):
    makeFolder(outFolder)
    savePath = os.path.join(".", outFolder, entranceFile)

    currentTime = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    return savePath.replace(".py", f"{currentTime}.py")


def main(entranceFilePath):
    # change folder to entranceFile's folder
    workFolder, entranceFile = getWorkFolderWithFile(entranceFilePath)
    os.chdir(workFolder)
    outFolder = "CombineOutput"
    savePath = getSavePath(outFolder, entranceFile)
    # 入口文件
    restImportModules, otherLines = recursiveParser(entranceFile, visited=[])
    combineRestImportLines = combineImportLines(restImportModules)
    showMainRests(entranceFilePath, combineRestImportLines, otherLines)
    outputPyFile(combineRestImportLines,
                 otherLines,
                 savePath)
    # print("saved to: ", savePath)
    return savePath


if __name__ == '__main__':
    # Test
    # main(r"D:\Project\mergeMultiPyFiles\examples\demo_simple\runMeSim.py")
    # main(r"D:\Project\mergeMultiPyFiles\examples\demo_complicate\runMeCom.py")
    print(main(r"..\\examples\\realProject1\\realTest.py"))
