# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 20:20
import os.path

from pyCombiner.echos.echo import showMainRests
from pyCombiner.lineParse import *
from pyCombiner.osFileOpr import *
from pyCombiner.parseCore import *


def recursiveParser(entranceFile, funcList=None) -> list:
    """
    单个文件的递归解析
    :param funcList: 需要导入的函数
    :param entranceFile:
    :return:
    """
    # 得到imports 和 fromImports
    lines = parseFile(entranceFile)
    importModules, fromModules, otherLines = parseSinglePy(lines)
    restImportModules = []
    restOtherLines = []
    # 处理import 语句的模块
    for module in importModules:
        pyFilePath = module + ".py"
        # 保存未找到的模块的原文
        if not isPyFileCanFind(pyFilePath):
            print("can't find module file: ", pyFilePath)
            restImportModules.append(importModules[module])
            continue
        childImportModule, childOtherLines = recursiveParser(pyFilePath)
        restImportModules += childImportModule
        restOtherLines += childOtherLines
    # 处理 from import 语句的模块
    for module in fromModules:
        pyFilePath = module + ".py"
        if not isPyFileCanFind(pyFilePath):
            print("can't find module file: ", pyFilePath)
            restImportModules += fromModules[module].values()
            continue
        childImportModule, childOtherLines = recursiveParser(pyFilePath)
        restImportModules += childImportModule
        restOtherLines += childOtherLines
        if funcList:
            importedPart = getPartImportedFuncs(childOtherLines, fromModules[module].keys())
            restOtherLines = importedPart + restOtherLines  # TODO:是否更改
    return [restImportModules, restOtherLines + otherLines]


def main(entrancePath):
    # change folder to entranceFile's folder
    workFolder, entranceFile = getWorkFolderWithFile(entrancePath)
    os.chdir(workFolder)
    outFolder = "CombineOutput"
    makeFolder(outFolder)
    savePath = os.path.join(".", outFolder, entranceFile)
    # 入口文件
    restImportModules, otherLines = recursiveParser(entranceFile)
    showMainRests(entrancePath, restImportModules, otherLines)
    outputPyFile(restImportModules,
                 otherLines,
                 savePath)
    print("saved to: ", savePath)


if __name__ == '__main__':
    # Test
    main(r"D:\Project\mergeMultiPyFiles\examples\demo_simple\runMeSim.py")
    main(r"D:\Project\mergeMultiPyFiles\examples\demo_complicate\runMeCom.py")
