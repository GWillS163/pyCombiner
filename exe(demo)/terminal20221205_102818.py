import re
import os
import sys
import time
import os.path
# Github: GWillS163
# User: 駿清清 
# Date: 12/11/2022 
# Time: 19:18

def colorPrint(*args, **kwargs):
    colorDict = {
        'black': 30,
        'red': 31,
        'green': 32,
        'yellow': 33,
        'blue': 34,
        'purple': 35,
        'cyan': 36,
        'white': 37,
        'blackBG': 40,
        'redBG': 41,
        'greenBG': 42,
        'yellowBG': 43,
        'blueBG': 44,
        'purpleBG': 45,
        'cyanBG': 46,
    }
    """Prints the given arguments in color.

    Args:
        *args: The arguments to print.
        **kwargs: The keyword arguments to pass to the print function.
    """
    # Get the color from the keyword arguments.
    colorStr = kwargs.pop("color", None)
    colorCode = colorDict.get(colorStr, None)
    # If the color is not None, then print the arguments in color.
    if colorCode:
        print(f"\033[{colorCode}m", end="")
        print(*args, **kwargs)
        print("\033[0m", end="")
    # Otherwise, print the arguments normally.
    else:
        print(*args, **kwargs)


# if __name__ == '__main__':
#
#     colorPrint("Hello World!", color="red")
# # Github: GWillS163
# # User: 駿清清
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


# if __name__ == '__main__':
#     importModules = {'update': 'import update as up'}
#     fromModules = {'bin': {'*': 'from bin import *'}}
#     otherLines = ['', '# 4', 'up.showBanner()', 'main()']
#     showRests("Mock Import data", importModules, fromModules, otherLines)
# # Github: GWillS163
# # User: 駿清清
# Date: 01/11/2022 
# Time: 21:52


def parseModules(importedStr):
    """
    获得import后面的模块名
    :param importedStr:
    :return:
    """
    # TODO: 没有处理 import *
    return [module.strip() for module in importedStr.split(',')]


def getImportLineModules(importLine: str) -> list:
    """
    single Import Line
    :param importLine:
    :return:
    """
    if 'as' in importLine:
        importLine = importLine[:importLine.find('as')]

    res = re.findall(r'import\s+(.*)', importLine)
    return parseModules(res[0])


def getFromLineModules(fromLine: str) -> list:
    """
    处理from [origin] import [func1, func2] 的情况
    :param fromLine:
    :return: [origin] [func1, func2]
    """
    importStartIndex = fromLine.find('import')
    # find the words after "from" and after "import" by regex
    res = re.findall(r'from\s+(.*)\s?', fromLine[:importStartIndex])
    if not res:
        return []

    moduleName = res[0].strip()
    funcsImported = getImportLineModules(fromLine[importStartIndex:])
    return [moduleName, funcsImported]


def getImportLineData(importLine):
    """
    return {Module: context, Module: context}
    :param importLine:
    :return: {'os': 'import os as o'}
             {'re': 'import re', 'utils': 'import utils'}
    """
    importLinesDict = {}
    moduleNames = getImportLineModules(importLine)
    if len(moduleNames) == 1:  # import os
        importLinesDict.update({moduleNames[0]: importLine})
    else:  # import os, datetime
        for moduleName in moduleNames:
            importLinesDict.update({moduleName: f'import {moduleName}'})
    return importLinesDict


def getFromLineData(fromLine) -> dict:
    """
    return {Module: {func: context, func: context}
    :param fromLine: from ... import ..., ... (as ...)
    :return: {'os': {'path': 'from os import path',
                     'sys': 'from os import sys'}},
            {'os': {'path': 'from os import path as p'}}
    """
    fromLines = {}
    [moduleName, funcsImported] = getFromLineModules(fromLine)
    if len(funcsImported) == 1:
        fromLines.update({moduleName: {funcsImported[0]: fromLine}})
    else:
        fromLines.update({moduleName: {}})
        for func in funcsImported:
            fromLines[moduleName].update({func: f"from {moduleName} import {func}"})
    return fromLines


def parseSinglePy(lines) -> list:
    """
    parse a single py file, get imports ,froms, others
    :param lines:
    :return: [{'os': 'import os'}, {}, []]
    """
    importModules = {}  # key is moduleName, value is import statement
    fromModules = {}
    otherLines = []
    for line in lines:
        if not line:
            continue
        lineStriped = line.strip()
        if line.startswith('import'):
            importModules.update(getImportLineData(lineStriped))
        elif line.startswith('from'):
            fromModules.update(getFromLineData(lineStriped))
        else:
            otherLines.append(line)

    return [importModules, fromModules, otherLines]


# if __name__ == '__main__':
#     print(getFromLineData("from core.main import path, sys"))
#     print(getImportLineData("import os"))
#
#     print(getImportLineData("import os, datetime"))
#     print(getFromLineData("from os import path, sys"))# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 21:20
# 系统文件操作



def outputPyFile(restImportLines, restCodeLines, path):
    # make new pyFile to write codes
    with open(path, 'w', encoding='utf-8') as f:
        for importLine in restImportLines:
            f.write(importLine)
            f.write("\n")
        f.writelines(restCodeLines)


def parseFile(filePath: str) -> list:
    # parse the python file
    # return a list of lines
    with open(filePath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    return lines


def makeFolder(path):
    if not os.path.exists(path):
        os.makedirs(path)


def getWorkFolderWithFile(entranceFile) -> list:
    """
    Get the folder of entranceFile, and change the workDir to it.
    :param entranceFile:
    :return:
    """
    if not "\\" in entranceFile or not "/" in entranceFile:
        entranceFile = os.path.abspath(entranceFile)
    divideNum = entranceFile.rfind('\\')
    return [entranceFile[:divideNum], entranceFile[divideNum + 1:]]


# if __name__ == '__main__':
#     outputPyFile(["import os", "import sys"],
#                  ["def Print():\n", "\tprint('hello world')\n"],
#                  "output\\readMe.py")
# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 20:59


# TODO: as xx 这种语法解析不了


def isModuleExist(moduleName):
    """
    检查模块是否存在当前目录
    :param moduleName:
    :return:
    """
    pyFilePath = moduleName + ".py"
    if not isPyFileCanFind(pyFilePath):
        print("can't find module file: ", pyFilePath)
        return False
    return True


def isPyFileCanFind(pyFilePath):
    """
    检查py文件是否存在
    :param pyFilePath:
    :return:
    """
    return os.path.exists(pyFilePath)


def getModuleAllFuncs(module) -> list:
    """
    得到模块中的所有函数
    :param module:
    :return:
    """
    allLines = parseFile(module)
    # return "\n".join([line for line in allLines])
    return allLines


def getPartImportedFuncs(module: str, funcs: list) -> list:
    """
    得到模块中的函数文本
    from module import func1, func2
    from module import *
    :param module: utils
    :param funcs:  [showLogin, showExit]
    :return: def showLogin() ... def showExit() ...
    """
    return ["部分导入的函数"]
    pass
    # TODO: 得到模块中的函数文本
    if not isModuleExist(module):
        return [""]
    pyFilePath = module + ".py"
    if len(funcs) == 1 and funcs[0] == "*":
        return getModuleAllFuncs(pyFilePath)
    lines = parseFile(pyFilePath)
# Github: GWillS163
# User: 駿清清 
# Date: 28/11/2022 
# Time: 12:20


def findFromOrigin(moduleFullName: str, workPath: str) -> str:
    """
    find the origin of from import
    :param moduleFullName:
    :param workPath:
    :return:
    """
    # os.chdir(workPath)
    allWays = [
        # 当前目录检测
        moduleFullName + ".py",
        # TODO:  本地库检测 比如 import numpy

        # core.main 对于含有"."的情况，需要检测是否是本地库
        moduleFullName.replace(".", os.sep) + ".py",
        # .utils .开头的情况

    ]
    if moduleFullName.startswith("."):
        allWays.append(moduleFullName[1:] + ".py")
        allWays.append(moduleFullName.replace(".", os.sep)[1:] + ".py")
    res = ""
    for pyFilePath in allWays:
        fullPath = os.path.join(workPath, pyFilePath)
        if os.path.isfile(fullPath):
            res = fullPath
        else:
            res = ""
    return res


# if __name__ == '__main__':
#     # print(findFromOrigin("os", r"D:\Project\mergeMultiPyFiles\examples\realProject1"))
#     # print(findFromOrigin("core.main", r"D:\Project\mergeMultiPyFiles\examples\realProject1"))
#     # print(findFromOrigin("core.config", r"D:\Project\mergeMultiPyFiles\examples\realProject1"))
#     # print(findFromOrigin("realTest", r"D:\Project\mergeMultiPyFiles\examples\realProject1"))
#     print(findFromOrigin(".utils", r"D:\Project\mergeMultiPyFiles\examples\realProject1\core"))
#
# # Github: GWillS163
# # User: 駿清清
# Date: 01/11/2022 
# Time: 20:20



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
    savePath = os.path.join(r".", outFolder, entranceFile)

    currentTime = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    return savePath.replace(".py", f"{currentTime}.py")


def core(entranceFilePath):
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


# if __name__ == '__main__':
#     # Test
#     # main(r"D:\Project\mergeMultiPyFiles\examples\demo_simple\runMeSim.py")
#     # main(r"D:\Project\mergeMultiPyFiles\examples\demo_complicate\runMeCom.py")
#     print(main(r"..\\examples\\realProject1\\realTest.py"))
# # Github: GWillS163
# # User: 駿清清
# # Date: 05/12/2022
# Time: 10:11

# read the inputs from the user in CLI
# and call the main function



def run():
    if len(sys.argv) == 1:
        # print red error message
        colorPrint("No input file", color="red")
        print("Please input the entrance file path.\n"
              "For example: ", end="")
        colorPrint("pyCombiner ./pyCombiner/main.py", color="green")
        sys.exit(0)

    entrancePath = sys.argv[1]
    if not os.path.exists(entrancePath):
        # print red error message
        colorPrint("[Error]", color="red")
        print("No such file or directory: ")
        sys.exit(0)

    outputPath = core(entrancePath)
    print("\nThe output file is: ", end="")
    colorPrint(outputPath, color="green")


def runForIPA(projectEntrancePyPath):
    """
    为IPA项目提供的接口
    :param projectEntrancePyPath:
    :return:
    """
    outputPath = core(projectEntrancePyPath)
    return outputPath


#run()
    # projectEntrancePath = "D:\\Project\\mergeMultiPyFiles\\examples\\demo_complicate\\runMeCom.py"
savePath = runForIPA(projectEntrancePath)

    # 请注意！
    # 本项目仅提供了一个简单的合并算法，仅供参考。
    # 通过处理import 语句，将多个py文件按照顺序合并为一个py文件。

    # 若要使用最新版，请pip install pycombiner
    # 用例: pycombiner <项目入口文件.py>
    # repo: https://github.com/GWillS163/pyCombiner/
    # pypi: https://pypi.org/project/pyCombiner/
