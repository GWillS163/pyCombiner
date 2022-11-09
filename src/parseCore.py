# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 20:59


# TODO: as xx 这种语法解析不了
import os

from src.osFileOpr import parseFile


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
    检查py文件是否是用户文件,
    排除模块文件
    :param pyFilePath:
    :return:
    """
    return os.path.isfile(pyFilePath)


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
