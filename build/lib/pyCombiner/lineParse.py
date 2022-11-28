# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 21:52
import re


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
    # TODO: strip "as" and following strings
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


if __name__ == '__main__':
    print(getFromLineData("from core.main import path, sys"))
    print(getImportLineData("import os"))

    print(getImportLineData("import os, datetime"))
    print(getFromLineData("from os import path, sys"))