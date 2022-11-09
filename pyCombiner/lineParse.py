# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 21:52
import re


def getModules(importedStr):
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
    return getModules(res[0])


def getFromLineModules(fromLine: str) -> list:
    """
    获取单行From 导入
    :param fromLine:
    :return:
    """
    importStart = fromLine.find('import')
    # find the words after "from" and after "import" by regex
    res = re.findall(r'from\s+(.*)\s?', fromLine[:importStart])
    if not res:
        return []

    moduleName = res[0].strip()
    funcsImported = getImportLineModules(fromLine[importStart:])
    return [moduleName, funcsImported]


def getImportModules(lines: list) -> list:
    """
    Get the modules that import directly.
    :param lines:
    :return: ['os', 'utils']
    """
    modules = []
    for line in lines:
        modules.extend(getImportLineModules(line))
    return modules


def getFromModules(lines: list) -> list:
    """
    Get the modules that import from other modules.
    :param lines:
    :return:
    """
    modules = []
    for line in lines:
        fromLineImport = getFromLineModules(line)
        if not fromLineImport:
            continue
        modules.append(fromLineImport)
    return modules


def getImportLineDict(importLine):
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


def getFromLineDict(fromLine) -> dict:
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
    :return: [{}, {}, []]
    """
    importModules = {}  # key is moduleName, value is import statement
    fromModules = {}
    otherLines = []
    for line in lines:
        if not line:
            continue
        lineStriped = line.strip()
        if line.startswith('import'):
            importModules.update(getImportLineDict(lineStriped))
        elif line.startswith('from'):
            fromModules.update(getFromLineDict(lineStriped))
        else:
            otherLines.append(line)

    return [importModules, fromModules, otherLines]

