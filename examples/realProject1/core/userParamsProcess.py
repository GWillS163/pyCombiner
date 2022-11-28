# Github: GWillS163
# User: 駿清清 
# Date: 11/10/2022 
# Time: 22:16
from .utils import getColLtr, getColNum
from .debugPrint import printAutoParamSht1, printAutoParamSht2, printAutoParamSht3, printAutoParamSht4

import datetime
import os

def autoGetSht1Params(sht1Module, sht1TitleCopyFromSttCol, sht1TitleCopyToSttCol):
    """
    通过少量用户输入, 自动获取问卷模板的参数.
    :param sht1Module: 问卷模板Sheet1
    :param sht1TitleCopyFromSttCol: "G" 给定模板的标题起始点，自动获取标题的范围
    :param sht1TitleCopyToSttCol: "G"  给定粘贴的起始点，自动获取标题粘贴和数据栏的范围
    :return:
    """
    # 获得模板表标题范围 - get the last column of the sht1Module
    sht1TitleCopyEndCol = sht1Module.used_range.last_cell.address.split("$")[1]
    sht1TitleCopyFromMdlScp = f"{sht1TitleCopyFromSttCol}1:{sht1TitleCopyEndCol}2"  # F1:MG2

    # 获得结果表标题最后列 - get the last column of the sht1Result
    sht1TitleCopyToEndCol = getColLtr(
        getColNum(sht1TitleCopyEndCol) - getColNum(sht1TitleCopyFromSttCol) + getColNum(sht1TitleCopyToSttCol))
    # 部门文件拆分用, 相同值不要合并，未经大量测试，不知道是否会出现问题
    sht1DeptTltRan = f"{sht1TitleCopyToSttCol}:{sht1TitleCopyToEndCol}"
    # 运行时自动获取 > sht1TitleCopyTo = f"{sht1TitleCopyToSttCol}1"
    sht1DataColRan = f"{sht1TitleCopyToSttCol}:{sht1TitleCopyToEndCol}"
    printAutoParamSht1(sht1TitleCopyFromMdlScp, sht1DeptTltRan, sht1DataColRan)
    return sht1TitleCopyFromMdlScp, sht1DeptTltRan, sht1DataColRan


def autoGetSht2Params(sht2Mdl, sht0Sur, sht2DeleteCopiedColScp, mdlTltStt):
    """
    通过少量用户输入, 自动获取sht2问卷模板的参数

    :param mdlTltStt:
    :param sht2DeleteCopiedColScp:
    :param sht0Sur:
    :param sht2Mdl:
    :return:
    """

    # 已自动获取 sht2WgtLstScp = "C3:C13"
    # 已自动获取 sht2DeleteCopiedRowScp = "A32:A55"  # 要删除的部分需要动态获取 党廉&纪检 的区间
    # 用户指定 sht2DeleteCopiedColScp = "C1:J1"  需要删除的列

    # Title标题 - 从模板中获取
    # 用户指定 mdlTltStt = "D"
    mdlTltEnd = sht2Mdl.used_range.last_cell.address.split("$")[1]  # "PO"
    sht2TitleCopyFromMdlScp = f"{mdlTltStt}1:{mdlTltEnd}2"  # autoGetSht2TltScp(self.sht2Module)
    sht2DeptTltRan = f"{mdlTltStt}:{mdlTltEnd}"  # 部门文件 拆分用

    # Index侧栏 - 从测试问卷中获取
    surLastCol = sht0Sur.used_range.last_cell.address.split("$")[1]  # 一般是 K
    surLastRow = sht0Sur.used_range.last_cell.address.split("$")[2]  # 一般是 32
    sht2IndexCopyFromSvyScp = f"A1:{surLastCol}{surLastRow}"  # "A1:K32"
    sht2IndexCopyTo = "A1"  # 默认粘贴在左上角 - default paste to the left top corner

    # 结果表 - 粘贴的起始点
    # 自动获取最后一列 > sht2TitleCopyTo = "D1"
    # 已自动获取 > sht2TitleDataColScp = "D1:L2"

    printAutoParamSht2(sht2DeleteCopiedColScp, sht2TitleCopyFromMdlScp,
                       sht2DeptTltRan, sht2IndexCopyFromSvyScp, sht2IndexCopyTo)
    return sht2DeleteCopiedColScp, sht2TitleCopyFromMdlScp, sht2DeptTltRan, sht2IndexCopyFromSvyScp, sht2IndexCopyTo


def autoGetSht3Params(sht3Mdl, sht0Sur, mdlTltStt, surLastCol, resTltStt):
    """
    通过少量用户输入, 自动获取sht3问卷模板的参数
    :param sht3Mdl: sht3模板
    :param sht0Sur: sht0问卷
    :param mdlTltStt: 模板表的标题起始列
    :param surLastCol: 测试问卷的最后一列
    :param resTltStt: 结果表的标题粘贴起始列
    :return:
    """
    # title标题 - 从模板中获取
    # mdlTltStt = "L"
    mdlTltEnd = sht3Mdl.used_range.last_cell.address.split("$")[1]  # "BA"
    sht3TitleCopyFromMdlScp = f"{mdlTltStt}1:{mdlTltEnd}2"

    # index侧栏 - 从测试问卷中获取
    # surLastCol = "J"
    surLastRow = sht0Sur.used_range.last_cell.address.split("$")[2]  # 54不是 "32"
    sht3IndexCopyFromSvyScp = f"A1:{surLastCol}32"

    # result结果表 - 粘贴的起始点
    # resTltStt = "K"
    resTltEnd = getColLtr(getColNum(mdlTltEnd) - getColNum(mdlTltStt) + getColNum(resTltStt))  # "AZ"
    sht3DataColRan = f"{resTltStt}:{resTltEnd}"
    sht3TitleCopyTo = f"{resTltStt}1"
    printAutoParamSht3(sht3IndexCopyFromSvyScp, sht3DataColRan, sht3TitleCopyFromMdlScp)
    return sht3IndexCopyFromSvyScp, sht3DataColRan, sht3TitleCopyFromMdlScp


def autoGetSht4Params(sht4, sht4IndexFromMdl4Scp, sht4TitleFromSht2Scp, sht4SumTitleFromMdlScp, sht4DataRowRan=None):
    """
    通过少量用户输入, 自动获取sht4问卷模板的参数
    :param sht4IndexFromMdl4Scp: 'A4:B52'  # Sht4固定两列AB公司, 长度是变量  title标题 - 从模板中获取
    :param sht4TitleFromSht2Scp: 'A1:C17'  # Sht2模板中的index侧栏，长度是固定的
    :param sht4SumTitleFromMdlScp: "P1:Q3"  # 汇总单元格（本线条、全公司排名）
    :param sht4DataRowRan: 扫描行范围
    :param sht4:
    :return:
    """
    # index侧栏 - 从sht2中获取
    # 用户输入：sht4IndexFromMdl4Scp =
    # 用户输入：sht4TitleFromSht2Scp =

    #
    # 用户输入： sht4SumTitleFromMdlScp =

    # result结果表 - 粘贴的起始点
    # 默认值 sht4TitleCopyTo = 'A1'  # 默认值 - 无需修改
    # 已自动获取最后一列 sht4SumTitleCopyTo = "R1"  # 有改动, 应该要根据sht2的标题行数自动调整

    printAutoParamSht4(sht4IndexFromMdl4Scp, sht4TitleFromSht2Scp, sht4SumTitleFromMdlScp,
                       )
    return sht4IndexFromMdl4Scp, sht4TitleFromSht2Scp, sht4SumTitleFromMdlScp,


def paramsCheckExist(surveyExlPath, partyAnsExlPh, peopleAnsExlPh):
    """
    检查输入文件是否存在, 并新建保存路径
    Check Input files are
    :param peopleAnsExlPh:
    :param surveyExlPath:
    :param partyAnsExlPh:
    :return:
    """
    if surveyExlPath == partyAnsExlPh:
        raise FileExistsError("文件名输出重复,", surveyExlPath, partyAnsExlPh)
    fileDict = {
        "问卷模板文件": surveyExlPath,
        "党员答题文件": partyAnsExlPh,
        "群众答题文件": peopleAnsExlPh
    }
    for name, path in fileDict.items():
        if not os.path.exists(path):
            raise FileNotFoundError(f"{name} 不存在:")


def getCurrentYear(userYear=None):
    # if userYear is a yearNum within 1970-2500, return year, otherwise return current year
    if not userYear or not userYear.isdigit():
        return datetime.datetime.now().year
    if 1970 <= int(userYear) <= 2500:
        return int(userYear)
    else:
        return datetime.datetime.now().year


def getSumSavePathNoSuffix(savePath, fileYear, fileName):
    """
    得到每一个文件全路径的输出的前缀
    :param savePath:
    :param fileYear:
    :param fileName:
    :return:
    """
    fileYear = getCurrentYear(fileYear)
    return os.path.join(savePath, f"{fileYear}_{fileName}")


def readLvDict(orgSht):
    """返回结构化的字典，key是部门名，value是部门下的子部门"""
    allOrg = {}
    row = 1
    while True:
        a = orgSht.range(f"A{row}").value
        b = orgSht.range(f"B{row}").value
        if not (a and b):
            break
        if a not in allOrg:
            allOrg.update({a: [b]})
        else:
            allOrg[a].append(b)
        row += 1
    return allOrg


def paramsCheckSurvey(surveyExl, shtNameList: list):
    """
    调查问卷 文件名检查
    Survey Excel params check
    :param surveyExl:
    :param shtNameList:
    :return:
    """
    for shtName in shtNameList:
        try:
            surveyExl.sheets[shtName]
        except Exception as e:
            raise Exception(f"{shtName} 不存在于{surveyExl.name}中.\n ({e})")
    print(f"{surveyExl.name} Sheet 参数检查通过")


def getSavePath(savePath, fileYear, fileName):
    """
    生成用户保存路径
    :param fileYear:
    :param fileName:
    :param savePath:
    :return:
    """
    # make an output dir with current time
    outputDir = os.path.join(savePath, "PartyBuildOutput_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    os.makedirs(outputDir)

    sumSavePathNoSuffix = getSumSavePathNoSuffix(outputDir, fileYear, fileName)
    return outputDir, sumSavePathNoSuffix
