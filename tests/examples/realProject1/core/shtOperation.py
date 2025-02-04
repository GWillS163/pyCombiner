#  Author : Github: @GWillS163
# Description: 用于 添加数据，增加行列
from .utils import *


def shtCopyTo(sht1, sht1Scp, sht2, sht2Start):
    """
    将sht1中的数据复制到sht2中
    copy data and format from sht1 to sht2
    :param sht1:
    :param sht1Scp:
    :param sht2:
    :param sht2Start:
    :return:
    """
    # sht1.activate()
    if not sht1.range(sht1Scp).value:
        print(f"\t{sht1.name} - {sht1Scp} - 无数据")
        return
    sht1.range(sht1Scp).api.Copy()
    # judge whether clipboard is empty
    sht2.activate()
    sht2.range(sht2Start).api.Select()
    sht2.api.Paste()  # 总是会报错，不知道为什么pywintypes.com_error: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel',
    # '类 Worksheet 的 Paste 方法无效', 'xlmain11.chm', 0, -2146827284), None)

    # if sht2.range(sht2Start).api.PasteSpecial(Paste=-4163):
    #     print("   复制成功")
    # else:
    #     print("   复制失败")

    # sht1.activate()

def getLastColCell(sht):
    """
    获取标题行的起始单元格
    :param sht:
    :return:
    """
    print("   自动获取标题粘贴起点 - ",
          end="")
    lastColNum = sht.used_range.last_cell.column  # 这个封装好了， 稳定
    # lastColNum = sht.range("A1").end("right").column  # 这个遇到空值会失效， 这个直观
    shtTitleCopyTo = f"{getColLtr(lastColNum)}1"
    print(f"{sht.name} - {shtTitleCopyTo}")
    return shtTitleCopyTo


def getLastRowCell(sht):
    print("   自动获取sht4标题粘贴起点 -",
          end="")
    lastRowNum = sht.used_range.last_cell.row  # 这个封装好了， 稳定
    # lastColNum = sht.range("A1").end("right").column  # 这个遇到空值会失效， 这个直观
    lastCell = f"A{lastRowNum + 1}"
    print(f"[{sht.name}] - {lastCell}")
    return lastCell


def sht2AddSummary(sht2_lv2Score, sht2WithLv, row, startCol, endCol):
    """
    为Sheet2 添加合计行
    :param sht2_lv2Score:
    :param sht2WithLv:
    :param row:
    :param startCol:
    :param endCol:
    :return:
    """
    unitScoreSum = 0
    while True:
        unit = sht2_lv2Score.range(f"A{row}").value
        unitScoreSum += sht2_lv2Score.range(f"{startCol}{row - 1}").value
        if unit:
            break
    sht2_lv2Score.range(f"{startCol}{row}").value = unitScoreSum


def addUnitScpOffsite(unitScp: list, offsite: int = 2) -> list:
    """
    为unitScp添加偏移量
    :param offsite: 默认2 是包含标题行
    :param unitScp:
    :return:
    """
    return [[node + offsite for node in unit] for unit in unitScp]


def sht2OprAddSummaryRows(sht2_lv2Score, unitScp: list):
    """
    为Sheet2 添加合计行
    add summary rows to sheet2
    :param sht2_lv2Score:
    :param unitScp:
    :return:
    """
    unitScpOffsite = addUnitScpOffsite(unitScp, 3)
    sumStr = ",".join([f"C{unitScp[0]}:C{unitScp[1]}" for unitScp in unitScpOffsite])
    endRow = unitScpOffsite[-1][-1] + 1  # 最后一行 - last row
    sht2_lv2Score.range(f"A{endRow}").api.EntireRow.Insert()
    sht2_lv2Score.range(f"B{endRow}").value = "总计"
    sht2_lv2Score.range(f"C{endRow}").value = f"=SUM({sumStr})"

    # 为Sheet2 添加合计行 - add summary rows to sheet2
    while unitScpOffsite:
        unit = unitScpOffsite.pop()
        insertRow = unit[-1] + 1
        sht2_lv2Score.range(f"B{insertRow}").api.EntireRow.Insert()
        sht2_lv2Score.range(f"B{insertRow}").value = "合计"
        sht2_lv2Score.range(f"C{insertRow}").value = f"=SUM(C{unit[0]}:C{unit[-1]})"

    # insertRow = []
    # for row in range(4, 20):
    #     unit = sht2_lv2Score.range(f"A{row}").value
    #     ques = sht2_lv2Score.range(f"B{row}").value
    #     if not ques:
    #         if not unit:
    #             insertRow.append(row)
    #             break
    #         continue
    #     if unit:
    #         insertRow.append(row)
    #
    # endRow = insertRow[-1]
    # # 增加每单元的合计行 - add summary for each unit
    # sumForAll = []
    # while insertRow:
    #     row = insertRow.pop()
    #     sht2_lv2Score.range(f"A{row}").api.EntireRow.Insert()
    #     sht2_lv2Score.range(f"B{row}").value = "合计"
    #     unitSumScp = fillSht2UnitSumLineHor(sht2_lv2Score, row, "C")
    #     sumForAll.append(unitSumScp)
    #     # 以下代码不需计算其他列的合计
    #     # sht2CalcSummary(sht2_lv2Score, sht2WithLv, row, "C", "L")
    # sht2_lv2Score.range(f"C{endRow}").value = f"=SUM({','.join(sumForAll)})"


def fillSht2UnitSumLineHor(sht2, summaryRow: int, sumCol: str, unitCol="A") -> str:
    """
    负责每个单元横向填充 的“合计” 及 求和权重，目前只需填充一列2022-10-1
    responsible for filling one line of summary and summing up the weight
    :param unitCol: 用来判断单元格大小
    :param sumCol: "D:L"
    :param summaryRow:
    :param sht2:"""
    # 获得最上面的范围 - get unit top scope
    unitRow = summaryRow
    while True:
        unitRow -= 1
        if sht2.range(f"{unitCol}{unitRow}").value:
            break
    # 单列填充 - Fill one Line
    unitScp = f"{sumCol}{unitRow}:{sumCol}{summaryRow - 1}"
    sht2.range(f"{sumCol}{summaryRow}").value = f"=SUM({unitScp})"
    return unitScp
    # 横向填充 - Fill one Line
    # fillRan = getTltColRange(sumCol)
    # for colNum in fillRan:
    #     colLtr = getColLtr(colNum)
    #     unitScp = f"{colLtr}{unitRow}:{colLtr}{summaryRow - 1}"
    #     sht2.range(f"{colLtr}{summaryRow}").value = f"=SUM({unitScp})"


def sht1SetData(sht1_lv2Result, sht1WithLv, titleRan):
    """
    Sheet1 中，纵向放入 每个lv2_3部门的数据(如果有) 包含“二级单位”
    :param titleRan:
    :param sht1_lv2Result:
    :param sht1WithLv:
    :return: None
    """
    lv2 = None
    n = 0
    for colNum in titleRan:
        n += 1
        colLtr = getColLtr(colNum)
        curLv2 = sht1_lv2Result.range(f"{colLtr}1").value
        lv3 = sht1_lv2Result.range(f"{colLtr}2").value
        if curLv2:
            lv2 = curLv2
        if not lv3:
            continue
        if lv2 not in sht1WithLv:
            continue
        if lv3 not in sht1WithLv[lv2]:
            continue
        # print(f"{lv2} & {lv3} exists")
        print(f"Sheet1 Progress:{n}/{len(titleRan)} - [{lv2}-{lv3}]")  # , end="\r")
        sht1_lv2Result.range(f"{colLtr}3"). \
            options(transpose=True).value = sht1WithLv[lv2][lv3]
    print("Sheet1 Progress:100% - [Done]")


def sht2SetData(sht2_lv2Score, sht2WithLv, titleRan: range):
    """
    Sheet2 中，纵向放入 每个lv3部门的数据(如果有) * 该lv3部门的权重
    :param titleRan: Range
    :param sht2_lv2Score:
    :param sht2WithLv:
    :return:
    """
    lv2 = None
    n = 0
    for colNum in titleRan:
        n += 1
        colLtr = getColLtr(colNum)
        curLv2 = sht2_lv2Score.range(f"{colLtr}1").value
        lv3 = sht2_lv2Score.range(f"{colLtr}2").value
        if curLv2:
            lv2 = curLv2
        if not lv3:
            continue
        if lv2 not in sht2WithLv:
            continue
        if lv3 not in sht2WithLv[lv2]:
            continue
        # print(f"{lv2} & {lv3} exists")
        print(f"Sheet2 Progress:{n}/{len(titleRan)} - [{lv2}-{lv3}]")  # , end="\r")
        sht2_lv2Score.range(f"{colLtr}3") \
            .options(transpose=True).value = sht2WithLv[lv2][lv3]
    print("Sheet2 Progress:100% - [Done]")


def sht3SetData(sht3, sht3WithLv: dict, titleRange: str, lv1Name: str):
    """
    Sheet3 中，纵向放入 每个lv2部门的数据(如果有)
    :param lv1Name:
    :param sht3:
    :param sht3WithLv:
    :param titleRange:
    :param lv1Name: 默认是 北京公司
    :return:
    """
    titleRan = getTltColRange(titleRange)
    # lv2Clz = None
    n = 0
    for col in titleRan:
        n += 1
        colLtr = getColLtr(col)
        lv2UpCurr = sht3.range(f"{colLtr}1").value
        lv2 = sht3.range(f"{colLtr}2").value
        # if lv2UpCurr:
        if lv2UpCurr == lv1Name:
            sht3.range(f"{colLtr}3").options(transpose=True).value = sht3WithLv[lv1Name]
        # lv2Clz = lv2UpCurr
        if not lv2 in sht3WithLv:
            continue
        # place score list vertically
        print(f"Sheet3 Progress:{n}/{len(titleRan)} - [{lv2UpCurr}-{lv2}]")  # , end="\r")
        sht3.range(f"{colLtr}3").options(transpose=True).value = sht3WithLv[lv2]
    print("Sheet3 Progress:100% - [Done]")


def sht4SetData(sht4, sht4WithLv, titleRan, lv1Name):
    """
    Sheet4 中，横放入 每个lv2部门的数据(如果有)
    :param lv1Name:
    :param sht4:
    :param sht4WithLv:
    :param titleRan:
    :return:
    """
    lv1 = None
    for row in titleRan:
        lv1Curr = sht4.range(f"A{row}").value
        lv2 = sht4.range(f"B{row}").value
        if lv1Curr:
            if lv1Curr == lv1Name:
                sht4.range(f"C{row}").value = sht4WithLv[lv1Name][lv2]
            lv1 = lv1Curr
        if not lv1 in sht4WithLv:
            continue
        if not lv2 in sht4WithLv[lv1]:
            continue
        # place score list vertically,  row object of type 'int' has no len()
        sht4.range(f"C{row}").value = sht4WithLv[lv1][lv2]
    print("Sheet4 Progress:100% - [Done]")


def addOneDptData(shtSum, scpLst, height,
                  shtDept, shtTitleTo):
    """
    粘贴一个部门的数据从 sht1Sum 到 sht1Dept
    :param shtDept:
    :param height:
    :param scpLst:
    :param shtTitleTo:
    :param shtSum:
    :param  从模板表中复制的部门范围
    :return:
    """
    # 数据栏 复制
    sht1BorderL, sht1BorderR = scpLst  # deptUnitSht1[deptName]
    sht1DataZone = f"{sht1BorderL}1:{sht1BorderR}{height}"
    # shtDept.activate()
    # pywintypes.com_error: (-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)
    shtCopyTo(shtSum, sht1DataZone,
              shtDept, shtTitleTo)
    return sht1BorderL, sht1BorderR


def getUnitScpRowList(sht, unitCol: str, baseCol: str, keywords: list) -> list:
    """
    找到关键词单位的范围, 多个单元必须是一片连续的范围
    :param sht: sheet
    :param unitCol: 单元列，用来识别关键词
    :param baseCol: 基准列
    :param keywords: 单元的关键词
    :return: [int, int]
    """
    res = [-1, -1]
    lastUnit = None
    lastRow = sht.used_range.last_cell.row + 1
    for row in range(0, lastRow):
        unit = sht.range(f"{unitCol}{row}").value
        base = sht.range(f"{baseCol}{row}").value
        if unit or not base:
            if lastUnit in keywords:
                res[1] = row
            break
        if unit in keywords:
            res[0] = row
            lastUnit = unit
    return res


def dltOneDptData(shtDept, shtTitleCopyTo, deptCopyHeight, shtBorderL, shtBorderR):
    """
    删除部门数据，用于重新填入
    :param shtDept:
    :param shtTitleCopyTo:
    :param deptCopyHeight:
    :param shtBorderR:
    :param shtBorderL:
    :return:
    """
    borderWidth = getColNum(shtBorderR) - getColNum(shtBorderL)
    borderStart = getColNum(shtTitleCopyTo[0])
    borderEnd = getColLtr(borderStart + borderWidth)
    shtDept.range(f"{shtTitleCopyTo}:{borderEnd}{deptCopyHeight}").api.Delete()


def getRuleByQuestionSurvey(name, questTitle, surveyTestSht, surveyQuesCol, surveyRuleCol):
    """

    # 找到所属规则 - locate which row is rule by questTitle in scoreExl
    :param name:
    :param questTitle:
    :param surveyTestSht:
    :param surveyQuesCol: 问题所在列
    :param surveyRuleCol:
    :return:
    """
    answerRow = -1
    for row in range(3, 40):
        # find the cell in the surveyExl by Question column
        ruleQuest = surveyTestSht.range(f"{surveyQuesCol}{row}").value
        if ruleQuest is None:
            continue
        if ruleQuest in questTitle:
            # print("Check_cellQuestion: ", ruleQuest)
            answerRow = row
            break
    if answerRow == -1:
        print(f"{name} [not found]: {questTitle}")
        return answerRow, None
    # get rule in the surveyExl App
    rule = surveyTestSht.range(f"{surveyRuleCol}{answerRow}").value
    return answerRow, rule


def getRuleByQuestionList( questTitle, surveyDataList):
    """
    读取到内存加速判断
    # 找到所属规则 - locate which row is rule by questTitle in scoreExl
    :param questTitle:
    :param surveyDataList:
    :return:
    """
    for row in surveyDataList:
        # find the cell in the surveyExl by Question column
        ruleQuest = row[0]  # 问题所在列
        if not ruleQuest:
            continue
        if ruleQuest.strip() in questTitle:
            return row[1], row[2]  # 类型， 规则

    return None, None
