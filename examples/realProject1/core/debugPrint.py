# Github: GWillS163
# User: 駿清清 
# Date: 12/10/2022 
# Time: 12:47

def printShtWithLv(shtName, shtWithLv):
    """
    print the sheet with level
    :param shtName:
    :param shtWithLv:
    :return:
    """
    for lv2 in shtWithLv:
        for lv3 in shtWithLv[lv2]:
            try:
                newName = shtName[:4]
                newLv2 = lv2[:4]
                newLv3 = lv3[:4]
            except Exception as e:
                newName = shtName
                newLv2 = lv2
                newLv3 = lv3 + str(e)
            print(f"{newName}-[{newLv2}]\t[{newLv3}]:{shtWithLv[lv2][lv3]}")


def printSht1Data(typ, staffWithLv, scoreExlTitle):
    print("staffWithLv:", staffWithLv)
    print("scoreExlTitle:", scoreExlTitle)
    print(f"{typ}分数：")
    # print(stuffWithLv) all score


def printSht1WithLv(sht1WIthLv):
    """
    输出Sht1WithLv
    :param sht1WIthLv:
    :return:
    """
    for lv2 in sht1WIthLv:
        for lv3 in sht1WIthLv[lv2]:
            print(f"[{lv2}][{lv3}] - {sht1WIthLv[lv2][lv3]}")


def printStaffWithLv(staffWithLv):
    for lv2 in staffWithLv:
        for lv3 in staffWithLv[lv2]:
            for stu in staffWithLv[lv2][lv3]:
                print(stu.name, stu.scoreLst)


def printScoreBugIf(answer, rule, questTitle, quesType, stuff, questionNum):
    if stuff.scoreLst[questionNum] == -100:
        print(f"answer: {answer}\n"
              f"rule: {rule}\n"
              f"questTitle: {questTitle}\n"
              f"quesType: {quesType}\n"
              f"score: {stuff.scoreLst[questionNum]}")


def printAutoParamSht1(sht1TitleCopyFromMdlScp, sht1DeptTltRan, sht1DataColRan):
    print("Sheet1 获取sheet1页填充范围数据 - I. get sheet data")
    print(f"\t从sht1模板这里复制:sht1TitleCopyFromMdlScp: {sht1TitleCopyFromMdlScp}")
    print(f"\tsht1模板标题至: sht1DeptTltRan: {sht1DeptTltRan}")
    print(f"\t数据列范围：sht1DataColRan: {sht1DataColRan}")


def printAutoParamSht2(sht2DeleteCopiedColScp, sht2TitleCopyFromMdlScp,
                       sht2DeptTltRan, sht2IndexCopyFromSvyScp, sht2IndexCopyTo):
    print(f"Sheet2 自Sht2模板自动获取的参数如下: ")
    print(f"\tsht2DeleteCopiedColScp: {sht2DeleteCopiedColScp}")
    print(f"\tsht2TitleCopyFromMdlScp: {sht2TitleCopyFromMdlScp}")
    print(f"\tsht2DeptTltRan: {sht2DeptTltRan}")
    print(f"\tsht2IndexCopyFromSvyScp: {sht2IndexCopyFromSvyScp}")
    print(f"\tsht2IndexCopyTo: {sht2IndexCopyTo}")


def printAutoParamSht3(sht3IndexCopyFromSvyScp, sht3DataColRan, sht3TitleCopyFromMdlScp):
    print(f"Sheet3 自Sht3模板自动获取的参数如下: ")
    print(f"\tsht3IndexCopyFromSvyScp: {sht3IndexCopyFromSvyScp}")
    print(f"\tsht3DataColRan: {sht3DataColRan}")
    print(f"\tsht3TitleCopyFromMdlScp: {sht3TitleCopyFromMdlScp}")


def printAutoParamSht4(sht4IndexFromSht2Scp, sht4TitleFromSht2Scp, sht4SumTitleFromMdlScp,
                       ):
    print(f"Sheet4 自Sht4模板自动获取的参数如下: ")
    print(f"\tsht4IndexFromSht2Scp: {sht4IndexFromSht2Scp}")
    print(f"\tsht4TitleFromSht2Scp: {sht4TitleFromSht2Scp}")
    print(f"\tsht4SumTitleFromMdlScp: {sht4SumTitleFromMdlScp}")
    # print(f"\tsht4DataRowRan: {sht4DataRowRan}")
