# GitHub: @GWillS163
# Start Time: 2022-08-18
# End Time:

from .getScoreData import *
from .shtDataCalc import *
from .shtOperation import *
from .userParamsProcess import *


class Excel_Operation:
    surveyRuleCol: str
    surveyQuesCol: str
    surveyQuesTypeCol: str
    sht4TitleFromSht2Scp: str

    def __init__(self, surveyExlPath, partyAnsExlPh, peopleAnsExlPh,
                 surveyTestShtName, sht1ModuleName, sht2ModuleName, sht3ModuleName, sht4ModuleName,
                 # ResultExl 配置
                 sht1Name, sht2Name, sht3Name, sht4Name,
                 # Sheet1 配置 : "F", "G"
                 sht1IndexScpFromSht0, sht1TitleCopyFromSttCol, sht1TitleCopyToSttCol,
                 # Sheet2 配置:  "C1:J1", "D"
                 sht2DeleteCopiedColScp, sht2MdlTltStt,
                 # Sheet3 配置:  "L", "J", "K"
                 sht3MdlTltStt, sht3SurLastCol, sht3ResTltStt,
                 # Sheet4 配置:
                 sht4IndexFromMdl4Scp, sht4TitleFromSht2Scp, sht4SumTitleFromMdlScp  # ,sht4DataRowRan
                 ):
        """
        param surveyExlPath: the path of the survey excel
        :param scrExlPh:
        """
        # Saving File Settings
        self.surveyWgtCol = "K"
        self.orgShtName = '行政机构组织'
        self.otherTitle = "其他人员"
        self.lv2AvgTitle = "二级单位成绩"
        paramsCheckExist(surveyExlPath, partyAnsExlPh, peopleAnsExlPh)

        # self.surveyExlPath = surveyExlPath
        # self.scoreExlPath = scrExlPh
        self.app4Survey1 = xw.App(visible=True, add_book=False)
        self.app4Result3 = xw.App(visible=True, add_book=False)
        self.app4Survey1.display_alerts = False
        self.app4Result3.display_alerts = False
        self.app4Survey1.api.CutCopyMode = False
        self.app4Result3.api.CutCopyMode = False
        # open the survey file with xlwings
        self.surveyExl = self.app4Survey1.books.open(surveyExlPath)
        self.resultExl = self.app4Result3.books.add()

        # 输入配置-规则表 修改后请核对以下配置
        paramsCheckSurvey(self.surveyExl, [surveyTestShtName, sht2ModuleName, sht3ModuleName, sht4ModuleName])
        self.sht0TestSurvey = self.surveyExl.sheets[surveyTestShtName]  # "测试问卷"
        # module settings
        self.sht1Mdl = self.surveyExl.sheets[sht1ModuleName]  # "调研结果_输出模板"
        self.sht2Mdl = self.surveyExl.sheets[sht2ModuleName]  # "调研成绩_输出模板"
        self.sht3Mdl = self.surveyExl.sheets[sht3ModuleName]  # "调研结果（2022年）_输出模板"
        self.sht4Mdl = self.surveyExl.sheets[sht4ModuleName]  # "调研成绩（2022年）_输出模板"
        # basic output settings 
        self.sht1NameRes = sht1Name
        self.sht2NameGrade = sht2Name
        self.sht3NameResYear = sht3Name
        self.sht4NameScoreYear = sht4Name

        # Module Sheet 0 TestSurvey
        self.deptCopyHeight = 35  # 从总表复制到 各个部门表时的高度
        self.lv1Name: str = "北京公司"
        self.sht0QuestionScp: str = "E3:E31"  # 调查结果表问题范围
        # TODO: 自动获取以下列字母？
        # sht0Res = autoGetSHt0Params()
        self.surveyQuesCol: str = "E"  # 调查规则表中的问题列
        self.surveyRuleCol: str = "J"  # 调查规则表中的规则列
        self.surveyQuesTypeCol: str = "G"  # 调查规则表中的问题类型列

        # Sheet1  每次更改模板表后之后请留意
        self.sht1IndexScpFromSht0 = sht1IndexScpFromSht0  # "A1:F31"
        sht1Res = autoGetSht1Params(self.sht1Mdl,
                                    sht1TitleCopyFromSttCol,
                                    sht1TitleCopyToSttCol)
        self.sht1TitleCopyFromMdlScp, self.sht1DeptTltRan, self.sht1DataColRan = sht1Res

        # Sheet2:
        self.sht2TitleCopyTo = None  # 自动获取sheet2 title 起始点
        sht2Res = autoGetSht2Params(self.sht2Mdl, self.sht0TestSurvey,
                                    sht2DeleteCopiedColScp,  # ="C1:J1"
                                    sht2MdlTltStt)  # ="D"
        self.sht2DeleteCopiedColScp, self.sht2TitleCopyFromMdlScp, self.sht2DeptTltRan,  self.sht2IndexCopyFromSvyScp, self.sht2IndexCopyTo = sht2Res

        # Sheet3:
        self.sht3TitleCopyTo = None  # 自动获取sheet3 title 起始点
        sht3Res = autoGetSht3Params(self.sht3Mdl, self.sht0TestSurvey, sht3MdlTltStt, sht3SurLastCol, sht3ResTltStt)  # "L", "J", "K"
        self.sht3IndexCopyFromSvyScp, self.sht3DataColRan, self.sht3TitleCopyFromMdlScp = sht3Res

        # Sheet4:
        self.sht4TitleCopyTo = None  # 自动获取sheet4 title 起始点
        sht4Res = autoGetSht4Params(self.sht4Mdl, sht4IndexFromMdl4Scp, sht4TitleFromSht2Scp, sht4SumTitleFromMdlScp)
        self.sht4IndexFromMdl4Scp, self.sht4TitleFromSht2Scp, self.sht4SumTitleFromMdlScp = sht4Res

        # When add Sheet 2
        self.lv2UnitColLtr = "B"
        self.weightColLtr = "K"
        self.sht2SttRow = 3
        self.sht2EndRow = 40

        # When generate department file
        self.sht1IndexScp4Depart = None  # 生成侧标题后 自动获取"A1:F32"  # 生成部门文件时，需要复制的范围
        self.sht2IndexScp4Depart = None  # 生成侧标题后 自动获取"A1:C20"  # 生成部门文件用，复制左侧栏
        self.sht1TitleCopyTo = None

    def addSheet1_surveyResult(self):
        """
        新增 Sheet1
        generate sheet1 of the surveyExl, which is the survey without weight
        :return:
        """
        # titleStart, r = self.sht1MdlTltScope.split(":")
        # dataStart = r + str(int(self.sht1MdlTltScope.split(":")[1][-1]) + 1)
        # Step1: add new sheet and define module sheet
        sht1_lv2Result = self.resultExl.sheets.add(self.sht1NameRes)

        # Step2.1: copy left column the surveySht to sht1 partially with style
        # UnitScpRowList = getUnitScpRowList(self.sht0TestSurvey, "A", "D", ["党建", "宣传文化"])

        print("Sheet1 Index Columns Processing...")
        shtCopyTo(self.sht0TestSurvey, self.sht1IndexScpFromSht0,
                  sht1_lv2Result, f"A1")
        # 记录左侧栏的范围 为生成部门文件时使用 - Record the left column range for generating department files
        self.sht1IndexScp4Depart = sht1_lv2Result.used_range

        # Step2.2: copy title
        print("Sheet1 Title Rows Processing...")
        self.sht1TitleCopyTo = getLastColCell(sht1_lv2Result)
        print("Sheet1 title Rows Processing...")
        shtCopyTo(self.sht1Mdl, self.sht1TitleCopyFromMdlScp,
                  sht1_lv2Result, self.sht1TitleCopyTo)  # sht1_tltScope = "F1:KZ2"
        return sht1_lv2Result

    # def addSheet1_singleDepartment(self, deptUnitScp):
    #     """
    #     add a single department to the sht1_lv2Result
    #     :param deptUnitScp:
    #     :return:
    #     """
    #     sht1_lv2Result = self.deptResultExl.sheets.add(self.sht1NameRes)
    #
    #     # Step2.1: copy left column the surveySht to sht1 partially with style
    #     shtCopyTo(self.surveyTestSht, self.sht1CopyFromMdlColScp,
    #               sht1_lv2Result, self.sht1CopyFromMdlColScp.split(":")[0])
    #     shtCopyTo(self.sht1Module, deptUnitScp,
    #               sht1_lv2Result, self.sht1TitleCopyTo)  # sht1_tltScope = "F1:KZ2"
    #
    #     return sht1_lv2Result

    def addSheet2_surveyGrade(self):
        """
        生成sheet2的成绩表无数据，仅有二级标题，新增二级求和与全部求和。
        generate sheet2 of the surveyExl, only with title and sum
        :return: None
        """
        # Step1: 新增Sheet2 - add new sheet and define module sheet
        sht2_lv2Score = self.resultExl.sheets.add(self.sht2NameGrade, after=self.sht1NameRes)

        print("Sheet2 表头和侧栏部分处理 - header and sidebar section")
        sht2_lv2Score.range("B1").column_width = 18.8
        print("Step2: 整体复制侧栏 - copy left column the surveySht to sht1, with style")
        shtCopyTo(self.sht0TestSurvey, self.sht2IndexCopyFromSvyScp, sht2_lv2Score, self.sht2IndexCopyTo)
        print("Step2.1 删除多余行(党廉&纪检) - delete the row of left column redundant")
        sht2DeleteCopiedRowScp = getSht2DeleteCopiedRowScp(sht2_lv2Score, ['党廉', '纪检'])
        sht2_lv2Score.range(sht2DeleteCopiedRowScp).api.EntireRow.Delete()

        print("Step2.2: 删除左侧多余单元行 - delete the row of left column redundant")
        print("Step2.2.1: 删除多余行 - delete the row of left column redundant")
        sht2UnitScp = getShtUnitScp(sht2_lv2Score, self.sht2SttRow, self.sht2EndRow, self.lv2UnitColLtr, "D")
        print("Step2.2.2: 为单元格重新计算权重 - recalculate the weight to the cell")
        # 给所有单元格边缘增加偏移 - add offset to all cells edge
        sht2UnitScpOffsite = [[edge + self.sht2SttRow for edge in eachUnit] for eachUnit in sht2UnitScp]
        resetUnitSum(sht2_lv2Score, sht2UnitScpOffsite, self.weightColLtr)
        print("Step2.2.3: 重设单元值完成 - Reset unit value completed")
        sht2DeleteRowLst = getSht2DeleteRowLst(sht2UnitScpOffsite)
        for _ in range(len(sht2DeleteRowLst)):
            row = sht2DeleteRowLst.pop()
            sht2_lv2Score.range(f"{self.lv2UnitColLtr}{row}").api.EntireRow.Delete()
        print("Step2.3 删除多余列 - delete the column C to I")
        sht2_lv2Score.range(self.sht2DeleteCopiedColScp).api.EntireColumn.Delete()

        # 记录左侧栏的范围 为生成部门文件时使用 - Record the left column range for generating department files
        self.sht2IndexScp4Depart = sht2_lv2Score.used_range

        print("Step3: 整体复制title - copy title")
        # get last column of the sht2
        self.sht2TitleCopyTo = getLastColCell(sht2_lv2Score)
        shtCopyTo(self.sht2Mdl, self.sht2TitleCopyFromMdlScp, sht2_lv2Score, self.sht2TitleCopyTo)

        print("Step4: 增加合计行 -  add summary row")
        lv2UnitScp = getShtUnitScp(sht2_lv2Score, self.sht2SttRow, self.sht2EndRow, "A", "B")
        sht2OprAddSummaryRows(sht2_lv2Score, lv2UnitScp)

        return sht2_lv2Score, lv2UnitScp

    # def addSheet2_singleDepartment(self, deptUnitScp):
    #     """
    #     add a single department to the sht2_lv2Score
    #     重复操作，可以考虑合并
    #     :param deptUnitScp:
    #     :return:
    #     """
    #     sht2_lv2Score = self.deptResultExl.sheets.add(self.sht2NameGrade, after=self.sht1NameRes)
    #
    #     """Sheet2 表头和侧栏部分处理"""
    #     # Step2: copy left column the surveySht to sht1, with style
    #     shtCopyTo(self.surveyTestSht, self.sht2IndexCopyFromColScp, sht2_lv2Score, self.sht2IndexCopyTo)
    #     # Step2.1: delete the row of left column redundant
    #     for row in self.sht2DeleteRowLst:
    #         sht2_lv2Score.range(f"B{row}").api.EntireRow.Delete()
    #     sht2_lv2Score.range("B1").column_width = 18.8
    #     # Step2.2 delete the column C to I
    #     sht2_lv2Score.range(self.sht2DeleteCopiedColScp).api.EntireColumn.Delete()
    #     sht2_lv2Score.range(self.sht2DeleteCopiedRowScp).api.EntireRow.Delete()
    #
    #     # Step3: copy title
    #     shtCopyTo(self.sht2Module, deptUnitScp, sht2_lv2Score, self.sht2TitleCopyTo)
    #     # Step2.3 add summary row
    #     sht2OprAddSummaryRows(sht2_lv2Score)
    #
    #     return sht2_lv2Score

    def addSheet3_surveyResultByYear(self):
        """
        generate sheet3 of the surveyExl, which is the survey without weight
        :return: None
        """

        # Step1: add new sheet and define module sheet
        sht3_ResYear = self.resultExl.sheets.add(self.sht3NameResYear, after=self.sht2NameGrade)

        # Step2: Set Index column
        shtCopyTo(self.sht0TestSurvey, self.sht3IndexCopyFromSvyScp, sht3_ResYear,
                  self.sht3IndexCopyFromSvyScp.split(":")[0])

        # Step3: copy title
        self.sht3TitleCopyTo = getLastColCell(sht3_ResYear)
        shtCopyTo(self.sht3Mdl, self.sht3TitleCopyFromMdlScp, sht3_ResYear, self.sht3TitleCopyTo)

        return sht3_ResYear

    def addSheet4_surveyGradeByYear(self):
        """
        generate sheet4 of the surveyExl, which is the survey with weight
        :return: None
        """
        print("Step1: 新增Sheet4 - add new sheet")
        sht4_surveyGradeByYear = self.resultExl.sheets.add(self.sht4NameScoreYear, after=self.sht3NameResYear)
        sht2_lv2Score = self.resultExl.sheets[self.sht2NameGrade]

        print("Step2: 粘贴问卷单元title - copy title of survey")
        sht2_lv2Score.api.Range(self.sht4TitleFromSht2Scp).Copy()
        sht4_surveyGradeByYear.api.Range("A1").PasteSpecial(Transpose=True)

        print("Step3: 复制模板的左侧线条公司栏 - copy left column of the template")
        sht4LineCompanyCopyTo = getLastRowCell(sht4_surveyGradeByYear)
        shtCopyTo(self.sht4Mdl, self.sht4IndexFromMdl4Scp,
                  sht4_surveyGradeByYear, sht4LineCompanyCopyTo)

        print("Step4: 粘贴汇总title - summarize Title patch")
        self.sht4TitleCopyTo = getLastColCell(sht4_surveyGradeByYear)
        shtCopyTo(self.sht4Mdl, self.sht4SumTitleFromMdlScp,
                  sht4_surveyGradeByYear, self.sht4TitleCopyTo)
        # set column B width = 20
        sht4_surveyGradeByYear.range("B1").column_width = 13
        return sht4_surveyGradeByYear

    def genDepartFile(self, departCode, sumSavePathNoSuffix):
        """ 通过汇总表生成所有部门文件, 生成固定的标题侧栏，然后填充数据-保存-清空数据
        generate the department file of the surveyExl， generate the fixed title sidebar, then fill data-save-clear data
        :return: None
        """
        # 从汇总表 获取 部门分类区间
        sht1Sum = self.resultExl.sheets[self.sht1NameRes]
        sht2Sum = self.resultExl.sheets[self.sht2NameGrade]
        deptUnitSht1 = getDeptUnit(sht1Sum, self.sht1DeptTltRan, 0)
        deptUnitSht2 = getDeptUnit(sht2Sum, self.sht2DeptTltRan, 0)

        print("新建部门文件 - create new excel with xlwings")
        app4Depart = xw.App(visible=True, add_book=False)
        deptResultExl = app4Depart.books.add()
        app4Depart.display_alerts = False
        app4Depart.api.CutCopyMode = False
        sht1Dept = deptResultExl.sheets.add(self.sht1NameRes)
        sht2Dept = deptResultExl.sheets.add(self.sht2NameGrade, after=sht1Dept)
        try:
            deptResultExl.sheets["Sheet1"].delete()
        except:
            pass
        print("从汇总表 复制 边栏")
        # sht1Dept.activate()
        shtCopyTo(sht1Sum, self.sht1IndexScp4Depart, sht1Dept, "A1")
        # sht2Dept.activate()
        shtCopyTo(sht2Sum, self.sht2IndexScp4Depart, sht2Dept, "A1")

        # 填充数据 - 保存 - 删除
        sht2BorderL, sht2BorderR = "G", "Z"
        departStt = time.time()
        n = 0
        for deptName in deptUnitSht1:
            n += 1
            # add department data
            sht1BorderL, sht1BorderR = addOneDptData(sht1Sum, deptUnitSht1[deptName], self.deptCopyHeight,
                                                     sht1Dept, self.sht1TitleCopyTo)
            if deptName in deptUnitSht2:
                sht2BorderL, sht2BorderR = addOneDptData(sht2Sum, deptUnitSht2[deptName], self.deptCopyHeight,
                                                         sht2Dept, self.sht2TitleCopyTo)

            # Save Excel 保存文件名需要 加上部门Code
            deptCode = ""
            try:
                deptCode = departCode[deptName]['departCode']
            except:
                pass
            # departFilePath = self.sumSavePath.replace(".xlsx", f"_{deptName}_{deptCode}.xlsx")
            departFilePath = sumSavePathNoSuffix + f"_{deptCode}.xlsx"
            print(f"\033[32m{n:2} - 正在保存:[{deptCode}] - [{int(time.time() - departStt)}s] {deptName} - \033[0m")
            deptResultExl.save(departFilePath)
            departStt = time.time()

            # Method1: Clear the sheet for next loop dynamically
            # sht1Dept.activate()
            # dltOneDptData(sht1Dept, self.sht1TitleCopyTo, self.deptCopyHeight,
            #               sht1BorderR, sht1BorderL)
            # if deptName in deptUnitSht2:
            #     sht2Dept.activate()
            #     dltOneDptData(sht2Dept, self.sht2TitleCopyTo, self.deptCopyHeight,
            #                   sht2BorderL, sht2BorderR)

            # Method2
            borderWidth = getColNum(sht1BorderR) - getColNum(sht1BorderL)
            borderStart = getColNum(self.sht1TitleCopyTo[0])
            borderEnd = getColLtr(borderStart + borderWidth)
            sht1Dept.activate()
            sht1Dept.range(f"{self.sht1TitleCopyTo}:{borderEnd}{self.deptCopyHeight}").api.Delete()
            if deptName in deptUnitSht2:
                borderWidth = getColNum(sht2BorderR) - getColNum(sht2BorderL)
                borderStart = getColNum(self.sht2TitleCopyTo[0])
                borderEnd = getColLtr(borderStart + borderWidth)
                sht2Dept.activate()
                sht2Dept.range(f"{self.sht2TitleCopyTo}:{borderEnd}{self.deptCopyHeight}").api.Delete()

        print("部门文件全部生成完毕，即将关闭文件")
        deptResultExl.close()
        app4Depart.quit()

        self.resultExl.close()
        self.app4Result3.quit()

    def placeBar(self) -> list:
        # Add Title & column
        sht1_lv2Result = self.addSheet1_surveyResult()
        print("sheet1 无数据页面生成完成")
        sht2_lv2Score, lv2UnitScp = self.addSheet2_surveyGrade()
        print("sheet2 无数据页面生成完成")
        sht3_ResYear = self.addSheet3_surveyResultByYear()
        print("sheet3 无数据页面生成完成")
        sht4_surveyGradeByYear = self.addSheet4_surveyGradeByYear()
        print("sheet4 无数据页面生成完成")
        return [sht1_lv2Result, sht2_lv2Score, sht3_ResYear, sht4_surveyGradeByYear, lv2UnitScp]

    def fillAllData(self, sht1WithLv, shtList, departCode: dict, sumSavePathNoSuffix: str):
        """
        使用sheet1的数据计算出来接下来的数据，然后填充到excel中, 获得汇总的1 个文件
        use sheet1 data to calculate the data of the next sheet, then fill it into the excel
        :param sumSavePathNoSuffix: 
        :param departCode: 
        :param shtList: sht1_lv2Result, sht2_lv2Score, sht3_ResYear, sht4_surveyGradeByYear 
        :param sht1WithLv: 
        :return: 
        """""
        sht1_lv2Result, sht2_lv2Score, sht3_ResYear, sht4_surveyGradeByYear, lv1UnitScp = shtList
        # Step3: Set data
        if sht1WithLv:
            sht1WithLv = sht1WithLv  # mock data, will be deleted

        print("sheet1 填充数据 - Sheet1 fill data vertically")
        sht1_lv2Result.activate()
        # two Methods to set data
        # sht1PlcScoreByPD(self.sht1Module, sht1_lv2Result, staffWithLv, self.sht1MdlTltScope, dataStart)
        printShtWithLv("sht1WithLv：", sht1WithLv)
        sht1SetData(sht1_lv2Result, sht1WithLv, getTltColRange(self.sht1DataColRan, 1))

        print("sheet2 填充数据 - sheet2 fill data vertically")
        sht2_lv2Score.activate()
        sht2WithLv = clacSheet2_surveyGrade(sht1_lv2Result, self.sht0TestSurvey, self.surveyWgtCol, sht1WithLv, lv1UnitScp, departCode)
        print("sht2WithLv：", sht2WithLv)
        sht2SetData(sht2_lv2Score, sht2WithLv, getTltColRange(self.sht2TitleCopyFromMdlScp, 1))

        print("sheet3 填充数据 - sheet3 fill data vertically")
        sht3_ResYear.activate()
        sht3WithLv = getSht3WithLv(sht1WithLv, self.lv1Name)
        print("sht3WithLv：", sht3WithLv)
        sht3SetData(sht3_ResYear, sht3WithLv, self.sht3DataColRan, self.lv1Name)
        # sht3 set conditional formatting partially

        print("sheet4 填充横向数据 - Sheet4 fill data horizontally")
        sht4_surveyGradeByYear.activate()
        sht4Hie = getSht4Hierarchy(sht4_surveyGradeByYear)
        sht4WithLv = getSht4WithLv(sht2WithLv, sht4Hie, self.lv1Name)
        print("sht4WithLv：", sht4WithLv)
        # Get last row Num by used_range
        sht4LastRow = sht4_surveyGradeByYear.used_range.last_cell.row
        sht4DataRowRan = range(4, sht4LastRow + 1)
        sht4SetData(sht4_surveyGradeByYear, sht4WithLv, sht4DataRowRan, self.lv1Name)

        print("删除无用的sheet - Delete useless sheet")
        self.resultExl.sheets["Sheet1"].delete()
        savePath = sumSavePathNoSuffix + "_200000000.xlsx"
        print(f"汇总表将保存在：{savePath}")
        self.resultExl.save(savePath)
        self.surveyExl.close()
        self.app4Survey1.quit()

    def getFirstSht1WithLvData(self, partyAnsExlPh, peopleAnsExlPh, savePath):
        """
        获取第一个sheet1的数据，用于计算其他sheet的数据， 使用后关闭答案
        :return:
        """
        print("开始获取第一个sheet1的数据")
        debugPath = os.path.join(savePath, "分数判断记录")
        judges = scoreJudgement(self.sht0TestSurvey, self.otherTitle, self.surveyQuesCol, self.surveyRuleCol, self.surveyQuesTypeCol)
        questionLst = self.sht0TestSurvey.range(self.sht0QuestionScp).value  # 问题列表
        print("开始获取群众数据 - Start to get people data")
        surveyData = getSurveyData(self.sht0TestSurvey)
        peopleQuesLst, sht1PeopleData = judges.getStaffData(peopleAnsExlPh, "群众", debugPath, surveyData, True)
        printSht1Data("群众", sht1PeopleData, peopleQuesLst)
        print("开始获取党员数据 - Start to get party data")
        partyQuesLst, sht1PartyData = judges.getStaffData(partyAnsExlPh, "党员", debugPath, surveyData, True)
        printSht1Data("党员", sht1PartyData, partyQuesLst)
        questionSortDebugPath = os.path.join(savePath, "题目对应记录")
        sht1WithLvCombine = combineMain(questionLst, peopleQuesLst, sht1PeopleData, partyQuesLst, sht1PartyData, questionSortDebugPath)
        judges.close()
        printSht1WithLv(sht1WithLvCombine)
        return sht1WithLvCombine

    def run(self, partyAnsExlPh, peopleAnsExlPh, outputDir, sumSavePathNoSuffix):
        """
        主程序
        run the whole process, the main function.
        :return:
        """

        stt = time.time()
        departCode = getAllOrgCode(self.surveyExl.sheets(self.orgShtName))
        print("一、获取答题数据，开始判分 - 0. get data of score sheet, start to calculate score")
        sht1WithLvCombine = self.getFirstSht1WithLvData(partyAnsExlPh, peopleAnsExlPh, outputDir)
        print(f"\n\033[33m\nGetScore time: {int(time.time() - stt)}s \033[0m")

        print("\n二、填充边栏 - II. fill the sidebar")
        shtList = self.placeBar()
        print(f"\n\033[33m\nPlaceBar time: {int(time.time() - stt)}s \033[0m")

        print("\n三、填充数据、生成汇总文件 - III. fill data, generate summary file")
        self.fillAllData(sht1WithLvCombine, shtList, departCode, sumSavePathNoSuffix)
        print(f"\n\033[33m\nFillData time: {int(time.time() - stt)}s \033[0m")

        print("\n四、生成各部门文件 - IV. generate each department file")
        self.genDepartFile(departCode, sumSavePathNoSuffix)

        print(f"\n\033[33m\nTotal time: {int(time.time() - stt)}s [All Done! Saved to: {outputDir}]\033[0m")

