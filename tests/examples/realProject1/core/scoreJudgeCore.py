#  Author : Github: @GWillS163
#  Time: $(Date)
import re


def judgeGradeSingle(answerIntLst, ruleSelect, ruleScore):
    # 1分:1
    # 2.2分
    # 9:9分
    # 10.10分
    if len(answerIntLst) != 1:
        # print("识别到的数字不止一个，请检查！", end="")
        # print(debugPrint)
        return -100

    # get the number in ruleSelect by regex
    numOfRuleDig = re.search(r"(\d{,3})([-|或])?(\d{,3})?", ruleSelect).groups()
    if not numOfRuleDig[1]:  # only one number
        if answerIntLst[0] == int(numOfRuleDig[0]):
            return ruleScore
    else:  # present scope or "或"
        if numOfRuleDig[1] == "或":
            if answerIntLst[0] == int(numOfRuleDig[0]) or answerIntLst[0] == int(numOfRuleDig[2]):
                return ruleScore
        elif numOfRuleDig[1] == "-":
            if int(numOfRuleDig[0]) <= answerIntLst[0] <= int(numOfRuleDig[2]):
                return ruleScore

    return -100


def judgeGradeMulti(answerIntLst, ruleSelect, ruleScore):
    # 已覆盖效果 (2022-8-29)
    # 单个选择: 10分：4
    # 单个任意: 10分：4或5
    # 全部选择: 8分：全选
    # 局部全选: 8分：1-4
    # 局部全选: 0分：1-4全选
    # 全局多选: 任选3个
    # 局部多选: 1-4任选3个
    # 局部多选以上: 1-4任选3个及以上

    # 识别效果(下行为标题):
    # 单选，4或5, 选择范围开始, 选择范围结束, 全选，任选，n个，n个及以上
    # ('4', None, None, None, None, None, None, None)
    # (None, '4或5', None, None, None, None, None, None)
    # (None, '4或5或20', None, None, None, None, None, None)
    # (None, None, None, None, '全选', None, None, None)
    # (None, None, '1', '4', None, None, None, None)
    # (None, None, '1', '4', '全选', None, None, None)
    # (None, None, None, None, None, '任选', '3', None)
    # (None, None, '1', '4', None, '任选', '3', None)
    # (None, None, '1', '4', None, '任选', '3', '及以上')

    # write a regex to identify above case
    scopeRan = re.search(
        r"(^\d{,3})$|(^\d{,3}(?:或\d{,3})+$)|(?:^(\d{,3})-(\d{,3}))?(?:(全选)|(任选)(\d{,3})-?(\d{,3})?个)?(及以上)?",
        ruleSelect).groups()
    sglSlt, dblSlts, permitNumStart, permitNumEnd, isAllSlt, isAnySlt, permitQtyMin, permitQtyMax, isMore = scopeRan

    # 判断是否在允许范围 judge each score if is not in permit scope
    if not permitNumStart and not permitNumEnd:
        permitNumStart, permitNumEnd = 1, 99  # default
        validAns = answerIntLst
    else:
        validAns = []
        permitScp = range(int(permitNumStart), int(permitNumEnd) + 1)
        for num in answerIntLst:
            if int(num) in permitScp:
                validAns.append(num)
        validAns = list(set(validAns))

    # e.g. 单个选择的规则: 4
    if sglSlt and len(validAns) == 1:
        if validAns[0] == int(sglSlt):
            return ruleScore

    # 单个任选: e.g. 4或5, 一旦选了就给分
    elif dblSlts:
        permitNumLst = list(map(int, dblSlts.split("或")))
        # check some element of the answerIntLst is duplicate with permitNumLst
        for num in validAns:
            if int(num) in permitNumLst:
                return ruleScore

    # 全部选择与部分全选:
    # e.g. 全选, 1-5, 1-5全选  (默认1-99)
    elif isAllSlt or (not isAllSlt and not isAnySlt):
        # return "多选 或部分全选"
        if not permitNumStart and permitNumEnd:
            print("全部选择与部分全选，需要指定范围！")
            return -100
        # 选择的数量与范围一致
        if len(validAns) == len(range(int(permitNumStart), int(permitNumEnd) + 1)):
            return ruleScore

    # 局部  (默认1-99)
    # 全局多选: 任选3个
    # 局部多选: 1-4任选3个
    # 局部多选: 1-4任选3-4个
    # 局部多选以上: 1-4任选3个及以上
    elif isAnySlt:
        # 判断数量是否足够
        # 匹配两种格式： 任选n(-m)个及以上
        if not permitQtyMax:
            permitQtyMax = permitQtyMin

        # 任选n个及以上
        if isMore:
            if len(validAns) >= int(permitQtyMin):
                return ruleScore
        # 任选n(-m)个
        else:
            if int(permitQtyMin) <= len(validAns) <= int(permitQtyMax):
                return ruleScore

    return -100


def judgeRulePreProcess(ruleP, debugPrint):
    ruleL, ruleR = 0, 0
    for char in [':', "：", "."]:
        if char in ruleP:
            ruleL, ruleR = ruleP.split(char)
            break

    if not (ruleL and ruleR):
        print("没有找到匹配的分隔字符，将处理为-100分 :")
        print(debugPrint)
        return -100, -100

    # get permit scope
    if "分" in ruleL:
        ruleScore = int(ruleL.strip("分"))
        ruleSelect = ruleR
    elif "分" in ruleR:
        ruleScore = int(ruleR.strip("分"))
        ruleSelect = ruleL
    else:
        print("无法分辨哪边是分数! 将处理为-100分 :", debugPrint)
        return -100, -100

    return ruleScore, ruleSelect


def judgeAnswerGrade(answer, rule, quesType):
    debugPrint = f"\tanswer: {answer}\n" \
                 f"\trule: {rule}\n" \
                 f"\tquesType: {quesType}"
    # 如果能直接匹配分数则返回 - match the number of digital in answer by regex, form is a number+分
    answerInt = re.findall(r"(\d{,3}?)分", answer.strip())
    if answerInt and answerInt[0]:
        return int(answerInt[0])

    if "开放题" in quesType:
        return 0
    elif "评分题" in quesType:
        return int(re.search(r"\d{0,3}", answer.strip()).group(0))

    # get RuleL & RuleR and verify
    answerIntLstDup = list(map(int, re.findall(r"(\d{,3})\..*?", answer)))
    answerIntLst = list(set(answerIntLstDup))
    # check each rule to answer
    for ruleP in rule.split("\n"):
        if not ruleP:
            continue
        ruleScore, ruleSelect = judgeRulePreProcess(ruleP, debugPrint)
        if ruleScore == -100 or ruleSelect == -100:
            continue

        # get answer scope with different type of ques
        score = -100
        if "单项" in quesType or "单选" in quesType:
            score = judgeGradeSingle(answerIntLst, ruleSelect, ruleScore)
        elif "不定项" in quesType:
            score = judgeGradeMulti(answerIntLst, ruleSelect, ruleScore)
        if score != -100:
            return score
    return -100
