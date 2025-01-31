# Github: GWillS163
# User: 駿清清 
# Date: 28/11/2022 
# Time: 12:20
import os


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
        found = False
        workDir = workPath

        # 支持在当前目录、上级目录和上上级目录查找
        for level in range(3):  # 当前目录 + 上级目录 + 上上级目录
            fullPath = os.path.join(workDir, pyFilePath)
            if os.path.isfile(fullPath):
                res = fullPath
                found = True
                break
            # 移动到上一级目录
            workDir = os.path.dirname(workDir)

        if found:
            break
    return res


if __name__ == '__main__':
    # print(findFromOrigin("os", r"D:\Project\mergeMultiPyFiles\examples\realProject1"))
    # print(findFromOrigin("core.main", r"D:\Project\mergeMultiPyFiles\examples\realProject1"))
    # print(findFromOrigin("core.config", r"D:\Project\mergeMultiPyFiles\examples\realProject1"))
    # print(findFromOrigin("realTest", r"D:\Project\mergeMultiPyFiles\examples\realProject1"))
    print(findFromOrigin(".utils", r"D:\Project\mergeMultiPyFiles\examples\realProject1\core"))

