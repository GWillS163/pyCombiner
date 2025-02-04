
def combine_import_lines(restImportModules):
    """
    将import语句合并
    import文を結合
    Combine import statements

    :param restImportModules: 未合并的import语句
                              未結合のimport文
                              Uncombined import statements
    :return: 合并后的import语句
             結合後のimport文
             Combined import statements
    """
    combineRestImportLines = list(set(restImportModules))
    combineRestImportLines.sort(key=lambda x: len(x))
    return combineRestImportLines
