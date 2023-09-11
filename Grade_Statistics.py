import pandas as pd
import os


# 获取要合并的文件
def getFiles():
    global path, files, logs
    if __name__ == "__main__":
        while True:
            path = input("请输入成绩文件夹路径（留空默认程序文件夹下的成绩文件夹）：")
            if path == "":
                # 在当前目录下创建成绩文件夹
                if not os.path.exists("成绩"):
                    os.mkdir("成绩")
                path = os.path.join(os.getcwd(), "成绩")
            try:
                files = os.listdir(path)
            except FileNotFoundError:
                print("您未输入正确的路径,请重新输入")
            else:
                return True
    else:
        try:
            files = os.listdir(path)
        except FileNotFoundError:
            logs.append("您未输入正确的合并路径,请重新输入")
            return False
        else:
            return True


# 向用户确认文件
def confirmFiles():
    global files
    if __name__ == "__main__":
        print("请确认文件路径正确包含所需文件:", files)
        print("输入y继续，输入其他任意字符重新输入")
        while input().lower() != "y":
            path = input("请输入成绩文件夹路径（留空默认程序文件夹下的成绩文件夹）：")
            files = os.listdir(path)
            print("请确认文件路径正确包含所需文件:", files)


def tryRow(rowName: str, df: pd.DataFrame):
    global logs
    try:
        return df[df.iloc[:, 0] == rowName].index[0], df
    except IndexError:
        logs.append(f"未找到{rowName}行，将自动添加至底部")
        df.loc[len(df.index)] = [None] * (len(df.columns))
        df.loc[len(df.index)] = [rowName] + [None] * (len(df.columns) - 1)
        return df[df.iloc[:, 0] == rowName].index[0], df


def tryMerge(keyword: str, df: pd.DataFrame, data: pd.DataFrame):
    global logs, keyWord, f, column_name
    if keyword in df.columns and keyword in data.columns:
        df[keyword] = df[keyword].astype(str)
        data[keyword] = data[keyword].astype(str)
        logs.append(f"正在尝试合并「{column_name}」中的「{keyword}」列")
        try:
            df = pd.merge(df, data[[keyword, keyWord]], on=keyword, how="left")
        except ValueError:
            logs.append(f"出于无法确定的原因合并失败, 请查看说明文档和日志以获取帮助")
            return False, df
    else:
        logs.append(f"「{column_name}」中未找到{keyword}列")
        return False, df
    return True, df


def getFilewithKeywordsColumns(keywords: list, data: pd.DataFrame):
    global logs, path, files, column_name
    while True:
        for keyword in keywords:
            if keyword in data.columns:
                return True, data
        try:
            data.columns = data.iloc[0]
            data = data.iloc[1:].reset_index(drop=True)
        except IndexError:
            logs.append(f"「{column_name}」中未找到姓名或学号列，请检查文件")
            return False


def calCoordinate(x, y):
    dic = {
        1: "A",
        2: "B",
        3: "C",
        4: "D",
        5: "E",
        6: "F",
        7: "G",
        8: "H",
        9: "I",
        10: "J",
        11: "K",
        12: "L",
        13: "M",
        14: "N",
        15: "O",
        16: "P",
        17: "Q",
        18: "R",
        19: "S",
        20: "T",
        21: "U",
        22: "V",
        23: "W",
        24: "X",
        25: "Y",
        26: "Z",
    }
    xstr = ""
    while x > 0:
        xstr = dic[x % 26] + xstr
        x = x // 26

    return xstr + "2:" + xstr + str(y + 2)


template = "姓名模版.xlsx"
path = os.path.join(os.getcwd(), "成绩")
outputPath = ""
keyWord = "得分"
logs = []
gauge = 0
files = []
maxRow = True
meanRow = True
noneList = []


def Merge():
    global logs, path, outputPath, template, keyWord, gauge, files, maxRow, meanRow, noneList, column_name
    # 导入姓名模版
    try:
        df = pd.read_excel(template)
    except FileNotFoundError:
        logs.append("未找到姓名模版，请输入正确的文件路径")
        return False
    logs.append("姓名模版:" + template)
    # 检测姓名模版中是否有姓名或学号列
    flag, df = getFilewithKeywordsColumns(["姓名", "学号"], df)
    if not flag:
        logs.append("姓名模版中未找到姓名或学号列，请检查模版，请输入正确的文件路径")
        return False

    # 若path是一个excel文件
    if path.endswith((".xls", ".xlsx")):
        files_xls = [path]
        try:
            data = pd.read_excel(path)
        except FileNotFoundError:
            logs.append("未找到文件，请输入正确的文件路径")
            return False
    else:
        # 获取文件
        if not getFiles():
            return False
        logs.append("要合并的文件夹路径:" + path)
        # 确认文件
        confirmFiles()
        # 获取所有excel文件
        files_xls = [f for f in files if f.endswith((".xls", ".xlsx"))]
    if files_xls == []:
        logs.append("未找到任何excel文件,请检查文件夹路径是否正确")
        return False

    # 合并所有excel文件
    # print("将要合并的文件：", files_xls)
    logs.append("将要合并的文件：" + str(files_xls))
    numFiles = len(files_xls)
    # print("共" + str(numFiles) + "个文件")
    logs.append("共" + str(numFiles) + "个文件")
    num = 0
    logs.append(f"关键字是「{keyWord}」")
    for f in files_xls:
        column_name = f.split(".")[0]
        if not column_name in df.columns:
            # 逐行检索是否存在一个包含「姓名」或「学号」的行
            data = pd.read_excel(os.path.join(path, f))
            flag, data = getFilewithKeywordsColumns(["姓名", "学号"], data)
            if not flag:
                logs.append(f"「{column_name}」中未找到姓名或学号列，请检查文件")
                return False
            logs.append(f"已在{column_name}找到姓名或学号列")
            data.columns = [str(column) for column in data.columns]
            if keyWord in data.columns:
                flag, df = tryMerge("学号", df, data)
                if not flag:
                    flag, df = tryMerge("姓名", df, data)
                    if not flag:
                        return False
            else:
                logs.append(f"「{column_name}」中未找到「{keyWord}」列")
                return False

            # 重命名列
            df.rename(columns={keyWord: column_name}, inplace=True)

            # 调整进度条
            num += 1
            gauge = int(num / (numFiles + maxRow + meanRow) * 100)
            logs.append("已合并 " + column_name + str(num) + "/" + str(numFiles))
    # print("合并完成")
    logs.append("合并完成")

    # 将未上传、nan等部分替换为空
    for key in noneList:
        df.replace(key, None, inplace=True)
    # print("未上传已替换为空")
    logs.append("未上传已替换为空")

    # 计算最高分和平均分并导入
    if maxRow or meanRow:
        rowEndIndex = 0
        # 逐行搜索获取姓名列包含的数量
        if "学号" in df.columns:
            for index in df.index:
                if df.loc[index, "学号"] == None:
                    rowEndIndex = index - 1
                    break
            else:
                rowEndIndex = len(df.index) - 1
        elif "姓名" in df.columns:
            for index in df.index:
                if df.loc[index, "姓名"] == None:
                    rowEndIndex = index - 1
                    break
            else:
                rowEndIndex = len(df.index) - 1
        columnNum = 0
        dic = {}
        for column in df.columns:
            columnNum += 1
            if column == "学号" or column == "姓名" or column in noneList:
                continue
            dic[column] = calCoordinate(columnNum, rowEndIndex)

    if maxRow:
        maxIndex, df = tryRow("班级最高分", df)
        for column in dic:
            df.loc[maxIndex, column] = f"=MAX({dic[column]})"
        gauge += int(1 / (numFiles + maxRow + meanRow) * 100)
        logs.append("最高分已导入")
    if meanRow:
        meanIndex, df = tryRow("班级平均分", df)
        for column in dic:
            df.loc[meanIndex, column] = f"=AVERAGE({dic[column]})"
        gauge += int(1 / (numFiles + maxRow + meanRow) * 100)
        logs.append("平均分已导入")

    try:
        df.to_excel(os.path.join(outputPath, "成绩汇总.xlsx"), index=False)
    except OSError:
        df.to_excel("成绩汇总.xlsx", index=False)
        logs.append("未找到输出路径，已输出至同目录下的成绩汇总.xlsx")
        return False
    # print("已输出至同目录下的成绩汇总.xlsx")
    gauge += int(1 / (numFiles + maxRow + meanRow) * 100)
    logs.append(f"已输出至{outputPath}{os.sep}成绩汇总.xlsx")

    return True
