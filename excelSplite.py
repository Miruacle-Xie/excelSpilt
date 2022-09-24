import pandas as pd
import os


def main():
    filePath = input("输入文件路径：\n")
    filePath = filePath.replace("\"", "").replace("\'", "")
    dirPath = os.path.splitext(filePath)
    splitNum = input("输入分隔数量：\n")
    splitNum = eval(splitNum)
    df = pd.read_excel(filePath)
    # print(df)
    tmpNum = len(df) // splitNum
    if tmpNum != 0:
        for i in range(tmpNum + 1):
            if i != tmpNum:
                tmpath = dirPath[0] + "(" + str(i * splitNum + 1) + "-" + str((i + 1) * splitNum) + ")" + dirPath[1]
                df_split = df.iloc[i * splitNum: (i + 1) * splitNum, :]
            else:
                tmpath = dirPath[0] + "(" + str(tmpNum * splitNum + 1) + "-" + str(len(df)) + ")" + dirPath[1]
                df_split = df.iloc[tmpNum * splitNum + 1: len(df), :]
            # print(df_split)
            writer = pd.ExcelWriter(tmpath)
            df_split.to_excel(writer, index=False)
            writer.save()
            writer.close()
            print(tmpath)
        input("分隔完成，按回车结束")
    else:
        input("总数小于分隔数量，无需分隔，按回车结束")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        repr(e)
        input("\n出现异常, 请确保表格格式为xlsx")
