import pandas
import os
def merge_excels(excels_path,output_path):
    """
    merge excels\n
    :excels_path : 列表，多个excel路径
    :output_path ：合并后输出的路径
    """
    df_list = []
    for excel_path in excels_path:
        if not os.path.exists(excel_path):
            continue

        excel_df = pandas.read_excel(excel_path)
        df_list.append(excel_df)

    df_all = pandas.concat(df_list)
    df_all.to_excel(output_path,index=False)


if __name__ =="__main__":
    excels_path = [r"C:/Users/Torre/Documents/Coding/python/Test/pyqt/dist\2.xlsx",
                    r"C:\Users\Torre\Documents\Coding\python\Test\pyqt\dist\1.xlsx"]
    merge_excels(excels_path,"merge.xlsx")


    