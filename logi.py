import pandas as pd
from tkinter import Tk, filedialog
import openpyxl

def select_files():
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_paths = filedialog.askopenfilenames(
        title="选择多个Excel文件",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    return list(file_paths)

def adjust_column_widths(sheet):
    for column in sheet.columns:
        column_letter = column[0].column_letter  # 获取列的字母
        # 设置第三列的宽度为5.13
        if column_letter == "C":
            sheet.column_dimensions[column_letter].width = 5.13
        else:
            max_length = 0
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = max_length + 2  # 增加一些缓冲空间
            sheet.column_dimensions[column_letter].width = adjusted_width

def merge_and_remove_duplicates(file_paths, reference_file, sheet_name="Sheet1"):
    # 读取参考文件的C列数据
    try:
        reference_df = pd.read_excel(reference_file, sheet_name=sheet_name, engine="openpyxl", dtype=str)
        reference_values = reference_df.iloc[:, 2].dropna().astype(str).str.strip().unique()  # 获取C列值并去重
        print("参考文件的C列值示例:")
        print(reference_values[:10])  # 输出前10个参考值供检查
    except Exception as e:
        print(f"读取参考文件失败: {e}")
        return

    combined_df = pd.DataFrame()

    for file in file_paths:
        try:
            df = pd.read_excel(file, engine="openpyxl", dtype=str)
            combined_df = pd.concat([combined_df, df], ignore_index=True)
        except Exception as e:
            print(f"读取文件 {file} 失败: {e}")

    print("合并后的数据预览:")
    print(combined_df.head())  # 输出合并后的前几行数据
    print(f"合并后的数据行数: {combined_df.shape[0]}")

    # 确保合并数据的第三列为字符串并清理空格
    if combined_df.shape[1] > 2:
        combined_df.iloc[:, 2] = combined_df.iloc[:, 2].astype(str).str.strip()

    # 删除基于合并后第三列的值，若其在参考文件的C列中存在
    initial_rows = combined_df.shape[0]
    filtered_df = combined_df[~combined_df.iloc[:, 2].isin(reference_values)]
    print(f"在参考文件中找到匹配的行，已删除 {initial_rows - filtered_df.shape[0]} 行")

    # 只保留第一列、第三列、第10列、第11列和第25列（注意索引是从0开始）
    columns_to_keep = [0, 2, 9, 10, 24]  # 第1列、第3列、第10列、第11列和第25列的索引
    filtered_df = filtered_df.iloc[:, columns_to_keep]

    # 删除重复行（基于新的第二列值）
    filtered_df = filtered_df.drop_duplicates(subset=filtered_df.columns[1])  # 基于第二列去重
    print(f"去重后的数据行数: {filtered_df.shape[0]}")

    # 插入一个新列到第二列后，并命名为“函数”
    filtered_df.insert(2, '函数', '')  # 在第二列之后插入新列

    # 保存结果并进行替换
    output_file = "合并并去重后的结果.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        sheet = writer.sheets["Sheet1"]

        # 在C1写入“函数”
        sheet["C1"] = "函数"

        # 替换D列中的 "00:00:00" 为空白并替换 "-" 为 "/"
        for row in range(2, sheet.max_row + 1):
            cell = sheet[f"D{row}"]
            if "00:00:00" in str(cell.value):
                cell.value = str(cell.value).replace("00:00:00", "").strip()
            if "-" in str(cell.value):
                cell.value = str(cell.value).replace("-", "/")

            # 尝试将值转换为日期对象
            try:
                date_obj = pd.to_datetime(cell.value, format='%Y/%m/%d', errors='coerce')
                if pd.notnull(date_obj):
                    cell.value = date_obj.date()  # 只保留日期部分
                    cell.number_format = 'yyyy/mm/dd'  # 设置为日期格式
            except Exception as e:
                print(f"日期转换失败: {e}")

        # 在C列添加公式，从C2开始
        for row in range(2, sheet.max_row + 1):
            cell = sheet[f"C{row}"]
            cell.value = f'=IF(OR(LEFT(B{row},1)="2",LEFT(B{row},1)="3"),"不要",IF(LEFT(B{row},1)="4","黑猫","投函"))'

        # 调整列宽
        adjust_column_widths(sheet)

    print(f"合并并去重并处理后的文件已保存到: {output_file}")

if __name__ == "__main__":
    reference_file_path = "Z:\\RS\\邮局数据处理\\LOGI 数据提取汇总.xlsx"
    files = select_files()
    if files:
        merge_and_remove_duplicates(files, reference_file_path)
    else:
        print("未选择文件。")
