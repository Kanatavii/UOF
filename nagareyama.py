import win32com.client as win32
from datetime import datetime
import pandas as pd
import os
import time

# 获取当前日期并设置文件名
current_datetime = datetime.now().strftime("%Y-%m-%d-%H%M")
excel_output_path = f'Z:\\UOF\\转运数据\\许可\\流山入库-{current_datetime}.xlsx'
pdf_output_path = f'Z:\\UOF\\转运数据\\许可\\流山入库-{current_datetime}.pdf'

# 读取 Excel 文件
file_path = r'Z:\UOF\转运数据\JJS出入库汇总表.xlsx'
df = pd.read_excel(file_path)

# 筛选第三列非空白的行
filtered_df = df[df.iloc[:, 2].notna()]

# 再筛选第四列为空白的行
final_df = filtered_df[filtered_df.iloc[:, 3].isna()]

# 只保留指定的列
columns_to_keep = ["许可时间", "回数", "送り状番号", "箱数", "转运公司", "转运备注", "现场用-函数对应", "入库时间", "取件地"]
final_df = final_df[columns_to_keep]

# 保存筛选后的数据为新的 Excel 文件
final_df.to_excel(excel_output_path, index=False)

# 使用 pywin32 控制 Excel 应用程序
excel = win32.DispatchEx('Excel.Application')  # 使用 DispatchEx 代替 Dispatch
time.sleep(1)  # 确保应用程序初始化完成
wb = excel.Workbooks.Open(excel_output_path)
ws = wb.Worksheets(1)

# 设置列宽并为第一列指定短日期格式
column_widths = {
    'A': 11, 'B': 4.63, 'C': 13.25, 'D': 4.63,
    'E': 18, 'F': 28, 'G': 58.13, 'H': 19.88, 'I': 15
}

for col, width in column_widths.items():
    ws.Columns(col).ColumnWidth = width

# 设置第一列为短日期格式（m/d/yyyy）
ws.Columns("A").NumberFormat = "m/d/yyyy"

# 添加全边框
thin_border = 2  # 边框粗细
used_range = ws.UsedRange
for row in range(1, used_range.Rows.Count + 1):  # 从第1行开始
    if ws.Cells(row, 1).Value is not None:
        for col in range(1, 10):  # 从A到I列
            ws.Cells(row, col).Borders.Weight = thin_border

# 设置打印区域
ws.PageSetup.PrintArea = f"A1:I{used_range.Rows.Count}"

# 设置页面布局
ws.PageSetup.Orientation = 2  # 横向
ws.PageSetup.PaperSize = 9  # A4纸
ws.PageSetup.Zoom = False
ws.PageSetup.FitToPagesWide = 1
ws.PageSetup.FitToPagesTall = 1

# 保存并导出为 PDF
wb.Save()
wb.ExportAsFixedFormat(0, pdf_output_path)  # 0 表示 PDF 格式
wb.Close(False)
excel.Application.Quit()

# 清理
del excel

print(f"Excel 文件已导出到: {excel_output_path}")
print(f"PDF 文件已导出到: {pdf_output_path}")
