import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit
import win32com.client
import pythoncom
import re

class SingleNumberInputApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle('单号输入 UI')
        self.layout = QVBoxLayout()
        
        self.label = QLabel('请输入单号：')
        self.layout.addWidget(self.label)
        
        self.input_box = QLineEdit()
        self.layout.addWidget(self.input_box)
        
        self.submit_button = QPushButton('提交单号')
        self.submit_button.clicked.connect(self.handle_input)
        self.layout.addWidget(self.submit_button)
        
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.layout.addWidget(self.log)
        
        self.setLayout(self.layout)

    def handle_input(self):
        input_number = self.input_box.text().strip()  # 去除输入的空格
        if input_number:
            self.log.append(f'处理单号：{input_number}')
            self.process_excel(input_number)
            self.input_box.clear()
        else:
            self.log.append('请输入有效的单号。')
    
    def process_excel(self, order_number):
        try:
            pythoncom.CoInitialize()  # 初始化 COM 库
            file_path = r'Z:\UOF\转运数据\市川\(社外)市川倉庫搬出申込書.xlsx'

            self.log.append(f'使用 win32com.client 读取和处理文件: {file_path}')

            # 使用 DispatchEx 创建独立的 Excel 实例
            excel = win32com.client.DispatchEx("Excel.Application")

            # 打开工作簿
            workbook = excel.Workbooks.Open(file_path, ReadOnly=True, Notify=False)
            self.log.append("文件成功打开（只读模式）。")

            # 获取目标工作表
            sheet = workbook.Sheets('市川社外用')

            # 在 O17 单元格输入单号
            sheet.Range('O17').Value = order_number
            self.log.append(f'O17 单元格已输入单号：{order_number}')

            # 等待计算完成
            excel.CalculateUntilAsyncQueriesDone()

            # 在其他工作表中查找单号，排除 '市川社外用'
            found_cell = None
            for ws in workbook.Sheets:
                if ws.Name == '市川社外用':
                    continue  # 跳过当前工作表

                self.log.append(f'正在搜索工作表: {ws.Name}')

                # 查找单号在第 O 列 (第 15 列)
                max_rows = ws.UsedRange.Rows.Count
                for row in range(1, max_rows + 1):
                    cell_value = ws.Cells(row, 15).Value
                    cell_display_value = ws.Cells(row, 15).Text  # 获取显示值

                    # 调试输出，查看每个单元格的值
                    self.log.append(f'行 {row} - 值: {cell_value} (显示值: {cell_display_value})')

                    # 去除空格和特殊字符进行比较
                    if cell_value:
                        cell_value_str = str(cell_value).strip()
                        cell_value_str = re.sub(r'\s+', '', cell_value_str)  # 去除所有空格
                        if cell_value_str == order_number or cell_display_value.strip() == order_number:
                            found_cell = ws.Cells(row, 15)
                            break
                if found_cell:
                    break

            if found_cell:
                self.log.append(f'找到单号在工作表 "{found_cell.Worksheet.Name}" 的第 {found_cell.Row} 行。')

                # 读取单号后面的单元格的值
                value_after_order = found_cell.Offset(0, 1).Value  # 单号后一个单元格
                value_after_order2 = found_cell.Offset(0, 2).Value  # 单号后两个单元格
                value_after_order3 = found_cell.Offset(0, 3).Value  # 单号后三个单元格
                self.log.append(f'单号后一个单元格的值：{value_after_order}')
                self.log.append(f'单号后两个单元格的值：{value_after_order2}')
                self.log.append(f'单号后三个单元格的值：{value_after_order3}')

                # 写入到 A23、O23 和 AD23
                sheet.Range('A23').Value = value_after_order
                sheet.Range('O23').Value = value_after_order2
                sheet.Range('AD23').Value = value_after_order3
                self.log.append(f'A23 填写：{value_after_order}')
                self.log.append(f'O23 填写：{value_after_order2}')
                self.log.append(f'AD23 填写：{value_after_order3}')

                # 生成新的文件名
                new_file_base = f'Z:\\UOF\\转运数据\\市川\\{order_number}-搬出票'
                new_file_path = f'{new_file_base}.xlsx'
                pdf_path = f'{new_file_base}.pdf'

                # 创建一个新工作簿并复制当前工作表
                new_workbook = excel.Workbooks.Add()
                sheet.Copy(Before=new_workbook.Sheets(1))  # 复制工作表到新工作簿

                # 保存新工作簿
                new_workbook.SaveAs(new_file_path)
                self.log.append(f'新工作簿已保存到: {new_file_path}')

                # 导出当前工作表为 PDF
                sheet.ExportAsFixedFormat(0, pdf_path)
                self.log.append(f'PDF 文件已导出到: {pdf_path}')

                # 关闭新工作簿
                new_workbook.Close(SaveChanges=False)
            else:
                self.log.append('未找到匹配的单号。')

            # 关闭原工作簿
            workbook.Close(SaveChanges=False)
        except Exception as e:
            self.log.append(f'处理出错: {e}')
        finally:
            # 关闭 Excel 应用
            try:
                excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()  # 释放 COM 库

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = SingleNumberInputApp()
    window.show()
    sys.exit(app.exec_())
