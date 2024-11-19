import os
import win32com.client
import datetime
import logging
import tkinter as tk
from tkinter import messagebox

# 配置日志记录
logging.basicConfig(filename='app.log', level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def check_and_create_folder(folder_path):
    """
    检查文件夹路径是否存在，如果不存在就创建它。
    """
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

def find_existing_subfolder(attachment_folder, received_time):
    """
    在附件保存的文件夹中查找与邮件接收时间相对应的子文件夹。
    """
    for subfolder in os.listdir(attachment_folder):
        subfolder_path = os.path.join(attachment_folder, subfolder)
        if os.path.isdir(subfolder_path) and subfolder.startswith(received_time):
            return subfolder
    return None

def delete_empty_folders(attachment_folder):
    """
    删除空文件夹。
    """
    for folder in os.listdir(attachment_folder):
        folder_path = os.path.join(attachment_folder, folder)
        if os.path.isdir(folder_path) and not os.listdir(folder_path):
            os.rmdir(folder_path)

try:
    # 创建 Outlook 应用程序对象
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # 获取默认文件夹（收件箱）
    inbox = namespace.GetDefaultFolder(6)  # 6 表示收件箱

    # 构建搜索过滤器
    filter_str = "@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%@plutus-thl.com%'"
    
    # 计算2天前的日期
    one_month_ago = datetime.datetime.now() - datetime.timedelta(days=2)
    filter_str += f" AND \"urn:schemas:httpmail:datereceived\" >= '{one_month_ago.strftime('%m/%d/%Y %H:%M %p')}'"

    # 在 Outlook 中执行搜索
    search_results = inbox.Items.Restrict(filter_str)

    # 打印搜索结果数量
    print(f"Found {len(search_results)} emails.")

    # 创建附件保存的文件夹
    attachment_folder = "Z:\\UOF\转运数据\JHSS"
    check_and_create_folder(attachment_folder)

    # 获取 JHSS 文件夹下所有子文件夹中已存在的附件文件名
    previous_attachments = set()
    for root, dirs, files in os.walk(attachment_folder):
        if root != attachment_folder:
            previous_attachments.update(files)

    # 新许可提示框
    new_license_message = ""

    # 遍历搜索结果
    for item in search_results:
        print("Subject:", item.Subject)
        print("Received Time:", item.ReceivedTime)
        print("Sender:", item.SenderEmailAddress)
        print("---------------------")

        # 获取邮件接收的时间，并转换为特定格式
        received_time = item.ReceivedTime.strftime("%Y%m%d_%H%M%S")

        # 在 JHSS 文件夹中以邮件接收时间创建子文件夹
        time_folder_path = os.path.join(attachment_folder, received_time)

        # 检查是否存在重名子文件夹
        existing_subfolder = find_existing_subfolder(attachment_folder, received_time)

        # 如果已经存在该子文件夹，则直接使用现有的文件夹
        if existing_subfolder:
            time_folder_path = os.path.join(attachment_folder, existing_subfolder)
        else:
            check_and_create_folder(time_folder_path)

        # 检查是否存在附件
        if item.Attachments.Count > 0:
            # 打印附件数量
            print(f"Found {item.Attachments.Count} attachments in the email.")
            # 遍历邮件中的附件
            for attachment in item.Attachments:
                attachment_filename = attachment.FileName
                print(f"Attachment file name: {attachment_filename}")
                print(f"Is the attachment previously saved? {attachment_filename in previous_attachments}")
                if attachment_filename not in previous_attachments:
                    save_path = os.path.join(time_folder_path, attachment_filename)
                    print(f"Save path: {save_path}")
                    if not os.path.exists(save_path):
                        try:
                            attachment.SaveAsFile(save_path)
                        except Exception as e:
                            print(f"Failed to save the file at {save_path}")
                            logging.error(f"Failed to save the file at {save_path}", exc_info=True)
                        else:
                            print("Attachment saved:", save_path)
                            # 添加到已存在附件集合
                            previous_attachments.add(attachment_filename)
                            # 添加新许可信息
                            new_license_message += f"新许可: {attachment_filename}\n"
                    else:
                        print("Attachment already exists:", attachment_filename)
                else:
                    print("Attachment already exists:", attachment_filename)
        else:
            print("No attachments found in this email or folder already exists.")

    # 删除空文件夹
    delete_empty_folders(attachment_folder)

    # 关闭 Outlook 应用程序
    outlook.Quit()

    # 显示新许可信息提示框
    if new_license_message:
        messagebox.showinfo("有新许可", new_license_message)
    else:
        messagebox.showinfo("没有新许可", "没有新许可.")
except Exception as e:
    print("An error occurred:", e)
    logging.error("An error occurred", exc_info=True)
