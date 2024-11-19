import time
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import threading

def run_tracking():
    # 获取用户输入的单号
    tracking_numbers = entry.get("1.0", "end-1c").splitlines()
    if not tracking_numbers:
        messagebox.showwarning("输入错误", "请输入至少一个单号")
        return
    
    # 创建浏览器实例
    driver = webdriver.Chrome()

    # 访问目标网页
    driver.get('https://toi.kuronekoyamato.co.jp/cgi-bin/tneko')  # 替换为您的实际目标 URL

    # 依次输入单号 - 适配 HTML 元素
    for i, number in enumerate(tracking_numbers):
        input_element_id = f"number{i+1:02}"  # 假设输入框的 name 属性为 'number01', 'number02' 等
        try:
            input_element = driver.find_element(By.NAME, input_element_id)
            input_element.clear()  # 清空原有值
            input_element.send_keys(number)
            # 如果是最后一个单号，发送回车键
            if i == len(tracking_numbers) - 1:
                input_element.send_keys(Keys.RETURN)
        except Exception as e:
            print(f"无法输入到元素 '{input_element_id}': {e}")

    # 添加延迟以查看结果
    time.sleep(3000)  # 可以根据需要调整等待时间
    driver.quit()

def start_tracking():
    # 在单独的线程中运行查询，以免阻塞主 UI
    threading.Thread(target=run_tracking).start()

# 创建简单的 UI
root = tk.Tk()
root.title("单号查询输入")
root.geometry("400x300")

tk.Label(root, text="请输入要查询的单号，每行一个:").pack(pady=10)
entry = tk.Text(root, height=10, width=40)
entry.pack(pady=10)

tk.Button(root, text="运行查询", command=start_tracking).pack(pady=10)

root.mainloop()
