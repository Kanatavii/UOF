import time
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import threading

def run_tracking(service):
    # 获取用户输入的单号
    tracking_numbers = entry.get("1.0", "end-1c").splitlines()
    if not tracking_numbers:
        messagebox.showwarning("输入错误", "请输入至少一个单号")
        return

    # 创建浏览器实例
    driver = webdriver.Chrome()

    # 选择服务并访问相应的网页
    if service == "黑猫":
        driver.get('https://toi.kuronekoyamato.co.jp/cgi-bin/tneko')  # 替换为黑猫的实际目标 URL
    elif service == "福山通运":
        driver.get('https://corp.fukutsu.co.jp/situation/tracking_no')  # 替换为福山通运的实际目标 URL

    # 依次输入单号 - 适配 HTML 元素（请根据服务修改具体输入逻辑）
    for i, number in enumerate(tracking_numbers):
        try:
            if service == "黑猫":
                input_element_id = f"number{i+1:02}"  # 黑猫假设输入框的 name 属性为 'number01', 'number02' 等
                input_element = driver.find_element(By.NAME, input_element_id)
            elif service == "福山通运":
                input_element_id = f"tracking_no{i+1}"  # 福山通运假设输入框的 id 为 'tracking_no1', 'tracking_no2' 等
                input_element = driver.find_element(By.ID, input_element_id)
            
            input_element.clear()  # 清空原有值
            input_element.send_keys(number)
            # 如果是最后一个单号，发送回车键
            if i == len(tracking_numbers) - 1:
                input_element.send_keys(Keys.RETURN)
        except Exception as e:
            print(f"无法输入到元素: {e}")

    # 添加延迟以查看结果
    time.sleep(3000)  # 可以根据需要调整等待时间
    driver.quit()

def start_tracking():
    # 获取选中的服务
    service = service_var.get()
    if service not in ["黑猫", "福山通运"]:
        messagebox.showwarning("选择错误", "请选择一个服务")
        return
    
    # 在单独的线程中运行查询，以免阻塞主 UI
    threading.Thread(target=run_tracking, args=(service,)).start()

# 创建简单的 UI
root = tk.Tk()
root.title("单号查询输入")
root.geometry("400x400")

tk.Label(root, text="请选择服务:").pack(pady=5)
service_var = tk.StringVar(value="黑猫")
tk.Radiobutton(root, text="黑猫", variable=service_var, value="黑猫").pack(anchor=tk.W)
tk.Radiobutton(root, text="福山通运", variable=service_var, value="福山通运").pack(anchor=tk.W)

tk.Label(root, text="请输入要查询的单号，每行一个:").pack(pady=10)
entry = tk.Text(root, height=10, width=40)
entry.pack(pady=10)

tk.Button(root, text="运行查询", command=start_tracking).pack(pady=10)

root.mainloop()
