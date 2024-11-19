import time
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading

def run_tracking(service):
    # 获取用户输入的单号并分批
    tracking_numbers = entry.get("1.0", "end-1c").splitlines()
    if not tracking_numbers:
        messagebox.showwarning("输入错误", "请输入至少一个单号")
        return

    # 将单号分成每批最多10个
    batch_size = 10
    batches = [tracking_numbers[i:i + batch_size] for i in range(0, len(tracking_numbers), batch_size)]

    # 创建浏览器实例
    driver = webdriver.Chrome()

    for batch in batches:
        try:
            # 选择服务并访问相应的网页
            if service == "黑猫":
                driver.get('https://toi.kuronekoyamato.co.jp/cgi-bin/tneko')  # 黑猫的实际目标 URL
            elif service == "福山通运":
                driver.get('https://corp.fukutsu.co.jp/situation/tracking_no')  # 福山通运的实际目标 URL
            elif service == "邮局":
                driver.get('https://trackings.post.japanpost.jp/services/srv/sequenceNoSearch?requestNo=&count=&moveIndividual.x=112&moveIndividual.y=27&locale=ja')  # 邮局的目标 URL
            elif service == "佐川":
                driver.get('https://k2k.sagawa-exp.co.jp/p/sagawa/web/okurijoinput.jsp')  # 佐川的目标 URL
            elif service == "Tonami":
                driver.get('https://trc1.tonami.co.jp/trc/search3/excSearch3')  # Tonami的目标 URL
            elif service == "GB":
                driver.get('https://www.gbtech722-tms.com/search.html')  # GB的目标 URL

            # 依次输入批次中的单号
            for i, number in enumerate(batch):
                try:
                    if service == "黑猫":
                        input_element_id = f"number{i+1:02}"  # 黑猫假设输入框的 name 属性为 'number01', 'number02' 等
                        input_element = driver.find_element(By.NAME, input_element_id)
                    elif service == "福山通运":
                        input_element_id = f"tracking_no{i+1}"  # 福山通运的输入框 ID 假设为 'tracking_no1', 'tracking_no2' 等
                        input_element = driver.find_element(By.ID, input_element_id)
                    elif service == "邮局":
                        input_element_name = f"requestNo{i+1}"  # 邮局的输入框 name 假设为 'requestNo1', 'requestNo2' 等
                        input_element = driver.find_element(By.NAME, input_element_name)
                    elif service == "佐川":
                        input_element_id = f"main:no{i+1}"  # 佐川的输入框 ID 假设为 'main:no1', 'main:no2' 等
                        input_element = driver.find_element(By.ID, input_element_id)
                    elif service == "Tonami":
                        input_element_id = f"excNum{i}"  # Tonami的输入框 ID 假设为 'excNum0', 'excNum1' 等
                        input_element = driver.find_element(By.ID, input_element_id)
                    elif service == "GB":
                        input_element_id = f"txtToiawaseNo{i+1}"  # GB的输入框 ID 假设为 'txtToiawaseNo1', 'txtToiawaseNo2' 等
                        input_element = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, input_element_id))
                        )
                    
                    input_element.clear()  # 清空原有值
                    input_element.send_keys(number)
                    # 如果是最后一个单号，发送回车键或特殊处理
                    if i == len(batch) - 1:
                        input_element.send_keys(Keys.RETURN)
                        # Tonami 特殊逻辑：点击提交按钮
                        if service == "Tonami":
                            submit_button = driver.find_element(By.NAME, "search")
                            submit_button.click()
                        # GB 特殊逻辑：点击提交按钮
                        elif service == "GB":
                            submit_button = driver.find_element(By.ID, "btnStart")
                            submit_button.click()
                except Exception as e:
                    print(f"无法输入到元素 '{input_element_id if service != '邮局' else input_element_name}': {e}")

            # 添加延迟以查看结果或处理页面的变化
            time.sleep(10)  # 根据需要调整等待时间以便完成批次查询

        except Exception as e:
            print(f"批次处理出错: {e}")
    
    # 关闭浏览器实例
    driver.quit()

def start_tracking():
    # 获取选中的服务
    service = service_var.get()
    if service not in ["黑猫", "福山通运", "邮局", "佐川", "Tonami", "GB"]:
        messagebox.showwarning("选择错误", "请选择一个服务")
        return
    
    # 在单独的线程中运行查询，以免阻塞主 UI
    threading.Thread(target=run_tracking, args=(service,)).start()

def clear_input():
    # 清空输入框中的内容
    entry.delete('1.0', tk.END)

# 创建简单的 UI
root = tk.Tk()
root.title("单号查询输入")
root.geometry("400x400")

# 创建服务选项
tk.Label(root, text="请选择服务:").pack(pady=5)
service_var = tk.StringVar(value="邮局")

# 创建一个框架用于分列布局
radio_frame = tk.Frame(root)
radio_frame.pack()

# 两列布局的单选按钮
tk.Radiobutton(radio_frame, text="邮局", variable=service_var, value="邮局").grid(row=0, column=0, sticky="w", padx=10, pady=2)
tk.Radiobutton(radio_frame, text="黑猫", variable=service_var, value="黑猫").grid(row=0, column=1, sticky="w", padx=10, pady=2)
tk.Radiobutton(radio_frame, text="佐川", variable=service_var, value="佐川").grid(row=1, column=0, sticky="w", padx=10, pady=2)
tk.Radiobutton(radio_frame, text="福山通运", variable=service_var, value="福山通运").grid(row=1, column=1, sticky="w", padx=10, pady=2)
tk.Radiobutton(radio_frame, text="Tonami", variable=service_var, value="Tonami").grid(row=2, column=0, sticky="w", padx=10, pady=2)
tk.Radiobutton(radio_frame, text="GB", variable=service_var, value="GB").grid(row=2, column=1, sticky="w", padx=10, pady=2)

tk.Label(root, text="请输入要查询的单号，每行一个:").pack(pady=5)
entry = tk.Text(root, height=10, width=40)
entry.pack(pady=5)

# 添加按钮布局
button_frame = tk.Frame(root)
button_frame.pack()
tk.Button(button_frame, text="运行查询", command=start_tracking).pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="清空", command=clear_input).pack(side=tk.LEFT, padx=5)

root.mainloop()
