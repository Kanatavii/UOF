import time
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import pandas as pd
import threading

def run_tracking(service):
    start_time = time.time()  # 记录开始时间
    
    # 获取用户输入的单号并分批
    tracking_numbers = entry.get("1.0", "end-1c").splitlines()
    if not tracking_numbers:
        messagebox.showwarning("输入错误", "请输入至少一个单号")
        return

    # 将单号分成每批最多10个
    batch_size = 10
    batches = [tracking_numbers[i:i + batch_size] for i in range(0, len(tracking_numbers), batch_size)]
    results = []  # 用于存储所有批次的查询结果

    # 设置无头模式
    chrome_options = Options()


    # 创建浏览器实例
    driver = webdriver.Chrome(options=chrome_options)

    try:
        for batch_index, batch in enumerate(batches):
            # 打开新分页
            if batch_index > 0:
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[batch_index])  # 切换到新的分页

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

            # 依次输入批次中的单号并获取结果
            for i, number in enumerate(batch):
                try:
                    # 输入单号
                    input_element_id = f"number{i+1:02}"  # 假设输入框的名称格式为 'number01', 'number02' 等
                    input_element = driver.find_element(By.NAME, input_element_id)
                    input_element.clear()
                    input_element.send_keys(number)
                    
                    # 确保页面完成加载
                    input_element.send_keys(Keys.RETURN)
                    #time.sleep(1)  # 添加适当的延迟，等待页面加载完成。根据需要调整时间。

                    # 提取时间和配送状况
                    try:
                        time_xpath = f"(//div[@class='data date pc-only'])[{i+1}]"
                        status_xpath = f"(//a[contains(@class, 'js-tracking-detail')])[{i+1}]"
                        delivery_time = driver.find_element(By.XPATH, time_xpath).text
                        delivery_status = driver.find_element(By.XPATH, status_xpath).text
                    except Exception as e:
                        delivery_time = "未找到时间"
                        delivery_status = "未找到配送状况"

                    # 将查询的单号和实际结果存储到 results 列表中
                    results.append({
                        "渠道": service,
                        "单号": number,
                        "时间": delivery_time,
                        "配送状况": delivery_status
                    })

                except Exception as e:
                    print(f"无法输入到元素 '{input_element_id}': {e}")
                    results.append({"渠道": service, "单号": number, "时间": "查询失败", "配送状况": f"错误: {e}"})

    except Exception as e:
        print(f"批次处理出错: {e}")

    finally:
        # 关闭浏览器实例
        driver.quit()

        # 将结果保存到 Excel 文件
        df = pd.DataFrame(results)
        df.to_excel("tracking_results.xlsx", index=False, engine='openpyxl')
        
        # 记录结束时间并计算耗时
        end_time = time.time()
        elapsed_time = end_time - start_time
        messagebox.showinfo("完成", f"查询完成，结果已保存到 'tracking_results.xlsx' 文件中。\n耗时：{elapsed_time:.2f} 秒")

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
