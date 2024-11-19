import time
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
import logging

# 设置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def query_batch(service, batch, batch_index):
    results = []  # 用于存储批次的查询结果
    batch_start_time = time.time()  # 记录批次开始时间
    logging.info(f"开始处理批次 {batch_index}")

    # 设置无头模式
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # 无头模式
    chrome_options.add_argument("--disable-gpu")

    # 创建浏览器实例
    driver = webdriver.Chrome(options=chrome_options)

    try:
        # 选择服务并访问相应的网页
        visit_start_time = time.time()
        if service == "黑猫":
            driver.get('https://toi.kuronekoyamato.co.jp/cgi-bin/tneko')
        elif service == "福山通运":
            driver.get('https://corp.fukutsu.co.jp/situation/tracking_no')
        elif service == "邮局":
            driver.get('https://trackings.post.japanpost.jp/services/srv/sequenceNoSearch?requestNo=&count=&moveIndividual.x=112&moveIndividual.y=27&locale=ja')
        elif service == "佐川":
            driver.get('https://k2k.sagawa-exp.co.jp/p/sagawa/web/okurijoinput.jsp')
        elif service == "Tonami":
            driver.get('https://trc1.tonami.co.jp/trc/search3/excSearch3')
        elif service == "GB":
            driver.get('https://www.gbtech722-tms.com/search.html')
        logging.info(f"页面加载耗时: {time.time() - visit_start_time:.2f} 秒")

        # 批量定位输入框并输入所有单号
        input_start_time = time.time()
        input_elements = driver.find_elements(By.CSS_SELECTOR, 'input[name^="number"]')
        for i, number in enumerate(batch):
            try:
                if i < len(input_elements):  # 确保索引在范围内
                    input_elements[i].clear()
                    input_elements[i].send_keys(number)
            except Exception as e:
                logging.error(f"无法输入到元素 '{input_elements[i].get_attribute('name') if i < len(input_elements) else 'unknown'}': {e}")
                results.append({"渠道": service, "单号": number, "时间": "查询失败", "配送状况": f"错误: {e}"})
        logging.info(f"输入单号耗时: {time.time() - input_start_time:.2f} 秒")

        # 一次性提交所有单号
        submit_start_time = time.time()
        try:
            input_elements[-1].send_keys(Keys.RETURN)  # 示例操作，可能需要根据页面实际情况修改
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.data')))
        except TimeoutException:
            logging.warning("页面加载超时。")
        logging.info(f"提交和页面加载耗时: {time.time() - submit_start_time:.2f} 秒")

        # 处理查询结果
        result_start_time = time.time()
        for i in range(len(batch)):
            try:
                time_xpath = f"(//div[@class='data date pc-only'])[{i+1}]"
                status_xpath = f"(//a[contains(@class, 'js-tracking-detail')])[{i+1}]"
                delivery_time = driver.find_element(By.XPATH, time_xpath).text
                delivery_status = driver.find_element(By.XPATH, status_xpath).text
                results.append({
                    "渠道": service,
                    "单号": batch[i],
                    "时间": delivery_time,
                    "配送状况": delivery_status
                })
            except Exception as e:
                results.append({"渠道": service, "单号": batch[i], "时间": "未找到时间", "配送状况": "未找到配送状况"})
        logging.info(f"结果处理耗时: {time.time() - result_start_time:.2f} 秒")

    except Exception as e:
        logging.error(f"批次处理出错: {e}")

    finally:
        driver.quit()
        logging.info(f"批次 {batch_index} 处理总耗时: {time.time() - batch_start_time:.2f} 秒")

    return results

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

    all_results = []  # 用于存储所有批次的查询结果

    # 使用多线程并行处理批次
    with ThreadPoolExecutor() as executor:
        futures = [executor.submit(query_batch, service, batch, batch_index) for batch_index, batch in enumerate(batches)]
        for future in futures:
            all_results.extend(future.result())

    # 将结果保存到 Excel 文件
    df = pd.DataFrame(all_results)
    df.to_excel("tracking_results.xlsx", index=False, engine='openpyxl')
    
    # 记录结束时间并计算耗时
    end_time = time.time()
    elapsed_time = end_time - start_time
    total_tracking_numbers = len(tracking_numbers)
    logging.info(f"总处理耗时: {elapsed_time:.2f} 秒")
    messagebox.showinfo("完成", f"查询完成，结果已保存到 'tracking_results.xlsx' 文件中。\n单号总数：{total_tracking_numbers}\n耗时：{elapsed_time:.2f} 秒")

# 创建简单的 UI
root = tk.Tk()
root.title("单号查询输入")
root.geometry("400x400")

tk.Label(root, text="请输入要查询的单号，每行一个:").pack(pady=5)
entry = tk.Text(root, height=10, width=40)
entry.pack(pady=5)

service_var = tk.StringVar(value="黑猫")  # 默认服务
tk.Label(root, text="请选择服务:").pack(pady=5)
tk.Radiobutton(root, text="黑猫", variable=service_var, value="黑猫").pack(anchor=tk.W)
tk.Radiobutton(root, text="福山通运", variable=service_var, value="福山通运").pack(anchor=tk.W)
tk.Radiobutton(root, text="邮局", variable=service_var, value="邮局").pack(anchor=tk.W)
tk.Radiobutton(root, text="佐川", variable=service_var, value="佐川").pack(anchor=tk.W)
tk.Radiobutton(root, text="Tonami", variable=service_var, value="Tonami").pack(anchor=tk.W)
tk.Radiobutton(root, text="GB", variable=service_var, value="GB").pack(anchor=tk.W)

tk.Button(root, text="运行查询", command=lambda: run_tracking(service_var.get())).pack(pady=10)

root.mainloop()
