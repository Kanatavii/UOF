import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# 创建浏览器实例
driver = webdriver.Chrome()

# 访问目标网页
driver.get('https://toi.kuronekoyamato.co.jp/cgi-bin/tneko')  # 替换为您的实际目标 URL

# 示例输入操作 - 替换为您的具体操作
tracking_numbers = [
    "3910-0320-3345",
    "3910-0320-3346",
    "3910-0320-3347",
    "3910-0320-3348",
    "3910-0320-3349",
    "3910-0320-3350",
    "3910-0320-3351",
    "3910-0320-3352",
    "3910-0320-3353",
    "3910-0320-3354"
]

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
time.sleep(30)  # 可以根据需要调整等待时间

# 关闭浏览器（如果需要）
# driver.quit()
