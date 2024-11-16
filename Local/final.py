import PySimpleGUI as sg
import pystray
from PIL import Image
import ctypes

# Set the icon for the application's system tray
def set_system_tray_icon(icon_path):
    # Load the icon image
    icon_image = Image.open(icon_path)

    # Convert the icon image to a format compatible with pystray
    icon_data = icon_image.tobytes()
    icon_width, icon_height = icon_image.size

    # Define the callback function for the system tray
    def on_system_tray_click(icon, item):
        pass  # Add your code here for the system tray click action

    # Create the system tray icon
    system_tray_icon = pystray.Icon("物流公司判断", icon=pystray.IconData(icon_data, icon_width, icon_height), title="物流公司判断")
    system_tray_icon.menu = pystray.Menu()

    # Assign the click callback function to the system tray icon
    system_tray_icon.run(on_click=on_system_tray_click)

sg.popup(
    "重要事项\n"
    "\n"
    "\n"
    "①混载配送的条件之一就是货物上面可以叠加其他货物，内有易碎品的须注意;\n"
    "②单件超过50KG，客户要有卸货能力，特别是超过80KG，须确认是否有フォークリフト（叉车），人力卸货恐有人身安全隐患;\n"
    "③个人件有叉车，会被派送公司质疑其真实性从而有被拒收的可能性，如果客户明确明复有叉车，但派送时没有的话，产生的二次派送费用等额外费用，由客户承担;\n"
    "④镜子或岩板等易碎商品,客户必须自己包装稳妥，如果在配送后外包装无破损但内装破碎，派送公司不赔付;\n"
    "⑤如签收时包装有异常，客户可以当面和派送公司确认货物是否完好或是直接拒收，派送公司会和我们直接确认中间环节，如不确认则意味派送正常完成，如有破碎，派送公司是概不负责的;\n"
    "\n"
    "\n"
    "以下条件不符合配送条件\n"
    "1、包车配送，须在到港前询价;\n"
    "2、客户可选择到大阪UOF泉佐野仓库自提;\n"
    "3、无卸货能力：包车公司准备人手帮忙、在车上拆木箱卸货、销毁木箱等费用需单独询价;\n"
    "4、如是多件独立捆包的托盘货物，数据需体现单件尺寸及重量，进大阪仓库后按捆包货物的实际情况安排转运;\n"
    "5、大阪仓库严禁拆木箱后转运内部货物;\n"
    "6、超大超重货物不搬楼;\n"
    )

font=("Helvetica", 100)  # 设置字体为Helvetica，大小为12


def check_logistics_company(person_type, address, length, width, height, weight, can_pass_4t, has_forklift,cardboard, pallet, wooden_pallet, wooden_box):
    logistics_company = set()

    if person_type == '个人':
        
        # 检查货物信息是否符合佐川要求
        if length + width + height <= 260 and weight <= 50 and cardboard:
            logistics_company.add('佐川')

        if length <= 400 and width + height <= 90 and weight <= 30 and cardboard:
            logistics_company.add('佐川')

        # 检查货物信息是否符合TONAMI要求
        if length <= 200 and width <= 200 and height <= 200 and can_pass_4t and (weight <= 85 or has_forklift):
            logistics_company.add('TONAMI')

        # 检查货物信息是否符合JHSS要求
        jhss_prefectures = ['東京都', '埼玉県', '千葉県', '神奈川県', '茨城県', '大阪府', '奈良県', '和歌山県', '京都府', '兵庫県']
        if address in jhss_prefectures:
            logistics_company.add('JHSS')

        # 检查货物信息是否符合大阪自派要求
        osaka_prefectures = ['東京都', '埼玉県', '千葉県', '神奈川県', '茨城県', '大阪府', '奈良県', '和歌山県', '京都府', '兵庫県', '愛知県', '岐阜県']
        if address in osaka_prefectures:
            logistics_company.add('大阪自派')

     

    elif person_type == '公司':
        # 检查货物信息是否符合佐川要求
        if length + width + height <= 260 and weight <= 50 and cardboard:
            logistics_company.add('佐川')

        if length <= 400 and width + height <= 90 and weight <= 30 and cardboard:
            logistics_company.add('佐川')

        # 检查货物信息是否符合TONAMI要求
        if length <= 200 and width <= 200 and height <= 200 and can_pass_4t and (weight <= 85 or has_forklift):
            logistics_company.add('TONAMI')

        # 检查货物信息是否符合JHSS要求
        jhss_prefectures = ['東京都', '埼玉県', '千葉県', '神奈川県', '茨城県', '大阪府', '奈良県', '和歌山県', '京都府', '兵庫県']
        if address in jhss_prefectures:
            logistics_company.add('JHSS')

        # 检查货物信息是否符合大阪自派要求
        osaka_prefectures = ['東京都', '埼玉県', '千葉県', '神奈川県', '茨城県', '大阪府', '奈良県', '和歌山県', '京都府', '兵庫県', '愛知県', '岐阜県']
        if address in osaka_prefectures:
            logistics_company.add('大阪自派')

        # 检查货物信息是否符合福山要求
        if length + width + height <= 500 and weight <= 1300:
            logistics_company.add('福山')

    # Convert the set back to a list if necessary
    logistics_company = list(logistics_company)


    return logistics_company

# 收件地址选项
prefectures = [
    '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県', '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都',
    '神奈川県', '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県', '静岡県', '愛知県', '三重県', '滋賀県', '京都府',
    '大阪府', '兵庫県', '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県', '徳島県', '香川県', '愛媛県', '高知県',
    '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'
]

layout = [
    [sg.Text('请输入以下信息：')],
    [sg.Text('收件人类型：'), sg.Radio('个人', 'PERSON', default=True, key='-PERSON-'), sg.Radio('公司', 'PERSON')],
    [sg.Text("收件地址"), sg.Combo(prefectures, key='-ADDRESS-', enable_events=True, size=(10, 1))],
    [sg.Text("路口能否进4T车"), sg.Radio("能", "CAN_DRIVE_4T", key='-CAN_DRIVE_4T_YES-', default=False), sg.Radio("否", "CAN_DRIVE_4T", key='-CAN_DRIVE_4T_NO-', default=True)],
    [sg.Text("收件人是否有叉车"), sg.Radio("有", "HAS_FORKLIFT", key='-HAS_FORKLIFT_YES-', default=False), sg.Radio("无", "HAS_FORKLIFT", key='-HAS_FORKLIFT_NO-', default=True)],
    [sg.Text('长(cm)：'), sg.Input(key='-LENGTH-')],
    [sg.Text('宽(cm)：'), sg.Input(key='-WIDTH-')],
    [sg.Text('高(cm)：'), sg.Input(key='-HEIGHT-')],
    [sg.Text('重量(kg)：'), sg.Input(key='-WEIGHT-')],
    [sg.Text("货物状态"), sg.Radio('纸箱', "CARGO_TYPE", key='-CARDBOARD-', default=True),
     sg.Radio('托盘', "CARGO_TYPE", key='-PALLET-'), sg.Radio('木托', "CARGO_TYPE", key='-WOODEN_PALLET-'),
     sg.Radio('木箱', "CARGO_TYPE", key='-WOODEN_BOX-')],
    [sg.Button('确定')]
]

window = sg.Window("物流公司判断", layout)
icon_path = r'D:\Python\UOF.ico'  # Replace with the actual path to your icon file
set_system_tray_icon(icon_path)

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break
             
    elif event == '确定':
        # 获取用户输入的信息
        person_type = '个人' if values['-PERSON-'] else '公司'
        address = values['-ADDRESS-']
        length = float(values['-LENGTH-']) if values['-LENGTH-'] else 0.0
        width = float(values['-WIDTH-']) if values['-WIDTH-'] else 0.0
        height = float(values['-HEIGHT-']) if values['-HEIGHT-'] else 0.0
        weight = float(values['-WEIGHT-']) if values['-WEIGHT-'] else 0.0
        can_pass_4t = values['-CAN_DRIVE_4T_YES-']
        has_forklift = values['-HAS_FORKLIFT_YES-']
        cardboard = values['-CARDBOARD-']
        pallet = values['-PALLET-']
        wooden_pallet = values['-WOODEN_PALLET-']
        wooden_box = values['-WOODEN_BOX-']

        if person_type == '个人' and has_forklift:
            sg.popup('个人真的有叉车吗')

        if not any(value == '' for value in [address, length, width, height, weight]):
            logistics_result = check_logistics_company(person_type, address, length, width, height, weight, can_pass_4t, has_forklift, cardboard, pallet, wooden_pallet, wooden_box)
        
            if logistics_result:
                sg.popup(f"推荐物流公司：{', '.join(logistics_result)}")
            else:
                sg.popup("没有找到适合的物流公司")
        else:
            sg.popup("请填写所有信息")
       
  
window.close()
