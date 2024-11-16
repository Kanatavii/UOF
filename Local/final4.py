import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import winsound


def show_info():
    info_text = (
        "重要事项\n\n"
        "①混载配送的条件之一就是货物上面可以叠加其他货物，内有易碎品的须注意;\n"
        "②单件超过50KG，客户要有卸货能力，特别是超过80KG，须确认是否有フォークリフト（叉车），人力卸货恐有人身安全隐患;\n"
        "③个人件有叉车，会被派送公司质疑其真实性从而有被拒收的可能性，如果客户明确明复有叉车，但派送时没有的话，产生的二次派送费用等额外费用，由客户承担;\n"
        "④镜子或岩板等易碎商品,客户必须自己包装稳妥，如果在配送后外包装无破损但内装破碎，派送公司是不赔付的c\n"
        "⑤如签收时包装有异常，客户可以当面和派送公司确认货物是否完好或是直接拒收，派送公司会和我们直接确认中间环节，如不确认则意味派送正常完成，如有破碎，派送公司是概不负责的；\n"
        "\n"
        "以下条件不符合配送条件\n"
        "1、包车配送，须在到港前询价；;\n"
        "2、客户可选择到大阪UOF泉佐野仓库自提；;\n"
        "3、无卸货能力：包车公司准备人手帮忙、在车上拆木箱卸货、销毁木箱等费用需单独询价；;\n"
        "4、如是多件独立捆包的托盘货物，数据需体现单件尺寸及重量，进大阪仓库后按捆包货物的实际情况安排转运；;\n"
        "5、大阪仓库严禁拆木箱后转运内部货物；;\n"
        "6、易碎品不搬楼；;\n"
    )

    # 显示重要事项的信息
    messagebox.showinfo("重要事项", info_text)

def check_logistics_company(person_type, address, length, width, height, weight, can_pass_4t, has_forklift,cardboard, pallet, wooden_pallet, wooden_box):
    logistics_company = []

    if person_type == '个人':
        
        # 检查货物信息是否符合佐川要求
        if length + width + height <= 260 and weight <= 50 and cardboard:
            logistics_company.append('佐川')

        if length <= 400 and width + height <= 90 and weight <= 30 and cardboard:
            logistics_company.append('佐川')

        # 检查货物信息是否符合TONAMI要求
        if length <= 200 and width <= 200 and height <= 200 and can_pass_4t and (weight <= 85 or has_forklift):
            logistics_company.append('TONAMI')

        # 检查货物信息是否符合JHSS要求
        jhss_prefectures = ['東京都', '埼玉県', '千葉県', '神奈川県', '茨城県', '大阪府', '奈良県', '和歌山県', '京都府', '兵庫県']
        if address in jhss_prefectures:
            logistics_company.append('JHSS')

        # 检查货物信息是否符合大阪自派要求
        osaka_prefectures = ['東京都', '埼玉県', '千葉県', '神奈川県', '茨城県', '大阪府', '奈良県', '和歌山県', '京都府', '兵庫県', '愛知県', '岐阜県']
        if address in osaka_prefectures:
            logistics_company.append('大阪自派')

     

    elif person_type == '公司':
        # 检查货物信息是否符合佐川要求
        if length + width + height <= 260 and weight <= 50 and cardboard:
            logistics_company.append('佐川')

        if length <= 400 and width + height <= 90 and weight <= 30 and cardboard:
            logistics_company.append('佐川')

        # 检查货物信息是否符合TONAMI要求
        if length <= 200 and width <= 200 and height <= 200 and can_pass_4t and (weight <= 85 or has_forklift):
            logistics_company.append('TONAMI')

        # 检查货物信息是否符合JHSS要求
        jhss_prefectures = ['東京都', '埼玉県', '千葉県', '神奈川県', '茨城県', '大阪府', '奈良県', '和歌山県', '京都府', '兵庫県']
        if address in jhss_prefectures:
            logistics_company.append('JHSS')

        # 检查货物信息是否符合大阪自派要求
        osaka_prefectures = ['東京都', '埼玉県', '千葉県', '神奈川県', '茨城県', '大阪府', '奈良県', '和歌山県', '京都府', '兵庫県', '愛知県', '岐阜県']
        if address in osaka_prefectures:
            logistics_company.append('大阪自派')

        # 检查货物信息是否符合福山要求
        if length + width + height <= 500 and weight <= 1300:
            logistics_company.append('福山')

    return logistics_company

def on_submit():
    # 获取用户输入的信息
    person_type = '个人' if person_var.get() else '公司'
    address = address_combo.get()
    length = float(length_entry.get()) if length_entry.get() else 0.0
    width = float(width_entry.get()) if width_entry.get() else 0.0
    height = float(height_entry.get()) if height_entry.get() else 0.0
    weight = float(weight_entry.get()) if weight_entry.get() else 0.0
    can_pass_4t = can_pass_4t_var.get()
    has_forklift = has_forklift_var.get()
    cardboard = cardboard_var.get()
    pallet = pallet_var.get()
    wooden_pallet = wooden_pallet_var.get()
    wooden_box = wooden_box_var.get()

    # 判断物流公司
    recommended_companies = check_logistics_company(person_type, address, length, width, height, weight, can_pass_4t, has_forklift, cardboard, pallet, wooden_pallet, wooden_box)

    # 显示结果
    if recommended_companies:
        result_label.config(text="推荐物流公司：{}".format(', '.join(recommended_companies)))
    else:
        result_label.config(text="没有符合要求的物流公司")

root = tk.Tk()

# 收件人类型
recipient_frame = ttk.LabelFrame(root, text='收件人类型')
recipient_frame.pack(pady=10)

person_var = tk.BooleanVar(value=True)
person_radio = ttk.Radiobutton(recipient_frame, text='个人', variable=person_var, value=True)
person_radio.grid(row=0, column=0, padx=5)
company_radio = ttk.Radiobutton(recipient_frame, text='公司', variable=person_var, value=False)
company_radio.grid(row=0, column=1, padx=5)

# 收件地址
address_label = ttk.Label(root, text='收件地址')
address_label.pack()

prefectures = [
    '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県', '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都',
    '神奈川県', '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県', '静岡県', '愛知県', '三重県', '滋賀県', '京都府',
    '大阪府', '兵庫県', '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県', '徳島県', '香川県', '愛媛県', '高知県',
    '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'
]

address_combo = ttk.Combobox(root, values=prefectures)
address_combo.pack()

# 路口能否进4T车
can_pass_4t_frame = ttk.LabelFrame(root, text='路口能否进4T车')
can_pass_4t_frame.pack(pady=10)

can_pass_4t_var = tk.BooleanVar(value=False)
can_pass_4t_yes_radio = ttk.Radiobutton(can_pass_4t_frame, text='能', variable=can_pass_4t_var, value=True)
can_pass_4t_yes_radio.grid(row=0, column=0, padx=5)
can_pass_4t_no_radio = ttk.Radiobutton(can_pass_4t_frame, text='否', variable=can_pass_4t_var, value=False)
can_pass_4t_no_radio.grid(row=0, column=1, padx=5)

# 收件人是否有叉车
has_forklift_frame = ttk.LabelFrame(root, text='收件人是否有叉车')
has_forklift_frame.pack(pady=10)

has_forklift_var = tk.BooleanVar(value=False)
has_forklift_yes_radio = ttk.Radiobutton(has_forklift_frame, text='有', variable=has_forklift_var, value=True)
has_forklift_yes_radio.grid(row=0, column=0, padx=5)
has_forklift_no_radio = ttk.Radiobutton(has_forklift_frame, text='无', variable=has_forklift_var, value=False)
has_forklift_no_radio.grid(row=0, column=1, padx=5)

# 尺寸和重量
size_weight_frame = ttk.LabelFrame(root, text='尺寸和重量')
size_weight_frame.pack(pady=10)

length_label = ttk.Label(size_weight_frame, text='长(cm):')
length_label.grid(row=0, column=0, padx=5, pady=5)
length_entry = ttk.Entry(size_weight_frame)
length_entry.grid(row=0, column=1, padx=5, pady=5)

width_label = ttk.Label(size_weight_frame, text='宽(cm):')
width_label.grid(row=1, column=0, padx=5, pady=5)
width_entry = ttk.Entry(size_weight_frame)
width_entry.grid(row=1, column=1, padx=5, pady=5)

height_label = ttk.Label(size_weight_frame, text='高(cm):')
height_label.grid(row=2, column=0, padx=5, pady=5)
height_entry = ttk.Entry(size_weight_frame)
height_entry.grid(row=2, column=1, padx=5, pady=5)

weight_label = ttk.Label(size_weight_frame, text='重量(kg):')
weight_label.grid(row=3, column=0, padx=5, pady=5)
weight_entry = ttk.Entry(size_weight_frame)
weight_entry.grid(row=3, column=1, padx=5, pady=5)

# 货物状态
cargo_frame = ttk.LabelFrame(root, text='货物状态')
cargo_frame.pack(pady=10)

cardboard_var = tk.BooleanVar(value=True)
cardboard_radio = ttk.Radiobutton(cargo_frame, text='纸箱', variable=cardboard_var, value=True)
cardboard_radio.grid(row=0, column=0, padx=5)
pallet_var = tk.BooleanVar(value=False)
pallet_radio = ttk.Radiobutton(cargo_frame, text='托盘', variable=pallet_var, value=True)
pallet_radio.grid(row=0, column=1, padx=5)
wooden_pallet_var = tk.BooleanVar(value=False)
wooden_pallet_radio = ttk.Radiobutton(cargo_frame, text='木托', variable=wooden_pallet_var, value=True)
wooden_pallet_radio.grid(row=0, column=2, padx=5)
wooden_box_var = tk.BooleanVar(value=False)
wooden_box_radio = ttk.Radiobutton(cargo_frame, text='木箱', variable=wooden_box_var, value=True)
wooden_box_radio.grid(row=0, column=3, padx=5)

# 提交按钮
submit_button = ttk.Button(root, text='确定', command=on_submit)
submit_button.pack(pady=10)

# 结果显示标签
result_label = ttk.Label(root, text='')
result_label.pack(pady=10)

root.mainloop()