import pandas as pd
from sqlalchemy import create_engine

# 文件路径和数据库配置
file_path = r'D:\UOF-JP\Work\SynologyDrive\UOF\转运数据\UOF出入库汇总表.xlsx'
sheet_name = '货物总单'

db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'osaka_main'
}

# 创建数据库引擎
def create_db_engine():
    engine_url = f"mysql+mysqlconnector://{db_config['user']}:{db_config['password']}@{db_config['host']}/{db_config['database']}"
    return create_engine(engine_url)

# 从数据库加载数据，并按主键排序
def load_data_from_db():
    engine = create_db_engine()
    query = "SELECT * FROM zongdan_2024"
    db_data = pd.read_sql(query, engine)
    db_data['送り状番号'] = db_data['送り状番号'].astype(str).str.strip()  # 保证主键类型一致
    print(f"数据库数据加载完成，总行数: {len(db_data)}")
    return db_data

# 从 Excel 文件加载数据，并按主键排序
def load_excel_data():
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
    df.columns = df.columns.astype(str).str.strip()
    df = df.where(pd.notnull(df), None).drop_duplicates()
    df['送り状番号'] = df['送り状番号'].astype(str).str.strip()  # 保证主键类型一致
    df = df.sort_values(by='送り状番号').reset_index(drop=True)
    print(f"Excel 数据加载完成，总行数: {len(df)}")
    return df

# 预处理日期格式和空值
def preprocess_dataframes(df1, df2):
    date_columns = ['到港时间', '搬入时间', '许可时间函数对应', '入库时间', 
                    '出库指示时间_入金确认时间', '转运日期']

    for col in date_columns:
        if col in df1.columns:
            df1[col] = pd.to_datetime(df1[col], format='%Y-%m-%d', errors='coerce').dt.date
        if col in df2.columns:
            df2[col] = pd.to_datetime(df2[col], format='%Y-%m-%d', errors='coerce').dt.date

    df1 = df1.where(df1.notna(), None)
    df2 = df2.where(df2.notna(), None)

    return df1, df2

# 对齐列和数据格式
def align_dataframes(df1, df2, primary_key):
    df1_aligned = df1.set_index(primary_key).sort_index()
    df2_aligned = df2.set_index(primary_key).sort_index()

    common_columns = df1_aligned.columns.intersection(df2_aligned.columns)
    df1_aligned = df1_aligned[common_columns]
    df2_aligned = df2_aligned[common_columns]

    return df1_aligned, df2_aligned

# 比较数据并显示不一致的部分
def compare_dataframes(df1, df2):
    try:
        pd.testing.assert_frame_equal(df1, df2, check_dtype=False, check_like=True)
        print("本地文件和数据库数据完全一致。")
    except AssertionError as e:
        print("检测到数据不一致。显示不一致的部分：")
        mismatches = df1.compare(df2, align_axis=0)
        print(mismatches)

# 验证数据一致性
def validate_data():
    excel_data = load_excel_data()
    db_data = load_data_from_db()

    if excel_data.empty or db_data.empty:
        print("数据加载失败，无法比较。")
        return

    primary_key = '送り状番号'

    excel_data, db_data = align_dataframes(excel_data, db_data, primary_key)
    excel_data, db_data = preprocess_dataframes(excel_data, db_data)

    compare_dataframes(excel_data, db_data)

if __name__ == "__main__":
    validate_data()
