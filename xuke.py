import pandas as pd
import mysql.connector
from mysql.connector import Error

# Excel 文件路径和数据库配置
file_path = r'Z:\UOF\转运数据\UOF出入库汇总表.xlsx'
sheet_name = '货物总单'

db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'osaka_main'
}

# 列名映射：将 Excel 的列名映射为数据库中的列名
column_mapping = {
    '许可时间\n函数对应': '许可时间函数对应',
    '出库指示时间/入金确认时间': '出库指示时间_入金确认时间',
    '空运/海运': '空运_海运',
    '会社名/个人': '会社名_个人',
    '乐天匹配\n用辅助函数': '乐天匹配用辅助函数',
    '乐天5位数\n担当者': '乐天5位数担当者'
}

def connect_db():
    """连接数据库"""
    try:
        connection = mysql.connector.connect(**db_config)
        if connection.is_connected():
            print("成功连接到数据库")
            return connection
    except Error as e:
        print(f"数据库连接错误: {e}")
    return None

def load_excel_data():
    """加载 Excel 数据并进行预处理"""
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
    df.columns = df.columns.astype(str).str.strip()
    df = df.rename(columns=column_mapping)  # 应用列名映射
    df = df.loc[:, ~df.columns.str.contains('^Unnamed', na=False)]
    df = df.where(pd.notnull(df), None).drop_duplicates()
    print(f"数据加载完成，总行数: {len(df)}")
    return df

def get_db_columns(connection):
    """获取数据库表的列名"""
    cursor = connection.cursor()
    cursor.execute("DESCRIBE zongdan_2024")
    db_columns = [row[0] for row in cursor.fetchall()]
    cursor.close()
    print(f"数据库中的列名: {db_columns}")
    return db_columns

def convert_datetime_columns(df):
    """转换时间列为日期格式"""
    datetime_columns = ['到港时间', '搬入时间', '许可时间函数对应', '入库时间', '出库指示时间_入金确认时间', '转运日期']
    for col in datetime_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.date  # 转为date格式
    return df

def clear_table(connection):
    """清空数据库表"""
    cursor = connection.cursor()
    try:
        cursor.execute("DELETE FROM zongdan_2024")
        connection.commit()
        print("数据库表已清空")
    except Error as e:
        print(f"清空表时出错: {e}")
    finally:
        cursor.close()

def fill_missing_values(df):
    """为非空约束列提供默认值"""
    if '请求含税' in df.columns:
        df['请求含税'] = df['请求含税'].fillna(0)
    if '成本含税' in df.columns:
        df['成本含税'] = df['成本含税'].fillna(0)
    return df

def insert_data_to_db(connection, df):
    """将数据插入数据库"""
    cursor = connection.cursor()
    columns = ', '.join([f"`{col}`" for col in df.columns])
    placeholders = ', '.join(['%s'] * len(df.columns))
    insert_query = f"INSERT INTO zongdan_2024 ({columns}) VALUES ({placeholders})"

    inserted = False
    for index, row in df.iterrows():
        row_values = [None if pd.isna(val) else val for val in row]
        # print(f"第 {index + 1} 行数据: {row_values}")

        if not any(row_values):
            print(f"第 {index + 1} 行为空或无效，跳过插入。")
            continue

        try:
            cursor.execute(insert_query, tuple(row_values))
            inserted = True
        except Error as e:
            print(f"跳过重复条目: {e}")

    if inserted:
        connection.commit()
        print("数据已成功插入数据库")
    else:
        print("没有数据插入，所有条目均为重复或无效。")
    cursor.close()

def remove_empty_primary_keys(df, primary_key):
    """删除主键为空的行"""
    if primary_key not in df.columns:
        print(f"错误: 主键列 '{primary_key}' 不存在。")
        return pd.DataFrame()
    df = df[~df[primary_key].astype(str).str.strip().eq('')]
    df = df[df[primary_key].notna()]
    print(f"过滤后的数据行数: {len(df)}")
    return df

def main():
    connection = connect_db()
    if connection:
        try:
            df = load_excel_data()

            # 打印 Excel 文件中的列名
            print(f"Excel 文件中的列名: {list(df.columns)}")

            # 获取数据库的列名
            db_columns = get_db_columns(connection)

            # 打印数据库中的列名
            print(f"数据库中的列名: {db_columns}")

            # 找出 Excel 和数据库列名的差异
            missing_columns = set(df.columns) - set(db_columns)
            if missing_columns:
                print(f"这些列在数据库中缺失: {missing_columns}")

            # 保留数据库中存在的列
            df = df[df.columns.intersection(db_columns)]

            primary_key = '送り状番号'
            df = remove_empty_primary_keys(df, primary_key)

            if df.empty:
                print("没有有效数据可插入。")
            else:
                # 转换日期列
                datetime_columns = ['到港时间', '搬入时间', '许可时间函数对应', '入库时间', 
                                    '出库指示时间_入金确认时间', '转运日期']
                for col in datetime_columns:
                    if col in df.columns:
                        df[col] = pd.to_datetime(df[col], format='%Y-%m-%d', errors='coerce').dt.date
                
                df = fill_missing_values(df)  # 填充非空列
                clear_table(connection)
                insert_data_to_db(connection, df)
                
        except ValueError as ve:
            print(f"处理数据时出错: {ve}")
        except Error as e:
            print(f"数据库操作时出错: {e}")
        finally:
            connection.close()
            print("数据库连接已关闭")
        print("操作完成")

if __name__ == "__main__":
    main()