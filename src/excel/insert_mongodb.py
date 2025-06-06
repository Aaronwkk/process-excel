import pandas as pd
from pymongo import MongoClient
from datetime import datetime
import math
from dotenv import load_dotenv
import os

load_dotenv() # 这会加载 .env 文件中的所有变量到 os.environ

def excel_to_mongodb(excel_file, mongodb_uri, db_name, collection_name):
    """
    将Excel农业损失数据导入MongoDB，采用扁平化数据结构。
    
    参数:
        excel_file: Excel文件路径
        mongodb_uri: MongoDB连接字符串
        db_name: 数据库名称
        collection_name: 集合名称
    """
    # 连接到MongoDB
    client = MongoClient(mongodb_uri)
    db = client[db_name]
    collection = db[collection_name]
    
    # 读取Excel文件
    try:
        df = pd.read_excel(excel_file, parse_dates=['出险时间'])
    except Exception as e:
        print(f"读取文件 {excel_file} 时出错: {e}")
        client.close()
        return
    
    # --- 列名清理修改开始 ---
    # 清理列名：去除空格、换行符，并将空格替换为下划线
    df.columns = df.columns.str.strip().str.replace(r'\s+', '_', regex=True)

    # --- 列名清理修改结束 ---

    # 预处理数据
    # 向前填充空的“村委”单元格
    # 使用清理后的列名
    df['村委'] = df['村委'].ffill()
    
    # 将数据转换为扁平的MongoDB文档格式
    documents = []
    for _, row in df.iterrows():
        # 处理可能的NaN值和特定数据类型
        def clean_value(value):
            if pd.isna(value):
                return None
            if isinstance(value, (int, float)) and math.isnan(value):
                return None
            return value
            
        doc = {
            # 基本信息 - 使用清理后的列名
            "township": clean_value(row['乡镇']),
            "village": clean_value(row['村委']),
            "risk_date": clean_value(row['出险时间']),
            "growth_stage": clean_value(row['出险时间对应生长时期']),
            "loss_level": clean_value(row['报损程度']),
            
            # 抽样信息 - 使用清理后的列名
            "farmer_name": clean_value(row['抽样农户名称']),
            "plot_name": clean_value(row['地块名称']),
            "average_spikes_per_mu": clean_value(row['平均亩穗（万/亩）']),
            "average_grains_per_spike": clean_value(row['平均穗粒数（粒/穗）']),
            "thousand_grain_weight": clean_value(row['平均千粒重（克）']),
            
            # 产量数据 - 使用清理后的列名
            "current_yield_kg_per_mu": clean_value(row['抽样地块平均产量（kg/亩）']),
            "historical_yield_kg_per_mu": clean_value(row['当地前三年平均产量（kg/亩）']),
            "loss_percentage": clean_value(row['损失程度%']),
            
            # 统计数据 - 使用清理后的列名
            "avg_loss_same_level": clean_value(row['相同报损程度平均损失率%']),
            
            # 元数据
            "source_file": os.path.basename(excel_file),
            "import_date": datetime.now(),
            # 这些标志通常需要检查原始Excel单元格公式，
            # 而 pandas.read_excel 不直接暴露。
            # 目前，我们默认将它们设置为False，或者您需要更
            # 高级的库（如openpyxl）来读取单元格公式。
            "is_calculated_yield": False, # 在没有公式访问权限的情况下，无法从pd.read_excel可靠检测
            "is_calculated_loss": False   # 在没有公式访问权限的情况下，无法从pd.read_excel可靠检测
        }
        documents.append(doc)
    
    # 批量插入数据
    if documents:
        try:
            result = collection.insert_many(documents)
            print(f"成功插入 {os.path.basename(excel_file)} 中的 {len(result.inserted_ids)} 条记录。")
        except Exception as e:
            print(f"插入文件 {os.path.basename(excel_file)} 的文档时出错: {e}")
    else:
        print(f"文件 {os.path.basename(excel_file)} 没有可插入的数据。")
    
    client.close()


def create_mongodb_indexes(mongodb_uri, db_name, collection_name):
    """
    创建MongoDB集合的索引。
    
    参数:
        mongodb_uri: MongoDB连接字符串
        db_name: 数据库名称
        collection_name: 集合名称
    """
    client = MongoClient(mongodb_uri)
    db = client[db_name]
    collection = db[collection_name]

    print("正在创建MongoDB索引...")
    # 为扁平化结构创建索引
    collection.create_index([("township", 1)])
    collection.create_index([("village", 1)])
    collection.create_index([("risk_date", 1)])
    collection.create_index([("farmer_name", 1)])
    collection.create_index([("loss_percentage", 1)])
    print("索引创建完成。")
    client.close()


def main():

    mongodb_uri = os.getenv("MONGODB_URI")
    db_name = os.getenv("DB_NAME")
    collection_name = os.getenv("COLLECTION_NAME")
    excel_directory = os.getenv("OUTPUT_DIRECTORY") # 存放Excel文件的目录

    # 在处理文件之前，先创建一次索引
    create_mongodb_indexes(mongodb_uri, db_name, collection_name)

    # 遍历指定目录中的文件
    for filename in os.listdir(excel_directory):
        if filename.endswith(".xls") or filename.endswith(".xlsx"):
            excel_file_path = os.path.join(excel_directory, filename)
            print(f"正在处理文件: {excel_file_path}")
            excel_to_mongodb(excel_file_path, mongodb_uri, db_name, collection_name)

if __name__ == "__main__":
    main()
