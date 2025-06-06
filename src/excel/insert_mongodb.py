import pandas as pd
from pymongo import MongoClient
from datetime import datetime
import math
from dotenv import load_dotenv
import os
from openpyxl import load_workbook

load_dotenv()  # 加载.env文件

def excel_to_mongodb(excel_file, mongodb_uri, db_name, collection_name):
    client = MongoClient(mongodb_uri)
    db = client[db_name]
    collection = db[collection_name]

    try:
        wb = load_workbook(excel_file, data_only=True)
        sheet = wb.active
        data = sheet.values
        columns = next(data)
        df = pd.DataFrame(data, columns=columns)
    except Exception as e:
        print(f"读取文件 {excel_file} 时出错: {e}")
        client.close()
        return

    df.columns = df.columns.str.strip().str.replace(r'\s+', '_', regex=True)
    df['村委'] = df['村委'].ffill()

    documents = []

    for _, row in df.iterrows():
        def clean_value(value):
            if pd.isna(value):
                return None
            if isinstance(value, (int, float)) and math.isnan(value):
                return None
            return value

        doc = {
            "township": clean_value(row.get('乡镇')),
            "village": clean_value(row.get('村委')),
            "risk_date": clean_value(row.get('出险时间')),
            "growth_stage": clean_value(row.get('出险时间对应生长时期')),
            "loss_level": clean_value(row.get('报损程度')),
            "farmer_name": clean_value(row.get('抽样农户名称')),
            "plot_name": clean_value(row.get('地块名称')),
            "average_spikes_per_mu": clean_value(row.get('平均亩穗（万/亩）')),
            "average_grains_per_spike": clean_value(row.get('平均穗粒数（粒/穗）')),
            "thousand_grain_weight": clean_value(row.get('平均千粒重（克）')),
            "current_yield_kg_per_mu": clean_value(row.get('抽样地块平均产量（kg/亩）')),
            "historical_yield_kg_per_mu": clean_value(row.get('当地前三年平均产量（kg/亩）')),
            "loss_percentage": clean_value(row.get('损失程度%')),
            "avg_loss_same_level": clean_value(row.get('相同报损程度平均损失率%')),
            "source_file": os.path.basename(excel_file),
            "import_date": datetime.now(),
            "is_calculated_yield": False,
            "is_calculated_loss": False
        }
        print(row)

        documents.append(doc)

    if documents:
        try:
            result = collection.insert_many(documents)
            print(f"✅ 成功插入 {os.path.basename(excel_file)} 中的 {len(result.inserted_ids)} 条记录。")
        except Exception as e:
            print(f"❌ 插入文档出错: {e}")
    else:
        print(f"⚠️ 没有可插入的数据: {excel_file}")

    client.close()


def create_mongodb_indexes(mongodb_uri, db_name, collection_name):
    client = MongoClient(mongodb_uri)
    db = client[db_name]
    collection = db[collection_name]

    print("正在创建MongoDB索引...")
    collection.create_index([("township", 1)])
    collection.create_index([("village", 1)])
    collection.create_index([("risk_date", 1)])
    collection.create_index([("farmer_name", 1)])
    collection.create_index([("loss_percentage", 1)])
    print("✅ 索引创建完成。")
    client.close()


def main():
    mongodb_uri = os.getenv("MONGODB_URI")
    db_name = os.getenv("DB_NAME")
    collection_name = os.getenv("COLLECTION_NAME")
    excel_directory = os.getenv("OUTPUT_DIRECTORY")

    if not all([mongodb_uri, db_name, collection_name, excel_directory]):
        print("❌ .env 配置项不完整，请确保包含 MONGODB_URI、DB_NAME、COLLECTION_NAME 和 OUTPUT_DIRECTORY。")
        return

    create_mongodb_indexes(mongodb_uri, db_name, collection_name)

    for filename in os.listdir(excel_directory):
        if filename.endswith(".xls") or filename.endswith(".xlsx"):
            file_path = os.path.join(excel_directory, filename)
            print(f"📄 正在处理: {file_path}")
            excel_to_mongodb(file_path, mongodb_uri, db_name, collection_name)


if __name__ == "__main__":
    main()
