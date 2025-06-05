import pandas as pd
from pymongo import MongoClient
from datetime import datetime
import math
import os

def excel_to_mongodb(excel_file, mongodb_uri, db_name, collection_name):
    """
    将Excel农业损失数据导入MongoDB，采用扁平化数据结构。
    
    参数:
        excel_file: Excel文件路径
        mongodb_uri: MongoDB连接字符串
        db_name: 数据库名称
        collection_name: 集合名称
    """
    # Connect to MongoDB
    client = MongoClient(mongodb_uri)
    db = client[db_name]
    collection = db[collection_name]
    
    # Read Excel file
    try:
        df = pd.read_excel(excel_file, parse_dates=['出险时间'])
    except Exception as e:
        print(f"Error reading {excel_file}: {e}")
        client.close()
        return
    
    # --- Start of modifications for column cleaning ---
    # Clean up column names: strip whitespace, newlines, and replace spaces with underscores
    df.columns = df.columns.str.strip().str.replace(r'\s+', '_', regex=True)
    # --- End of modifications for column cleaning ---

    # Preprocess data
    # Forward-fill empty '村委' cells
    # Use the cleaned column name
    df['村委'] = df['村委'].ffill()
    
    # Calculate average loss rate for the same '报损程度'
    # Use the cleaned column names
    df['损失程度%'] = pd.to_numeric(df['损失程度%'], errors='coerce')
    df['相同报损程度平均损失率%'] = df.groupby('报损程度')['损失程度%'].transform('mean')
    
    # Convert data to flat MongoDB document format
    documents = []
    for _, row in df.iterrows():
        # Handle potential NaN values and specific data types
        def clean_value(value):
            if pd.isna(value):
                return None
            if isinstance(value, (int, float)) and math.isnan(value):
                return None
            return value
            
        doc = {
            # Basic Info - using cleaned column names
            "township": clean_value(row['乡镇']),
            "village": clean_value(row['村委']),
            "risk_date": clean_value(row['出险时间']),
            "growth_stage": clean_value(row['出险时间对应生长时期']),
            "loss_level": clean_value(row['报损程度']),
            
            # Sampling Info - using cleaned column names
            "farmer_name": clean_value(row['抽样农户名称']),
            "plot_name": clean_value(row['地块名称']),
            "average_spikes_per_mu": clean_value(row['平均亩穗（万/亩）']),
            "average_grains_per_spike": clean_value(row['平均穗粒数（粒/穗）']),
            "thousand_grain_weight": clean_value(row['平均千粒重（克）']),
            
            # Yield Data - using cleaned column names
            "current_yield_kg_per_mu": clean_value(row['抽样地块平均产量（kg/亩）']),
            "historical_yield_kg_per_mu": clean_value(row['当地前三年平均产量（kg/亩）']),
            "loss_percentage": clean_value(row['损失程度%']),
            
            # Statistics - using cleaned column name
            "avg_loss_same_level": clean_value(row['相同报损程度平均损失率%']),
            
            # Metadata
            "source_file": os.path.basename(excel_file),
            "import_date": datetime.now(),
            # These flags usually require inspecting the original Excel cell formula,
            # which pandas.read_excel doesn't expose directly.
            # For now, we'll set them to False as a default, or you'd need a more
            # advanced library like openpyxl to read cell formulas.
            "is_calculated_yield": False, # Cannot reliably detect from pd.read_excel without formula access
            "is_calculated_loss": False   # Cannot reliably detect from pd.read_excel without formula access
        }
        documents.append(doc)
    
    # Bulk insert data
    if documents:
        try:
            result = collection.insert_many(documents)
            print(f"Successfully inserted {len(result.inserted_ids)} records from {os.path.basename(excel_file)}")
        except Exception as e:
            print(f"Error inserting documents from {os.path.basename(excel_file)}: {e}")
    else:
        print(f"No data to insert from {os.path.basename(excel_file)}")
    
    client.close()


def create_mongodb_indexes(mongodb_uri, db_name, collection_name):
    """
    创建MongoDB集合的索引。
    """
    client = MongoClient(mongodb_uri)
    db = client[db_name]
    collection = db[collection_name]

    print("Creating MongoDB indexes...")
    # Using the flat structure for indexes
    collection.create_index([("township", 1)])
    collection.create_index([("village", 1)])
    collection.create_index([("risk_date", 1)])
    collection.create_index([("farmer_name", 1)])
    collection.create_index([("loss_percentage", 1)])
    print("Indexes created.")
    client.close()


def main():
    excel_directory = "/Users/a1/理赔文件/temp/" # Directory containing the Excel files
    mongodb_uri = "mongodb://localhost:27017/"
    db_name = "agricultural_insurance"
    collection_name = "loss_records"

    # Create indexes once before processing files
    create_mongodb_indexes(mongodb_uri, db_name, collection_name)

    # Iterate through files in the specified directory
    for filename in os.listdir(excel_directory):
        if filename.endswith(".xls") or filename.endswith(".xlsx"):
            excel_file_path = os.path.join(excel_directory, filename)
            print(f"Processing file: {excel_file_path}")
            excel_to_mongodb(excel_file_path, mongodb_uri, db_name, collection_name)

if __name__ == "__main__":
    main()