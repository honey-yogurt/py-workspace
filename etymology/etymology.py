import pandas as pd
import yaml
from openpyxl import load_workbook
import os

def classify_vehicle_type(input_path, sheet_names, columns_to_check_map, output_path, new_column_name):
    writer = pd.ExcelWriter(output_path, engine='openpyxl')

    # 读取 Excel 文件以获取所有 sheet 名称
    xls = pd.ExcelFile(input_path)

    total_processed_rows = 0  # 总处理行数

    for sheet_name in sheet_names:
        if sheet_name not in xls.sheet_names:
            print(f"Sheet '{sheet_name}' 不存在，跳过该 sheet")
            continue

        df = pd.read_excel(input_path, sheet_name=sheet_name)
        print(f"Processing sheet: {sheet_name}")
        print(f"DataFrame columns: {df.columns.tolist()}")

        # 检查新列名是否存在，如果不存在则添加
        if new_column_name not in df.columns:
            df[new_column_name] = None

        print(f"New column name to be used: {new_column_name}")

        # 创建一个包含所有需要匹配的列的DataFrame
        combined_columns_df = df[list(columns_to_check_map.keys())].astype(str)

        # 创建一个全为空的Series用于存储结果
        result_series = pd.Series([None] * len(df), index=df.index)

        # 遍历指定的列，检查是否包含关键字
        for column, keyword_map in columns_to_check_map.items():
            if column not in df.columns:
                print(f"Column '{column}' 不存在于 sheet '{sheet_name}'，跳过该列")
                continue

            # 创建一个空的Series用于存储当前列的匹配结果
            match_series = pd.Series([None] * len(df), index=df.index)

            for keyword, vehicle_type in keyword_map.items():
                match_indices = combined_columns_df[column].str.contains(keyword)
                match_series.loc[match_indices] = vehicle_type

            # 仅对result_series中为空的项进行更新
            result_series = result_series.combine_first(match_series)

        # 将结果Series更新到DataFrame的目标列中
        df[new_column_name] = result_series

        # 打印处理进度
        total_rows = len(df)
        total_processed_rows += total_rows
        for i in range(0, total_rows, 1000):
            print(f"Processed {min(i + 1000, total_rows)} rows out of {total_rows} in sheet '{sheet_name}'")

        # 将 DataFrame 写入 Excel 文件
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Finished processing {total_rows} rows in sheet '{sheet_name}'.")

    # 保存所有处理后的数据到 Excel
    writer.book.save(output_path)

    print(f"Total rows processed: {total_processed_rows}")
    print("Success!")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description='对 Excel 文件的某几列进行关键字匹配，并添加业务类型列')
    parser.add_argument('-c', '--config_file', help='配置文件的路径', default='config.yaml')

    args = parser.parse_args()

    config_file = args.config_file

    # 检查配置文件是否存在
    if not os.path.exists(config_file):
        raise FileNotFoundError(f"配置文件 {config_file} 不存在")

    # 读取配置文件
    with open(config_file, 'r', encoding='utf-8') as file:
        config = yaml.safe_load(file)

    input_path = config['input_path']
    sheet_names = config['sheet_names']
    columns_to_check_map = config['columns_to_check_map']
    output_path = config['output_path']
    new_column_name = config['new_column_name']

    classify_vehicle_type(input_path, sheet_names, columns_to_check_map, output_path, new_column_name)

    input("Press Enter to exit...")
