import pandas as pd
import argparse
from enum import Enum
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

class VehicleType(Enum):
    COMPLETE = '整车'
    INCOMPLETE = '非整车'

def classify_vehicle_type(file_path, sheet_names, columns_to_check, keyword, output_path):
    writer = pd.ExcelWriter(output_path, engine='openpyxl')

    # 读取 Excel 文件以获取所有 sheet 名称
    xls = pd.ExcelFile(file_path)

    for sheet_name in sheet_names:
        if sheet_name not in xls.sheet_names:
            print(f"Sheet '{sheet_name}' 不存在，跳过该 sheet")
            continue

        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"Processing sheet: {sheet_name}")
        print(f"DataFrame columns: {df.columns.tolist()}")

        # 检查是否存在指定的列
        valid_columns = []
        for column in columns_to_check:
            if column in df.columns:
                valid_columns.append(column)
            else:
                print(f"Column '{column}' 不存在于 sheet '{sheet_name}'，跳过该列")

        # 初始化类型列为 '非整车'
        df['类型'] = VehicleType.INCOMPLETE.value

        # 遍历指定的列，检查是否包含关键字
        for index, row in df.iterrows():
            for column in valid_columns:
                if keyword in str(row[column]):
                    df.at[index, '类型'] = VehicleType.COMPLETE.value
                    break

        # 将DataFrame写入Excel文件
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    writer.book.save(output_path)

    # 打开写入类型列后的Excel文件，添加数据验证
    wb = load_workbook(output_path)

    for sheet_name in sheet_names:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        dv = DataValidation(type="list", formula1='"整车,非整车"', showDropDown=True)

        # 添加数据验证到类型列
        for row in range(2, ws.max_row + 1):
            cell = ws[f"{chr(65 + df.columns.get_loc('类型'))}{row}"]
            dv.add(cell)

        ws.add_data_validation(dv)

    wb.save(output_path)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='对 Excel 文件的某几列进行关键字匹配，并添加类型列')
    parser.add_argument('file_path', help='输入的 Excel 文件路径')
    parser.add_argument('-s', '--sheets', required=True, help='要处理的 sheet 名称列表，用逗号分隔')
    parser.add_argument('-c', '--columns', required=True, help='需要检查的列名列表，用逗号分隔')
    parser.add_argument('keyword', help='要匹配的关键字')
    parser.add_argument('output_path', help='输出的 Excel 文件路径')

    args = parser.parse_args()

    sheet_names = args.sheets.split(',')
    columns_to_check = args.columns.split(',')

    classify_vehicle_type(args.file_path, sheet_names, columns_to_check, args.keyword, args.output_path)
