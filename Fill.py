import pandas as pd
import json
import requests

# 读取CSV或Excel文件并返回表头
def read_csv_or_excel(file_path):
    try:
        df = pd.read_csv(file_path, encoding='utf-8-sig')
        return df.columns.tolist(), 'csv'
    except Exception:
        try:
            df = pd.read_excel(file_path)
            return df.columns.tolist(), 'excel'
        except Exception as ex:
            raise Exception(f"无法读取文件 {file_path}: {ex}")

def download_json_from_url(url):
    try:
        if url.startswith("http://") or url.startswith("https://"):
            response = requests.get(url)
            response.raise_for_status()  # 检查请求是否成功
            json_data = response.text
        else:
            with open(url, 'r', encoding='utf-8') as file:
                json_data = file.read()
        return parse_json(json_data)
    except (requests.RequestException, FileNotFoundError, IOError) as e:
        raise Exception(f"无法从URL或文件下载JSON数据: {e}")
        
# 解析JSON数据并返回字典列表
def parse_json(json_data):
    try:
        data = json.loads(json_data)
        if isinstance(data, dict):
            data = [data]
        return data
    except json.JSONDecodeError as e:
        raise Exception(f"无效的JSON数据: {e}")

# 根据表头将JSON数据填入DataFrame
def fill_dataframe(headers, json_data):
    rows = [{header: item.get(header, None) for header in headers} for item in json_data]
    df = pd.DataFrame(rows)
    return df

# 将DataFrame保存为CSV或Excel文件
def save_to_csv_or_excel(df, output_path, file_type):
    if file_type == 'csv':
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
    elif file_type == 'excel':
        df.to_excel(output_path, index=False)
    else:
        raise ValueError("文件类型必须是 'csv' 或 'excel'")

# 主函数，用于处理文件上传和JSON数据下载
def process_upload_and_json(file, json_url, output_path, desired_output_type):
    try:
        headers, file_type = read_csv_or_excel(file)
        print(f"读取的表头: {headers}, 文件类型: {file_type}")
    except Exception as e:
        print(f"读取文件时发生错误: {e}")
        return

    try:
        json_data_parsed = download_json_from_url(json_url)
        print(f"解析的JSON数据: {json_data_parsed}")
    except Exception as e:
        print(f"下载或解析JSON数据时发生错误: {e}")
        return

    try:
        df = fill_dataframe(headers, json_data_parsed)
        print(f"填充后的表格:\n{df}")
    except Exception as e:
        print(f"填充DataFrame时发生错误: {e}")
        return

    try:
        save_to_csv_or_excel(df, output_path, desired_output_type if desired_output_type in ['csv', 'excel'] else file_type)
        print(f"数据已成功保存到 {output_path}")
    except Exception as e:
        print(f"保存文件时发生错误: {e}")

if __name__ == "__main__":
    
    input_file = 'test.csv'  # 或 'test.xlsx'
    # json_url = 'http://'
    json_url = r'C:\Users\mooncell\Desktop\test\test.json'
    output_file = 'output_data.csv'  # 或 'output_data.xlsx'
    output_type = 'csv'  # 或 'excel'

    process_upload_and_json(input_file, json_url, output_file, output_type)