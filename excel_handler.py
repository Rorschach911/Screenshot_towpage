import pandas as pd

def read_excel(file_path, sheet_name='Sheet1'):
    """
    读取Excel文件并返回DataFrame
    
    参数:
    file_path - Excel文件路径
    sheet_name - 工作表名称，默认为'Sheet1'
    
    返回:
    包含Excel数据的DataFrame
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        # 确保必要的列存在
        required_columns = ["媒体名称","发布时间","链接"]
        missing_columns = validate_excel_columns(df, required_columns)
        if missing_columns:
            raise Exception(f"Excel文件中缺少以下列: {', '.join(missing_columns)}")
        return df
    except Exception as e:
        raise Exception(f"读取Excel文件时出错: {str(e)}")

def validate_excel_columns(df, required_columns):
    """
    验证DataFrame是否包含所需的列
    
    参数:
    df - 要验证的DataFrame
    required_columns - 必需的列名列表
    
    返回:
    缺失的列名列表，如果没有缺失则返回空列表
    """
    missing_columns = [col for col in required_columns if col not in df.columns]
    return missing_columns

def get_media_info_from_excel(file_path, sheet_name='Sheet1'):
    """
    从Excel文件中提取媒体信息（媒体名称、链接和发布时间）
    
    参数:
    file_path - Excel文件路径
    sheet_name - 工作表名称，默认为'Sheet1'
    
    返回:
    包含媒体信息的DataFrame
    """
    df = read_excel(file_path, sheet_name)
    required_columns = ["媒体名称","发布时间","链接"]
    for col in required_columns:
        if col not in df.columns:
            raise Exception(f"Excel文件中不存在列 '{col}'")
    
    return df[required_columns]