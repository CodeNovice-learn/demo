import pandas as pd
import os
from typing import List, Dict, Union, Optional

class ExcelReader:
    """
    Excel文件读取插件类
    支持读取Excel文件内容并返回结构化数据
    """
    
    def __init__(self):
        self.supported_formats = ['.xlsx', '.xls', '.xlsm']
    
    def read_excel(self, file_path: str, sheet_name: Union[str, int, List, None] = None, 
                   header: Union[int, List[int]] = 0, index_col: Union[int, None] = None) -> Dict:
        """
        读取Excel文件
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称或索引，None表示读取所有工作表
            header: 行索引，指定哪一行作为列名
            index_col: 列索引，指定哪一列作为索引
            
        Returns:
            Dict: 包含Excel数据的字典
            
        Raises:
            FileNotFoundError: 文件不存在
            ValueError: 文件格式不支持
        """
        # 检查文件是否存在
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        # 检查文件格式
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in self.supported_formats:
            raise ValueError(f"不支持的文件格式: {file_ext}. 支持格式: {self.supported_formats}")
        
        try:
            # 使用pandas读取Excel文件
            if sheet_name is None:
                # 读取所有工作表
                excel_data = pd.read_excel(file_path, sheet_name=None, header=header, index_col=index_col)
                result = {
                    'sheets': {},
                    'total_sheets': 1
                }
                
                # 处理返回的数据结构
                if isinstance(excel_data, dict):
                    # 多个工作表的情况
                    result['total_sheets'] = len(excel_data)
                    for sheet_name_key, data in excel_data.items():
                        result['sheets'][sheet_name_key] = {
                            'data': data.to_dict('records'),
                            'columns': list(data.columns),
                            'shape': data.shape,
                            'sheet_name': sheet_name_key
                        }
                else:
                    # 单个工作表的情况
                    result['sheets']['Sheet1'] = {
                        'data': excel_data.to_dict('records'),
                        'columns': list(excel_data.columns),
                        'shape': excel_data.shape,
                        'sheet_name': 'Sheet1'
                    }
            else:
                # 读取指定工作表
                data = pd.read_excel(file_path, sheet_name=sheet_name, header=header, index_col=index_col)
                result = {
                    'sheets': {
                        sheet_name if isinstance(sheet_name, str) else f"Sheet_{sheet_name}": {
                            'data': data.to_dict('records'),
                            'columns': list(data.columns),
                            'shape': data.shape,
                            'sheet_name': sheet_name if isinstance(sheet_name, str) else f"Sheet_{sheet_name}"
                        }
                    },
                    'total_sheets': 1
                }
            
            return result
            
        except Exception as e:
            raise ValueError(f"读取Excel文件失败: {str(e)}")
    
    def get_sheet_info(self, file_path: str) -> List[Dict]:
        """
        获取Excel文件中所有工作表的信息
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            List[Dict]: 工作表信息列表
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        try:
            # 获取所有工作表名称
            excel_file = pd.ExcelFile(file_path)
            sheets_info = []
            
            for sheet_name in excel_file.sheet_names:
                # 读取工作表获取行列数
                data = pd.read_excel(excel_file, sheet_name=sheet_name)
                sheets_info.append({
                    'sheet_name': sheet_name,
                    'rows': len(data),
                    'columns': len(data.columns),
                    'shape': data.shape
                })
            
            return sheets_info
            
        except Exception as e:
            raise ValueError(f"获取工作表信息失败: {str(e)}")
    
    def search_data(self, file_path: str, search_text: str, sheet_name: Union[str, None] = None) -> List[Dict]:
        """
        在Excel文件中搜索指定文本
        
        Args:
            file_path: Excel文件路径
            search_text: 搜索文本
            sheet_name: 工作表名称，None表示在所有工作表中搜索
            
        Returns:
            List[Dict]: 包含搜索结果的列表
        """
        try:
            result = []
            
            if sheet_name is None:
                # 在所有工作表中搜索
                excel_data = pd.read_excel(file_path, sheet_name=None)
                for sheet_name_key, data in excel_data.items():
                    for row_idx, row in data.iterrows():
                        for col_idx, cell_value in enumerate(row):
                            if search_text in str(cell_value):
                                result.append({
                                    'sheet_name': sheet_name_key,
                                    'row': row_idx + 1,  # Excel行号从1开始
                                    'column': col_idx + 1,  # Excel列号从1开始
                                    'value': cell_value,
                                    'column_name': data.columns[col_idx]
                                })
            else:
                # 在指定工作表中搜索
                data = pd.read_excel(file_path, sheet_name=sheet_name)
                for row_idx, row in data.iterrows():
                    for col_idx, cell_value in enumerate(row):
                        if search_text in str(cell_value):
                            result.append({
                                'sheet_name': sheet_name,
                                'row': row_idx + 1,
                                'column': col_idx + 1,
                                'value': cell_value,
                                'column_name': data.columns[col_idx]
                            })
            
            return result
            
        except Exception as e:
            raise ValueError(f"搜索数据失败: {str(e)}")
    
    def validate_file(self, file_path: str) -> Dict:
        """
        验证Excel文件格式和内容
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            Dict: 验证结果
        """
        if not os.path.exists(file_path):
            return {'valid': False, 'error': '文件不存在'}
        
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in self.supported_formats:
            raise ValueError(f'不支持的文件格式: {file_ext}')
        
        try:
            excel_file = pd.ExcelFile(file_path)
            data = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0])
            
            return {
                'valid': True,
                'file_path': file_path,
                'file_size': os.path.getsize(file_path),
                'total_sheets': len(excel_file.sheet_names),
                'sheet_names': excel_file.sheet_names,
                'first_sheet_shape': data.shape,
                'first_sheet_columns': list(data.columns),
                'is_readable': True
            }
            
        except Exception as e:
            return {
                'valid': False,
                'error': f'文件验证失败: {str(e)}'
            }


def main():
    """
    主函数 - 使用示例
    """
    # 创建Excel读取器实例
    reader = ExcelReader()
    
    # 示例文件路径（需要替换为实际的Excel文件路径）
    example_file = "example.xlsx"
    
    print("=== Excel文件读取插件使用示例 ===\n")
    
    try:
        # 1. 验证文件
        print("1. 验证Excel文件...")
        validation_result = reader.validate_file(example_file)
        
        if validation_result['valid']:
            print(f"[成功] 文件验证成功")
            print(f"   文件大小: {validation_result['file_size']} 字节")
            print(f"   工作表数量: {validation_result['total_sheets']}")
            print(f"   工作表名称: {validation_result['sheet_names']}")
            print(f"   第一工作表形状: {validation_result['first_sheet_shape']}")
            
            # 2. 获取工作表信息
            print("\n2. 获取工作表信息...")
            sheets_info = reader.get_sheet_info(example_file)
            for sheet in sheets_info:
                print(f"   工作表 '{sheet['sheet_name']}': {sheet['rows']} 行 × {sheet['columns']} 列")
            
            # 3. 读取Excel数据
            print("\n3. 读取Excel数据...")
            excel_data = reader.read_excel(example_file)
            print(f"   读取到 {excel_data['total_sheets']} 个工作表")
            
            # 打印第一个工作表的前5行数据
            first_sheet_name = list(excel_data['sheets'].keys())[0]
            first_sheet_data = excel_data['sheets'][first_sheet_name]['data']
            print(f"   第一工作表 '{first_sheet_name}' 的前5行数据:")
            for i, row in enumerate(first_sheet_data[:5]):
                print(f"     行 {i+1}: {row}")
            
            # 4. 搜索数据
            print("\n4. 搜索数据...")
            search_results = reader.search_data(example_file, "测试")
            print(f"   找到 {len(search_results)} 个包含 '测试' 的单元格")
            for result in search_results[:3]:  # 只显示前3个结果
                print(f"     工作表: {result['sheet_name']}, 行: {result['row']}, 列: {result['column']}, 值: {result['value']}")
                
        else:
            print(f"[失败] 文件验证失败: {validation_result['error']}")
            
    except Exception as e:
        print(f"[失败] 错误: {str(e)}")


if __name__ == "__main__":
    main()