# Excel读取插件

一个功能强大的Python Excel文件读取插件，支持读取.xlsx、.xls、.xlsm格式的Excel文件。

## 功能特性

- ✅ 支持读取Excel文件内容
- ✅ 获取工作表信息和结构
- ✅ 支持搜索Excel中的文本
- ✅ 文件格式验证
- ✅ 错误处理和异常捕获
- ✅ 支持多工作表读取
- ✅ 类型注解和完整文档

## 安装依赖

```bash
pip install -r requirements.txt
```

依赖包：
- pandas>=1.5.0
- openpyxl>=3.0.0
- xlrd>=2.0.0

## 使用方法

### 基本使用

```python
from excel_reader import ExcelReader

# 创建读取器实例
reader = ExcelReader()

# 验证文件
validation = reader.validate_file("your_file.xlsx")
if validation['valid']:
    print("文件验证成功")
    
# 读取Excel文件
data = reader.read_excel("your_file.xlsx")
print(f"读取到 {data['total_sheets']} 个工作表")

# 获取工作表信息
sheets_info = reader.get_sheet_info("your_file.xlsx")
for sheet in sheets_info:
    print(f"工作表 {sheet['sheet_name']}: {sheet['rows']} 行 × {sheet['columns']} 列")

# 搜索数据
search_results = reader.search_data("your_file.xlsx", "搜索文本")
print(f"找到 {len(search_results)} 个匹配项")
```

### 详细API

#### ExcelReader()

初始化Excel读取器实例。

#### read_excel(file_path, sheet_name=None, header=0, index_col=None)

- `file_path`: Excel文件路径
- `sheet_name`: 工作表名称或索引，None表示读取所有工作表
- `header`: 行索引，指定哪一行作为列名
- `index_col`: 列索引，指定哪一列作为索引

#### get_sheet_info(file_path)

获取Excel文件中所有工作表的基本信息。

#### search_data(file_path, search_text, sheet_name=None)

在Excel文件中搜索指定文本。

#### validate_file(file_path)

验证Excel文件格式和内容是否有效。

## 运行测试

```bash
python test_excel_reader.py
```

## 支持的文件格式

- .xlsx - Excel 2007+
- .xls - Excel 97-2003
- .xlsm - Excel启用宏的工作簿

## 返回数据格式

### read_excel() 返回格式

```python
{
    'sheets': {
        'Sheet1': {
            'data': [{'列1': '值1', '列2': '值2'}, ...],
            'columns': ['列1', '列2', ...],
            'shape': (行数, 列数),
            'sheet_name': 'Sheet1'
        }
    },
    'total_sheets': 1
}
```

### get_sheet_info() 返回格式

```python
[
    {
        'sheet_name': 'Sheet1',
        'rows': 10,
        'columns': 5,
        'shape': (10, 5)
    }
]
```

### search_data() 返回格式

```python
[
    {
        'sheet_name': 'Sheet1',
        'row': 1,
        'column': 1,
        'value': '找到的值',
        'column_name': '列名'
    }
]
```

## 示例代码

参见 `test_excel_reader.py` 中的完整示例。

## 错误处理

插件会抛出以下异常：
- `FileNotFoundError`: 文件不存在
- `ValueError`: 文件格式不支持或读取失败

## 许可证

MIT License