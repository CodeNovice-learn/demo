import sys
import os
import pandas as pd

# 添加当前目录到Python路径，以便导入模块
sys.path.append(os.path.dirname(__file__))

from excel_reader import ExcelReader

def test_excel_reader():
    """
    测试Excel读取插件
    """
    print("=== Excel读取插件测试 ===\n")
    
    # 创建测试Excel文件
    print("1. 创建测试Excel文件...")
    test_data = {
        '姓名': ['张三', '李四', '王五', '赵六'],
        '年龄': [25, 30, 35, 28],
        '部门': ['技术部', '市场部', '财务部', '人事部'],
        '入职日期': pd.to_datetime(['2020-01-01', '2019-03-15', '2018-07-20', '2021-02-10'])
    }
    
    df = pd.DataFrame(test_data)
    test_file = "test_employee.xlsx"
    df.to_excel(test_file, index=False)
    
    try:
        # 创建Excel读取器实例
        reader = ExcelReader()
        
        # 测试1: 验证文件
        print("\n2. 测试文件验证...")
        validation = reader.validate_file(test_file)
        assert validation['valid'], f"文件验证失败: {validation.get('error')}"
        print("[成功] 文件验证通过")
        
        # 测试2: 获取工作表信息
        print("\n3. 测试获取工作表信息...")
        sheets_info = reader.get_sheet_info(test_file)
        assert len(sheets_info) > 0, "没有获取到工作表信息"
        print(f"[成功] 成功获取 {len(sheets_info)} 个工作表")
        
        # 测试3: 读取Excel数据
        print("\n4. 测试读取Excel数据...")
        data = reader.read_excel(test_file)
        assert data['total_sheets'] > 0, "没有读取到工作表"
        print(f"Available sheets: {list(data['sheets'].keys())}")
        sheet_name = list(data['sheets'].keys())[0]
        assert len(data['sheets'][sheet_name]['data']) == 4, f"数据行数不正确，预期4行，实际{len(data['sheets'][sheet_name]['data'])}行"
        print("[成功] 成功读取Excel数据")
        
        # 测试4: 搜索数据
        print("\n5. 测试搜索数据...")
        search_results = reader.search_data(test_file, '张三')
        assert len(search_results) > 0, "没有搜索到数据"
        print(f"[成功] 成功搜索到 {len(search_results)} 条结果")
        
        # 测试5: 错误处理
        print("\n6. 测试错误处理...")
        try:
            reader.read_excel("nonexistent_file.xlsx")
            assert False, "应该抛出FileNotFoundError"
        except FileNotFoundError:
            print("[成功] 正确处理了文件不存在的情况")
        
        # 测试6: 不支持的文件格式
        print("\n7. 测试不支持的文件格式...")
        # 创建一个不支持的文件格式来测试
        with open("test.txt", "w") as f:
            f.write("This is not an Excel file")
        try:
            reader.validate_file("test.txt")
            assert False, "应该抛出错误"
        except ValueError:
            print("[成功] 正确处理了不支持的文件格式")

        print("\n[完成] 所有测试通过！")

    finally:
        # 清理测试文件
        if os.path.exists(test_file):
            os.remove(test_file)
            print(f"\n[清理] 已清理测试文件: {test_file}")


if __name__ == "__main__":
    test_excel_reader()