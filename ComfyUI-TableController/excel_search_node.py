import openpyxl
import csv

class ExcelSearchNode:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "file_path": ("STRING", {"default": ""}),
                "keyword": ("STRING", {"default": ""}),
                "search_column_number": ("INT", {"default": 1, "min": 1, "max": 100, "step": 1}),
                "return_column_number": ("INT", {"default": 1, "min": 1, "max": 100, "step": 1}),
            }
        }

    RETURN_TYPES = ("STRING",)
    RETURN_NAMES = ("result",)
    CATEGORY = "excel"
    FUNCTION = "search_file"

    def search_file(self, file_path, keyword, search_column_number, return_column_number):
        # 将用户输入的列号（从1开始）转换为程序使用的索引（从0开始）
        search_index = search_column_number - 1
        return_index = return_column_number - 1

        # 根据文件扩展名判断文件类型
        file_lower = file_path.lower()
        try:
            if file_lower.endswith('.xlsx') or file_lower.endswith('.xls'):
                return self._search_excel(file_path, keyword, search_index, return_index, search_column_number, return_column_number)
            elif file_lower.endswith('.csv'):
                return self._search_csv(file_path, keyword, search_index, return_index, search_column_number, return_column_number)
            else:
                return (f"Error: Unsupported file format. Please use .xlsx, .xls or .csv files",)
        except FileNotFoundError:
            return (f"Error: File '{file_path}' not found",)
        except Exception as e:
            return (f"Error: {str(e)}",)

    def _search_excel(self, file_path, keyword, search_index, return_index, search_column_number, return_column_number):
        # 加载 Excel 工作簿
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook.active
        
        # 遍历行
        for row in sheet.iter_rows():
            if not row:  # 跳过空行
                continue
                
            # 获取搜索列的值
            if search_index >= len(row):
                return (f"Error: Search column {search_column_number} is out of range",)
            
            search_cell_value = str(row[search_index].value if row[search_index].value is not None else "")
            
            if search_cell_value == keyword:  # 在指定列中查找关键字
                if 0 <= return_index < len(row):
                    value = row[return_index].value
                    return (str(value) if value is not None else "",)
                else:
                    return (f"Error: Return column {return_column_number} is out of range",)
        
        return (f"Error: Keyword '{keyword}' not found in column {search_column_number}",)

    def _search_csv(self, file_path, keyword, search_index, return_index, search_column_number, return_column_number):
        # 尝试不同的编码格式
        encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as file:
                    csv_reader = csv.reader(file, delimiter=';')  # 使用分号作为分隔符
                    
                    for row in csv_reader:
                        if not row:  # 跳过空行
                            continue
                            
                        # 检查列索引是否有效
                        if search_index >= len(row):
                            return (f"Error: Search column {search_column_number} is out of range",)
                        
                        # 获取搜索列的值
                        search_cell_value = str(row[search_index] if row[search_index] is not None else "")
                        
                        if search_cell_value == keyword:
                            if 0 <= return_index < len(row):
                                value = row[return_index]
                                return (str(value) if value is not None else "",)
                            else:
                                return (f"Error: Return column {return_column_number} is out of range",)
                    
                    # 如果能成功读取文件但没找到关键字，就跳出编码尝试循环
                    return (f"Error: Keyword '{keyword}' not found in column {search_column_number}",)
                    
            except UnicodeDecodeError:
                continue  # 如果编码错误，尝试下一个编码
            
        return ("Error: Unable to read the CSV file with any supported encoding",) 