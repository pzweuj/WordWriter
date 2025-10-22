# WordWriter

[English](README.md) | 简体中文

## 项目简介

WordWriter 是一个用于处理 Word 文档模板的 Python 工具库，可以方便地替换文档中的预留标签。

支持替换段落文本、单元格文本、文本框、页眉页脚，还可以插入表格和图片。

## ✨ v4.0 新特性

- **🎯 面向对象 API**：全新的 `WordWriter` 类，支持流式接口
- **⛓️ 链式调用**：链式操作让代码更简洁优雅
- **🔍 标签检查**：使用 `get_tags()` 获取所有标签列表
- **🛡️ 更好的错误处理**：清晰的异常提示和错误信息
- **📦 上下文管理器**：支持 `with` 语句
- **🔄 完全向后兼容**：所有 v3.x 代码无需修改即可运行！

## 安装要求

```bash
pip install python-docx
pip install pandas
python3
```

## 快速开始（v4.0 面向对象 API）- 推荐

### 方式1：链式调用（推荐）

```python
from WordWriter import WordWriter

# 最简洁的方式
WordWriter("template.docx") \
    .replace({
        "#[title]#": "我的报告",
        "#[date]#": "2025-10-22"
    }) \
    .save("output.docx")
```

### 方式2：分步调用

```python
from WordWriter import WordWriter

# 创建实例
writer = WordWriter("template.docx")
writer.load()

# 查看找到的标签（v4.0 新功能）
tags = writer.get_tags()
print(f"找到 {len(tags)} 个标签")

# 替换并保存
writer.replace({"#[title]#": "我的报告"})
writer.save("output.docx")
```

### 方式3：上下文管理器

```python
from WordWriter import WordWriter

with WordWriter("template.docx") as writer:
    writer.replace({"#[title]#": "我的报告"})
    writer.save("output.docx")
```

### 方式4：类方法（一步完成）

```python
from WordWriter import WordWriter

WordWriter.process("template.docx", "output.docx", 
                   {"#[title]#": "我的报告"})
```

## 经典用法（v3.x）- 仍然支持

```python
from WordWriter import word_writer

resultsDict = {}
resultsDict["#[testheader1]#"] = "测试页眉1"
resultsDict["#[testheader2]#"] = "测试页眉2"
resultsDict["#[testString]#"] = "测试文本"
resultsDict["#[testfooter]#"] = "测试页脚"
resultsDict["#[TX-testString2]#"] = "文本框文本"
resultsDict["#[testTableString1]#"] = "单元格文本1"
resultsDict["#[testTableString2]#"] = "单元格文本2"
resultsDict["#[IMAGE-test1-(30,30)]#"] = "testPicture.png"
resultsDict["#[IMAGE-test2]#"] = "testPicture2.png"
resultsDict["#[IMAGE-test3-(10,10)]#"] = "testPicture.png"
resultsDict["#[TABLE-test1]#"] = "testTable.txt"

word_writer("test.docx", "output.docx", resultsDict)
```

## 标签格式说明

### 文本标签
```
#[标签名]#
```
用于替换段落文本、单元格文本、页眉页脚等。

### 图片标签
```
#[IMAGE-图片名]#                    # 自动大小
#[IMAGE-图片名-(宽,高)]#            # 指定大小（单位：厘米）
#[TBIMG-图片名-(宽,高)]#            # 表格中的图片
```

### 表格标签
```
#[TABLE-表格名]#
```
表格数据文件应为 tab 分隔的文本文件（.txt）。

### 文本框标签
```
#[TX-文本框名]#
```
用于替换文本框中的内容。

## 特殊值

- `#DELETETHISPARAGRAPH#` - 删除包含标签的段落
- `#DELETETHISTABLE#` - 删除包含标签的表格

## 完整示例

```python
from WordWriter import WordWriter

# 准备替换数据
replace_dict = {
    # 文本替换
    "#[title]#": "年度工作报告",
    "#[author]#": "张三",
    "#[date]#": "2025年10月22日",
    
    # 页眉页脚
    "#[header]#": "机密文件",
    "#[footer]#": "第1页",
    
    # 图片插入
    "#[logo-(5,5)]#": "company_logo.png",
    "#[chart]#": "sales_chart.png",
    
    # 表格插入
    "#[TABLE-sales]#": "sales_data.txt",
    
    # 文本框
    "#[TX-note]#": "重要提示：本文件仅供内部使用",
    
    # 删除段落
    "#[draft_watermark]#": "#DELETETHISPARAGRAPH#"
}

# 使用链式调用处理
WordWriter("template.docx") \
    .replace(replace_dict) \
    .save("annual_report.docx")

print("✓ 报告生成完成！")
```

## 高级用法

### 批量处理多个文档

```python
from WordWriter import WordWriter

templates = ["template1.docx", "template2.docx", "template3.docx"]
data_list = [data1, data2, data3]

for template, data in zip(templates, data_list):
    WordWriter(template) \
        .replace(data) \
        .save(f"output_{template}")
```

### 条件替换

```python
writer = WordWriter("template.docx")
writer.load()

# 根据找到的标签决定替换内容
tags = writer.get_tags()
replace_dict = {}

if "#[date]#" in tags:
    from datetime import datetime
    replace_dict["#[date]#"] = datetime.now().strftime("%Y年%m月%d日")
    
if "#[title]#" in tags:
    replace_dict["#[title]#"] = "自动生成的报告"

writer.replace(replace_dict).save("output.docx")
```

### 错误处理

```python
from WordWriter import WordWriter

try:
    writer = WordWriter("template.docx")
    writer.load()
    writer.replace(replace_dict)
    writer.save("output.docx")
    print("✓ 处理成功！")
except FileNotFoundError as e:
    print(f"✗ 文件不存在: {e}")
except RuntimeError as e:
    print(f"✗ 运行时错误: {e}")
except Exception as e:
    print(f"✗ 未知错误: {e}")
```

## 表格合并

WordWriter 还提供了表格行合并功能：

```python
from WordWriter import merge_table_row
from docx import Document

doc = Document("document.docx")
table = doc.tables[0]

# 按第一列的内容合并相同的行
merge_table_row(table, 0)

doc.save("merged.docx")
```

## API 参考

### WordWriter 类

#### 构造函数
```python
WordWriter(template_path: str)
```

#### 方法

- `load() -> WordWriter` - 加载模板（支持链式调用）
- `replace(replace_dict: Dict[str, str], logs: bool = True) -> WordWriter` - 替换标签（支持链式调用）
- `save(output_path: str) -> None` - 保存文档
- `get_tags() -> List[str]` - 获取所有标签列表
- `process(template_path, output_path, replace_dict, logs=True)` - 类方法，一步完成

#### 特殊方法
- `__enter__` / `__exit__` - 支持上下文管理器
- `__repr__` - 对象字符串表示

### 函数式 API（向后兼容）

```python
word_writer(input_docx: str, output_docx: str, 
            replace_dict: Dict[str, str], logs: bool = True) -> None
```

```python
merge_table_row(table: Table, col_index: int, 
                remove_other_row_text: bool = True) -> None
```

## 常见问题

### Q: v4.0 与 v3.x 有什么区别？
**A**: v4.0 引入了面向对象 API，提供更现代化的使用方式。但完全向后兼容，所有 v3.x 代码无需修改即可运行。

### Q: 我应该使用哪种 API？
**A**: 新项目推荐使用 v4.0 的面向对象 API（链式调用）。现有项目可以继续使用函数式 API。

### Q: 如何查看我的版本？
```python
import WordWriter
print(WordWriter.__version__)  # 输出: 4.0.0
```

### Q: 标签没有被替换怎么办？
1. 检查标签格式是否正确（`#[标签名]#`）
2. 使用 `get_tags()` 查看实际找到的标签
3. 确保模板文件中确实存在该标签

### Q: 支持哪些图片格式？
支持 python-docx 支持的所有格式：PNG, JPG, JPEG, GIF, BMP, TIFF 等。

### Q: 表格文件格式是什么？
Tab 分隔的文本文件（.txt），每行一条记录，字段之间用 Tab 键分隔。

## 许可证

MIT License
