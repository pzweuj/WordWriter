#### 依赖
```bash
pip install python-docx
pip install pandas
```

#### 详细说明

[详细说明](https://pzweuj.github.io/2021/06/07/WordWriter.html)

#### 基本使用
```python
## python3
import WordWriter3
```

在模板word中创建标签，形式为**#[xxxx]#**。

其中表格标签必须是#[TABLE-xxx]#，

表格内容标签必须是#[TBS-xxx]#，

文本框中内容标签必须是#[TX-xxx]#，

图片标签必须是#[IMAGE-xxx]#，支持定义图片大小#[IMAGE-xxx-(30,40)]#，

单元格中的图片标签必须是#[TBIMG-xxxx]#，支持定义图片格式大小#[TBIMG-xxxx-(30,40)]#，

页眉标签必须是#[HEADER-xxx]#，页脚标签必须是#[FOOTER-xxx]#，

其他文本内容标签可自定义#[xxxx]#。


#### 实例
test.docx是自定义模板。


python3

```python
# 测试脚本
testDict = {}
testDict["#[HEADER-1]#"] = "模板测试"
testDict["#[HEADER-2]#"] = "2019年7月18日"
testDict["#[NAME]#"] = "测试模板"
testDict["#[fullParagraph]#"] = "这是一段测试段落，通过WordWriter输入。"
testDict["#[TBS-1]#"] = "未突变"
testDict["#[TX-1]#"] = "文本框测试成功"
testDict["#[TX-2]#"] = "文本框测试很成功"
testDict["#[FOOTER]#"] = "页脚测试"

# 此处输入的是文件路径
testDict["#[TABLE-1]#"] = "test/testTable.txt"
testDict["#[IMAGE-1-(30,30)]#"] = "test/testPicture.png"
testDict["#[IMAGE-2]#"] = "test/testPicture.png"
testDict["#[TBIMG-3-(20,20)]#"] = "test/testPicture.png"

# 使用主函数进行报告填充
WordWriter("test/test.docx", "test/testOut.docx", testDict)
```

#### 注意事项
word中以**run**的形式储存每次的输入内容。如果输入不连贯，同一行文本会被分为不同的run。这会导致标签无法被正确识别。因此，建议将每个标签一次性的粘贴进模板中，再全选标签进行格式调整。
可以通过把docx文件另存为xml文件，查看内容看标签是否完整。
