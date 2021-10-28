## 用途
用于替换docx模板中的预留标签。支持替换段落字符串，单元格字符串，文本框字符串，页眉，页脚；插入表格，插入图片。



#### 依赖

```bash
pip install python-docx
pip install pandas
```

#### 详细描述
使用方式点击[这里](https://pzweuj.github.io/2021/06/07/WordWriter.html)。

#### 基本使用

```python
## python3
import WordWriter as ww
```

在模板word中创建标签，形式为**#[xxxx]#**。

其中表格标签必须是#[TABLE-xxx]#，

文本框中内容标签必须是#[TX-xxx]#，

图片标签必须是#[IMAGE-xxx]#，支持定义图片大小#[IMAGE-xxx-(30,40)]#，

其他文本内容标签，如段落字符串、单元格中的字符串、页眉、页脚等均可自定义#[xxxx]#。



#### 实例

test.docx是自定义模板。


python3

```python
# 测试脚本
import WordWriter as ww

resultsDict = {}
resultsDict["#[testheader1]#"] = "测试页眉1"
resultsDict["#[testheader2]#"] = "页眉测试2"
resultsDict["#[testString]#"] = "，文本替换成功"
resultsDict["#[testfooter]#"] = "测试页脚"
resultsDict["#[TX-testString2]#"] = "，文本框文本替换成功"
resultsDict["#[testTableString1]#"] = "单元格文本替换成功"
resultsDict["#[testTableString2]#"] = "单元格文本替换成功"
resultsDict["#[IMAGE-test1-(30,30)]#"] = "testPicture.png"
resultsDict["#[IMAGE-test2]#"] = "testPicture2.png"
resultsDict["#[IMAGE-test3-(10,10)]#"] = "testPicture.png"
resultsDict["#[TABLE-test1]#"] = "testTable.txt"

ww.WordWriter("test.docx", "output.docx", resultsDict)
```

#### 注意事项
word中以**run**的形式储存每次的输入内容。如果输入不连贯，同一行文本会被分为不同的run。这会导致标签无法被正确识别。
