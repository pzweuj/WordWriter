# coding=utf-8
# pzw
# 20211028

# import WordWriter as ww

import sys
sys.path.insert(0, '..')  # 添加父目录到路径
from WordWriter import word_writer


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

# ww.WordWriter("test.docx", "output.docx", resultsDict)
word_writer("test.docx", "output.docx", resultsDict)
