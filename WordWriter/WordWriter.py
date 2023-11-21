# coding=utf-8
# pzw
# 20231019
# v3.1   修复bottom无法找到border的问题
# v3.0.3 修复部分bug
# v3.0   解决run不完整的问题

import os
import pandas as pd
from docx import Document
from docx.oxml.ns import qn as nsqn
from docx.oxml import OxmlElement

# 通用搜索循环
## 形成的是类似{tag1: [p, r1, r2, r3], tag2: [p, r1]}这样的字典
def searchTag(tagDict, paragraphs):
    for p in paragraphs:
        if "#[" in p.text and "]#" in p.text:
            tag_name = ""
            run_list = []
            for r in p.runs:
                text = r.text.strip()
                # 如果tag_name不为空以及不是终止
                if tag_name and not "]#" in text:
                    tag_name += text
                    run_list.append(r)
                else:
                    # tag_name是空 或 是终止位
                    if "#[" in text:
                        # 此时是完整的tag
                        if "]#" in text:
                            tagDict.setdefault(text, []).append([p, [r]])
                        # 此时仅仅是起始
                        else:
                            tag_name = text
                            run_list = [r]
                    elif "]#" in text:
                        # tag_name不为空，同时是结束位置
                        if tag_name:
                            tag_name += text
                            run_list.append(r)
                            tagDict.setdefault(tag_name, []).append([p, run_list])
                            # 重新初始化
                            tag_name = ""
                            run_list = []

# 建立各类tag字典
## 遍历模板，从模板中寻找完整的tag
def searchTemplateTag(document):
    tagDict = {}

    ### 页眉页脚
    sectionsList = []
    for s in document.sections:
        sectionsList.extend([s.header, s.first_page_header, s.footer, s.first_page_footer])
    for sl in sectionsList:
        searchTag(tagDict, sl.paragraphs)

    ### 段落
    searchTag(tagDict, document.paragraphs)

    ### 表格
    for t in document.tables:
        rows = t.rows
        for r in range(len(rows)):
            cells = rows[r].cells
            for c in range(len(cells)):
                cell = cells[c]
                if "#[" in cell.text and "]#" in cell.text:
                    if "#[TABLE" in cell.text and "]#" in cell.text:
                        tag = "#[TABLE-" + cell.text.split("#[TABLE-")[1].split("]#")[0] + "]#"
                        tagDict.setdefault(tag, []).append([t, r, c])
                    else:
                        # 单元格中的字符串tag
                        searchTag(tagDict, cell.paragraphs)

    ### 文本框，仅支持整个文本框中的内容替换
    children = document.element.body.iter()
    for child in children:
        if child.tag.endswith(("AlternateContent", "textbox")):
            for ci in child.iter():
                if ci.tag.endswith(("main}r", "main}pPr")):
                    if ci.text != None:
                        if "#[TX" in ci.text and "]#" in ci.text:
                            tagDict.setdefault(ci.text.strip(), []).append(ci)
    
    return tagDict

# 获得指定行号表格边框底线格式
def get_table_bottom_border_details(tableObj, row_index, cell_index):    
    # 获取表格的指定行号
    last_row = tableObj.rows[row_index]

    # 默认空样式
    ## val: single 实线；dashed 虚线；nil 隐藏
    default_border_details = {
        'size': '0',
        'color': 'auto',
        'space': '0',
        'val': 'single'
    }

    # 获取表格的边框格式
    tbl_borders = tableObj._tbl.tblPr.first_child_found_in("w:tblBorders")
    if tbl_borders:
        tbl_bottom_border = tbl_borders.find(nsqn("w:bottom"))
        tbl_border_details = {
            'size': tbl_bottom_border.get(nsqn('w:sz'), '0'),
            'color': tbl_bottom_border.get(nsqn('w:color'), 'auto'),
            'space': tbl_bottom_border.get(nsqn('w:space'), '0'),
            'val': tbl_bottom_border.get(nsqn('w:val'), 'single'),
        }
    else:
        tbl_border_details = default_border_details

    # 获取指定行号中所有单元格的底线边框格式
    bottom_border_details = []
    for cell in last_row.cells[cell_index:]:
        # 获取单元格的底线边框格式
        tc_borders = cell._tc.get_or_add_tcPr().first_child_found_in("w:tcBorders")
        bottom_border = tc_borders.find(nsqn("w:bottom")) if tc_borders != None else None
        if bottom_border is not None:
            border_details = {
                'size': bottom_border.get(nsqn('w:sz'), '0'),
                'color': bottom_border.get(nsqn('w:color'), 'auto'),
                'space': bottom_border.get(nsqn('w:space'), '0'),
                'val': bottom_border.get(nsqn('w:val'), 'single'),
            }
        else:
            border_details = default_border_details
        bottom_border_details.append(border_details)
    return bottom_border_details, tbl_border_details

# 设置单元格底线边框格式
def set_cell_bottom_border(cell, styleList):
    size = styleList["size"]
    color = styleList["color"]
    space = styleList["space"]
    border_type = styleList["val"]
    tcPr = cell._tc.get_or_add_tcPr()
    tc_borders = tcPr.first_child_found_in("w:tcBorders")
    if tc_borders == None:
        tc_borders = OxmlElement("w:tcBorders")
        tcPr.append(tc_borders)

    bottom_border = OxmlElement("w:bottom") if tc_borders.find(nsqn("w:bottom")) == None else tc_borders.find(nsqn("w:bottom"))
    
    # 设置底线边框的属性
    bottom_border.set(nsqn('w:sz'), size)
    bottom_border.set(nsqn('w:color'), color)
    bottom_border.set(nsqn('w:space'), space)
    bottom_border.set(nsqn('w:val'), border_type)
    tc_borders.append(bottom_border)

# 设置表格的底线边框格式
def set_table_bottom_border(table, styleList):
    size = styleList["size"]
    color = styleList["color"]
    space = styleList["space"]
    border_type = styleList["val"]
    tblPr = table._tbl.tblPr
    tbl_borders = tblPr.first_child_found_in("w:tblBorders")
    if tbl_borders == None:
        tbl_borders = OxmlElement("w:tblBorders")
        tblPr.append(tbl_borders)
    bottom_border = OxmlElement("w:bottom") if tbl_borders.find(nsqn("w:bottom")) == None else tbl_borders.find(nsqn("w:bottom"))

    # 设置底线边框的属性
    bottom_border.set(nsqn('w:sz'), size)
    bottom_border.set(nsqn('w:color'), color)
    bottom_border.set(nsqn('w:space'), space)
    bottom_border.set(nsqn('w:val'), border_type)
    tbl_borders.append(bottom_border)

## 字符串替换，适用于表格单元格中的字符串/页眉页脚字符串/段落字符串
def replaceParagraphString(run_list, replaceString):
    run_list[0].text = replaceString
    for i, r in enumerate(run_list):
        if i != 0:
            r.clear()
    if replaceString == "#DELETETHISPARAGRAPH#":
        paragraph = run_list[0]._element.getparent()
        remove_ele(paragraph)

## 图片插入，适用于表格中的图片和段落中的图片
def insertPicture(run_list, tag, picturePath):
    if os.path.isfile(picturePath):
        for r in run_list:
            r.text = ""
        if "(" in tag and ")" in tag:
            width = int(tag.split("(")[1].split(",")[0])
            height = int(tag.split(")")[0].split(",")[1])
            run_list[0].add_picture(picturePath, width*100000, height*100000)
        else:
            run_list[0].add_picture(picturePath)
    else:
        if picturePath == "#DELETETHISPARAGRAPH#":
            paragraph = run_list[0]._element.getparent()
            remove_ele(paragraph)
        else:
            run_list[0].text = picturePath
            for i, r in enumerate(run_list):
                if i != 0:
                    r.clear()

## 文本框中字符串替换，仅适合于文本框内字符串
def replaceTextBoxString(childList, replaceString):
    for c in childList:
        c.text = replaceString

## 表格插入，通过插入一个以tab分割的txt文件插入表格
### 表格初始化
def OriginTableReadyToFill(tableFile):
    table = pd.read_csv(tableFile, header=None, sep="\t", dtype=str)
    return table

### 删除表格行
def remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)

### 获取表格格式
def table_style_list(table, rowIdx, cellIdx):
    cell_list = table.rows[rowIdx].cells[cellIdx:]
    style_list = []
    for cell in cell_list:
        p0 = cell.paragraphs[0]
        lineSpacingRule = p0.paragraph_format.line_spacing
        spaceAfter = p0.paragraph_format.space_after
        r0 = p0.runs[0]
        font = r0.font
        style_list.append([cell.vertical_alignment, p0.style, p0.alignment, r0.bold, r0.italic, r0.underline, font.name, font.size, font.color.rgb, font.highlight_color, lineSpacingRule, spaceAfter])
    return style_list

### 表格内容填充及调整格式
def fill_table_text_and_style(table, row_id, fill_table_id, fill_cell_id, fill_row_id, fill_col_id, style_list):
    start = 0
    run_row = row_id
    while row_id <= fill_row_id + run_row - 1:
        for co in range(fill_col_id):
            tc = table.cell(row_id, co + fill_cell_id)
            tc.text = str(fill_table_id.iloc[start, co]).replace("\\x0a", "\n")
            tc.vertical_alignment = style_list[co][0]
            tc.paragraphs[0].style = style_list[co][1]
            tc.paragraphs[0].alignment = style_list[co][2]
            tc.paragraphs[0].paragraph_format.line_spacing = style_list[co][10]
            tc.paragraphs[0].paragraph_format.space_after = style_list[co][11]
            r = tc.paragraphs[0].runs[0]
            r.bold = style_list[co][3]
            r.italic = style_list[co][4]
            r.underline = False if style_list[co][5] != True else True
            r.font.name = style_list[co][6]
            if not r._element.rPr.rFonts == None:
                r._element.rPr.rFonts.set(nsqn("w:eastAsia"), r.font.name)
            r.font.size = style_list[co][7]
            r.font.color.rgb = style_list[co][8]
            r.font.highlight_color = style_list[co][9]

        start += 1
        row_id += 1

### 表格插入
def fillTable(table, row_id, cell_id, insertTable):
    tableToFill = OriginTableReadyToFill(insertTable)
    rowToFill = tableToFill.shape[0]
    columnToFill = tableToFill.shape[1]

    # 格式刷
    styleList = table_style_list(table, row_id, cell_id)

    # 获得标签行及最后一行的底边样式
    tagBottomStyle, tagTableStyle = get_table_bottom_border_details(table, row_id, cell_id)
    lastLineBottomStyle, lastLineTableStyle = get_table_bottom_border_details(table, -1, cell_id)

    # 将当前的最后一行的底边样式先处理为正常格式
    currentLastLine = table.rows[-1].cells[cell_id:]
    if len(currentLastLine) != len(tagBottomStyle):
        if len(tagBottomStyle) != 0:
            set_cell_bottom_border(currentLastLine[c], tagBottomStyle[0])
    else:
        for c in range(len(currentLastLine)):
            set_cell_bottom_border(currentLastLine[c], tagBottomStyle[c])

    # 判断行数是否足够，不够就添加
    if len(table.rows) - row_id < rowToFill:
        addRowAmount = rowToFill - len(table.rows) + row_id
        i = 0
        while i < addRowAmount:
            table.add_row()
            i += 1

    # 填充内容
    fill_table_text_and_style(table, row_id, tableToFill, cell_id, rowToFill, columnToFill, styleList)
    
    # 删除空行
    for row in table.rows:
        pString = ""
        for cell in row.cells:
            for p in cell.paragraphs:
                pString = pString + p.text
        if pString == "":
            remove_row(table, row)

    # 处理表格的边框底线样式
    set_table_bottom_border(table, lastLineTableStyle)

    # 处理此时最后一行的边框底线样式
    newLastLine = table.rows[-1].cells[cell_id:]
    if len(newLastLine) != len(lastLineBottomStyle):
        if len(lastLineBottomStyle) != 0:
            if (lastLineBottomStyle[0]["size"] == "0") and (lastLineBottomStyle[0]["color"] == "auto"):
                # 这种情况认为是默认状态，那就优先表格底线样式
                set_cell_bottom_border(newLastLine[c], lastLineTableStyle)
            else:
                set_cell_bottom_border(newLastLine[c], lastLineBottomStyle[0])
    else:
        for c in range(len(newLastLine)):
            if (lastLineBottomStyle[0]["size"] == "0") and (lastLineBottomStyle[0]["color"] == "auto"):
                set_cell_bottom_border(newLastLine[c], lastLineTableStyle)
            else:
                set_cell_bottom_border(newLastLine[c], lastLineBottomStyle[c])

### 删除元素
def remove_ele(ele):
    if str(type(ele)) == "<class 'docx.oxml.text.paragraph.CT_P'>":
        ele.getparent().remove(ele)
    else:
        parent = ele._element.getparent()
        parent.remove(ele._element)

# 函数合并
def WordWriter(inputDocx, outputDocx, replaceDict, logs=True):
    template = Document(inputDocx)
    templateTagDict = searchTemplateTag(template)
    for k in replaceDict:
        if not k in templateTagDict:
            if logs:
                print("【Missing Tag】 " + k)
        else:
            if logs:
                print("【Filling Tag】 " + k)
            if "#[TABLE" in k:
                if replaceDict[k] == "#DELETETHISTABLE#":
                    for i in templateTagDict[k]:
                        tableID = i[0]
                        remove_ele(tableID)
                else:
                    for i in templateTagDict[k]:
                        fillTable(i[0], i[1], i[2], replaceDict[k])
            elif "#[TX" in k:
                for i in templateTagDict[k]:
                    replaceTextBoxString(i, replaceDict[k])
            elif "#[IMAGE" in k or "#[TBIMG" in k:
                for i in templateTagDict[k]:
                    insertPicture(i[1], k, replaceDict[k])
            else:
                for i in templateTagDict[k]:
                    replaceParagraphString(i[1], replaceDict[k])
    template.save(outputDocx)

# 合并内容相同的行，这些行需要是排好序的
def MergeTableRow(tableObj, colIndex, remove_other_row_text=True):
    # 获得需要合并的行
    rowLen = len(tableObj.rows)
    mergeList = []
    nowText = ""
    mergeStartPoint = mergeEndPoint = 0
    for i in range(rowLen):
        currentText = tableObj.rows[i].cells[colIndex].text
        if currentText != nowText:
            if mergeEndPoint > mergeStartPoint:
                mergeList.append([mergeStartPoint, mergeEndPoint])
            mergeEndPoint = i - 1
            mergeStartPoint = i
            nowText = currentText
        else:
            mergeEndPoint = i
    if mergeEndPoint > mergeStartPoint:
        mergeList.append([mergeStartPoint, mergeEndPoint])

    # 合并
    for m in mergeList:
        if remove_other_row_text:
            for j in range(m[0], m[1] + 1):
                cell = tableObj.cell(j, colIndex)
                if j != m[0]:
                    cell.text = ""
                    for p in cell.paragraphs:
                        p.clear()
        tableObj.cell(m[0], colIndex).merge(tableObj.cell(m[1], colIndex))


