# -*- coding: utf-8 -*-
# @createTime    : 2019/7/6 10:13
# @author  : Huanglg
# @fileName: main.py
# @email: luguang.huang@mabotech.com


import pdfplumber
import pandas as pd
import re
import uuid
from datetime import datetime

header = "2018年年度报告"
noise = ["□适用 √不适用 ", "√适用 □不适用", "2018年年度报告"]

def MergeTables(allPageTables, meragePageNums):
    """
    Merge cross-page tables
    :param tables: tables object list
    :return:
    """
    for i in meragePageNums:
        currentTables = allPageTables[i]
        preTables = allPageTables[i - 1]
        perPageLastTable = preTables[len(preTables) - 1]
        currentPageFirstTable = currentTables[0]
        perPageLastTable.extend(currentPageFirstTable)
        currentTables.remove(currentPageFirstTable)
    return allPageTables


def GetsTableAcrossPages(pages):

    """
    Returns the page index of the table that needs to be removed and merged
    :param pages:
    :return:
    """
    result = []
    for index in range(len(pages) - 1):
        currentPageTables = pages[index].extract_tables()
        nextPageTables = pages[index + 1].extract_tables()
        currentPageLastTable = currentPageTables[len(currentPageTables) - 1]
        nextPageFirstTable = nextPageTables[0]
        text = pages[index + 1].extract_text().replace(header,"").lstrip()
        tableText = nextPageTables[0][0][0].lstrip()
        if len(currentPageLastTable[0]) == len(
                nextPageFirstTable[0]) and text[0:2] == tableText[0:2]:
            result.append(index + 1)
    return result


def HandlerText(text):

    # Remove the page number
    text = re.sub(r"\d+\s/\s\d+", "", text)

    # Get rid of the noise
    for s in noise:
        text = text.replace(s, "")

    # Gets table filename and units
    title = re.findall(r"\((.*)\s+单位", text)
    unit = re.findall(r"单位(.*)\s", text)

    # handler
    resTitle = []
    step = 0
    for i in title:
        index = text.find(i, step)
        step = index
        resTitle.append(text[index - 18:index + 10].replace(" ",
                                                            "").replace("\n", "").replace("其他说明：", ""))
    for i in range(len(unit)):
        unit[i] = "单位" + unit[i]
    return zip(resTitle, unit)


def generateExcel(path, out, stockCode=str(uuid.uuid1()), fileFormat='.xlsx'):
    """

    :param path: pdf's path
    :param stockCode:
    :param out: out path
    :param fileFormat: defalut .xlsx
    :return:
    """

    pdf = pdfplumber.open(path)
    text = ""
    allPageTables = []
    pages = pdf.pages
    meragePageNums = GetsTableAcrossPages(pages)
    for page in pdf.pages:
        text += page.extract_text()
        allPageTables.append(page.extract_tables())

    text = HandlerText(text)
    res = MergeTables(allPageTables, meragePageNums)
    for pageTables in res:
        for table in pageTables:
            s = text.__next__()
            lenth = len(table[0])
            temp = ["" for i in range(lenth)]
            temp[0] = s[0]
            temp[1] = s[1]
            table.insert(0, temp)
            tb = pd.DataFrame(table[1:], columns=table[0], index=None)
            year = str(datetime.now().year)
            fileName = stockCode + "_" + year + "_" + s[0] + fileFormat
            tb.to_excel(out + fileName, encoding="utf-8", index=False)

