# -*- coding: utf-8 -*-
# @createTime    : 2019/7/7 10:29
# @author  : Huanglg
# @fileName: demo.py
# @email: luguang.huang@mabotech.com
from main import generateExcel
import os

if __name__ == '__main__':
    path = os.getcwd() + "/example/11111.pdf"
    out = os.getcwd() + "/out/"
    if not os.path.exists(out):
        os.makedirs(out)
    stockCode = "0001"

    generateExcel(path=path, out=out, stockCode=stockCode)
