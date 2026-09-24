# -*- coding: utf-8 -*-
"""生成 GUI 专用的边界用例模板（从「唛头模板表.xlsx」派生）。

只跑一次，产物进 `测试用例/`。跑法：
    python make_gui_cases.py
"""
import os
import shutil

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "唛头模板表.xlsx")
OUT = os.path.join(HERE, "测试用例")
SHEET = "唛头输入表"

# 列号：A..M = 1..13；N=14 票名 / O=15 厂家英文 / P=16 发货人 / Q=17 发货人地址
#        R=18 收货人 / S=19 收货人地址 / T=20 原产国
C = {"票名": 14, "厂家英文": 15, "发货人": 16, "发货人地址": 17,
     "收货人": 18, "收货人地址": 19, "原产国": 20,
     "品名英文": 2, "件数": 7}


def base():
    return openpyxl.load_workbook(SRC)


def save(wb, name):
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, name)
    wb.save(p)
    print("已生成", p)
    return p


def clear(ws, field, row):
    ws.cell(row, C[field]).value = None


# 1) 必填缺失：空票名（且无法继承）+ 空品名 + 空收货人地址（首行，无处可继承）
wb = base()
ws = wb[SHEET]
clear(ws, "票名", 2)          # 第一票的票名被清掉 → 整票无归属
clear(ws, "收货人地址", 2)     # 第 2 行收货人地址为空，且它是第一行 → 没得继承
clear(ws, "品名英文", 3)       # 第 3 行品名为空（品名不参与向下继承）
save(wb, "负面3_必填缺失.xlsx")

# 2) 件数非法：文本 + 0
wb = base()
ws = wb[SHEET]
ws.cell(3, C["件数"]).value = "abc"
ws.cell(4, C["件数"]).value = 0
save(wb, "负面4_件数非法.xlsx")

# 3) 缺列：把「Packages」表头删掉 → 每行都报「缺字段」
wb = base()
ws = wb[SHEET]
ws.cell(1, C["件数"]).value = None
save(wb, "负面2_缺件数列.xlsx")

# 4) 表头有杂七杂八的未识别列 → 应静默忽略/只提示，不影响生成
wb = base()
ws = wb[SHEET]
ws.cell(1, 21).value = "备注随便写"
ws.cell(1, 22).value = "Remarks"
ws.cell(2, 21).value = "这列程序不认识，应该被忽略"
save(wb, "测试9_多余列.xlsx")

# 5) 只有一行、只出一张唛头（最小可用表）
wb = base()
ws = wb[SHEET]
for r in range(ws.max_row, 2, -1):
    ws.delete_rows(r)
ws.cell(2, C["件数"]).value = 1
save(wb, "测试10_最小单张.xlsx")

# 6) 件数写全角数字 + 文本数字
wb = base()
ws = wb[SHEET]
ws.cell(3, C["件数"]).value = "２"
save(wb, "测试11_全角件数.xlsx")

print("完成")
