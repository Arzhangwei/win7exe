#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations  # 新增此行，解决Python 3.8的类型注解兼容性问题

"""
读取 peopleList.csv → 生成步步高 USB 电话通讯录专用 CSV
打包：pyinstaller -F -w bbk_csv_tool.py
"""
import os
import sys
import csv
import pandas as pd
from phone import Phone
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout,
                             QPushButton, QTextEdit, QMessageBox)

# ----------------------------------------------------------
# 路径工具：打包后 exe 同目录，开发时脚本目录
# ----------------------------------------------------------
def current_dir() -> str:
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.realpath(__file__))

def res_path(*names) -> str:
    full = os.path.join(current_dir(), *names)
    os.makedirs(os.path.dirname(full), exist_ok=True)
    return full

# ----------------------------------------------------------
# 文件常量
# ----------------------------------------------------------
CSV_FILE   = res_path('peopleList.csv')     # 输入
DAT_PATH   = res_path('phone.dat')          # 归属地库
VCARD_CSV  = res_path('名片.csv')           # 输出

BBK_HEADER = ["姓名", "移动电话", "办公电话", "家庭电话", "备注"]

# ----------------------------------------------------------
# 写 CSV 工具：GB18030 无 BOM
# ----------------------------------------------------------
def write_bbk_csv(path, rows: list[list[str]]):
    with open(path, 'w', newline='', encoding='gb18030') as f:
        writer = csv.writer(f)
        writer.writerow(BBK_HEADER)
        writer.writerows(rows)

# ----------------------------------------------------------
# Qt 界面
# ----------------------------------------------------------
class MainWin(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('利辛农商行电话回访辅助（CSV→步步高）')
        self.resize(420, 300)

        self.run_btn = QPushButton('开始处理')
        self.log = QTextEdit(readOnly=True)

        layout = QVBoxLayout(self)
        layout.addWidget(self.run_btn)
        layout.addWidget(self.log)

        self.run_btn.clicked.connect(self.work)

        if not os.path.exists(DAT_PATH):
            QMessageBox.warning(self, '提示', f'请将 phone.dat 放于：\n{current_dir()}')

    # -------------- 工具 --------------
    def need_prefix(self, mobile: str) -> bool:
        info = Phone(dat_file=DAT_PATH).find(mobile) or {}
        return info.get('area_code') != '0558' or info.get('city') != '亳州'

    # -------------- 主流程 --------------
    def work(self):
        if not os.path.exists(CSV_FILE):
            QMessageBox.critical(self, '错误', f'当前目录未找到：\n{CSV_FILE}')
            return
        try:
            # 自动推断编码，兼容 UTF-8 / GBK
            df = pd.read_csv(CSV_FILE, dtype=str, encoding='gb18030').fillna('')
        except Exception as e:
            QMessageBox.critical(self, '读 csv 失败', str(e))
            return

        # 逐行从左到右扫描：客户姓名/手机/担保人姓名/手机
        rows = []
        for _, line in df.iterrows():
            # 客户
            name1 = str(line.get('客户姓名', '')).strip()
            mob1  = str(line.get('手机号码', '')).strip()
            if mob1:
                if self.need_prefix(mob1):
                    mob1 = '0' + mob1
                rows.append([f"{name1}_{mob1}", mob1])
            # 担保人
            name2 = str(line.get('担保人姓名', '')).strip()
            mob2  = str(line.get('手机号码.1', '')).strip()  # 第二列手机号标题
            if mob2:
                if self.need_prefix(mob2):
                    mob2 = '0' + mob2
                rows.append([f"{name2}_担_{mob2}", mob2])

        if not rows:
            QMessageBox.information(self, '提示', '未找到任何有效手机号')
            return

        # 写步步高 CSV
        try:
            bbk_rows = [[n, m, '', '', ''] for n, m in rows]
            write_bbk_csv(VCARD_CSV, bbk_rows)
        except PermissionError:
            QMessageBox.warning(self, '文件被占用',
                                f'请先关闭正在打开的文件再试：\n{VCARD_CSV}')
            return
        except Exception as e:
            QMessageBox.critical(self, '写CSV失败', str(e))
            return

        # 日志
        text = '\n'.join(['\t'.join(r) for r in rows])
        self.log.setPlainText(f'共生成 {len(rows)} 条记录，已写入步步高兼容CSV：\n\n' + text)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWin()
    w.show()
    sys.exit(app.exec_())
