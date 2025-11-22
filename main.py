# =============================
# 导入模块
# =============================

import tkinter as tk

from datetime import date, datetime

from pathlib import Path
import shutil

from openpyxl import load_workbook, Workbook

import pandas as pd
from pandas.io.formats.excel import ExcelFormatter

import pymupdf as fitz

# 导入uploader.py
from services.uploader import Uploader
from services.update_inv import InvUpdater


# =============================
# 配置区
# =============================
def main():
    # 设置pandas输出格式
    ExcelFormatter.header_style = None

    # 设置路径
    TEMPLATES_DIR = Path(r"D:\共享文件夹\金丽华")
    PARAM_FILE = TEMPLATES_DIR / "参数表.xlsx"
    BODY_TEMPLATE = TEMPLATES_DIR / "发票导入模板.xlsx"
    HEAD_TEMPLATE = TEMPLATES_DIR / "接单信息模板.xlsx"
    INV_HEAD = TEMPLATES_DIR / "抬头.xlsx"
    STAMP = TEMPLATES_DIR / "贝克休斯章PNG.png"

    # 待处理文件夹，用户输入
    # src_folder = r"D:\共享文件夹\金丽华\59-251017-004603 分拨数据 20251001-20251016"

    # 兜底的打印方法
    printlog = PrintLog()
    # 错误日志
    exception = ExceptionLogger()

    # 导入 GUI 界面
    from gui import create_gui
    # 完整配置
    config = Config(
        param_file=PARAM_FILE,
        body_template=BODY_TEMPLATE,
        head_template=HEAD_TEMPLATE,
        inv_head=INV_HEAD,
        stamp=STAMP,
        log=printlog,   # 默认是 PrintLog，实际由 GUI 回传
        exception=exception
        # src_folder=src_folder # 可选，默认是 None，因为是 GUI 提供
    )
    # 初始化产线
    lines = [
        Line("004", "FBIT", config),
        Line("603", "FBSN", config),
    ]
    actions = Actions(config, lines)

    create_gui(config, actions)


    # =============================
    # 开始处理
    # =============================
    # for line in lines:
    #     # 查找Excel
    #     line.find_datasets()
    #     # 查找预览文件
    #     line.find_docs()
    #     # 生成预览文件
    #     line.export_datasets()
    #     # 创建invoice模板
    #     line.build_invoice_bodys()
    #     # 创建表头模板
    #     line.build_invoice_heads()


    # =============================
    # 写入日志
    # =============================
    # log_file = config.src_folder / "操作记录_log.txt"
    # with open(log_file, "w", encoding="utf-8") as log_file:
    #     if exception.logs:
    #         for line in exception.logs:
    #             log_file.write(line + "\n")
    #         print(f"⚠️ 发现异常，请查看日志: {log_file}")
    #     else:
    #         log_file.write("🎉 所有文件均已成功匹配并移动！")
    #         print("🎉 所有文件均已成功匹配并移动！")
    

if __name__ == "__main__":
    main()