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
from actions.uploader import Uploader
from actions.update_inv import InvUpdater

# =============================
# 定义类
# =============================
    
# 模拟一个print
class PrintLog:
    SYMBOLS ={
        "info": "ℹ️ ",
        "ok": "✅",
        "warn": "⚠️",
        "error": "❌",
        "path": "📂",
    }

    def _write(self, level, msg):
        prefix = self.SYMBOLS.get(level, "")
        print(f"{prefix} {msg}")

    def info(self, msg):        self._write("info", msg)
    def ok(self, msg):          self._write("ok", msg)
    def warn(self, msg):        self._write("warn", msg)
    def error(self, msg):       self._write("error", msg)
    def path(self, path_obj):   self._write("path", path_obj)

# 日志类
class ExceptionLogger:
    def __init__(self):
        self.logs = []
        self.log_no = 0

    def add(self, msg):
        self.log_no += 1
        self.logs.append(f"{self.log_no:02d}: {msg}")
    
    def __str__(self):        
        return "\n".join(self.logs) if self.logs else "(没有日志)"
    
    def __repr__(self):
        return f"<Logger count={self.log_no}>"

# 配置类
class Config:
    def __init__(self, param_file, body_template, head_template, inv_head, stamp, log, exception, src_folder=None):
        # GUI 运行时更新的状态（初始值为空）
        self.src_folder = src_folder
        self.serial_code = ""

        self.inv_fields_map = {}
        self.base_fields_map = {}
        self.head_fields_map = {}
        self.country_dict = {}
        self.currency_dict = {}
        self._load_dependents(param_file)   # 加载上述参数
        # 保存模板
        self.param_file = param_file
        self.body_template = body_template
        self.head_template = head_template
        self.inv_head = inv_head
        self.stamp = stamp

        # 日志接口（GUI 会稍后替换）
        self.log = log or PrintLog()
        self.exception = exception        

    def set_src_folder(self, folder):
        """由 GUI 调用，选择文件夹后更新路径 + 解析流水号"""
        path = Path(folder)
        if not (path.exists() and path.is_dir()):
            self.log.error(f"无效文件夹：{path}")
            return

        self.src_folder = path
        # 解析文件夹名称中的流水号
        self.serial_code = self._parse_serial_code(path.name)
        # self.log.ok(f"已选择文件夹: {self.src_folder} (流水号: {self.serial_code})")
        self.log.ok(f"已选择文件夹：")
        self.log.path(self.src_folder)
        self.log.ok(f"提取的流水号: {self.serial_code}")

    def set_log(self, log_view):
        """由 GUI 调用，替换日志输出接口（PrintLog -> LogView）"""
        self.log = log_view

    # 获取年月流水号
    def _parse_serial_code(self, folder_name):
        parts = folder_name.split("-")[:2] # 用 - 分割，然后取前两个段
        # parts 有一个是6位的，另一个是2位的，假设顺序不确定。2位的是流水号（可以补0前缀），6位取前4位为年月
        # 提取2位的流水号
        try:
            flow_no = next(s for s in parts if len(s) == 2 and s.isdigit())
            flow_no = f"{int(flow_no):02d}"  # 补0，转化为字符串
        except StopIteration:
            flow_no = "00"
            msg = "未找到有效流水号（2位数字），已使用00，请检查文件夹命名"
            self.log.warn(msg)
            self.exception.add(msg)
        # 提取年月
        def is_yymm(s4):
            try:
                datetime.strptime(s4, "%y%m")
                return True
            except ValueError:
                return False
        try:
            yymm = next(s[:4] for s in parts if len(s) == 6 and is_yymm(s[:4]))
        except StopIteration:
            yymm = date.today().strftime("%y%m") # 兜底：当前年月
            msg = "找到有效年月编码，已使用当前年月，请检查文件夹命名"
            self.log.warn(msg)
            self.exception.add(msg)

        return f"{yymm}{flow_no}"

    def _load_table_as_df(self, wb, sheet_name, tbl_name=None):
        if tbl_name is None:
            tbl_name = sheet_name
        
        ws = wb[sheet_name]
        tbl = ws.tables[tbl_name]
        ref = tbl.ref   # e.g. 'A1:G10'
        rng = ws[ref]
        data = [[c.value for c in row] for row in rng]
        df = pd.DataFrame(data[1:], columns=data[0])
        return df

    def _load_dependents(self, file_path):
        wb_dep = load_workbook(file_path, data_only=True)

        def map_from(table):
            df = self._load_table_as_df(wb_dep, table)
            return dict(zip(
                    df["目标字段名"],
                    zip(df["类型"], df["源字段名"])
            ))
        # 发票字段映射表
        self.inv_fields_map = map_from("发票字段映射表")
        # 接单字段映射表
        self.base_fields_map = map_from("接单字段映射表")
        # 表头字段映射表        
        self.head_fields_map = map_from("表头字段映射表")

        # 国家代码表
        country_df = self._load_table_as_df(wb_dep, "国家代码表")
        self.country_dict = dict(zip(country_df["国别中文"], country_df["ISO代码"]))
        # 币制代码表
        currency_df = self._load_table_as_df(wb_dep, "币制代码表")
        self.currency_dict = dict(zip(currency_df["货币中文"], currency_df["货币代码"]))

        wb_dep.close()

# 数据集类
class DataSet:
    def __init__(self, internal_code, config):
        self.internal_code = internal_code

        self.config = config
        self.qd_code = ""  # QD号

        # 源Excel
        self.src_excel = None   # Path
        self.ws_customs = None
        self.ws_qd = None
        # 实际数据
        self.df_customs = None   # pandas dataframe
        self.gross_weight = 0
        self.pcs_count = 0
        self.shipping_fee = 0
        # 源doc
        self.doc_path = None    # Path 对象
        self.doc_type = ""      # .xps/.pdf
        # 导出文件
        self.dst_dir = None      # Path
        self.dst_excel = None    # Path
        self.dst_doc = None      # Path
        self.dst_pdf = None      # Path
        # 模板文件
        self.body_template = None
        self.head_template = None
        self.uploader = Uploader()
        # 最终结果
        self.inv_excel = None
        self.ccs_excel = None
        # 更新发票
        self.inv_updater = InvUpdater()

    @property
    def src_folder(self):
        return self.config.src_folder
    
    @property
    def exception(self):
        return self.config.exception

    def __repr__(self):
        return (
            f"<DataSet {self.qd_code}\n"
            f"  internal_code={self.internal_code}\n"
            f"  src_excel={self.src_excel}\n"
            f">"
        )
    
    def find_doc(self):
        if not self.qd_code:
            msg = f"{self.internal_code} 无QD号，无法查找预览文件"
            self.exception.add(msg)
            return

        for file in self.src_folder.iterdir():
            if file.suffix.lower() not in [".xps", ".pdf"]:
                continue

            try:
                with fitz.open(file) as doc:
                    text = doc.get_page_text(0)
                    if self.qd_code in text:
                        self.doc_path = file.resolve()
                        self.doc_type = file.suffix.lower()
                        return
            except Exception as e:
                msg = f"无法解析文档 {file}，已跳过"
                self.exception.add(msg)
                continue

        msg = f"{self.internal_code} 未找到对应的的预览文件（.xps/.pdf）"
        self.exception.add(msg)

    def export(self):
        # 创建目录
        self.dst_dir = self.src_folder / self.internal_code
        self.dst_dir.mkdir(exist_ok=True)
        self.dst_excel = self.dst_dir / f"{self.internal_code}.xlsx"

        # 导出 Excel
        # 先用 pandas 写入 customs 数据，注意按照报关单序号排序
        src_ws = self.ws_customs
        title = src_ws.title
        data = [[c.value for c in row] for row in src_ws]
        df = pd.DataFrame(data[1:], columns=data[0])
        # 不能无脑fillna(0)，先这样吧
        '''不能无脑fillna(0），先这样吧'''
        df["报关单商品序号"] = pd.to_numeric(df["报关单商品序号"], errors="coerce").fillna(0).astype(int)
        df = df.sort_values(by=['报关单商品序号'])

        with pd.ExcelWriter(self.dst_excel, engine="openpyxl") as writer:    # mode w 表示覆盖, a 表示追加, exist_ok=True 表示不报错
            df.to_excel(writer, sheet_name=title, index=False)

        # 然后用 openpyxl 写入 qd 数据
        if self.ws_qd:
            src_ws = self.ws_qd
            title = src_ws.title

            wb = load_workbook(self.dst_excel)
            ws = wb.create_sheet(title)
            for row in src_ws:
                row_values = [cell.value for cell in row]
                ws.append(row_values)

            ws_customs = wb.worksheets[0]
            wb.move_sheet(ws_customs, offset=1)
            wb.active = ws_customs
            
            wb.save(self.dst_excel)
            wb.close()

        # 设置df
        self.df_customs = df
        # 计算毛重
        gross_weight = pd.to_numeric(df["毛重"], errors="coerce").fillna(0).sum()
        self.gross_weight = gross_weight        
        # 计算大件数
        pcs_count = round(gross_weight / 150)
        pcs_count = max(pcs_count, 1)
        self.pcs_count = str(pcs_count)
        # 计算运费
        shipping_fee = gross_weight * 3.5
        self.shipping_fee = f"{shipping_fee:.2f}"

        # 复制预览文件
        if self.doc_path:
            doc_name = Path(self.doc_path).name # 包含后缀
            self.dst_doc = self.dst_dir / doc_name
            shutil.copy2(self.doc_path, self.dst_doc)

            # 转换pdf
            self.dst_pdf = self.dst_dir / f"{self.internal_code}.pdf"
            if self.doc_type == ".xps":
                with fitz.open(self.doc_path) as xps:
                    pdfbytes = xps.convert_to_pdf()
                with fitz.open("pdf", pdfbytes) as pdf:
                    pdf.save(self.dst_pdf)
            else:
                shutil.copy2(self.doc_path, self.dst_pdf)

    def build_invoice_head(self):
        template_path = self.config.head_template   # 是个 Path 对象
        dst_path = self.dst_dir / f"{self.internal_code}-{template_path.name}"
        shutil.copy2(template_path, dst_path)

        wb = load_workbook(dst_path)
        # 写入接单信息
        ws_base = wb["关务信息"]
        if ws_base.max_row > 1:
            ws_base.delete_rows_base_info(2, ws_base.max_row - 1)

        headers = [cell.value for cell in ws_base[1]]
        mapping = self.config.base_fields_map.copy()    # 避免污染全局mapping，其实问题不大，因为下面会覆盖
        mapping["企业内部编号"] = ("CONST", self.internal_code)

        result = {}
        for key, (_type, data) in mapping.items():
            if _type != "EMPTY": # 提取非空值字段的数据
                result[key] = data

        row_data = [result.get(h, None) for h in headers]   # 不存在的列名设置为 None
        ws_base.append(row_data)

        # 写入基本信息
        ws_head = wb["基本信息"]
        if ws_head.max_row > 1:
            ws_head.delete_rows_head_info(2, ws_head.max_row - 1)

        headers = [cell.value for cell in ws_head[1]]
        mapping = self.config.head_fields_map.copy()    # 避免污染全局mapping，其实问题不大，因为下面会覆盖
        mapping["企业内部编号"] = ("CONST", self.internal_code)
        mapping["运费数值"] = ("CONST", self.shipping_fee)
        mapping["件数"] = ("CONST", self.pcs_count)
        mapping["毛重（KG）"] = ("CONST", self.gross_weight)

        result = {}
        for key, (_type, data) in mapping.items():
            if _type != "EMPTY": # 提取非空值字段的数据
                result[key] = data

        row_data = [result.get(h, None) for h in headers]   # 不存在的列名设置为 None
        ws_head.append(row_data)

        wb.save(dst_path)
        wb.close()

        self.head_template = dst_path

    def build_invoice_body(self):
        country_dict = self.config.country_dict
        currency_dict = self.config.currency_dict

        df = self.df_customs.copy() # 避免污染原始 df
        df["发票原产国"] = df["发票原产国"].map(country_dict) # map表示用新值替换旧值
        df["币制"] = df["币制"].map(currency_dict)            

        mapping = self.config.inv_fields_map.copy() # 避免污染全局mapping，其实问题不大，因为下面会覆盖
        mapping["发票号"] = ("CONST", self.internal_code)
        mapping["大件数"] = ("CONST", self.pcs_count)

        result = {}
        for key, (_type, src) in mapping.items():
            if _type == "VAR":
                result[key] = df[src].tolist()
            elif _type == "CONST":
                result[key] = [src] * len(df)

        out_df = pd.DataFrame(result)
        
        template_path = self.config.body_template   # 是个 Path 对象
        dst_path = self.dst_dir / f"{self.internal_code}-{template_path.name}"
        shutil.copy2(template_path, dst_path)

        # 写入数据
        wb = load_workbook(dst_path)
        ws = wb.active

        header_row = 2
        headers = [cell.value for cell in ws[header_row]]

        if ws.max_row > header_row:
            ws.delete_rows(header_row + 1, ws.max_row - header_row)

        # 写入新数据
        for idx, row in out_df.iterrows():
            row_data = [row.get(h, None) for h in headers]  # 不存在的列名设置为 None
            ws.append(row_data)
        
        wb.save(dst_path)
        wb.close()

        self.body_template = dst_path

    def upload_and_download(self):
        inv_excel, ccs_excel = self.uploader.upload_and_download(
            head_excel=self.head_template,
            body_excel=self.body_template,
            internal_code=self.internal_code,
            dst_dir=self.dst_dir
        )
        self.inv_excel = inv_excel
        self.ccs_excel = ccs_excel

    def update_inv(self):
        self.inv_updater.update_inv(
            inv_excel=self.inv_excel,
            template_excel=self.config.inv_head,
            stamp=self.config.stamp
        )

# 产线类
class Line:
    def __init__(self, line_no, prefix, config):
        self.line_no = line_no
        self.prefix = prefix
        self.config = config    # 只存引用，不复制状态

        # 运行状态（任务完成标记）
        self.datasets_found = False
        self.docs_found = False
        self.datasets_exported = False
        self.body_templates_built = False
        self.head_templates_built = False
        self.invoice_downloaded = False
        self.header_updated = False

        # 数据容器对象
        self.excel_path = None  # Path
        self.customs_sheets = []
        self.datasets = []

    # ---------- 派生属性 (实时读取 config) ----------
    @property
    def src_folder(self):
        return self.config.src_folder
    
    @property
    def internal_code(self):
        return f"{self.prefix}{self.config.serial_code}"
    
    @property
    def exception(self):
        return self.config.exception
    
    @property
    def src_folder_set(self):
        return self.src_folder and self.src_folder.exists()
    
    @property   # 可以自动调用，不需要加括号
    def templates_built(self):
        return self.body_templates_built and self.head_templates_built

    def __repr__(self):
        return (
            f"<Line {self.line_no}\n"
            f"  internal_code={self.internal_code}\n"
            f"  excel_path={self.excel_path}\n"
            # f"  datasets={self.datasets}\n"
            f">"
        )

    def find_datasets(self):
        if not self.src_folder_set:
            return "error", f"请先选择文件夹"
        
        ws_pairs = []

        keys_qd = ["备案序号", "报关单商品序号", "商品编码", "商品名称", "申报数量", "法定数量", "申报单价"]
        keys_customs = ["原始货物备件号", "报关单商品序号", "中文品名", "数量", "净重", "金额", "指令单号"]

        for file in self.src_folder.glob("*.xlsx"):
            if self.line_no not in file.name:   # 如果文件名不包含产线号，跳过，这种情况必须人工确认
                continue
            
            try:    # 有必要try catch，防止单个文件读取失败导致整体中断？
                wb = load_workbook(file, data_only=True)

                # 第一遍，识别qd sheet
                qd_map = {}
                for ws in wb.worksheets:
                    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
                    is_qd = sum(key in headers for key in keys_qd) >= len(keys_qd) * 0.7
                    if is_qd:
                        col = 1 # 暂时固定核注清单在第一列
                        qd_code = ws.cell(row=2, column=col).value
                        if qd_code:
                            qd_code = qd_code.upper().strip()
                            qd_map[qd_code] = ws
                
                # 第二遍，识别customs sheet，创建DataSet，再匹配qd sheet
                for ws in wb.worksheets:
                    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
                    is_customs = sum(key in headers for key in keys_customs) >= len(keys_customs) * 0.7
                    if is_customs:
                        self.excel_path = file.resolve()
                        # 识别 QD号
                        col = headers.index("指令单号") + 1
                        qd_code = ws.cell(row=2, column=col).value
                        if qd_code:
                            qd_code = qd_code.upper().strip()
                            # 匹配 QD sheet
                        matched_qd = qd_map.get(qd_code, None)    # ✅ 直接匹配，不需要再循环
                        # 记录匹配结果
                        ws_pairs.append((ws, matched_qd, qd_code))

            except Exception as e:
                path = file.resolve()
                msg = f"Excel文件读取异常: {e}"
                self.exception.add(f"{msg}{path}")
            finally:
                wb.close()
            
            # 处理完一个Excel文件后，如果有匹配结果，就不用再循环了
            if ws_pairs:
                break
        # ✅ 创建DataSet
        datasets = []
        for idx, (ws, ws_qd, qd_code) in enumerate(ws_pairs, start=1):
            ds = DataSet(self.internal_code, config=self.config)
            ds.index = idx
            ds.internal_code = f"{self.internal_code}-{idx}" if len(ws_pairs) > 1 else self.internal_code
            ds.ws_customs = ws
            ds.ws_qd = ws_qd
            ds.qd_code = qd_code
            datasets.append(ds)

        # ✅ 更新状态
        self.datasets = datasets
        self.datasets_found = bool(datasets)    # True or False

        # 返回结果（不负责 UI）
        if not datasets:
            msg = f"产线 {self.line_no} 未找到任何底账数据"
            self.exception.add(msg)
            return "error", f"产线 {self.line_no} 未找到任何底账数据"
        
        # 可能存在部分 QD 匹配失败 -> 非致命（允许继续）
        unmatched = [ds.qd_code for ds in datasets if not ds.ws_qd]
        if unmatched:
            msg = f"产线 {self.line_no} 的 {unmatched} 未找到匹配的核注清单sheet"
            self.exception.add(msg)
            return "warn", f"产线 {self.line_no} 的 {unmatched} 未找到匹配的核注清单sheet"
        
        return "ok", f"产线 {self.line_no} 找到 {len(datasets)} 套底账数据"

    def find_docs(self):
        if not self.datasets_found:            
            return "error", f"产线 {self.line_no} 未找到底账数据，无法查找预览文件"

        for ds in self.datasets:
            ds.find_doc()
        
        self.docs_found = True  # 没找到doc也没关系

        found = [ds for ds in self.datasets if ds.doc_path]
        not_found = [ds for ds in self.datasets if not ds.doc_path]

        if not found:
            msg = f"产线 {self.line_no} 没有匹配到任何预览文件"
            self.exception.add(msg)
            return "warning", msg
        
        if not_found:
            msg = f"产线 {self.line_no} 中部分数据未匹配到预览文件: {[ds.qd_code for ds in not_found]}"
            self.exception.add(msg)
            return "warning", msg
        
        return "ok", f"产线 {self.line_no} 全部数据匹配到预览文件"

    def export_datasets(self):
        if not self.datasets_found:
            return "error", f"产线 {self.line_no} 未找到底账数据，无法导出"

        for ds in self.datasets:
            ds.export()

        exported = [ds for ds in self.datasets if ds.dst_excel.exists()]
        not_exported = [ds for ds in self.datasets if not ds.dst_excel.exists()]

        if not exported:
            msg = f"产线 {self.line_no} 底账文件全部导出失败"
            self.exception.add(msg)
            return "warning", msg
        
        self.datasets_exported = True

        if not_exported:
            msg = f"产线 {self.line_no} 中部分底账未导出: {[ds.internal_code for ds in not_exported]}"
            self.exception.add(msg)
            return "warning", msg
        
        msg = (
            f"产线 {self.line_no} 的底账文件已导出:\n"
            + "\n".join(
                f"内部编号：{ds.internal_code}\n毛重：{round(ds.gross_weight, 10)}，运费：{ds.shipping_fee}，件数：{ds.pcs_count}"
                for ds in exported
            )
        )
        return "ok", msg

    def build_invoice_bodys(self):
        if not self.datasets_exported:
            return "error", f"产线 {self.line_no} 还未导出底账，无法生成发票表体"

        for ds in self.datasets:
            ds.build_invoice_body()

        ok = [ds for ds in self.datasets if ds.body_template.exists()]
        fail = [ds for ds in self.datasets if not ds.body_template.exists()]

        if not ok:
            msg = f"产线 {self.line_no} 发票表体全部生成失败"
            self.exception.add(msg)
            return "warning", msg

        self.body_templates_built = True

        if fail:
            msg = f"产线 {self.line_no} 中部分发票表体生成失败: {[ds.internal_code for ds in fail]}"
            self.exception.add(msg)
            return "warning", msg

        return "ok", f"产线 {self.line_no} 发票表体生成完成"

    def build_invoice_heads(self):
        if not self.datasets_exported:
            return "error", f"产线 {self.line_no} 还未导出底账，无法生成发票表头"

        for ds in self.datasets:
            ds.build_invoice_head()

        ok = [ds for ds in self.datasets if ds.head_template.exists()]
        fail = [ds for ds in self.datasets if not ds.head_template.exists()]

        if not ok:
            msg = f"产线 {self.line_no} 发票表头全部生成失败"
            self.exception.add(msg)
            return "warning", msg

        self.head_templates_built = True

        if fail:
            msg = f"产线 {self.line_no} 中部分发票表头生成失败: {[ds.internal_code for ds in fail]}"
            self.exception.add(msg)
            return "warning", msg

        return "ok", f"产线 {self.line_no} 发票表头生成完成"
    
    def upload_and_download(self):
        if not self.templates_built:
            return "error", f"产线 {self.line_no} 的导入表体或者表头还未生成，无法上传"
        
        for ds in self.datasets:
            ds.upload_and_download()

        downloaded = [ds for ds in self.datasets if ds.inv_excel.exists()]
        not_downloaded = [ds for ds in self.datasets if not ds.ccs_excel.exists()]

        if not downloaded:
            msg = f"产线 {self.line_no} 发票全部导出失败"
            self.exception.add(msg)
            return "warning", msg

        self.invoice_downloaded = True

        if not_downloaded:
            msg = f"产线 {self.line_no} 中部分发票上传失败: {[ds.internal_code for ds in not_downloaded]}"
            self.exception.add(msg)
            return "warning", msg
        
        return "ok", f"产线 {self.line_no} 数据已导入并下载完成"
    
    def update_inv(self):
        if not self.invoice_downloaded:
            return "error", f"产线 {self.line_no} 还未导出发票，无法更新"
        
        for ds in self.datasets:
            ds.update_inv()
        
        return "ok", f"产线 {self.line_no} 发票表头更新完成"

# 动作类
class Actions:
    def __init__(self, config, lines):
        self.config = config
        self.lines = lines  # Line对象的list

        self.datasets_exported = False
        self.templates_built = False
        self.invoice_downloaded = False

    @property
    def log(self):
        return self.config.log

    @property
    def src_folder(self):
        return self.config.src_folder

    @property
    def src_folder_set(self):
        return self.src_folder and self.src_folder.exists()
    
    def _run(self, fn):
        """统一执行并捕获 (level, msg) 返回结果"""
        """接收执行结果，如果ok或warn，返回True，否则返回False"""
        # try 可以捕获函数中的错误
        try:
            level, msg = fn()   # 获取函数的运行结果
        except Exception as e:
            level, msg = "error", f"执行 {fn.__name__} 出错: {e}"    # 如果函数运行出错，改写运行结果

        if level == "ok":   # 如果运行成功，传递成功信息msg
            self.log.ok(msg)
            return True     # 传递出记录运行结果的状态
        elif level == "warn":
            self.log.warn(msg)
            return True
        else:   # "error"   # 如果运行失败，传递失败信息msg
            self.log.error(msg)
            return False    
        
    # 动作
    def export_datasets(self):
        if not self.src_folder_set:
            self.log.warn("请先选择文件夹!")
            return

        for line in self.lines:
            if not self._run(line.find_datasets):   # 先运行函数，运行的过程中会打印日志，然后检查运行结果，如果运行失败，并结束
                return
            if not self._run(line.find_docs):  # 如果运行失败
                return
            if not self._run(line.export_datasets):  # 如果运行失败
                return

        self.datasets_exported = True
        # self.log.ok("数据集已成功提取\n")

    def build_templates(self):
        if not self.datasets_exported:
            self.log.warn("请先执行：提取数据集")
            return
        
        for line in self.lines:
            if not self._run(line.build_invoice_heads):
                return
            if not self._run(line.build_invoice_bodys):
                return

        self.templates_built = True
        self.log.ok("发票模板生成完成\n")

    def upload_and_download(self):
        if not self.templates_built:
            self.log.warn("请先执行：生成发票模板")
            return
        
        for line in self.lines:
            if not self._run(line.upload_and_download):
                return
        
        self.invoice_downloaded = True
        # self.log.ok("数据已导入并下载完成\n")

    def update_inv(self):
        if not self.invoice_downloaded:
            self.log.warn("请先执行：导出发票")
            return
        
        for line in self.lines:
            if not self._run(line.update_inv):
                return

        # self.log.error("“修改表头”功能尚未实现")
        # print("“修改表头”功能尚未实现")

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