
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
