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
