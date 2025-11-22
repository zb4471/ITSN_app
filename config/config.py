
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
