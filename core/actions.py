
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
