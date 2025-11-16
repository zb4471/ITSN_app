import customtkinter as ctk
from tkinter import filedialog
from pathlib import Path
import os
import subprocess
import threading

ctk.set_appearance_mode("dark")      # "dark" / "light" / "system"
ctk.set_default_color_theme("blue")  # 或 "green", "dark-blue"

# log 样式（简单表示：颜色只是前景色）
COLORS = {
    "info": "#DDE6FF", # 淡蓝
    "ok": "#66FF66", # 绿
    "warn": "#FFD966", # 黄
    "error": "#FF6A6A", # 红
    "path": "#9CD7FF", # 天蓝
}

class LogView:
    def __init__(self, log_widget, colors):
        self.log = log_widget   # 日志控件
        self.colors = colors

        # 绑定 路径打开函数
        self.log.tag_bind("path", "<Button-1>", self._open_path)

    def info(self, text):

        self._write(text, "info")

    def ok(self, text):
        self._write(text, "ok")

    def warn(self, text):
        self._write(text, "warn")

    def error(self, text):
        self._write(text, "error")
    
    def path(self, path_obj):
        self._write_clickable_path(path_obj)    # Path 对象
    
    # === 内部方法 ===    
    def _write(self, text, level="info"):
        # 若不在主线程，则用 after 调度到主线程执行
        if threading.current_thread() != threading.main_thread():
            self.log.after(0, lambda: self._write(text, level))
            return

        self.log.configure(state="normal")

        # 颜色处理
        self.log.insert("end", f"[{level.upper()}]\t{text}\n", (level,)) # (tag,) 是元组，因为可能有很多tag
        color = self.colors.get(level, self.colors["info"])
        self.log.tag_config(level, foreground=color)

        self.log.configure(state="disabled")
        self.log.see("end")

        # # 🔔 自动弹窗提示
        # if level == "error":
        #     self._show_popup("错误", text, self.colors["error"])
        #     # messagebox.showerror("错误", text, parent=self.log)
        # elif level == "warn":
        #     self._show_popup("警告", text, self.colors["warn"])
        #     # messagebox.showwarning("警告", text)

    def _write_clickable_path(self, path):
        path = Path(path)
        if not path.exists():
            self._write(f'路径不存在: 📂 "{path}"', "error")
            return
        tag = f"path_{id(path)}" # 例如：path_2071992375152, 为什么这样命名呢？因为这样才能区分不同的路径
        text = path.name if path.is_file() else f"{path.name}/"

        self.log.configure(state="normal")
        # 先写入一个图标
        folderico_tag = "folderico"
        self.log.insert("end", "📂 ", (folderico_tag,))
        self.log.tag_config(folderico_tag, foreground=self.colors["info"])
        # 再写入路径
        self.log.insert("end", f" {text}\n", (tag,)) # (tag,) 是元组
        self.log.tag_config(tag, foreground=self.colors["path"], underline=True)
        self.log.tag_bind(tag, f"<Button-1>", lambda e, p=path: self._open_path(p))
        self.log.tag_bind(tag, f"<Enter>", lambda e, p=path: self.log.configure(cursor="hand2"))
        self.log.tag_bind(tag, f"<Leave>", lambda e, p=path: self.log.configure(cursor="arrow"))
        self.log.configure(state="disabled")
        self.log.see("end")

    # ✅ 打开路径（文件或文件夹）
    def _open_path(self, path):
        path = Path(path)
        if path.is_file():
            subprocess.Popen(["explorer", "/select,", str(path)])    # path 不用转str
        elif path.is_dir():
            os.startfile(path)

def _wrap_in_thread(fn):
    # 简单把耗时任务放到线程，避免 GUI 卡死
    def run():
        try:
            fn()
        except Exception as e:
            # log_view.error(f"处理失败：{e}")
            print(f"[THREAD ERROR] {e}")
    t = threading.Thread(target=run, daemon=True)
    t.start()

ACTIONS = [
    ("提取数据集", "export_datasets"),
    ("生成发票模板", "build_templates"),
    ("上传系统并下载", "upload_and_download"),
    ("修改表头", "update_inv"),
]

def create_gui(config, actions):
    root = ctk.CTk()
    root.title("分拨数据处理工具")
    root.geometry("650x400+50+500")
    root.resizable(False, False) 

    # 左侧按钮区
    left_frame = ctk.CTkFrame(root)
    left_frame.pack(side="left", fill="y", padx=10, pady=10)

    def choose_folder():
        f = filedialog.askdirectory(title="请选择要处理的文件夹", initialdir=r"D:\共享文件夹\金丽华")
        if f:
            label_folder.configure(text=f'已选择文件夹：{Path(f).name}')
            config.set_src_folder(f)

    def reset_all():
        config.src_folder = None
        config.serial_code = ""
        label_folder.configure(text="未选择文件夹")
        log_widget.configure(state="normal")
        log_widget.delete("1.0", "end")
        log_widget.configure(state="disabled")
    
    # 创建选择文件夹按钮
    ctk.CTkButton(left_frame, text="选择文件夹", command=choose_folder).pack(padx=20, pady=(20, 10))
    # 批量创建动作按钮
    for text, action_name in ACTIONS:
        action = getattr(actions, action_name)   # ✅ 动态取方法
        ctk.CTkButton(
            left_frame,
            text=text,
            fg_color="green",
            command=lambda a=action: _wrap_in_thread(a)
        ).pack(padx=20, pady=10)
    # 创建退出按钮
    ctk.CTkButton(left_frame, text="退出", fg_color="red", command=root.destroy).pack(side="bottom", padx=20, pady=10)
    # 创建重置按钮
    ctk.CTkButton(left_frame, text="重置", fg_color="gray", command=reset_all).pack(side="bottom", padx=20, pady=10)

    
    # 右侧信息/日志区
    right_frame = ctk.CTkFrame(root)
    right_frame.pack(side="right", fill="both", expand=True, padx=(0, 10), pady=10)

    label_header = ctk.CTkLabel(
        right_frame,
        text="欢迎使用分拨数据处理工具！",
        font=("Roboto", 20, "bold"),
        text_color=("#66FF66"),
        justify="center"
    )
    label_header.pack(padx=10, pady=10)

    label_folder = ctk.CTkLabel(right_frame, text="未选择文件夹")
    label_folder.pack(anchor="w", padx=10, pady=4)

    log_widget = ctk.CTkTextbox(right_frame, width =500, height=200, state="disabled", cursor="arrow")
    log_widget.pack(fill="both", expand=True) #, padx=10, pady=10)
    # 禁止选择、编辑、复制
    for event in ["<Button-1>", "<B1-Motion>", "<Double-Button-1>", "<Key>"]: # 只禁止拖动和双击
        log_widget.bind(event, lambda e: "break")

    log_view = LogView(log_widget, COLORS)

    config.set_log(log_view)

    # log_view.warn("⚠️ 请不要在处理过程中关闭窗口")
    # log_view.error("请先选择要处理的文件夹！")
    # log_view.path(r"文件路径：D:\共享文件夹\金丽华\参数表.xlsx")
    # log_view.path(r"D:\共享文件夹\金丽华\59-251017-004603 分拨数据 20251001-20251016")

    root.mainloop()


if __name__ == "__main__":

    # 测试代码
    class DummyConfig:
        src_folder = None
        log = print   # 简单使用 print 代替 LogView

    class DummyLine:
        def __init__(self, name):
            self.name = name
            self.config = DummyConfig()

        def find_datasets(self): print(f"[{self.name}] find_datasets")
        def find_docs(self): print(f"[{self.name}] find_docs")
        def export_datasets(self): print(f"[{self.name}] export_datasets")
        def build_invoice_bodys(self): print(f"[{self.name}] build_invoice_bodys")
        def build_invoice_heads(self): print(f"[{self.name}] build_invoice_heads")

    class DummyActions:
        def __init__(self):
            self.line = DummyLine("TEST")

        def export_datasets(self):
            self.line.find_datasets()
            self.line.find_docs()
            self.line.export_datasets()

        def build_templates(self):
            self.line.build_invoice_bodys()
            self.line.build_invoice_heads()

        def upload_and_download(self):
            print("upload_and_download")

        def adjust_headers(self):
            print("adjust_headers")

    a = DummyActions()

    config = DummyConfig()

    create_gui(a, config)


    # a.export_datasets()
    # a.build_templates()
    # print(a.name)

    # log_view.path(r"文件路径：D:\共享文件夹\金丽华\参数表.xlsx")
    # log_view.path(r"D:\共享文件夹\金丽华\59-251017-004603 分拨数据 20251001-20251016")

    # log.configure(state="normal")
    # log.insert("end", "日志:\n")
    # log.insert("end", "1. 选择文件夹\n")
    # log.insert("end", "2. 开始处理\n")
    # log.insert("end", "3. 退出\n")
    # log.configure(state="disabled")
    # log.see("end")


    # buttons = [
    #     ("选择文件夹", choose_folder),
    #     ("开始处理", start),
    #     ("退出", root.destroy)
    # ]

    # for text, func in buttons:
    #     ctk.CTkButton(left, text=text, command=func).pack(padx=20, pady=10)

    # ACTIONS = [
    #     ("提取数据集", "export_datasets"),
    #     ("生成发票模板", "build_templates"),
    #     ("上传系统并下载", "upload_and_download"),
    #     ("修改表头", "adjust_headers"),
    # ]
    # for text, method_name in ACTIONS:
    # method = getattr(actions, method_name)   # ✅ 动态取方法
    # ctk.CTkButton(
    #     left_frame,
    #     text=text,
    #     fg_color="green",
    #     command=lambda m=method: _wrap_in_thread(m)
    # ).pack(padx=20, pady=10)
