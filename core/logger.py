# 模拟一个print
class PrintLog:
    COLORS = {
        "reset": "\033[0m",
        "info": "",    # 默认颜色
        "ok": "\033[92m",      # 绿色
        "warn": "\033[93m",    # 黄色
        "error": "\033[91m",   # 红色
        "path": "",    # 默认颜色
    }

    SYMBOLS ={
        "info": "ℹ️",
        "ok": "✅",
        "warn": "⚠️",
        "error": "❌",
        "path": "📂",
    }

    def _write(self, level, msg):
        color = self.COLORS.get(level, "")
        reset = self.COLORS["reset"]
        prefix = self.SYMBOLS.get(level, "")
        # print(f"{prefix} {msg}")
        print(f"{color}{prefix} {msg}{reset}")

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
    

if __name__ == "__main__":
    log = PrintLog()

    log.info("This is an info")
    log.ok("This is an ok")
    log.warn("This is a warn")
    log.error("This is an error")
    log.path("This is a path")
