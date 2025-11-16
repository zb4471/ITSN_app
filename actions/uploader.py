from pathlib import Path
from playwright.sync_api import sync_playwright
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import expect
import subprocess
import pygetwindow as gw
import time

LOG_IN_URL = "https://ab.eptrade.cn/swgd-imap-user-center/#/login"
TAEGET = "https://ab.eptrade.cn/swgd-imap-user-center/#/cusDeclareJDList"

CORP_ID = "bakerhughes"
USER_ID = "ITSN503016233"
PASSWORD = "888888aA"
# “调试”的英文是：
CHROME_OPTION = r'--remote-debugging-port=9222 --user-data-dir="C:\PlaywrightProfiles\fenbo"'
CHROME_OPTION_FULL = r'"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\PlaywrightProfiles\fenbo"'

class Uploader:
    def __init__(self):
        PROFILE_DIR = r"C:\PlaywrightProfiles\fenbo"
        
        self.CHROME_CMD = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            "--remote-debugging-port=9222",
            f"--user-data-dir={PROFILE_DIR}"
        ]
        self.LOG_IN_URL = "https://ab.eptrade.cn/swgd-imap-user-center/#/login"
        self.TAEGET = "https://ab.eptrade.cn/swgd-imap-user-center/#/cusDeclareJDList"

        self.CORP_ID = "bakerhughes"
        self.USER_ID = "ITSN503016233"
        self.PASSWORD = "888888aA"

    def upload_and_download(self, head_excel, body_excel, internal_code, dst_dir):
        
        def wait_for_browser_ready(p, retry=10, delay=1):   # 等待10秒
            for i in range(retry):
                try:
                    browser = p.chromium.connect_over_cdp("http://localhost:9222")
                    return browser
                except PlaywrightError as e:
                    time.sleep(delay)
            raise Exception("等待超时，启动浏览器失败，请手动打开关务自动化专用浏览器")

        with sync_playwright() as p:
            try:
                browser = p.chromium.connect_over_cdp("http://localhost:9222")
            except PlaywrightError as e:
                # 自动打开调试浏览器
                subprocess.Popen(self.CHROME_CMD, shell=False)
                browser = wait_for_browser_ready(p)

            ctx = browser.contexts[0]   # 复用第一个浏览器上下文
            page = None
            blank_page = None
            for pg in ctx.pages:
                title = pg.title()
                if "磐石关务" in title:
                    page = pg
                    break
                elif "新标签页" in title:
                    blank_page = pg

            page = page or blank_page or ctx.new_page()

            # 将标签页和窗口都激活
            page.bring_to_front()
            time.sleep(0.5)
            for w in gw.getWindowsWithTitle("磐石关务"):
                w.activate()
                break
            
            page.goto(self.TAEGET)
            # 等待网络空闲（所有请求完成）
            page.wait_for_load_state('networkidle', timeout=30_000) # 意思是等待直到没有网络请求？
            # 检查是否会跳登录页面
            if page.url == self.LOG_IN_URL:
                # 帮助用户输入用户名密码
                try:
                    page.get_by_placeholder("请输入企业代码").fill(self.CORP_ID)
                    page.get_by_placeholder("请输入账号/手机号").fill(self.USER_ID)
                    page.get_by_placeholder("请输入密码").fill(self.PASSWORD)
                    page.get_by_placeholder("请输入验证码").focus()
                except Exception as e:
                    print(e)
                    pass

                raise Exception("已打开登录页面，请手动输入验证码并登录，然后再重新执行")            

            # 等待数据加载
            corp_input = page.locator("label:text-is('经营单位') + div input")
            # 不为空才等待
            if corp_input.input_value() != "":
                first_tr = page.locator("table.el-table__body tbody tr").nth(0)
                first_tr.wait_for(state="visible", timeout=5000)
                page.wait_for_timeout(500)

            # 设置搜索条件
            corp_input = page.locator("label:text-is('经营单位') + div input")
            v = "上海上实外联发进出口有限公司"
            if corp_input.input_value() != v:
                corp_input.click()
                page.wait_for_timeout(300)
                li = page.locator(f"li:text-is('{v}')")
                li.click(timeout=5000)
            
            user_input = page.locator("label:text-is('操作人') + div input")
            v = "ITSN503016233"
            if user_input.input_value() != v:
                user_input.click()
                page.wait_for_timeout(300)
                li = page.locator(f"li:text-is('{v}')")
                li.click(timeout=5000)

            # 查询
            page.wait_for_timeout(300)
            query_button = page.get_by_role("button", name="查询")
            query_button.click(timeout=5000)
            page.wait_for_timeout(300)

            def upload_file(page, button_name, file_path):
                # 拦截file_chooser
                with page.expect_file_chooser() as fc:
                    page.get_by_role("button", name=button_name).click()
                fc.value.set_files(file_path)

            upload_file(page, "表头批量导入", head_excel)
            page.wait_for_timeout(1000)

            # 读取回执
            msg_div = page.locator("div.ep-message--success").filter(visible=True)
            try:
                msg_div.wait_for(state="visible", timeout=5000)
            except Exception:
                raise Exception("发票表头导入失败，请自行检查")

            # 点击编辑按钮
            row = page.locator("tr", has_text=internal_code)
            edit_button = row.get_by_role("button", name="编辑")
            edit_button.click()
            page.wait_for_timeout(200)

            # 点击保存按钮
            header = page.locator("div.ep-modal-header :text-is('接单')")   # 空格表示子元素
            modal = page.locator("div.ep-modal").filter(has=header)
            modal.get_by_role("button", name="保存").click()
            page.wait_for_timeout(200)

            # 等待frame内部的元素准备好
            frame = page.frame(name="outFrame", url="*cusDeclarationSingle*")
            input = frame.locator("label:text('监管方式') + div input")
            input.wait_for(state="visible")
            # print("等待单证明细页面:", input.input_value())

            page.wait_for_timeout(200)
            
            # # 将按钮显示区域扩大，但是如果用了 dispatch_event click 的话就不需要了
            # block = frame.locator("div.imgd-cargo[componentkey='cargoWholeCountry'] > div.block")
            # block.evaluate("el => el.style.height = '60px'")

            # 点击发票导入
            import_button = frame.get_by_role("button", name="发票导入")
            import_button.dispatch_event("click")

            # 点击归并导入
            header = frame.locator("div.ep-modal-header :text-is('提示')")   # 空格表示子元素
            modal = frame.locator("div.ep-modal").filter(has=header)
            merge_import = modal.get_by_role("button", name="归并导入", exact=True)

            with page.expect_file_chooser() as fc: # 必须用page
                merge_import.click()
            fc.value.set_files(body_excel)

            # 读取回执
            modal = frame.locator("div.ep-modal", has_text="校验结果展示")
            msg = modal.get_by_role("paragraph").first.inner_text()
            print("发票表头导入结果:", msg)

            if "导入成功" not in msg:
                raise Exception("发票表头导入失败，请自行检查")

            print("✅ 已上传发票模板")
            page.wait_for_timeout(500)

        
            # 下载模板
            page.goto(TAEGET)

            # 勾选目标行
            row = page.locator("tr").filter(has_text=internal_code, visible=True).nth(0)
            checkbox = row.get_by_role("checkbox")
            is_checked = checkbox.is_checked(timeout=5000)
            label = row.locator("label.el-checkbox")
            if not is_checked:
                label.click(timeout=5000)

            page.wait_for_timeout(200)

            # 发票导出button
            inv_exp_button = page.get_by_role("button", name="发票导出")

            with page.expect_download() as dl:
                inv_exp_button.click()
            download = dl.value
            inv_excel = dst_dir / download.suggested_filename
            download.save_as(inv_excel)
            page.wait_for_timeout(300)

            if not is_checked:
                label.click(timeout=2000)

            # CCS导出button
            ccs_exp_button = page.get_by_role("button", name="CCS导出")
            ccs_exp_button.click()

            # 点击EXCEL选项
            page.wait_for_timeout(300)
            xl_button = page.get_by_role("button", name="EXCEL导出")

            with page.expect_download() as dl:
                xl_button.click()
            download = dl.value
            ccs_excel = dst_dir / download.suggested_filename
            download.save_as(ccs_excel)

            page.wait_for_timeout(1000)

            return inv_excel, ccs_excel

if __name__ == "__main__":
    uploader = Uploader()

    dir = Path(r"D:\共享文件夹\金丽华\59-251017-004603 分拨数据 20251001-20251016")
    internal_code = "FBIT251059"
    dst_dir = dir / internal_code
    head_excel = dir / internal_code / f"{internal_code}-接单信息模板.xlsx"
    body_excel = dir / internal_code / f"{internal_code}-发票导入模板.xlsx"

    uploader.upload_and_download(head_excel, body_excel, internal_code, dst_dir)
