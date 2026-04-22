import os
import sys
import time
import hashlib
import requests
import configparser
import subprocess
from datetime import datetime
from playwright.sync_api import sync_playwright

# ================= 1. 环境与配置自适应 =================
if getattr(sys, 'frozen', False):
    # 【生产环境】如果是打包后的 exe，获取 exe 所在的绝对目录
    PROG_DIR = os.path.dirname(sys.executable)
    # 仅在 exe 模式下，才强制重定向浏览器内核到本地文件夹，实现免安装便携化
    BROWSERS_PATH = os.path.join(PROG_DIR, "pw-browsers")
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = BROWSERS_PATH
else:
    # 【开发环境】如果是源码运行，获取 main.py 所在的绝对目录
    PROG_DIR = os.path.dirname(os.path.abspath(__file__))
    # 源码模式下不设置环境变量，让 Playwright 自动去读取你电脑里之前已经装好的全局浏览器内核
    BROWSERS_PATH = None

CONFIG_PATH = os.path.join(PROG_DIR, "config.ini")

def load_config():
    if not os.path.exists(CONFIG_PATH):
        print(f"\n[错误] 未找到配置文件: {CONFIG_PATH}")
        print(f"请确保 config.ini 与程序放在同一个文件夹内。")
        input("\n按回车键退出...")
        sys.exit(1)

    _config = configparser.ConfigParser()
    _config.read(CONFIG_PATH, encoding="utf-8")
    return _config

# 加载配置变量
try:
    conf = load_config()
    ACCOUNT = conf.get("Credentials", "account")
    RAW_PASSWORD = conf.get("Credentials", "password")
    START_TIME = conf.get("Query", "start_time")
    END_TIME = conf.get("Query", "end_time")
except Exception as e:
    print(f"[错误] 读取 config.ini 格式失败: {e}")
    input("\n按回车键退出...")
    sys.exit(1)

BASE_URL = "http://sw.cesaas.com:81"

# ================= 2. 核心业务类 =================
class OrderHistoryRPA:
    def __init__(self):
        self.session = requests.Session()
        self.token = ""
        self.headers = {"Content-Type": "application/json"}

    def login(self):
        """登录并自动处理密码加密"""
        url = f"{BASE_URL}/admin/login/login"

        # 自动将明文密码转化为 MD5 密文
        md5_obj = hashlib.md5()
        md5_obj.update(RAW_PASSWORD.encode('utf-8'))
        encrypted_pwd = md5_obj.hexdigest()

        payload = {
            "ati": "3643730255656",
            "IsOms": 1,
            "Account": ACCOUNT,
            "Password": encrypted_pwd,
            "Date": datetime.now().strftime("%Y-%m-%d")
        }

        print(f"[*] 正在登录账号: {ACCOUNT} ...")
        try:
            resp = self.session.post(url, json=payload, timeout=15).json()
            if resp.get("IsSuccess"):
                self.token = resp["TModel"]["Token"]
                self.headers["Authorization"] = f"Bearer {self.token}"
                print("[+] 登录成功，获取到有效授权 Token")
            else:
                raise Exception(resp.get("Message", "未知错误"))
        except Exception as e:
            raise Exception(f"登录接口连接失败: {e}")

    def fetch_all_tickets(self):
        """获取全量工单，基于 PageCount 进行精确分页"""
        url = f"{BASE_URL}/workflow/order_history/List"
        all_items = []
        page_index = 1
        page_size = 50
        total_pages = 1

        while page_index <= total_pages:
            payload = {
                "PageIndex": page_index,
                "PageSize": page_size,
                "Sort": [{"Field": "TicketId", "Value": "DESC"}],
                "Filter": [
                    {"Field": "Status", "Value": "(10,20,30)", "Operator": "in", "Connector": "and"},
                    {"Field": "CategoryId", "Value": "(10,11,12,13,14,15,18,21,17)", "Operator": "in", "Connector": "and"},
                    {"Field": "CreateTimeStart", "Value": START_TIME, "Operator": ">=", "Connector": "and"},
                    {"Field": "CreateTimeEnd", "Value": END_TIME, "Operator": "<=", "Connector": "and"},
                    {"Field": "MyTicket", "Value": 1, "Operator": "=", "Connector": "and"}
                ]
            }

            resp = self.session.post(url, json=payload, headers=self.headers).json()

            if page_index == 1:
                total_pages = resp.get("PageCount", 1)
                record_count = resp.get("RecordCount", 0)
                print(f"[!] 检索完成: 时间段内共发现 {record_count} 条工单，分 {total_pages} 页处理")
                if record_count == 0: break

            # 兼容多种返回结构的提取逻辑
            items = []
            t_model = resp.get("TModel")
            if isinstance(t_model, list):
                items = t_model
            elif isinstance(t_model, dict):
                items = t_model.get("Items", [])

            all_items.extend(items)
            print(f"[*] 正在抓取第 {page_index}/{total_pages} 页数据...")
            page_index += 1

        return all_items

    def save_pdf(self, page, item):
        """导出PDF并按责任原因分类"""
        # 安全提取字段
        ticket_id = item.get("TicketId")
        category_id = item.get("CategoryId")
        ticket_no = str(item.get("TicketNo") or f"T{int(time.time())}").strip()
        resp_name = str(item.get("ResponsibleName") or "未分类").strip()
        reason_one = str(item.get("ReasonOne") or "其他原因").strip()

        if not ticket_id: return

        # 清洗非法字符，防止 Windows/Mac 建文件夹报错
        safe_resp = resp_name.replace('/', '_').replace('\\', '_')
        safe_reason = reason_one.replace('/', '_').replace('\\', '_')

        folder_path = os.path.join(PROG_DIR, "工单导出结果", safe_resp, safe_reason)
        os.makedirs(folder_path, exist_ok=True)

        file_path = os.path.join(folder_path, f"{ticket_no}.pdf")
        print_url = f"{BASE_URL}/workflow/order_print?categoryId={category_id}&ticketId={ticket_id}&nos={ticket_no}"

        try:
            page.goto(print_url, wait_until="networkidle", timeout=60000)
            time.sleep(1) # 等待样式渲染
            page.pdf(path=file_path, format="A4", print_background=True)
            print(f"    [成功] {safe_resp} > {safe_reason} > {ticket_no}")
        except Exception as e:
            print(f"    [失败] 工单 {ticket_no}: {e}")

    def run(self):
        self.login()
        tickets = self.fetch_all_tickets()
        if not tickets:
            print("[!] 未查询到符合条件的工单，请检查配置的时间范围。")
            return

        print(f"[+] 启动无头浏览器，准备下载 {len(tickets)} 份 PDF ...")
        with sync_playwright() as p:
            # 兼容某些环境下的执行文件路径
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            page = context.new_page()

            for index, item in enumerate(tickets):
                print(f"[*] 进度 [{index+1}/{len(tickets)}] No: {item.get('TicketNo')}")
                self.save_pdf(page, item)

            browser.close()

# ================= 3. 执行入口与交互处理 =================
def install_chromium():
    """自动化安装 Chromium 浏览器内核"""
    try:
        if getattr(sys, 'frozen', False):
            print(f"\n[*] 正在安装便携版 Chromium 浏览器组件到: {BROWSERS_PATH}")
            print("[*] 文件体积较大（约 100MB+），请耐心等待...")
            from playwright._impl._driver import compute_driver_executable, get_driver_env
            env = get_driver_env()
            executable = compute_driver_executable()
            subprocess.check_call([executable, 'install', 'chromium'], env=env)
        else:
            print(f"\n[*] 正在安装全局 Chromium 浏览器组件到系统默认目录...")
            print("[*] 文件体积较大（约 100MB+），请耐心等待...")
            subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
        print("\n[+] 浏览器组件下载安装成功！")
    except Exception as e:
        print(f"\n[致命错误] 浏览器安装失败: {e}")
        print("提示：请检查网络连接，或尝试关闭系统代理后重试。")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "install":
        install_chromium()
        input("\n按回车键退出...")
        sys.exit(0)

    try:
        rpa = OrderHistoryRPA()
        rpa.run()
        print("\n" + "="*30)
        print(">>> 导出任务全部完成！")
        print(f">>> 文件保存在: {os.path.join(PROG_DIR, '工单导出结果')}")
    except Exception as e:
        error_msg = str(e)
        if "Executable doesn't exist" in error_msg or "playwright install" in error_msg:
            print("\n[!] 运行中断：未检测到 Chromium 浏览器运行环境。")
            choice = input("是否立即自动下载所需的浏览器组件？(Y/N): ").strip().upper()
            if choice == 'Y':
                install_chromium()
                print("\n[!] 准备就绪，请重新运行本程序执行导出任务。")
            else:
                print("\n[!] 取消安装。如需手动安装，请在命令行执行: 你的程序.exe install")
        else:
            print(f"\n[致命错误]: {error_msg}")

    input("\n程序执行结束，按回车键退出...")
