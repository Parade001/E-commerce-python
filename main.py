import os
import sys
import time
import hashlib
import requests
import configparser
import subprocess
from datetime import datetime

# ================= 1. 核心环境劫持 =================
if getattr(sys, 'frozen', False):
    PROG_DIR = os.path.dirname(sys.executable)
else:
    PROG_DIR = os.path.dirname(os.path.abspath(__file__))

BROWSERS_PATH = os.path.join(PROG_DIR, "pw-browsers")
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = BROWSERS_PATH

from playwright.sync_api import sync_playwright

# ================= 2. 配置加载 =================
CONFIG_PATH = os.path.join(PROG_DIR, "config.ini")

def load_config():
    if not os.path.exists(CONFIG_PATH):
        print(f"\n[错误] 未找到配置文件: {CONFIG_PATH}")
        input("\n按回车键退出...")
        sys.exit(1)
    _config = configparser.ConfigParser()
    _config.read(CONFIG_PATH, encoding="utf-8")
    return _config

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

# ================= 3. 核心业务类 =================
class OrderHistoryRPA:
    def __init__(self):
        self.session = requests.Session()
        self.token = ""
        self.headers = {
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept": "application/json, text/plain, */*",
            "Origin": BASE_URL,
            "Referer": f"{BASE_URL}/"
        }

    def login(self):
        url = f"{BASE_URL}/admin/login/login"
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
            resp = self.session.post(url, json=payload, headers=self.headers, timeout=60).json()
            if resp.get("IsSuccess"):
                self.token = resp["TModel"]["Token"]
                self.headers["Authorization"] = f"Bearer {self.token}"
                print("[+] 登录成功")
            else:
                raise Exception(resp.get("Message", "未知错误"))
        except Exception as e:
            raise Exception(f"登录接口连接失败: {e}")

    def fetch_all_tickets(self):
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
                    {"Field": "CreateTimeStart", "Value": START_TIME, "Operator": ">=", "Connector": "and"},
                    {"Field": "CreateTimeEnd", "Value": END_TIME, "Operator": "<=", "Connector": "and"},
                    {"Field": "MyTicket", "Value": 1, "Operator": "=", "Connector": "and"}
                ]
            }

            resp = self.session.post(url, json=payload, headers=self.headers).json()
            if page_index == 1:
                total_pages = resp.get("PageCount", 1)
                record_count = resp.get("RecordCount", 0)
                print(f"[!] 检索完成: 共发现 {record_count} 条工单，分 {total_pages} 页处理")
                if record_count == 0: break

            t_model = resp.get("TModel")
            items = t_model if isinstance(t_model, list) else (t_model.get("Items", []) if t_model else [])
            all_items.extend(items)
            print(f"[*] 正在抓取第 {page_index}/{total_pages} 页数据...")
            page_index += 1

        return all_items

    def save_pdf(self, page, item):
        ticket_id = item.get("TicketId")
        category_id = item.get("CategoryId")
        ticket_no = str(item.get("TicketNo") or f"T{int(time.time())}").strip()
        resp_name = str(item.get("ResponsibleName") or "未分类").strip()
        reason_one = str(item.get("ReasonOne") or "其他原因").strip()

        if not ticket_id: return

        safe_resp = resp_name.replace('/', '_').replace('\\', '_')
        safe_reason = reason_one.replace('/', '_').replace('\\', '_')

        folder_path = os.path.join(PROG_DIR, "工单导出结果", safe_resp, safe_reason)
        os.makedirs(folder_path, exist_ok=True)
        file_path = os.path.join(folder_path, f"{ticket_no}.pdf")

        # 【优化1】：加入时间戳，强制打破 Vue 路由缓存，解决同一页面问题
        print_url = f"{BASE_URL}/workflow/order_print?categoryId={category_id}&ticketId={ticket_id}&nos={ticket_no}&_t={int(time.time()*1000)}"

        try:
            # 【优化2】：将 networkidle 降级为 domcontentloaded，极大提升加载速度
            page.goto(print_url, wait_until="domcontentloaded", timeout=30000)

            # 【优化3】：精准等待“导出PDF”按钮出现
            btn = page.locator("button").filter(has_text="导出PDF").first
            btn.wait_for(state="visible", timeout=15000)

            # 【优化4】：精准等待表格数据渲染（尝试等待 table 或 tr 元素）
            # 不再使用死板的 sleep(3)
            try:
                page.wait_for_selector("table tr", state="attached", timeout=5000)
            except:
                pass # 如果没表格也继续，防卡死

            # 极短的动画缓冲
            page.wait_for_timeout(500)

            with page.expect_download(timeout=30000) as download_info:
                btn.click()

            download = download_info.value
            download.save_as(file_path)
            print(f"    [成功] {safe_resp} > {safe_reason} > {ticket_no}")
        except Exception as e:
            print(f"    [失败] 工单 {ticket_no}: {e}")

    def run(self):
        self.login()
        tickets = self.fetch_all_tickets()
        if not tickets:
            print("[!] 未查询到符合条件的工单。")
            return

        print(f"[+] 启动无头浏览器，准备下载 {len(tickets)} 份 PDF ...")
        with sync_playwright() as p:
            try:
                # 【优化5】：加入 Chromium 底层性能参数，砍掉无关的渲染和沙盒开销
                browser = p.chromium.launch(
                    headless=True,
                    args=[
                        "--disable-dev-shm-usage",
                        "--disable-gpu",
                        "--no-sandbox",
                        "--disable-software-rasterizer"
                    ]
                )
            except Exception as e:
                raise Exception(f"BROWSER_INIT_FAILED|{str(e)}")

            context = browser.new_context(
                extra_http_headers={"Authorization": f"Bearer {self.token}"},
                accept_downloads=True
            )

            cookies = [{"name": k, "value": v, "url": BASE_URL} for k, v in self.session.cookies.items()]
            if cookies:
                context.add_cookies(cookies)

            # 初始化 LocalStorage
            init_page = context.new_page()
            try:
                init_page.goto(f"{BASE_URL}/admin/login/login", wait_until="domcontentloaded", timeout=30000)
                init_page.evaluate(f"""
                    window.localStorage.setItem('Token', '{self.token}');
                    window.localStorage.setItem('token', '{self.token}');
                    window.localStorage.setItem('Authorization', 'Bearer {self.token}');
                """)
            except Exception as e:
                print(f"[!] 警告: 同步 LocalStorage 失败: {e}")
            init_page.close()

            # 【优化6】：全局复用唯一标签页 (Master Page)，彻底干掉 new_page() 的巨额开销
            master_page = context.new_page()

            for index, item in enumerate(tickets):
                print(f"[*] 进度 [{index+1}/{len(tickets)}] No: {item.get('TicketNo')}")
                self.save_pdf(master_page, item)

            master_page.close()
            browser.close()


# ================= 4. 执行入口 =================
def install_chromium():
    try:
        print(f"\n[*] 正在安装便携版 Chromium 浏览器组件到: {BROWSERS_PATH}")
        print("[*] 文件体积较大（约 100MB+），请耐心等待...")
        if getattr(sys, 'frozen', False):
            from playwright._impl._driver import compute_driver_executable, get_driver_env
            env = get_driver_env()
            executable = compute_driver_executable()
            subprocess.check_call([executable, 'install', 'chromium'], env=env)
        else:
            subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
        print("\n[+] 浏览器组件下载安装成功！")
    except Exception as e:
        print(f"\n[致命错误] 浏览器安装失败: {e}")
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
        if "BROWSER_INIT_FAILED" in error_msg:
            print("\n[!] 运行中断：未检测到 Chromium 浏览器内核或初始化失败。")
            choice = input("是否立即自动下载所需的浏览器组件？(Y/N): ").strip().upper()
            if choice == 'Y':
                install_chromium()
                print("\n[!] 准备就绪，请重新运行本程序执行导出任务。")
            else:
                print("\n[!] 取消安装。")
        else:
            print(f"\n[致命错误]: {error_msg}")

    input("\n程序执行结束，按回车键退出...")
