import os
import sys
import time
import hashlib
import requests
import configparser
import subprocess
from datetime import datetime
from playwright.sync_api import sync_playwright

# ================= 1. 核心环境劫持 =================
if getattr(sys, 'frozen', False):
    PROG_DIR = os.path.dirname(sys.executable)
else:
    PROG_DIR = os.path.dirname(os.path.abspath(__file__))

BROWSERS_PATH = os.path.join(PROG_DIR, "pw-browsers")
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = BROWSERS_PATH

# ================= 2. 配置加载与参数解耦 =================
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
    TARGET_CATEGORIES = conf.get("Query", "categories", fallback="全部").strip()

    # 【核心新增】读取 5 个工单状态开关
    STATUS_FLAGS = {
        "0": conf.get("Query", "status_0", fallback="1").strip(),
        "10": conf.get("Query", "status_10", fallback="1").strip(),
        "20": conf.get("Query", "status_20", fallback="1").strip(),
        "30": conf.get("Query", "status_30", fallback="1").strip(),
        "40": conf.get("Query", "status_40", fallback="1").strip(),
    }

    # 一次性加载 9 个布尔值查询开关
    QUERY_FLAGS = {
        "MyTicket": conf.get("Query", "my_ticket", fallback="0").strip(),
        "IsAbnormal": conf.get("Query", "is_abnormal", fallback="0").strip(),
        "IsOvertime": conf.get("Query", "is_overtime", fallback="0").strip(),
        "IsUrgent": conf.get("Query", "is_urgent", fallback="0").strip(),
        "NeedClaim": conf.get("Query", "need_claim", fallback="0").strip(),
        "IsReject": conf.get("Query", "is_reject", fallback="0").strip(),
        "IsBack": conf.get("Query", "is_back", fallback="0").strip(),
        "HasRecvice": conf.get("Query", "has_recvice", fallback="0").strip(),
        "HasSpell": conf.get("Query", "has_spell", fallback="0").strip(),
    }
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

        self.category_name_to_id = {}
        self.category_id_to_name = {}

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

    def fetch_category_map(self):
        print("[*] 正在同步系统工单类型字典...")
        dict_url = f"{BASE_URL}/workflow/setting_flow/List"

        try:
            resp = self.session.post(dict_url, json={}, headers=self.headers, timeout=15).json()
            items = resp.get("TModel", [])
            for item in items:
                cid = item.get("CategoryId")
                title = item.get("Title")
                if cid is not None and title:
                    self.category_name_to_id[title] = cid
                    self.category_id_to_name[cid] = title
            print(f"[+] 字典同步成功，共获取到 {len(self.category_name_to_id)} 种工单类型")
        except Exception as e:
            print(f"[!] 警告: 字典同步失败，启用本地静态字典。")
            fallback_map = {
                8: "赠品工单", 9: "发票工单", 10: "维修工单", 11: "退货退款工单", 12: "以货换货工单",
                13: "换货补发工单", 14: "供应商补件工单", 15: "部分退款工单", 16: "自制配件工单",
                17: "自制补件工单", 18: "投诉工单", 19: "催发货工单", 21: "差评工单", 23: "转保工单",
                25: "补发工单", 26: "系统运维工单", 27: "物流提货测试工单", 46: "退货异常反馈工单",
                47: "分销商售后", 48: "投诉催发", 49: "下沉售后工单", 50: "工商投诉工单", 51: "逾期判责工单",
                52: "物流对接工单", 53: "物流异常反馈工单", 54: "国补异常登记工单", 55: "跨工厂发货登记工单",
                56: "占货补发工单", 57: "物流散单发货工单", 58: "赠品异常稽查工单", 59: "格调补件工单", 60: "店铺罚款工单"
            }
            for cid, title in fallback_map.items():
                self.category_id_to_name[cid] = title
                self.category_name_to_id[title] = cid

    def fetch_all_tickets(self):
        url = f"{BASE_URL}/workflow/order_history/List"
        all_items = []
        page_index = 1
        page_size = 50
        total_pages = 1

        # 1. 动态生成 CategoryId 的 Filter 节点
        target_ids = []
        if TARGET_CATEGORIES == "全部":
            target_ids = [str(cid) for cid in self.category_id_to_name.keys()]
        else:
            target_names = [n.strip() for n in TARGET_CATEGORIES.split(",")]
            for name in target_names:
                if name in self.category_name_to_id:
                    target_ids.append(str(self.category_name_to_id[name]))

        base_filters = []
        if target_ids:
            ids_str = f"({','.join(target_ids)})"
            base_filters.append({"Field": "CategoryId", "Value": ids_str, "Operator": "in", "Connector": "and"})

        # 2. 动态组装 5个工单状态节点
        active_statuses = [k for k, v in STATUS_FLAGS.items() if v == "1"]
        if active_statuses:
            status_str = f"({','.join(active_statuses)})"
            base_filters.append({"Field": "Status", "Value": status_str, "Operator": "in", "Connector": "and"})
            print(f"[*] 当前限定的工单状态: {status_str}")
        else:
            print("[!] 警告: 未勾选任何工单状态开关，可能导致查询不到数据！")

        # 3. 追加固定的日期 Filter 节点
        base_filters.extend([
            {"Field": "CreateTimeStart", "Value": START_TIME, "Operator": ">=", "Connector": "and"},
            {"Field": "CreateTimeEnd", "Value": END_TIME, "Operator": "<=", "Connector": "and"}
        ])

        # 4. 遍历 QUERY_FLAGS 字典，动态判定并拼装 9 个条件开关
        active_flags = []
        for field, value in QUERY_FLAGS.items():
            if value == "1":
                base_filters.append({"Field": field, "Value": 1, "Operator": "=", "Connector": "and"})
                active_flags.append(field)

        if active_flags:
            print(f"[*] 当前生效的高级过滤项: {', '.join(active_flags)}")

        while page_index <= total_pages:
            payload = {
                "PageIndex": page_index,
                "PageSize": page_size,
                "Sort": [{"Field": "TicketId", "Value": "DESC"}],
                "Filter": base_filters
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

        category_title = self.category_id_to_name.get(category_id, f"未知类型({category_id})")
        safe_type = category_title.replace('/', '_').replace('\\', '_')
        safe_resp = resp_name.replace('/', '_').replace('\\', '_')
        safe_reason = reason_one.replace('/', '_').replace('\\', '_')

        folder_path = os.path.join(PROG_DIR, "工单导出结果", safe_type, safe_resp, safe_reason)
        os.makedirs(folder_path, exist_ok=True)
        file_path = os.path.join(folder_path, f"{ticket_no}.pdf")

        print_url = f"{BASE_URL}/workflow/order_print?categoryId={category_id}&ticketId={ticket_id}&nos={ticket_no}&_t={int(time.time()*1000)}"

        try:
            page.goto(print_url, wait_until="domcontentloaded", timeout=30000)

            btn = page.locator("button").filter(has_text="导出PDF").first
            btn.wait_for(state="visible", timeout=15000)

            try:
                page.wait_for_selector("table tr", state="attached", timeout=5000)
            except:
                pass

            page.wait_for_timeout(500)

            with page.expect_download(timeout=30000) as download_info:
                btn.click()

            download = download_info.value
            download.save_as(file_path)
            print(f"    [成功] {safe_type} > {safe_resp} > {safe_reason} > {ticket_no}")
        except Exception as e:
            print(f"    [失败] 工单 {ticket_no}: {e}")

    def run(self):
        self.login()
        self.fetch_category_map()

        tickets = self.fetch_all_tickets()
        if not tickets:
            print("[!] 未查询到符合条件的工单。")
            return

        print(f"[+] 启动无头浏览器，准备下载 {len(tickets)} 份 PDF ...")
        with sync_playwright() as p:
            try:
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
