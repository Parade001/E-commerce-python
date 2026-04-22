import os, time, requests, configparser
from datetime import datetime
from playwright.sync_api import sync_playwright
import hashlib

# ================= 1. 配置加载 =================
config = configparser.ConfigParser()
config.read("config.ini", encoding="utf-8")

ACCOUNT = config.get("Credentials", "account")
PASSWORD = config.get("Credentials", "password")
START_TIME = config.get("Query", "start_time")
END_TIME = config.get("Query", "end_time")
BASE_URL = "http://sw.cesaas.com:81"

class SmartOrderRPA:
    def __init__(self):
        self.session = requests.Session()
        self.headers = {"Content-Type": "application/json"}

    def login(self):
        """登录并获取Token"""
        url = f"{BASE_URL}/admin/login/login"

        # ================= 新增：密码自动 MD5 加密逻辑 =================
        # 1. 尝试使用标准的纯净 MD5 计算
        md5_obj = hashlib.md5()
        md5_obj.update(PASSWORD.encode('utf-8'))
        encrypted_pwd = md5_obj.hexdigest()

        # 打印出来比对：看算出来的值是否恰好等于 2131f8ecf18db66a758f718dc729e00e
        print(f"[*] 已将明文密码转化为密文: {encrypted_pwd}")
        # ==============================================================

        payload = {
            "ati": "3643730255656",
            "IsOms": 1,
            "Account": ACCOUNT,
            "Password": encrypted_pwd,  # <--- 使用 Python 刚刚自动计算好的密文
            "Date": datetime.now().strftime("%Y-%m-%d")
        }

        print(f"[*] 正在尝试登录账号: {ACCOUNT}...")
        resp = self.session.post(url, json=payload).json()
        if resp.get("IsSuccess"):
            self.token = resp["TModel"]["Token"]
            self.headers["Authorization"] = f"Bearer {self.token}"
            print("[+] 登录成功，Token获取完成")
        else:
            raise Exception(f"登录失败: {resp.get('Message')}")

    def get_all_data(self):
        """核心逻辑：基于服务器真实 PageCount 进行安全遍历，杜绝死循环"""
        url = f"{BASE_URL}/workflow/order_history/List"
        all_items = []
        page_index = 1
        page_size = 50
        total_pages = 1  # 初始预设1页，第一次请求后会动态更新

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
                    # [关键修正] 根据您的真实抓包，去掉选中我的工单，对应的值是 1
                    {"Field": "MyTicket", "Value": 1, "Operator": "=", "Connector": "and"}
                ]
            }

            resp = self.session.post(url, json=payload, headers=self.headers).json()

            # 【进度透明化】在第一页时，读取服务器返回的总页数和总条数
            if page_index == 1:
                total_pages = resp.get("PageCount", 1)
                record_count = resp.get("RecordCount", 0)
                print(f"[!] 接口响应: 时间段内共发现 {record_count} 条数据，服务器下发需抓取 {total_pages} 页")
                if record_count == 0:
                    break

            # 自动提取数据列表
            curr_list = []
            t_model = resp.get("TModel", [])
            if isinstance(t_model, list):
                curr_list = t_model
            elif isinstance(t_model, dict):
                curr_list = t_model.get("Items", [])

            all_items.extend(curr_list)
            print(f"[*] 正在抓取第 {page_index}/{total_pages} 页... 本页成功获取 {len(curr_list)} 条")

            page_index += 1

        return all_items

    def export_pdf(self, page, item):
        """本地分类逻辑：根据数据中的字段动态决定存放路径，增加强容错机制"""
        # 1. 动态提取分类名称与核心ID（全部改用 .get() 安全提取，防止 KeyError）
        resp_name = str(item.get("ResponsibleName") or "未归类").strip()
        reason_one = str(item.get("ReasonOne") or "其他原因").strip()
        ticket_no = str(item.get("TicketNo") or f"未知单号_{int(time.time())}").strip()

        category_id = item.get("CategoryId", "")
        ticket_id = item.get("TicketId", "")

        # 边界防御：如果没有 TicketId，说明这条数据是服务器产生的完全无效的残缺数据，直接跳过
        if not ticket_id:
            print(f"    [!] 跳过一条无效脏数据（缺少TicketId）。")
            return

        # 2. 建立多级文件夹 (过滤掉可能导致系统建文件夹报错的非法字符)
        safe_resp = resp_name.replace('/', '_').replace('\\', '_')
        safe_reason = reason_one.replace('/', '_').replace('\\', '_')
        save_dir = os.path.join("工单导出", safe_resp, safe_reason)
        os.makedirs(save_dir, exist_ok=True)

        # 3. 打印导出
        file_path = os.path.join(save_dir, f"{ticket_no}.pdf")
        print_url = f"{BASE_URL}/workflow/order_print?categoryId={category_id}&ticketId={ticket_id}&nos={ticket_no}"

        try:
            # 访问打印页面，等待网络空闲
            page.goto(print_url, wait_until="networkidle")
            # 存为PDF
            page.pdf(path=file_path, format="A4", print_background=True)
            print(f"    [OK] 导出到: {safe_resp} -> {safe_reason} (单号: {ticket_no})")
        except Exception as e:
            print(f"    [FAIL] 工单 {ticket_no} 导出失败: {e}")

    def start(self):
        self.login()
        tickets = self.get_all_data()
        if not tickets:
            print("[!] 未查询到任何数据，请检查时间范围。")
            return

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            page = context.new_page()

            for index, item in enumerate(tickets):
                print(f"[*] 处理进度 [{index+1}/{len(tickets)}] No: {item.get('TicketNo')}")
                self.export_pdf(page, item)

            browser.close()
            print("\n>>> 全部任务已完成。")

if __name__ == "__main__":
    SmartOrderRPA().start()
