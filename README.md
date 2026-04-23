-----

* [中文版 (Chinese)](https://www.google.com/search?q=%23%E4%B8%AD%E6%96%87%E7%89%88-chinese)
* [English Version](https://www.google.com/search?q=%23english-version)

-----

# 中文版 (Chinese)

# 📄 工单导出助手 (Ticket Export RPA)

本项目是一个自动化的工单抓取与原生 PDF 导出工具。它能够根据指定的时间范围、工单类型及 9 大高级业务状态，自动登录 OMS 系统，精确抓取符合条件的工单，并使用 Playwright 无头浏览器接管底层下载流，导出排版完美的原生 PDF 文件。导出的文件会自动按照 **“工单类型 \> 责任人 \> 责任原因”** 的三级目录结构进行智能归档。

**核心特性：**

* **全自动化与动态配置**：自动登录并加密密码，支持在配置文件中按名称动态组合工单类型及 9 大业务开关。
* **原生无损下载**：彻底抛弃截屏逻辑，智能破解 Vue 路由缓存，物理模拟点击接管前端文件流，确保 PDF 原生排版。
* **极致性能与便携**：抛弃单文件打包导致的 I/O 阻塞，采用文件夹级打包。内置浏览器环境隔离（离线内核），打包后无需客户机联网下载，完美绕过企业内网网络限制及杀毒软件扫描瓶颈。

-----

## 🛠️ 1. 开发环境配置 (开发者专属)

如果你需要修改代码并在本地测试，请按照以下步骤初始化环境：

### 1.1 安装依赖

确保电脑已安装 Python (推荐 3.9 - 3.12)。在终端中执行以下命令安装必要的 Python 库：

```bash
pip install requests playwright pyinstaller
```

### 1.2 初始化便携版浏览器内核

这是**极其重要**的一步。为了保证程序独立性，强制将浏览器内核安装在项目目录下的 `pw-browsers` 文件夹中：

```bash
python main.py install
```

*(注：该命令会下载约 100MB+ 的 Chromium 内核。成功后项目目录中会出现 `pw-browsers` 文件夹。)*

### 1.3 源码运行测试

确保同级目录下存在填写正确的 `config.ini` 文件，然后执行：

```bash
python main.py
```

-----

## ⚙️ 2. 配置文件说明 (`config.ini`)

程序高度依赖同级目录下的 `config.ini` 文件。它支持动态查询控制，请确保内容格式如下：

```ini
[Credentials]
# 登录账号
account = 15819893256
# 登录密码（填入明文即可，代码会自动转 MD5 加密）
password = 123456

[Query]
# 查询的起始与结束时间 (格式: YYYY-MM-DD HH:mm:ss)
start_time = 2026-01-23 00:00:00
end_time = 2026-04-23 23:59:59

# 需要导出的工单类型，多个类型用英文逗号隔开，全量抓取请直接填 "全部"
categories = 维修工单,退货退款工单,以货换货工单,换货补发工单,供应商补件工单,部分退款工单,自制补件工单,投诉工单,差评工单

# ================= 9大业务状态开关 =================
# 配置说明：1 = 勾选该条件（仅查询该状态），0 = 不限制（查全部）
# 我的工单 (1=仅查自己名下)
my_ticket = 1
# 异常工单
is_abnormal = 0
# 逾期工单
is_overtime = 0
# 加急工单
is_urgent = 0
# 需理赔工单
need_claim = 0
# 驳回工单
is_reject = 0
# 退回工单
is_back = 0
# 已收货工单
has_recvice = 0
# 拼单工单
has_spell = 0
```

-----

## 📦 3. 打包与分发流程 (生成发布版本)

为了彻底解决 Windows 环境下打包程序的运行缓慢及杀毒软件拦截问题，**必须采用文件夹模式（-D）打包**：

### 第一步：执行打包命令

在终端执行以下命令（注意使用 `--collect-all` 载入底层驱动）：

```bash
pyinstaller --onefile --name 工单导出助手 main.py
```

*编译完成后，会在 `dist` 目录下生成一个名为 `工单导出助手` 的大型文件夹。*

### 第二步：组装“绿色免安装包”

为了让同事**双击即用（不触发任何网络下载）**，你需要补全依赖文件：

1.  进入 `dist/工单导出助手` 目录。
2.  将项目根目录下的 `config.ini` 文件拷贝进去。
3.  将第一步中生成的包含了离线内核的 `pw-browsers` 文件夹也拷贝进去。

### 第三步：压缩分发

将整个组装好的 `工单导出助手` 文件夹打成 `.zip` 压缩包，发送给业务同事。

-----

## 🚀 4. 最终用户使用指南 (面向业务人员)

1.  **解压**：将 `.zip` 压缩包完整解压到电脑本地（建议解压到 D 盘或桌面，**绝对不要**在压缩包内直接双击运行）。
2.  **配置**：用记事本打开目录中的 `config.ini`，修改你需要的时间、类型及 `1` 或 `0` 的业务开关。
3.  **运行**：双击运行 `工单导出助手.exe`。
4.  **获取结果**：程序运行结束后，当前目录下会自动生成 `工单导出结果` 文件夹，按照 `工单类型 > 责任人 > 原因` 完美分类。

-----

## ⚠️ 5. 常见问题排查 (Troubleshooting)

| 报错现象 | 原因及解决办法 |
| :--- | :--- |
| **未找到配置文件** | 用户没有将 `config.ini` 和 `.exe` 放在同一个文件夹内。 |
| **未检测到浏览器内核** | 打包分发时，忘记将 `pw-browsers` 文件夹拷入 `dist/工单导出助手` 目录中。 |
| **登录接口连接失败** | 检查网络连接，若开启了 VPN 或代理导致超时，请确保目标域名的直连连通性。 |
| **查不到任何数据** | 检查 `config.ini` 中是否勾选了相互矛盾的开关（如某些单不可能同时是退回又同时是已收货）。 |
| **下载耗时极长(\>10分钟)** | 属于杀毒软件底层 I/O 拦截。请将整个 `工单导出助手` 文件夹加入 Windows Defender 或安全软件的白名单。 |

<br><br>

-----

# English Version

# 📄 Ticket Export RPA (Order History Assistant)

This project is an automated tool for fetching tickets and exporting them as native PDF files. Based on specified time ranges, ticket types, and 9 advanced business status flags, it logs into the OMS system, accurately retrieves matching tickets, and uses Playwright to intercept the underlying download stream for perfectly formatted native PDFs. Exported files are intelligently archived into a 3-level nested directory structure: **"Category \> Responsible Person \> Reason"**.

**Core Features:**

* **Fully Automated & Dynamic Configuration:** Auto-login with password encryption. Dynamically combine ticket types and 9 business switches directly from the configuration file.
* **Native Lossless Download:** Abandons the standard screenshot approach. It bypasses Vue routing caches and physically simulates button clicks to intercept frontend file streams, ensuring native PDF layouts.
* **Ultimate Performance & Portability:** Uses directory-mode packing to prevent I/O blocking typical of single-file executables. Incorporates isolated browser environments (offline binaries), allowing it to bypass enterprise network restrictions and antivirus scanning bottlenecks without requiring client-side downloads.

-----

## 🛠️ 1. Development Environment Setup (For Developers)

If you need to modify the code and test locally, initialize your environment with the following steps:

### 1.1 Install Dependencies

Ensure Python (3.9 - 3.12 recommended) is installed. Run the following command in your terminal:

```bash
pip install requests playwright pyinstaller
```

### 1.2 Initialize Portable Browser Binaries

This is a **critical step**. To ensure standalone execution, we force the browser engine to install locally into a `pw-browsers` folder:

```bash
python main.py install
```

*(Note: This downloads a 100MB+ Chromium binary. A `pw-browsers` folder will appear in your project root upon success.)*

### 1.3 Run Locally

Ensure a properly configured `config.ini` file exists in the same directory, then execute:

```bash
python main.py
```

-----

## ⚙️ 2. Configuration Guide (`config.ini`)

The program strictly relies on `config.ini` to read account credentials and query conditions. Make sure it looks like this:

```ini
[Credentials]
# Login Account
account = 15819893256
# Login Password (Enter plain text; the program encrypts it via MD5 automatically)
password = 123456

[Query]
# Time Range (Format: YYYY-MM-DD HH:mm:ss)
start_time = 2026-01-23 00:00:00
end_time = 2026-04-23 23:59:59

# Ticket categories to export (comma-separated). Use "全部" (All) for full extraction.
categories = 维修工单,退货退款工单,以货换货工单,换货补发工单,供应商补件工单,部分退款工单,自制补件工单,投诉工单,差评工单

# ================= 9 Advanced Business Flags =================
# Note: 1 = Enable condition, 0 = No restriction (Fetch all)
# My Tickets Only
my_ticket = 1
# Abnormal
is_abnormal = 0
# Overtime
is_overtime = 0
# Urgent
is_urgent = 0
# Needs Claim
need_claim = 0
# Rejected
is_reject = 0
# Returned/Backed
is_back = 0
# Received
has_recvice = 0
# Spelled/Merged
has_spell = 0
```

-----

## 📦 3. Packing and Distribution

To resolve severe execution delays and antivirus interceptions on Windows, you **MUST use Directory Mode (`-D`)** when packing:

### Step 1: Execute PyInstaller

Run the following command (ensure `--collect-all` is used to load underlying drivers):

```bash
pyinstaller -D main.py --collect-all playwright -n 工单导出助手
```

*Once compiled, a large folder named `工单导出助手` will be generated in the `dist` directory.*

### Step 2: Assemble the "Portable Package"

To make it **plug-and-play without network downloads**, you must include the dependencies:

1.  Navigate into the `dist/工单导出助手` directory.
2.  Copy the `config.ini` file from your project root into this folder.
3.  Copy the `pw-browsers` folder (generated in Step 1.2) into this folder.

### Step 3: Compress and Distribute

Zip the assembled `工单导出助手` folder and send it to the operations team.

-----

## 🚀 4. End-User Guide (For Business Operations)

1.  **Unzip:** Extract the `.zip` archive completely to a local drive (e.g., Desktop or D: Drive. **NEVER** run the `.exe` directly from within the zip preview).
2.  **Configure:** Open `config.ini` with Notepad to set your desired time range, categories, and `1`/`0` status switches.
3.  **Execute:** Double-click `工单导出助手.exe`.
4.  **Retrieve Results:** Once finished, an output folder (`工单导出结果`) will automatically appear in the same directory, categorized perfectly by `Category > Responsible > Reason`.

-----

## ⚠️ 5. Troubleshooting

| Error / Symptom | Cause & Solution |
| :--- | :--- |
| **Config file not found** | The user did not place `config.ini` in the exact same directory as the `.exe`. |
| **Browser initialization failed** | The `pw-browsers` folder was missing from the packaged directory during distribution. |
| **Login connection failed** | Check network connection. If using a VPN/Proxy, ensure direct routing to the target domain is allowed. |
| **Zero data fetched** | Check `config.ini` for mutually exclusive flags (e.g., a ticket cannot be both 'Returned' and 'Received' simultaneously). |
| **Extremely slow download (\>10m)** | This is caused by underlying Antivirus I/O interception. Add the entire application folder to the Windows Defender / Antivirus whitelist. |
