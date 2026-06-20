import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import Text
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog
import threading
import datetime
import os
import requests

# ======================== 项目信息 ========================

APP_NAME = "MoonOrange Bulletin PDF Downloader"
APP_VERSION = "1.1.0"
GITHUB_URL = "https://github.com/Orangewang124/Bulletin-PDF-Download"
COPYRIGHT_TEXT = f"{APP_NAME} v{APP_VERSION}  |  Free & Open Source  |  GitHub: {GITHUB_URL}"
COPYRIGHT_SHORT = f"Free & Open Source  |  GitHub: {GITHUB_URL}"

# ======================== 白底黑字配色方案 ========================

CLR_BG_DARK = "#ffffff"
CLR_BG_MID = "#f0f0f0"
CLR_BG_LIGHT = "#e0e0e0"
CLR_BG_HOVER = "#d0d0d0"
CLR_FG_WHITE = "#1a1a1a"
CLR_FG_LIGHT = "#333333"
CLR_FG_DIM = "#888888"
CLR_ACCENT = "#c0c0c0"
CLR_BORDER = "#cccccc"
CLR_SUCCESS = "#666666"
CLR_FAIL = "#333333"
CLR_SKIP = "#aaaaaa"

# ======================== 核心下载逻辑 (原 tryGet.py) ========================

BULLETIN_URL = "https://vip.stock.finance.sina.com.cn/corp/view/vCB_AllBulletin.php"
PDF_BASE_URL = "https://file.finance.sina.com.cn/211.154.219.97:9494/MRGG"


def get_market_code(stock_id):
    """根据股票代码判断市场代码。

    Args:
        stock_id: 6位股票代码字符串

    Returns:
        str: 市场代码，如 "CNSESH_STOCK"（沪市）或 "CNSESZ_STOCK"（深市）
    """
    if stock_id.startswith("6"):
        return "CNSESH_STOCK"
    elif stock_id.startswith("0") or stock_id.startswith("3"):
        return "CNSESZ_STOCK"
    else:
        return "CNSESH_STOCK"


def validate_stock_id(stock_id):
    """验证股票代码是否有效。

    通过新浪财经接口查询股票信息，验证代码是否存在。

    Args:
        stock_id: 6位股票代码字符串

    Returns:
        tuple: (valid: bool, stock_name: str or None, message: str)
    """
    if not stock_id or not stock_id.isdigit() or len(stock_id) != 6:
        return False, None, "股票代码应为6位数字"

    market = get_market_code(stock_id)
    # 使用新浪实时行情接口验证
    prefix = "sh" if stock_id.startswith("6") else "sz"
    url = f"https://hq.sinajs.cn/list={prefix}{stock_id}"
    try:
        headers = {"Referer": "https://finance.sina.com.cn"}
        response = requests.get(url, timeout=10, headers=headers)
        response.encoding = "gbk"
        content = response.text
        # 有效股票返回格式: var hq_str_sh600388="..."
        # 无效股票返回: var hq_str_sh999999="";
        if '=""' in content or content.strip().endswith('=""'):
            return False, None, f"股票代码 {stock_id} 不存在"
        # 提取股票名称（第一个字段）
        quote_start = content.find('"')
        if quote_start == -1:
            return False, None, f"无法解析股票代码 {stock_id} 的信息"
        quote_end = content.find('"', quote_start + 1)
        fields = content[quote_start + 1:quote_end].split(",")
        if len(fields) < 1 or not fields[0]:
            return False, None, f"股票代码 {stock_id} 不存在"
        stock_name = fields[0]
        return True, stock_name, f"验证通过: {stock_name} ({prefix.upper()}{stock_id})"
    except requests.RequestException as e:
        return False, None, f"网络请求失败: {e}"


def fetch_bulletin_list(stock_id, max_pages=10):
    """从新浪财经获取所有公告信息列表。

    Args:
        stock_id: 6位股票代码
        max_pages: 最大爬取页数

    Returns:
        list[dict]: 每个元素包含 id, date, year, year_month, name, url 字段
    """
    market = get_market_code(stock_id)
    detail_prefix = f"vCB_AllBulletinDetail.php?stockid={stock_id}&id="
    pdf_base = f"{PDF_BASE_URL}/{market}"

    bulletin_list = []
    for page in range(1, max_pages + 1):
        url = f"{BULLETIN_URL}?stockid={stock_id}&Page={page}"
        try:
            response = requests.get(url, timeout=15)
            response.encoding = "gbk"
        except requests.RequestException:
            break

        work_text = response.text
        index = work_text.find(detail_prefix)
        if index == -1:
            break

        while index != -1:
            before_text = work_text[:index]
            after_text = work_text[index:]

            id_start = after_text.find("&id=")
            id_end = after_text.find("'>")
            str_id = after_text[id_start + 4:id_end]

            nbsp_index = before_text.rfind("&nbsp")
            str_date = before_text[nbsp_index - 10:nbsp_index]
            str_year = str_date[:4]
            str_year_month = str_date
            if str_date[5] == "0":
                str_year_month = str_year_month[:5] + str_year_month[6:7]
            else:
                str_year_month = str_year_month[:7]

            name_start = after_text.find(">")
            name_end = after_text.find("<")
            name_temp = after_text[name_start + 1:name_end]

            url_pdf = (
                f"{pdf_base}/{str_year}/{str_year_month}/{str_date}/{str_id}.PDF"
            )

            bulletin_list.append(
                {
                    "id": str_id,
                    "date": str_date,
                    "year": str_year,
                    "year_month": str_year_month,
                    "name": name_temp,
                    "url": url_pdf,
                }
            )

            work_text = after_text[5:]
            index = work_text.find(detail_prefix)

    return bulletin_list


def filter_by_date(bulletin_list, start_date, end_date):
    """按日期范围过滤公告列表。

    Args:
        bulletin_list: fetch_bulletin_list 返回的列表
        start_date: 起始日期字符串，格式 "YYYY-MM-DD"
        end_date: 终止日期字符串，格式 "YYYY-MM-DD"

    Returns:
        list[dict]: 过滤后的公告列表
    """
    start_dt = datetime.datetime.strptime(start_date, "%Y-%m-%d")
    end_dt = datetime.datetime.strptime(end_date, "%Y-%m-%d")

    filtered = []
    for item in bulletin_list:
        try:
            item_dt = datetime.datetime.strptime(item["date"], "%Y-%m-%d")
        except ValueError:
            continue
        if start_dt <= item_dt <= end_dt:
            filtered.append(item)
    return filtered


def generate_filename(item):
    """根据公告信息生成文件名。"""
    return f"{item['date']}-{item['name']}-{item['id']}.pdf"


def download_pdf(item, save_path, progress_callback=None):
    """下载单个 PDF 文件。

    Args:
        item: 公告信息字典
        save_path: 保存目录路径
        progress_callback: 进度回调函数，接收 (item, status, message) 参数
            status: "success" | "fail" | "skip"

    Returns:
        tuple: (success: bool, message: str)
    """
    os.makedirs(save_path, exist_ok=True)
    file_name = generate_filename(item)
    save_name = os.path.join(save_path, file_name)

    if os.path.exists(save_name):
        msg = f"{file_name} 已存在，跳过"
        if progress_callback:
            progress_callback(item, "skip", msg)
        return True, msg

    try:
        response = requests.get(item["url"], timeout=30)
        if response.status_code == 200:
            with open(save_name, "wb") as f:
                f.write(response.content)
            msg = f"{file_name} 下载成功"
            if progress_callback:
                progress_callback(item, "success", msg)
            return True, msg
        else:
            msg = f"{file_name} 下载失败，状态码: {response.status_code}"
            if progress_callback:
                progress_callback(item, "fail", msg)
            return False, msg
    except requests.RequestException as e:
        msg = f"{file_name} 下载异常: {e}"
        if progress_callback:
            progress_callback(item, "fail", msg)
        return False, msg


def download_batch(bulletin_list, save_path, progress_callback=None):
    """批量下载 PDF 文件。

    Args:
        bulletin_list: 公告信息列表
        save_path: 保存目录路径
        progress_callback: 进度回调函数，接收 (index, total, item, status, message)

    Returns:
        dict: 包含 success_count, fail_count, skip_count, results 信息
    """
    total = len(bulletin_list)
    success_count = 0
    fail_count = 0
    skip_count = 0
    results = []

    for idx, item in enumerate(bulletin_list):
        def single_callback(it, status, message):
            nonlocal success_count, fail_count, skip_count
            if status == "success":
                success_count += 1
            elif status == "fail":
                fail_count += 1
            elif status == "skip":
                skip_count += 1
            if progress_callback:
                progress_callback(idx, total, it, status, message)

        ok, msg = download_pdf(item, save_path, single_callback)
        results.append({"item": item, "success": ok, "message": msg})

    return {
        "success_count": success_count,
        "fail_count": fail_count,
        "skip_count": skip_count,
        "results": results,
    }


# ======================== GUI 界面 ========================

class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("960x860")
        self.resizable(True, True)
        self.minsize(800, 700)
        self.configure(bg=CLR_BG_DARK)

        self._apply_monochrome_style()

        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_penguin.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.stock_id = ttk.StringVar(value="600388")
        self.stock_name = ttk.StringVar(value="")
        self.stock_valid = False
        self.save_path = ttk.StringVar(value="D:/BulletinPDFDownload/")
        self.start_date = ttk.StringVar(value=f"{datetime.date.today().year}-01-01")
        self.end_date = ttk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.bulletin_list = []
        self.filtered_list = []
        self.is_downloading = False

        self._build_ui()
        self.after(50, self._force_log_text_style)

    def _force_log_text_style(self):
        try:
            self.log_text.configure(
                background="#ffffff", foreground="#1a1a1a",
                insertbackground="#1a1a1a", selectbackground="#d0d0d0",
                selectforeground="#1a1a1a"
            )
        except Exception:
            pass

    def _apply_monochrome_style(self):
        s = ttk.Style()
        s.configure(".", background=CLR_BG_DARK, foreground=CLR_FG_WHITE,
                     fieldbackground=CLR_BG_MID, bordercolor=CLR_BORDER,
                     darkcolor=CLR_BG_DARK, lightcolor=CLR_BG_DARK,
                     troughcolor=CLR_BG_MID, focuscolor=CLR_ACCENT,
                     selectbackground=CLR_BG_HOVER, selectforeground=CLR_FG_WHITE,
                     insertcolor=CLR_FG_WHITE, font=("Segoe UI", 10))

        s.configure("TFrame", background=CLR_BG_DARK)
        s.configure("TLabel", background=CLR_BG_DARK, foreground=CLR_FG_LIGHT, font=("Segoe UI", 10))
        s.configure("Bold.TLabel", font=("Segoe UI", 10, "bold"), foreground=CLR_FG_WHITE)
        s.configure("Header.TLabel", font=("Segoe UI", 18, "bold"), foreground=CLR_FG_WHITE, background=CLR_BG_DARK)
        s.configure("SubHeader.TLabel", font=("Segoe UI", 11), foreground=CLR_FG_DIM, background=CLR_BG_DARK)
        s.configure("Desc.TLabel", font=("Segoe UI", 10), foreground=CLR_FG_DIM, background=CLR_BG_DARK)
        s.configure("Status.TLabel", font=("Segoe UI", 9), foreground=CLR_FG_DIM, background=CLR_BG_MID, relief="sunken")
        s.configure("Footer.TLabel", font=("Segoe UI", 8), foreground=CLR_FG_DIM, background=CLR_BG_DARK)
        s.configure("Progress.TLabel", font=("Segoe UI", 10, "bold"), foreground=CLR_FG_WHITE, background=CLR_BG_DARK)
        s.configure("Count.TLabel", font=("Segoe UI", 10), foreground=CLR_FG_DIM, background=CLR_BG_DARK)

        s.configure("V.TLabel", font=("Segoe UI", 10), foreground=CLR_FG_DIM, background=CLR_BG_DARK)
        s.configure("VOK.TLabel", font=("Segoe UI", 10), foreground=CLR_FG_WHITE, background=CLR_BG_DARK)
        s.configure("VWarn.TLabel", font=("Segoe UI", 10), foreground=CLR_FG_DIM, background=CLR_BG_DARK)
        s.configure("VFail.TLabel", font=("Segoe UI", 10), foreground=CLR_FG_LIGHT, background=CLR_BG_DARK)

        s.configure("HeaderBar.TFrame", background=CLR_BG_DARK)
        s.configure("Card.TLabelframe", background=CLR_BG_DARK, foreground=CLR_FG_WHITE,
                     bordercolor=CLR_BORDER, relief="groove")
        s.configure("Card.TLabelframe.Label", background=CLR_BG_DARK, foreground=CLR_FG_WHITE,
                     font=("Segoe UI", 10, "bold"))

        s.configure("Mono.TButton", background=CLR_BG_LIGHT, foreground=CLR_FG_WHITE,
                     bordercolor=CLR_BORDER, focuscolor=CLR_BG_HOVER, font=("Segoe UI", 10),
                     relief="raised", padding=(10, 5))
        s.map("Mono.TButton",
              background=[("active", CLR_BG_HOVER), ("disabled", CLR_BG_MID)],
              foreground=[("disabled", CLR_FG_DIM), ("active", CLR_FG_WHITE)])

        s.configure("Accent.TButton", background=CLR_FG_WHITE, foreground=CLR_BG_DARK,
                     bordercolor=CLR_FG_WHITE, font=("Segoe UI", 10, "bold"),
                     relief="raised", padding=(10, 5))
        s.map("Accent.TButton",
              background=[("active", CLR_ACCENT), ("disabled", CLR_BG_MID)],
              foreground=[("disabled", CLR_FG_DIM), ("active", CLR_BG_DARK)])

        s.configure("Mono.TEntry", fieldbackground=CLR_BG_MID, foreground=CLR_FG_WHITE,
                     bordercolor=CLR_BORDER, insertcolor=CLR_FG_WHITE,
                     selectbackground=CLR_BG_HOVER, selectforeground=CLR_FG_WHITE)

        s.configure("Mono.Treeview", background=CLR_BG_MID, foreground=CLR_FG_LIGHT,
                     fieldbackground=CLR_BG_MID, bordercolor=CLR_BORDER,
                     rowheight=28, font=("Segoe UI", 9))
        s.configure("Mono.Treeview.Heading", background=CLR_BG_LIGHT, foreground=CLR_FG_WHITE,
                     bordercolor=CLR_BORDER, font=("Segoe UI", 9, "bold"), relief="flat")
        s.map("Mono.Treeview",
              background=[("selected", CLR_BG_HOVER)],
              foreground=[("selected", CLR_FG_WHITE)])

        s.configure("Mono.Vertical.TScrollbar", background=CLR_BG_LIGHT,
                     troughcolor=CLR_BG_MID, bordercolor=CLR_BORDER,
                     arrowcolor=CLR_FG_DIM)
        s.map("Mono.Vertical.TScrollbar",
              background=[("active", CLR_BG_HOVER)])

        s.configure("Mono.Horizontal.TProgressbar", background=CLR_FG_WHITE,
                     troughcolor=CLR_BG_MID, bordercolor=CLR_BORDER, thickness=6)

        s.configure("Mono.TSeparator", background=CLR_BORDER)

    def _build_ui(self):
        self._build_header()
        self._build_config_section()
        self._build_action_bar()
        self._build_list_section()
        self._build_progress_section()
        self._build_footer()

    def _build_header(self):
        header = ttk.Frame(self, style="HeaderBar.TFrame")
        header.pack(fill=X, padx=0, pady=0)

        inner = ttk.Frame(header, padding=(20, 14), style="HeaderBar.TFrame")
        inner.pack(fill=X)

        ttk.Label(
            inner, text="Bulletin PDF Downloader",
            font=("Segoe UI", 18, "bold"), style="Header.TLabel"
        ).pack(side=LEFT)

        ttk.Label(
            inner, text=f"v{APP_VERSION}",
            font=("Segoe UI", 11), style="SubHeader.TLabel"
        ).pack(side=LEFT, padx=(12, 0), pady=(6, 0))

        ttk.Label(
            inner, text="新浪财经公告PDF批量下载工具",
            font=("Segoe UI", 10), style="Desc.TLabel"
        ).pack(side=RIGHT)

        ttk.Separator(self, orient=HORIZONTAL).pack(fill=X)

    def _build_config_section(self):
        config_card = ttk.Labelframe(self, text="  参数配置  ", padding=15, style="Card.TLabelframe")
        config_card.pack(fill=X, padx=15, pady=(10, 5))

        row1 = ttk.Frame(config_card, style="TFrame")
        row1.pack(fill=X, pady=(0, 10))

        ttk.Label(row1, text="股票代码", style="Bold.TLabel").pack(side=LEFT, padx=(0, 6))
        self.stock_entry = ttk.Entry(row1, textvariable=self.stock_id, width=10, style="Mono.TEntry")
        self.stock_entry.pack(side=LEFT, padx=(0, 8))
        self.stock_entry.bind("<Return>", lambda e: self._validate_stock())

        self.btn_validate = ttk.Button(
            row1, text="验证", command=self._validate_stock,
            style="Mono.TButton", width=8
        )
        self.btn_validate.pack(side=LEFT, padx=(0, 12))

        self.stock_info_label = ttk.Label(
            row1, text="未验证", style="V.TLabel"
        )
        self.stock_info_label.pack(side=LEFT, padx=4)

        ttk.Separator(row1, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=12)

        ttk.Label(row1, text="日期范围", style="Bold.TLabel").pack(side=LEFT, padx=(0, 6))
        self.start_date_entry = ttk.Entry(row1, textvariable=self.start_date, width=12, style="Mono.TEntry")
        self.start_date_entry.pack(side=LEFT, padx=(0, 4))
        ttk.Label(row1, text="至").pack(side=LEFT, padx=(0, 4))
        self.end_date_entry = ttk.Entry(row1, textvariable=self.end_date, width=12, style="Mono.TEntry")
        self.end_date_entry.pack(side=LEFT, padx=(0, 4))

        row2 = ttk.Frame(config_card, style="TFrame")
        row2.pack(fill=X, pady=(0, 0))

        ttk.Label(row2, text="保存路径", style="Bold.TLabel").pack(side=LEFT, padx=(0, 6))
        ttk.Entry(row2, textvariable=self.save_path, style="Mono.TEntry").pack(
            side=LEFT, fill=X, expand=True, padx=(0, 8)
        )
        ttk.Button(row2, text="浏览", command=self._browse_path, style="Mono.TButton", width=8).pack(
            side=LEFT
        )

    def _build_action_bar(self):
        bar = ttk.Frame(self, padding=(15, 8), style="TFrame")
        bar.pack(fill=X)

        self.btn_preview = ttk.Button(
            bar, text="预览列表", command=self._preview,
            style="Accent.TButton", width=14
        )
        self.btn_preview.pack(side=LEFT, padx=(0, 10))

        self.btn_download = ttk.Button(
            bar, text="确认下载", command=self._start_download,
            style="Mono.TButton", width=14, state=DISABLED
        )
        self.btn_download.pack(side=LEFT, padx=(0, 10))

        self.progress_label = ttk.Label(bar, text="", style="Progress.TLabel")
        self.progress_label.pack(side=LEFT, padx=8)

        self.count_label = ttk.Label(bar, text="", style="Count.TLabel")
        self.count_label.pack(side=RIGHT, padx=4)

    def _build_list_section(self):
        list_card = ttk.Labelframe(self, text="  公告列表  ", padding=5, style="Card.TLabelframe")
        list_card.pack(fill=BOTH, expand=True, padx=15, pady=(0, 5))

        columns = ("date", "name", "filename")
        self.tree = ttk.Treeview(
            list_card, columns=columns, show="headings",
            selectmode="extended", style="Mono.Treeview", height=12
        )
        self.tree.heading("date", text="日期", anchor=W)
        self.tree.heading("name", text="公告名称", anchor=W)
        self.tree.heading("filename", text="文件名", anchor=W)
        self.tree.column("date", width=110, anchor=CENTER, minwidth=90)
        self.tree.column("name", width=300, anchor=W, minwidth=150)
        self.tree.column("filename", width=420, anchor=W, minwidth=200)

        scrollbar = ttk.Scrollbar(list_card, orient=VERTICAL, command=self.tree.yview, style="Mono.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

    def _build_progress_section(self):
        progress_card = ttk.Labelframe(self, text="  下载进度  ", padding=8, style="Card.TLabelframe")
        progress_card.pack(fill=X, padx=15, pady=(0, 5))

        self.progress_bar = ttk.Progressbar(
            progress_card, style="Mono.Horizontal.TProgressbar", mode=DETERMINATE
        )
        self.progress_bar.pack(fill=X, pady=(0, 6))

        log_frame = ttk.Frame(progress_card, style="TFrame")
        log_frame.pack(fill=BOTH, expand=True)

        self.log_text = Text(
            log_frame, height=6,
            background="#ffffff", foreground="#1a1a1a",
            insertbackground="#1a1a1a", selectbackground="#d0d0d0",
            selectforeground="#1a1a1a", borderwidth=1,
            relief="sunken", font=("Consolas", 9),
            wrap="word"
        )
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)

        log_scrollbar = ttk.Scrollbar(log_frame, orient=VERTICAL, command=self.log_text.yview, style="Mono.Vertical.TScrollbar")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        log_scrollbar.pack(side=RIGHT, fill=Y)

        self.log_text.tag_configure("ok", foreground="#1a1a1a", font=("Consolas", 9))
        self.log_text.tag_configure("fail", foreground="#cc0000", font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("skip", foreground="#888888", font=("Consolas", 9))

    def _build_footer(self):
        footer = ttk.Frame(self, padding=(15, 4), style="TFrame")
        footer.pack(fill=X, side=BOTTOM)

        self.status_var = ttk.StringVar(value=f"就绪  |  {COPYRIGHT_SHORT}")
        status_label = ttk.Label(
            footer, textvariable=self.status_var,
            style="Status.TLabel", anchor=W, padding=(8, 3), cursor="hand2"
        )
        status_label.pack(fill=X)
        status_label.bind("<Button-1>", lambda e: self._open_github())

    def _open_github(self):
        import webbrowser
        webbrowser.open(GITHUB_URL)

    def _browse_path(self):
        path = filedialog.askdirectory(title="选择保存路径")
        if path:
            self.save_path.set(path.replace("/", "\\") + "\\")

    def _log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    def _validate_dates(self):
        try:
            start = datetime.datetime.strptime(self.start_date.get(), "%Y-%m-%d")
            end = datetime.datetime.strptime(self.end_date.get(), "%Y-%m-%d")
            if start > end:
                Messagebox.show_error("起始日期不能晚于终止日期", title="日期错误", parent=self)
                return None, None
            return start, end
        except ValueError:
            Messagebox.show_error("日期格式不正确，请使用 YYYY-MM-DD 格式", title="日期错误", parent=self)
            return None, None

    def _validate_stock(self):
        sid = self.stock_id.get().strip()
        if not sid:
            self.stock_info_label.configure(text="请输入股票代码", style="VFail.TLabel")
            self.stock_valid = False
            return

        self.btn_validate.configure(state=DISABLED)
        self.stock_info_label.configure(text="验证中...", style="VWarn.TLabel")
        self.update_idletasks()

        def worker():
            valid, name, msg = validate_stock_id(sid)
            self.after(0, self._show_validate_result, valid, name, msg)

        threading.Thread(target=worker, daemon=True).start()

    def _show_validate_result(self, valid, name, msg):
        self.btn_validate.configure(state=NORMAL)
        self.stock_valid = valid
        if valid:
            self.stock_name.set(name)
            self.stock_info_label.configure(text=f"✔ {msg}", style="VOK.TLabel")
        else:
            self.stock_name.set("")
            self.stock_info_label.configure(text=f"✖ {msg}", style="VFail.TLabel")

    def _preview(self):
        sid = self.stock_id.get().strip()
        if not sid:
            Messagebox.show_warning("请先输入股票代码", title="提示", parent=self)
            return

        start, end = self._validate_dates()
        if start is None:
            return

        self.btn_preview.configure(state=DISABLED)
        self.btn_validate.configure(state=DISABLED)
        self.status_var.set(f"正在验证股票代码并获取公告列表，请稍候...  |  {COPYRIGHT_SHORT}")
        self.stock_info_label.configure(text="验证中...", style="VWarn.TLabel")
        self.count_label.configure(text="")
        self.update_idletasks()

        def worker():
            valid, name, vmsg = validate_stock_id(sid)
            if not valid:
                self.after(0, self._show_validate_result, valid, name, vmsg)
                self.after(0, lambda: self.status_var.set(f"股票代码验证失败  |  {COPYRIGHT_SHORT}"))
                self.after(0, lambda: self.btn_preview.configure(state=NORMAL))
                self.after(0, lambda: self.btn_validate.configure(state=NORMAL))
                return

            self.after(0, self._show_validate_result, valid, name, vmsg)

            try:
                self.bulletin_list = fetch_bulletin_list(sid, max_pages=20)
                self.filtered_list = filter_by_date(
                    self.bulletin_list,
                    self.start_date.get(),
                    self.end_date.get(),
                )
                self.after(0, self._show_preview_result)
            except Exception as e:
                self.after(0, lambda: self._show_preview_error(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _show_preview_result(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        for item in self.filtered_list:
            filename = generate_filename(item)
            self.tree.insert("", END, values=(item["date"], item["name"], filename))

        count = len(self.filtered_list)
        self.status_var.set(f"共找到 {count} 个公告  |  {COPYRIGHT_SHORT}")
        self.count_label.configure(text=f"{count} 条记录")
        self.btn_preview.configure(state=NORMAL)
        self.btn_validate.configure(state=NORMAL)
        self.btn_download.configure(state=NORMAL if count > 0 else DISABLED)

    def _show_preview_error(self, err):
        self.status_var.set(f"获取公告列表失败  |  {COPYRIGHT_SHORT}")
        self.btn_preview.configure(state=NORMAL)
        self.btn_validate.configure(state=NORMAL)
        Messagebox.show_error(f"获取公告列表失败:\n{err}", title="错误", parent=self)

    def _start_download(self):
        if self.is_downloading:
            return
        if not self.filtered_list:
            Messagebox.show_warning("请先预览文件列表", title="提示", parent=self)
            return

        save = self.save_path.get().strip()
        if not save:
            Messagebox.show_error("请设置保存路径", title="错误", parent=self)
            return

        self.is_downloading = True
        self.btn_download.configure(state=DISABLED)
        self.btn_preview.configure(state=DISABLED)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(self.filtered_list)

        self.log_text.delete("1.0", END)
        self.status_var.set(f"正在下载...  |  {COPYRIGHT_SHORT}")

        def worker():
            def on_progress(idx, total, item, status, message):
                self.after(0, self._update_download_progress, idx, total, status, message)

            result = download_batch(self.filtered_list, save, progress_callback=on_progress)
            self.after(0, self._download_finished, result)

        threading.Thread(target=worker, daemon=True).start()

    def _update_download_progress(self, idx, total, status, message):
        self.progress_bar["value"] = idx + 1
        self.progress_label.configure(text=f"{idx + 1}/{total}")

        tag = ""
        if status == "success":
            tag = "ok"
        elif status == "fail":
            tag = "fail"
        elif status == "skip":
            tag = "skip"

        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")

    def _download_finished(self, result):
        self.is_downloading = False
        self.btn_download.configure(state=NORMAL)
        self.btn_preview.configure(state=NORMAL)

        s = result["success_count"]
        f = result["fail_count"]
        k = result["skip_count"]
        total = s + f + k

        summary = f"下载完成: 共 {total} 个, 成功 {s} 个, 失败 {f} 个, 跳过 {k} 个"
        self.status_var.set(f"{summary}  |  {COPYRIGHT_SHORT}")
        self.progress_label.configure(text="完成" if f == 0 else "有失败")
        self._log(f"\n{summary}")

        if f > 0:
            Messagebox.show_warning(
                f"下载完成\n成功: {s}\n失败: {f}\n跳过: {k}",
                title="下载完成", parent=self
            )
        else:
            Messagebox.show_info(
                f"下载完成\n成功: {s}\n跳过: {k}",
                title="下载完成", parent=self
            )


if __name__ == "__main__":
    app = App()
    app.mainloop()
