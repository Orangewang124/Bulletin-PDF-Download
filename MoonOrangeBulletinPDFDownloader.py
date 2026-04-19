import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import datetime
import os
import requests

# ======================== 项目信息 ========================

APP_NAME = "MoonOrange Bulletin PDF Downloader"
APP_VERSION = "1.0.0"
GITHUB_URL = "https://github.com/Orangewang124/Bulletin-PDF-Download"
COPYRIGHT_TEXT = f"{APP_NAME} v{APP_VERSION}  |  Free & Open Source  |  GitHub: {GITHUB_URL}"
COPYRIGHT_SHORT = f"Free & Open Source  |  GitHub: {GITHUB_URL}"

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

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("820x820")
        self.resizable(True, True)

        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_penguin.png")
        if os.path.exists(icon_path):
            try:
                icon_img = tk.PhotoImage(file=icon_path)
                self.iconphoto(True, icon_img)
                self._icon_img = icon_img  # 保持引用防止被回收
            except tk.TclError:
                pass

        self.stock_id = tk.StringVar(value="600388")
        self.stock_name = tk.StringVar(value="")
        self.stock_valid = False
        self.save_path = tk.StringVar(value="D:/BulletinPDFDownload/")
        self.start_date = tk.StringVar(value="2026-01-01")
        self.end_date = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.bulletin_list = []
        self.filtered_list = []
        self.is_downloading = False

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # ---- 股票代码 ----
        stock_frame = ttk.LabelFrame(self, text="股票代码", padding=8)
        stock_frame.pack(fill="x", padx=10, pady=(10, 4))

        ttk.Label(stock_frame, text="股票代码:").pack(side="left")
        self.stock_entry = ttk.Entry(stock_frame, textvariable=self.stock_id, width=10)
        self.stock_entry.pack(side="left", padx=(4, 8))

        self.btn_validate = ttk.Button(
            stock_frame, text="验证", command=self._validate_stock
        )
        self.btn_validate.pack(side="left", padx=(0, 8))

        self.stock_info_label = ttk.Label(stock_frame, text="未验证", foreground="gray")
        self.stock_info_label.pack(side="left", padx=4)

        # ---- 保存路径 ----
        path_frame = ttk.LabelFrame(self, text="保存路径", padding=8)
        path_frame.pack(fill="x", padx=10, pady=(10, 4))

        ttk.Entry(path_frame, textvariable=self.save_path, width=60).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(path_frame, text="浏览...", command=self._browse_path).pack(
            side="left", padx=(8, 0)
        )

        # ---- 日期范围 ----
        date_frame = ttk.LabelFrame(self, text="日期范围", padding=8)
        date_frame.pack(fill="x", padx=10, pady=4)

        ttk.Label(date_frame, text="起始日期:").pack(side="left")
        ttk.Entry(date_frame, textvariable=self.start_date, width=12).pack(
            side="left", padx=(4, 16)
        )
        ttk.Label(date_frame, text="终止日期:").pack(side="left")
        ttk.Entry(date_frame, textvariable=self.end_date, width=12).pack(
            side="left", padx=(4, 16)
        )

        # ---- 操作按钮 ----
        btn_frame = ttk.Frame(self, padding=4)
        btn_frame.pack(fill="x", padx=10, pady=4)

        self.btn_preview = ttk.Button(
            btn_frame, text="预览文件列表", command=self._preview
        )
        self.btn_preview.pack(side="left", padx=(0, 8))

        self.btn_download = ttk.Button(
            btn_frame, text="确认下载", command=self._start_download, state="disabled"
        )
        self.btn_download.pack(side="left", padx=(0, 8))

        self.progress_label = ttk.Label(btn_frame, text="")
        self.progress_label.pack(side="left", padx=8)

        # ---- 文件列表 ----
        list_frame = ttk.LabelFrame(self, text="文件列表", padding=4)
        list_frame.pack(fill="both", expand=True, padx=10, pady=4)

        columns = ("date", "name", "filename")
        self.tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", selectmode="extended"
        )
        self.tree.heading("date", text="日期")
        self.tree.heading("name", text="公告名称")
        self.tree.heading("filename", text="文件名")
        self.tree.column("date", width=100, anchor="center")
        self.tree.column("name", width=280, anchor="w")
        self.tree.column("filename", width=380, anchor="w")

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ---- 下载进度 ----
        progress_frame = ttk.LabelFrame(self, text="下载进度", padding=4)
        progress_frame.pack(fill="x", padx=10, pady=4)

        self.progress_bar = ttk.Progressbar(
            progress_frame, orient="horizontal", mode="determinate"
        )
        self.progress_bar.pack(fill="x", pady=(0, 4))

        self.log_text = tk.Text(progress_frame, height=6, state="disabled", wrap="word")
        log_scroll = ttk.Scrollbar(
            progress_frame, orient="vertical", command=self.log_text.yview
        )
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        # 日志颜色标签
        self.log_text.tag_configure("ok", foreground="green")
        self.log_text.tag_configure("fail", foreground="red")
        self.log_text.tag_configure("skip", foreground="gray")

        # ---- 底部状态栏 ----
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w").pack(
            fill="x", padx=10, pady=(0, 2)
        )

        # ---- 底部版权信息 ----
        copyright_frame = tk.Frame(self)
        copyright_frame.pack(fill="x", padx=10, pady=(0, 6))

        copyright_label = tk.Label(
            copyright_frame,
            text=COPYRIGHT_SHORT,
            fg="gray",
            anchor="center",
            font=("", 9),
        )
        copyright_label.pack(fill="x")

        # 点击版权信息打开 GitHub
        copyright_label.bind("<Button-1>", lambda e: self._open_github())
        copyright_label.configure(cursor="hand2")

    # ------------------------------------------------------------------ helpers
    def _open_github(self):
        """在浏览器中打开 GitHub 仓库页面。"""
        import webbrowser
        webbrowser.open(GITHUB_URL)

    def _browse_path(self):
        path = filedialog.askdirectory(title="选择保存路径")
        if path:
            self.save_path.set(path.replace("/", "\\") + "\\")

    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _validate_dates(self):
        try:
            start = datetime.datetime.strptime(self.start_date.get(), "%Y-%m-%d")
            end = datetime.datetime.strptime(self.end_date.get(), "%Y-%m-%d")
            if start > end:
                messagebox.showerror("错误", "起始日期不能晚于终止日期")
                return None, None
            return start, end
        except ValueError:
            messagebox.showerror("错误", "日期格式不正确，请使用 YYYY-MM-DD 格式")
            return None, None

    def _validate_stock(self):
        """验证股票代码是否有效。"""
        sid = self.stock_id.get().strip()
        if not sid:
            self.stock_info_label.configure(text="请输入股票代码", foreground="red")
            self.stock_valid = False
            return

        self.btn_validate.configure(state="disabled")
        self.stock_info_label.configure(text="验证中...", foreground="gray")
        self.update_idletasks()

        def worker():
            valid, name, msg = validate_stock_id(sid)
            self.after(0, self._show_validate_result, valid, name, msg)

        threading.Thread(target=worker, daemon=True).start()

    def _show_validate_result(self, valid, name, msg):
        self.btn_validate.configure(state="normal")
        self.stock_valid = valid
        if valid:
            self.stock_name.set(name)
            self.stock_info_label.configure(text=msg, foreground="green")
        else:
            self.stock_name.set("")
            self.stock_info_label.configure(text=msg, foreground="red")

    # ------------------------------------------------------------------ preview
    def _preview(self):
        # 先验证股票代码
        sid = self.stock_id.get().strip()
        if not sid:
            messagebox.showwarning("提示", "请先输入股票代码")
            return

        start, end = self._validate_dates()
        if start is None:
            return

        self.btn_preview.configure(state="disabled")
        self.btn_validate.configure(state="disabled")
        self.status_var.set("正在验证股票代码并获取公告列表，请稍候...")
        self.stock_info_label.configure(text="验证中...", foreground="gray")
        self.update_idletasks()

        def worker():
            # 先验证股票代码
            valid, name, vmsg = validate_stock_id(sid)
            if not valid:
                self.after(0, self._show_validate_result, valid, name, vmsg)
                self.after(0, lambda: self.status_var.set("股票代码验证失败"))
                self.after(0, lambda: self.btn_preview.configure(state="normal"))
                self.after(0, lambda: self.btn_validate.configure(state="normal"))
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
            self.tree.insert(
                "", "end", values=(item["date"], item["name"], filename)
            )

        count = len(self.filtered_list)
        self.status_var.set(f"共找到 {count} 个公告")
        self.btn_preview.configure(state="normal")
        self.btn_validate.configure(state="normal")
        self.btn_download.configure(state="normal" if count > 0 else "disabled")

    def _show_preview_error(self, err):
        self.status_var.set("获取公告列表失败")
        self.btn_preview.configure(state="normal")
        self.btn_validate.configure(state="normal")
        messagebox.showerror("错误", f"获取公告列表失败:\n{err}")

    # ------------------------------------------------------------------ download
    def _start_download(self):
        if self.is_downloading:
            return
        if not self.filtered_list:
            messagebox.showwarning("提示", "请先预览文件列表")
            return

        save = self.save_path.get().strip()
        if not save:
            messagebox.showerror("错误", "请设置保存路径")
            return

        self.is_downloading = True
        self.btn_download.configure(state="disabled")
        self.btn_preview.configure(state="disabled")
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(self.filtered_list)

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.status_var.set("正在下载...")

        def worker():
            def on_progress(idx, total, item, status, message):
                self.after(0, self._update_download_progress, idx, total, status, message)

            result = download_batch(
                self.filtered_list, save, progress_callback=on_progress
            )
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

        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _download_finished(self, result):
        self.is_downloading = False
        self.btn_download.configure(state="normal")
        self.btn_preview.configure(state="normal")

        s = result["success_count"]
        f = result["fail_count"]
        k = result["skip_count"]
        total = s + f + k

        summary = f"下载完成: 共 {total} 个, 成功 {s} 个, 失败 {f} 个, 跳过 {k} 个"
        self.status_var.set(summary)
        self._log(f"\n{summary}")

        if f > 0:
            messagebox.showwarning(
                "下载完成",
                f"下载完成\n成功: {s}\n失败: {f}\n跳过: {k}",
            )
        else:
            messagebox.showinfo(
                "下载完成",
                f"下载完成\n成功: {s}\n跳过: {k}",
            )


if __name__ == "__main__":
    app = App()
    app.mainloop()
