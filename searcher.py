#!/usr/bin/env python3
"""
Google 批量搜索工具 — 从 Excel 读取搜索词，批量搜索，关键词过滤，导出结果
v4: ttkbootstrap 专业主题 UI + undetected-chromedriver 真实浏览器搜索
"""

import os
import sys
import time
import threading
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import urllib.parse

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText

import openpyxl

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


THEMES = ["flatly", "litera", "cosmo", "minty", "journal", "lumen",
          "darkly", "cyborg", "superhero", "solar", "vapor"]


class GoogleBatchSearcher:
    def __init__(self):
        self.root = tb.Window(themename="flatly")
        self.root.title("Google Batch Searcher")
        self.root.geometry("920x760")
        self.root.minsize(750, 600)

        self._center_window()

        self.is_running = False
        self.all_results = []
        self.driver = None
        self.driver_lock = threading.Lock()

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── 窗口居中 ──
    def _center_window(self):
        self.root.update_idletasks()
        w, h = 920, 760
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ── UI 构建 ──
    def _build_ui(self):
        # ── 工具栏 ──
        toolbar = ttk.Frame(self.root, style="primary.TFrame")
        toolbar.pack(fill=tk.X)

        inner = ttk.Frame(toolbar, style="primary.TFrame")
        inner.pack(fill=tk.X, padx=16, pady=6)

        tb_label = ttk.Label(inner, text="🔍  Google Batch Searcher",
                             font=("Microsoft YaHei", 16, "bold"),
                             foreground="white", background=tb.Style().colors.get("primary"),
                             style="primary.Inverse.TLabel")
        tb_label.pack(side=tk.LEFT)

        # 主题选择
        theme_frame = ttk.Frame(inner, style="primary.TFrame")
        theme_frame.pack(side=tk.RIGHT)

        ttk.Label(theme_frame, text="主题  ", foreground="white",
                  background=tb.Style().colors.get("primary")).pack(side=tk.LEFT)
        self.theme_cb = ttk.Combobox(theme_frame, values=THEMES, width=12, state="readonly")
        self.theme_cb.set("flatly")
        self.theme_cb.pack(side=tk.LEFT)
        self.theme_cb.bind("<<ComboboxSelected>>", self._on_theme_change)

        # ── 主内容 ──
        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True, padx=16, pady=(10, 6))

        # ── 第 1 行: 输入配置 | 搜索参数 ──
        row1 = ttk.Frame(main)
        row1.pack(fill=tk.X, pady=(0, 8))

        # 输入配置卡片
        config_lf = ttk.Labelframe(row1, text="📂 输入配置", padding=(12, 8))
        config_lf.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        # Excel 文件行
        ef = ttk.Frame(config_lf)
        ef.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(ef, text="Excel 文件", width=9).pack(side=tk.LEFT, padx=(0, 4))
        self.excel_path_var = tk.StringVar()
        ttk.Entry(ef, textvariable=self.excel_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(ef, text="浏览", bootstyle="outline-secondary", width=6,
                   command=self._browse_excel).pack(side=tk.RIGHT)

        # 列配置行
        cf = ttk.Frame(config_lf)
        cf.pack(fill=tk.X)

        ttk.Label(cf, text="Sheet", width=5).pack(side=tk.LEFT, padx=(0, 2))
        self.sheet_var = tk.StringVar(value="Sheet1")
        ttk.Entry(cf, textvariable=self.sheet_var, width=9).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(cf, text="搜索词列").pack(side=tk.LEFT, padx=(0, 2))
        self.qcol_var = tk.StringVar(value="A")
        ttk.Entry(cf, textvariable=self.qcol_var, width=4, justify=tk.CENTER).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(cf, text="过滤列").pack(side=tk.LEFT, padx=(0, 2))
        self.fcol_var = tk.StringVar(value="B")
        ttk.Entry(cf, textvariable=self.fcol_var, width=4, justify=tk.CENTER).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Label(cf, text="(可选)", foreground="gray").pack(side=tk.LEFT)

        # 搜索参数卡片
        param_lf = ttk.Labelframe(row1, text="⚙️ 搜索参数", padding=(12, 8))
        param_lf.pack(side=tk.RIGHT, fill=tk.Y)

        pf = ttk.Frame(param_lf)
        pf.pack()

        ttk.Label(pf, text="结果数").pack(side=tk.LEFT, padx=(0, 4))
        self.num_var = tk.IntVar(value=10)
        ttk.Spinbox(pf, from_=1, to=30, textvariable=self.num_var, width=5).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(pf, text="间隔(秒)").pack(side=tk.LEFT, padx=(0, 4))
        self.delay_var = tk.DoubleVar(value=3.0)
        ttk.Spinbox(pf, from_=1, to=15, increment=0.5, textvariable=self.delay_var, width=5).pack(side=tk.LEFT)

        # ── 第 2 行: 操作按钮 + 进度 ──
        action_bar = ttk.Frame(main)
        action_bar.pack(fill=tk.X, pady=(0, 8))

        self.start_btn = ttk.Button(action_bar, text="🔍 开始批量搜索",
                                    bootstyle="primary", command=self.do_search)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 4))

        self.stop_btn = ttk.Button(action_bar, text="⏹ 停止",
                                   bootstyle="danger", state=tk.DISABLED, command=self._stop)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 4))

        self.export_btn = ttk.Button(action_bar, text="📥 导出 Excel",
                                     bootstyle="success", state=tk.DISABLED, command=self.do_export)
        self.export_btn.pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(action_bar, mode="determinate", length=200,
                                        bootstyle="primary-striped")
        self.progress.pack(side=tk.RIGHT, padx=(8, 4))
        self.progress["maximum"] = 100

        self.progress_label = ttk.Label(action_bar, text="就绪", foreground="gray")
        self.progress_label.pack(side=tk.RIGHT, padx=(0, 4))

        # ── 第 3 行: 搜索结果 ──
        result_lf = ttk.Labelframe(main, text="📋 搜索结果", padding=(12, 8))
        result_lf.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        cols = ("query", "title", "url", "kept")
        self.result_tree = ttk.Treeview(result_lf, columns=cols, show="headings",
                                        height=10, bootstyle="primary")
        self.result_tree.heading("query", text="搜索词")
        self.result_tree.heading("title", text="标题")
        self.result_tree.heading("url", text="链接")
        self.result_tree.heading("kept", text="保留")
        self.result_tree.column("query", width=100, minwidth=80)
        self.result_tree.column("title", width=200, minwidth=100)
        self.result_tree.column("url", width=380, minwidth=150)
        self.result_tree.column("kept", width=50, anchor=tk.CENTER, minwidth=40)

        sy = ttk.Scrollbar(result_lf, orient=tk.VERTICAL, command=self.result_tree.yview)
        sx = ttk.Scrollbar(result_lf, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx.pack(side=tk.BOTTOM, fill=tk.X)

        self.result_tree.bind("<Double-1>", self._copy_url)

        # ── 第 4 行: 日志 ──
        log_lf = ttk.Labelframe(main, text="📜 运行日志", padding=(12, 6))
        log_lf.pack(fill=tk.X, pady=(0, 4))

        self.log_text = ScrolledText(log_lf, height=4, autohide=True,
                                     font=("Consolas", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ── 主题切换 ──
    def _on_theme_change(self, event=None):
        theme = self.theme_cb.get()
        tb.Style().theme_use(theme)

    # ── 日志 ──
    def _log(self, msg: str):
        def _write():
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
            self.log_text.see(tk.END)
        self.root.after(0, _write)

    # ── 文件浏览 ──
    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.excel_path_var.set(path)

    # ── 关闭 ──
    def _on_close(self):
        self._cleanup_driver()
        self.root.destroy()

    # ── 浏览器管理 ──
    def _init_driver(self):
        if self.driver is not None:
            return
        self._log("🚀 正在启动 Chrome 浏览器...")
        options = uc.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-gpu")
        try:
            self.driver = uc.Chrome(options=options, version_main=None)
            self.driver.set_page_load_timeout(30)
            self._log("✅ Chrome 浏览器已启动")
        except Exception as e:
            self._log(f"❌ 启动浏览器失败: {e}")
            raise

    def _cleanup_driver(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
            self._log("🧹 Chrome 浏览器已关闭")

    # ── Google 搜索核心 ──
    def _google_search(self, query: str, num: int = 10, lang: str = "zh") -> list[str]:
        with self.driver_lock:
            if self.driver is None:
                self._init_driver()
            urls = []
            params = {"q": query, "num": min(num, 30), "hl": lang}
            search_url = f"https://www.google.com/search?{urllib.parse.urlencode(params)}"

            try:
                self.driver.get(search_url)
            except Exception as e:
                raise Exception(f"页面加载失败: {e}")

            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#search"))
                )
            except Exception:
                pt = self.driver.page_source[:2000].lower()
                if "captcha" in pt or "unusual traffic" in pt:
                    raise Exception("Google 触发 CAPTCHA，请手动在浏览器中完成验证后重试")
                if "consent.google" in self.driver.current_url or "before you continue" in pt:
                    raise Exception("Google 弹出同意页，请手动点击 Accept all 后重试")
                time.sleep(2)

            try:
                for link in self.driver.find_elements(By.CSS_SELECTOR, 'a[jsname="UWckNb"]'):
                    href = link.get_attribute("href")
                    if href and href.startswith("http") and "google.com" not in href:
                        if href not in urls:
                            urls.append(href)
                        if len(urls) >= num:
                            return urls
            except Exception:
                pass

            if not urls:
                try:
                    for h3 in self.driver.find_elements(By.TAG_NAME, "h3"):
                        try:
                            a = h3.find_element("xpath", "./ancestor::a")
                            href = a.get_attribute("href")
                            if href and href.startswith("http") and "google.com" not in href:
                                if href not in urls:
                                    urls.append(href)
                        except Exception:
                            continue
                        if len(urls) >= num:
                            return urls
                except Exception:
                    pass

            if not urls:
                try:
                    for link in self.driver.find_elements(By.CSS_SELECTOR, 'a[href^="http"]'):
                        href = link.get_attribute("href")
                        if href and not any(d in href for d in [
                            "google.com", "googleadservices.com", "youtube.com",
                            "accounts.google", "policies.google", "support.google",
                        ]):
                            if href not in urls:
                                urls.append(href)
                            if len(urls) >= num:
                                return urls
                except Exception:
                    pass

            return urls

    # ── 工具方法 ──
    def _col_to_index(self, col_letter: str) -> int:
        col_letter = col_letter.strip().upper()
        r = 0
        for c in col_letter:
            r = r * 26 + (ord(c) - ord('A') + 1)
        return r - 1

    def _stop(self):
        self.is_running = False
        self._log("⏹ 用户停止了搜索")

    def _copy_url(self, event):
        sel = self.result_tree.selection()
        if sel:
            url = self.result_tree.item(sel[0], "values")[2]
            self.root.clipboard_clear()
            self.root.clipboard_append(url)
            self._log(f"📋 已复制: {url}")

    # ── 批量搜索 ──
    def do_search(self):
        excel_path = self.excel_path_var.get().strip()
        if not excel_path or not Path(excel_path).exists():
            messagebox.showerror("错误", "请先选择一个有效的 Excel 文件")
            return

        sheet_name = self.sheet_var.get().strip() or "Sheet1"
        query_col = self.qcol_var.get().strip() or "A"
        filter_col = self.fcol_var.get().strip() or ""

        try:
            num_results = self.num_var.get()
            delay = self.delay_var.get()
        except (tk.TclError, ValueError):
            messagebox.showerror("错误", "条数或间隔格式不对")
            return

        try:
            q_idx = self._col_to_index(query_col)
            f_idx = self._col_to_index(filter_col) if filter_col else None
        except Exception:
            messagebox.showerror("错误", f"列号格式错误: {query_col} / {filter_col}")
            return

        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True)
            if sheet_name not in wb.sheetnames:
                messagebox.showerror("错误", f"Sheet '{sheet_name}' 不存在。可用: {', '.join(wb.sheetnames)}")
                return
            ws = wb[sheet_name]
        except Exception as e:
            messagebox.showerror("错误", f"无法读取 Excel: {e}")
            return

        rows_data = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) <= q_idx:
                continue
            q = str(row[q_idx]).strip() if row[q_idx] else ""
            if not q:
                continue
            fk = ""
            if f_idx is not None and len(row) > f_idx and row[f_idx]:
                fk = str(row[f_idx]).strip()
            rows_data.append((q, fk))
        wb.close()

        if not rows_data:
            messagebox.showinfo("提示", "Excel 中没有找到搜索词（从第2行开始读取）")
            return

        self._log(f"📖 读取到 {len(rows_data)} 条搜索词")
        self._log(f"⚙️ 每条 {num_results} 个结果 | 间隔 {delay}s | 浏览器模式")

        self.all_results.clear()
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self.is_running = True
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self.export_btn.configure(state=tk.DISABLED)
        self.progress["value"] = 0

        def _run():
            total = len(rows_data)
            for i, (query, filter_kw) in enumerate(rows_data):
                if not self.is_running:
                    break

                pct = int((i + 1) / total * 100)
                self.progress_label.configure(text=f"{i+1}/{total}")
                self.progress["value"] = pct
                self.root.update_idletasks()

                self._log(f"🔍 [{i+1}/{total}] 搜索: {query}")

                include_kw = (
                    [k.strip().lower() for k in filter_kw.replace("，", ",").split(",") if k.strip()]
                    if filter_kw else []
                )

                try:
                    results = self._google_search(query, num=num_results, lang="zh")
                except Exception as e:
                    self._log(f"⚠️ 搜索失败: {e}")
                    time.sleep(delay)
                    continue

                kept = 0
                for j, url in enumerate(results):
                    title = f"结果 {j+1}"
                    is_kept = True
                    if include_kw:
                        is_kept = any(kw in url.lower() for kw in include_kw)

                    self.all_results.append({"query": query, "title": title, "url": url, "kept": is_kept})
                    if is_kept:
                        kept += 1

                    status = "✅" if is_kept else "⏭️"
                    self.root.after(0, lambda q=query, t=title, u=url, s=status:
                                    self.result_tree.insert("", tk.END, values=(q, t, u, s)))

                msg = f"    → {len(results)} 条结果"
                if include_kw:
                    msg += f", 保留 {kept} 条"
                self._log(msg)
                time.sleep(delay)

            self._cleanup_driver()
            kept_total = sum(1 for r in self.all_results if r["kept"])
            self._log(f"✅ 搜索完成 — 共 {len(self.all_results)} 条, 保留 {kept_total} 条")
            self.progress_label.configure(text=f"完成 {len(self.all_results)} 条")
            self.progress["value"] = 100
            self.start_btn.configure(state=tk.NORMAL)
            self.stop_btn.configure(state=tk.DISABLED)
            self.export_btn.configure(state=tk.NORMAL if self.all_results else tk.DISABLED)
            self.is_running = False

        threading.Thread(target=_run, daemon=True).start()

    # ── 导出 Excel ──
    def do_export(self):
        kept = [r for r in self.all_results if r["kept"]]
        if not kept:
            messagebox.showinfo("提示", "没有可导出的结果")
            return

        out_path = filedialog.asksaveasfilename(
            title="保存结果", defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile=f"google_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not out_path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "搜索结果"
        ws.append(["搜索词", "结果链接", "标题", "是否保留"])
        for r in kept:
            ws.append([r["query"], r["url"], r["title"], "✅" if r["kept"] else "⏭️"])

        ws2 = wb.create_sheet("全部结果")
        ws2.append(["搜索词", "结果链接", "标题", "是否保留"])
        for r in self.all_results:
            ws2.append([r["query"], r["url"], r["title"], "✅" if r["kept"] else "⏭️"])

        wb.save(out_path)
        self._log(f"📥 已导出: {out_path}")
        messagebox.showinfo("导出成功", f"结果已保存到:\n{out_path}\n\n{len(kept)} 条保留 / {len(self.all_results)} 条全部")

    def run(self):
        self.root.mainloop()


def main():
    GoogleBatchSearcher().run()


if __name__ == "__main__":
    main()
