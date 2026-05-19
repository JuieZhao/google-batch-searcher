#!/usr/bin/env python3
"""
Google 批量搜索工具 — 从 Excel 读取搜索词，批量搜索，关键词过滤，导出结果
v3: CustomTkinter 现代 UI + undetected-chromedriver 真实浏览器搜索
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

import customtkinter as ctk
import openpyxl

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ── 全局主题配置 ────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT = "#3b82f6"
ACCENT_HOVER = "#2563eb"
BG_DARK = "#1a1a2e"
BG_CARD = "#16213e"
BG_CARD2 = "#0f3460"
TEXT_PRIMARY = "#e2e8f0"
TEXT_SECONDARY = "#94a3b8"
GREEN = "#22c55e"
RED = "#ef4444"
YELLOW = "#eab308"


class GoogleBatchSearcher(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Google Batch Searcher")
        self.geometry("900x780")
        self.resizable(True, True)
        self.minsize(750, 600)

        # 窗口图标 / 居中
        self._center_window()

        self.is_running = False
        self.all_results = []
        self.driver = None
        self.driver_lock = threading.Lock()

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── 窗口居中 ──────────────────────────────────────
    def _center_window(self):
        self.update_idletasks()
        w, h = 900, 780
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── 辅助控件 ──────────────────────────────────────
    def _make_spin_entry(self, parent, default, width=64, from_val=1, to_val=50):
        """创建带校验的数字输入框（替代 ttk.Spinbox）"""
        var = ctk.StringVar(value=str(default))
        entry = ctk.CTkEntry(parent, width=width, height=32, textvariable=var,
                             justify="center")

        def validate(char):
            if char == "":
                return True
            if not char.isdigit():
                return False
            try:
                v = int(var.get() + char)
                return from_val <= v <= to_val
            except ValueError:
                return False

        vcmd = self.register(validate)
        entry.configure(validate="key", validatecommand=(vcmd, "%S"))
        return entry, var

    # ── UI 构建 ────────────────────────────────────────
    def _build_ui(self):
        # ── 顶栏 ──
        topbar = ctk.CTkFrame(self, height=44, fg_color=BG_DARK)
        topbar.pack(fill=tk.X, padx=0, pady=0)

        ctk.CTkLabel(topbar, text="🔍 Google Batch Searcher",
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side=tk.LEFT, padx=20, pady=8)

        # 主题切换
        self.theme_btn = ctk.CTkButton(
            topbar, text="☀️", width=36, height=36,
            fg_color="transparent", hover_color=BG_CARD2,
            command=self._toggle_theme
        )
        self.theme_btn.pack(side=tk.RIGHT, padx=12, pady=6)

        self.appearance_label = ctk.CTkLabel(
            topbar, text="Dark", font=ctk.CTkFont(size=11),
            text_color=TEXT_SECONDARY
        )
        self.appearance_label.pack(side=tk.RIGHT, padx=(0, 4), pady=8)

        # ── 主内容滚动区域 ──
        content = ctk.CTkScrollableFrame(self, fg_color="transparent")
        content.pack(fill=tk.BOTH, expand=True, padx=16, pady=(10, 6))

        # ── 卡片 1: 输入配置 ──
        card1 = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12)
        card1.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(card1, text="📂 输入配置", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor=tk.W, padx=18, pady=(14, 10))

        # Excel 选择
        row_f = ctk.CTkFrame(card1, fg_color="transparent")
        row_f.pack(fill=tk.X, padx=18, pady=(0, 8))

        ctk.CTkLabel(row_f, text="Excel 文件", width=80, anchor="w",
                     text_color=TEXT_SECONDARY).pack(side=tk.LEFT, padx=(0, 8))
        self.excel_path_var = ctk.StringVar()
        ctk.CTkEntry(row_f, textvariable=self.excel_path_var, height=32).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ctk.CTkButton(row_f, text="浏览", width=70, height=32,
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=self._browse_excel).pack(side=tk.RIGHT)

        # 列配置
        row_c = ctk.CTkFrame(card1, fg_color="transparent")
        row_c.pack(fill=tk.X, padx=18, pady=(0, 8))

        ctk.CTkLabel(row_c, text="Sheet 名", width=80, anchor="w",
                     text_color=TEXT_SECONDARY).pack(side=tk.LEFT, padx=(0, 8))
        self.sheet_var = ctk.StringVar(value="Sheet1")
        ctk.CTkEntry(row_c, textvariable=self.sheet_var, width=100, height=32).pack(
            side=tk.LEFT, padx=(0, 16))

        ctk.CTkLabel(row_c, text="搜索词列", anchor="w",
                     text_color=TEXT_SECONDARY).pack(side=tk.LEFT, padx=(0, 6))
        self.qcol_var = ctk.StringVar(value="A")
        ctk.CTkEntry(row_c, textvariable=self.qcol_var, width=50, height=32,
                     justify="center").pack(side=tk.LEFT, padx=(0, 16))

        ctk.CTkLabel(row_c, text="过滤列", anchor="w",
                     text_color=TEXT_SECONDARY).pack(side=tk.LEFT, padx=(0, 6))
        self.fcol_var = ctk.StringVar(value="B")
        ctk.CTkEntry(row_c, textvariable=self.fcol_var, width=50, height=32,
                     justify="center").pack(side=tk.LEFT, padx=(0, 8))

        ctk.CTkLabel(row_c, text="(可选, 保留关键词用)", font=ctk.CTkFont(size=10),
                     text_color=TEXT_SECONDARY).pack(side=tk.LEFT)

        # 搜索参数
        row_p = ctk.CTkFrame(card1, fg_color="transparent")
        row_p.pack(fill=tk.X, padx=18, pady=(4, 14))

        ctk.CTkLabel(row_p, text="条数", width=50, anchor="w",
                     text_color=TEXT_SECONDARY).pack(side=tk.LEFT, padx=(0, 6))
        self.num_entry, self.num_var = self._make_spin_entry(row_p, 10, from_val=1, to_val=30)
        self.num_entry.pack(side=tk.LEFT, padx=(0, 16))

        ctk.CTkLabel(row_p, text="间隔(秒)", anchor="w",
                     text_color=TEXT_SECONDARY).pack(side=tk.LEFT, padx=(0, 6))
        self.delay_entry, self.delay_var = self._make_spin_entry(row_p, 3, from_val=1, to_val=15)
        self.delay_entry.pack(side=tk.LEFT)

        ctk.CTkLabel(row_p, text="(防封)", font=ctk.CTkFont(size=10),
                     text_color=TEXT_SECONDARY).pack(side=tk.LEFT, padx=6)

        # ── 卡片 2: 操作区 ──
        card2 = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12)
        card2.pack(fill=tk.X, pady=(0, 10))

        btn_row = ctk.CTkFrame(card2, fg_color="transparent")
        btn_row.pack(fill=tk.X, padx=18, pady=14)

        self.start_btn = ctk.CTkButton(btn_row, text="🔍 开始批量搜索", height=38,
                                        fg_color=ACCENT, hover_color=ACCENT_HOVER,
                                        font=ctk.CTkFont(size=14, weight="bold"),
                                        command=self.do_search)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = ctk.CTkButton(btn_row, text="⏹ 停止", height=38,
                                       fg_color=RED, hover_color="#dc2626",
                                       state=tk.DISABLED, command=self._stop)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.export_btn = ctk.CTkButton(btn_row, text="📥 导出 Excel", height=38,
                                         fg_color=GREEN, hover_color="#16a34a",
                                         state=tk.DISABLED, command=self.do_export)
        self.export_btn.pack(side=tk.LEFT)

        # 进度条
        self.progress = ctk.CTkProgressBar(btn_row, width=200, height=12,
                                            fg_color=BG_CARD2, progress_color=ACCENT)
        self.progress.pack(side=tk.RIGHT, padx=(8, 8))
        self.progress.set(0)

        self.progress_label = ctk.CTkLabel(btn_row, text="", text_color=TEXT_SECONDARY,
                                           font=ctk.CTkFont(size=11))
        self.progress_label.pack(side=tk.RIGHT)

        # ── 卡片 3: 搜索结果 ──
        card3 = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12)
        card3.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        ctk.CTkLabel(card3, text="📋 搜索结果", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor=tk.W, padx=18, pady=(14, 8))

        # Treeview 容器
        tree_frame = tk.Frame(card3, bg=BG_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 8))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Results.Treeview",
                        background="#1e293b",
                        foreground="#e2e8f0",
                        fieldbackground="#1e293b",
                        borderwidth=0,
                        rowheight=28)
        style.configure("Results.Treeview.Heading",
                        background=BG_CARD2,
                        foreground=TEXT_PRIMARY,
                        borderwidth=0,
                        font=("Microsoft YaHei", 10, "bold"))
        style.map("Results.Treeview",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "#ffffff")])

        cols = ("query", "title", "url", "kept")
        self.result_tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                        height=10, style="Results.Treeview")
        self.result_tree.heading("query", text="搜索词")
        self.result_tree.heading("title", text="标题")
        self.result_tree.heading("url", text="链接")
        self.result_tree.heading("kept", text="保留")
        self.result_tree.column("query", width=100, minwidth=80)
        self.result_tree.column("title", width=200, minwidth=100)
        self.result_tree.column("url", width=360, minwidth=150)
        self.result_tree.column("kept", width=50, anchor=tk.CENTER, minwidth=40)

        sy = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        sx = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx.pack(side=tk.BOTTOM, fill=tk.X)

        # 双击复制
        self.result_tree.bind("<Double-1>", self._copy_url)

        # ── 卡片 4: 日志 ──
        card4 = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12)
        card4.pack(fill=tk.X, pady=(0, 6))

        ctk.CTkLabel(card4, text="📜 运行日志", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=TEXT_SECONDARY).pack(anchor=tk.W, padx=18, pady=(10, 4))

        self.log_box = ctk.CTkTextbox(card4, height=90, fg_color=BG_DARK,
                                       corner_radius=8, border_width=0,
                                       font=ctk.CTkFont(family="Consolas", size=11),
                                       text_color=TEXT_SECONDARY,
                                       activate_scrollbars=True)
        self.log_box.pack(fill=tk.X, padx=18, pady=(0, 12))
        self.log_box.configure(state=tk.DISABLED)

    # ── 主题切换 ──────────────────────────────────────
    def _toggle_theme(self):
        current = ctk.get_appearance_mode()
        new = "Light" if current == "Dark" else "Dark"
        ctk.set_appearance_mode(new)
        self.appearance_label.configure(text=new)
        self.theme_btn.configure(text="🌙" if new == "Dark" else "☀️")

        # 更新 Treeview 颜色以适应主题
        style = ttk.Style()
        if new == "Light":
            style.configure("Results.Treeview",
                            background="#f8fafc",
                            foreground="#1e293b",
                            fieldbackground="#f8fafc")
            style.configure("Results.Treeview.Heading",
                            background="#e2e8f0",
                            foreground="#1e293b")
        else:
            style.configure("Results.Treeview",
                            background="#1e293b",
                            foreground="#e2e8f0",
                            fieldbackground="#1e293b")
            style.configure("Results.Treeview.Heading",
                            background=BG_CARD2,
                            foreground=TEXT_PRIMARY)

    # ── 日志 ──────────────────────────────────────────
    def _log(self, msg: str):
        def _write():
            self.log_box.configure(state=tk.NORMAL)
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_box.insert(tk.END, f"[{ts}] {msg}\n")
            self.log_box.see(tk.END)
            self.log_box.configure(state=tk.DISABLED)
        self.after(0, _write)

    # ── 文件浏览 ──────────────────────────────────────
    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.excel_path_var.set(path)

    # ── 关闭 ──────────────────────────────────────────
    def _on_close(self):
        self._cleanup_driver()
        self.destroy()

    # ── 浏览器管理 ─────────────────────────────────────
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

    # ── Google 搜索核心 ─────────────────────────────────
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

            # 策略 1: a[jsname="UWckNb"]
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

            # 策略 2: h3 parent a
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

            # 策略 3: 宽泛提取
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

    # ── 列号转换 ──────────────────────────────────────
    def _col_to_index(self, col_letter: str) -> int:
        col_letter = col_letter.strip().upper()
        r = 0
        for c in col_letter:
            r = r * 26 + (ord(c) - ord('A') + 1)
        return r - 1

    # ── 停止 ──────────────────────────────────────────
    def _stop(self):
        self.is_running = False
        self._log("⏹ 用户停止了搜索")

    # ── 双击复制链接 ──────────────────────────────────
    def _copy_url(self, event):
        sel = self.result_tree.selection()
        if sel:
            url = self.result_tree.item(sel[0], "values")[2]
            self.clipboard_clear()
            self.clipboard_append(url)
            self._log(f"📋 已复制: {url}")

    # ── 批量搜索 ──────────────────────────────────────
    def do_search(self):
        excel_path = self.excel_path_var.get().strip()
        if not excel_path or not Path(excel_path).exists():
            messagebox.showerror("错误", "请先选择一个有效的 Excel 文件")
            return

        sheet_name = self.sheet_var.get().strip() or "Sheet1"
        query_col = self.qcol_var.get().strip() or "A"
        filter_col = self.fcol_var.get().strip() or ""

        try:
            num_results = int(self.num_var.get())
            delay = float(self.delay_var.get())
        except ValueError:
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
        self.progress.set(0)

        def _run():
            total = len(rows_data)
            for i, (query, filter_kw) in enumerate(rows_data):
                if not self.is_running:
                    break

                self.progress_label.configure(text=f"{i+1}/{total}")
                self.progress.set((i + 1) / total)
                self.update_idletasks()

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
                    self.after(0, lambda q=query, t=title, u=url, s=status:
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
            self.start_btn.configure(state=tk.NORMAL)
            self.stop_btn.configure(state=tk.DISABLED)
            self.export_btn.configure(state=tk.NORMAL if self.all_results else tk.DISABLED)
            self.is_running = False

        threading.Thread(target=_run, daemon=True).start()

    # ── 导出 Excel ────────────────────────────────────
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


def main():
    app = GoogleBatchSearcher()
    app.mainloop()


if __name__ == "__main__":
    main()
