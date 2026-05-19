#!/usr/bin/env python3
"""
Google 批量搜索工具 — 从 Excel 读取搜索词，批量搜索，关键词过滤，导出结果
v2: 使用 undetected-chromedriver（真实 Chrome 浏览器）绕过 Google 反爬
"""

import os
import sys
import time
import threading
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

import openpyxl
import urllib.parse

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class GoogleBatchSearcher:
    def __init__(self, root):
        self.root = root
        self.root.title("Google 批量搜索工具 🦇")
        self.root.geometry("840x720")
        self.root.resizable(True, True)
        self.root.minsize(700, 550)

        self.is_running = False
        self.all_results = []  # [{query, title, url, desc, kept}]
        self.driver = None
        self.driver_lock = threading.Lock()

        self._build_ui()

        # 关闭窗口时清理浏览器
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI ──────────────────────────────────────────────
    def _build_ui(self):
        # 顶部标题
        header = tk.Frame(self.root, bg="#fafafa", height=36)
        header.pack(fill=tk.X, padx=16, pady=(10, 0))
        tk.Label(header, text="Google 批量搜索工具 (浏览器版)", font=("Microsoft YaHei", 15, "bold"),
                 fg="#1d1d1f", bg="#fafafa").pack(side=tk.LEFT)

        # ---- 设置区域 ----
        card = tk.Frame(self.root, bg="white", highlightbackground="#e0e0e0", highlightthickness=1)
        card.pack(fill=tk.X, padx=16, pady=10)

        ttk.Label(card, text="📂 输入配置", font=("Microsoft YaHei", 10, "bold"),
                  background="white").pack(anchor=tk.W, padx=12, pady=(10, 6))

        # Excel 文件
        row1 = tk.Frame(card, bg="white")
        row1.pack(fill=tk.X, padx=12, pady=3)
        ttk.Label(row1, text="Excel 文件:", width=10, background="white").pack(side=tk.LEFT)
        self.excel_path_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.excel_path_var, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(row1, text="浏览...", command=self._browse_excel).pack(side=tk.LEFT)

        # Excel 配置行
        row2 = tk.Frame(card, bg="white")
        row2.pack(fill=tk.X, padx=12, pady=3)
        ttk.Label(row2, text="Sheet 名:", width=10, background="white").pack(side=tk.LEFT)
        self.sheet_name_var = tk.StringVar(value="Sheet1")
        ttk.Entry(row2, textvariable=self.sheet_name_var, width=15).pack(side=tk.LEFT, padx=5)

        ttk.Label(row2, text="搜索词列:", background="white").pack(side=tk.LEFT, padx=(15, 0))
        self.query_col_var = tk.StringVar(value="A")
        ttk.Entry(row2, textvariable=self.query_col_var, width=5).pack(side=tk.LEFT, padx=3)

        ttk.Label(row2, text="过滤关键词列:", background="white").pack(side=tk.LEFT, padx=(10, 0))
        self.filter_col_var = tk.StringVar(value="B")
        ttk.Entry(row2, textvariable=self.filter_col_var, width=5).pack(side=tk.LEFT, padx=3)

        ttk.Label(row2, text="(该列填需保留的关键词，逗号分隔，可选)", font=("Microsoft YaHei", 7),
                  foreground="#999", background="white").pack(side=tk.LEFT, padx=5)

        # 搜索参数
        row3 = tk.Frame(card, bg="white")
        row3.pack(fill=tk.X, padx=12, pady=(3, 10))
        ttk.Label(row3, text="每条搜", width=10, background="white").pack(side=tk.LEFT)
        self.num_results_var = tk.IntVar(value=10)
        ttk.Spinbox(row3, from_=1, to=30, textvariable=self.num_results_var, width=5).pack(side=tk.LEFT)
        ttk.Label(row3, text="条结果", background="white").pack(side=tk.LEFT)
        ttk.Label(row3, text="搜索间隔:", background="white").pack(side=tk.LEFT, padx=(15, 0))
        self.delay_var = tk.DoubleVar(value=3.0)
        ttk.Spinbox(row3, from_=1, to=15, increment=0.5, textvariable=self.delay_var, width=5).pack(side=tk.LEFT)
        ttk.Label(row3, text="秒（防封）", background="white").pack(side=tk.LEFT)

        # ---- 操作按钮 + 进度 ----
        btn_row = tk.Frame(card, bg="white")
        btn_row.pack(fill=tk.X, padx=12, pady=(0, 10))

        self.start_btn = ttk.Button(btn_row, text="🔍 开始批量搜索", command=self.do_search)
        self.start_btn.pack(side=tk.LEFT)
        self.stop_btn = ttk.Button(btn_row, text="⏹ 停止", command=self._stop, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(btn_row, text="📥 导出结果到 Excel", command=self.do_export, state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(btn_row, mode="determinate", length=200)
        self.progress.pack(side=tk.RIGHT, padx=10)

        self.progress_label = ttk.Label(btn_row, text="", background="white", foreground="#86868b")
        self.progress_label.pack(side=tk.RIGHT)

        # ---- 结果预览 ----
        result_card = tk.Frame(self.root, bg="white", highlightbackground="#e0e0e0", highlightthickness=1)
        result_card.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 5))

        ttk.Label(result_card, text="📋 搜索结果预览", font=("Microsoft YaHei", 10, "bold"),
                  background="white").pack(anchor=tk.W, padx=12, pady=(10, 6))

        cols = ("query", "title", "url", "kept")
        self.result_tree = ttk.Treeview(result_card, columns=cols, show="headings", height=12)
        self.result_tree.heading("query", text="搜索词")
        self.result_tree.heading("title", text="标题")
        self.result_tree.heading("url", text="链接")
        self.result_tree.heading("kept", text="保留")
        self.result_tree.column("query", width=100, minwidth=80)
        self.result_tree.column("title", width=220, minwidth=120)
        self.result_tree.column("url", width=300, minwidth=150)
        self.result_tree.column("kept", width=50, anchor=tk.CENTER, minwidth=40)

        sy = ttk.Scrollbar(result_card, orient=tk.VERTICAL, command=self.result_tree.yview)
        sx = ttk.Scrollbar(result_card, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=12)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx.pack(side=tk.BOTTOM, fill=tk.X, padx=12)

        self.result_tree.bind("<Double-1>", self._copy_url)

        # ---- 日志 ----
        log_frame = tk.Frame(self.root, bg="white", highlightbackground="#e0e0e0", highlightthickness=1)
        log_frame.pack(fill=tk.X, padx=16, pady=(0, 12))
        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=4, wrap=tk.WORD, font=("Consolas", 9), state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

    def _log(self, msg):
        """线程安全地写入日志"""
        def _write():
            self.log_text.configure(state=tk.NORMAL)
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
        self.root.after(0, _write)

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.excel_path_var.set(path)

    def _on_close(self):
        """关闭窗口前清理浏览器"""
        self._cleanup_driver()
        self.root.destroy()

    # ── 浏览器管理 ─────────────────────────────────────
    def _init_driver(self):
        """初始化 undetected Chrome 浏览器（只调一次）"""
        if self.driver is not None:
            return

        self._log("🚀 正在启动 Chrome 浏览器...")
        options = uc.ChromeOptions()
        # 一些反检测和性能优化
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-gpu")
        # 如果想隐藏窗口（无头模式存在风险，更容易被检测），不要加 --headless
        # options.add_argument("--headless=new")   # 不推荐，除非确认能工作

        try:
            self.driver = uc.Chrome(options=options, version_main=None)
            self.driver.set_page_load_timeout(30)
            self._log("✅ Chrome 浏览器已启动")
        except Exception as e:
            self._log(f"❌ 启动浏览器失败: {e}")
            raise

    def _cleanup_driver(self):
        """关闭浏览器"""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
            self._log("🧹 Chrome 浏览器已关闭")

    # ── Google 搜索核心（Selenium 版）────────────────────
    def _google_search(self, query: str, num: int = 10, lang: str = "zh") -> list[str]:
        """
        使用真实 Chrome 浏览器搜索 Google，返回结果 URL 列表。
        线程安全：通过 driver_lock 确保同一时间只有一个线程操作浏览器。
        """
        with self.driver_lock:
            if self.driver is None:
                self._init_driver()

            urls = []
            params = {
                "q": query,
                "num": min(num, 30),
                "hl": lang,
            }
            encoded = urllib.parse.urlencode(params)
            search_url = f"https://www.google.com/search?{encoded}"

            try:
                self.driver.get(search_url)
            except Exception as e:
                raise Exception(f"页面加载失败: {e}")

            # 等待搜索结果出现
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#search"))
                )
            except Exception:
                # 可能碰到了 Google consent / CAPTCHA 页面
                page_text = self.driver.page_source[:2000].lower()
                if "captcha" in page_text or "unusual traffic" in page_text:
                    raise Exception("Google 触发了 CAPTCHA 验证，请手动在浏览器中完成验证后重试")
                if "consent.google" in self.driver.current_url or "before you continue" in page_text:
                    raise Exception("Google 弹出同意页面，请手动在浏览器中点击「Accept all」后重试")
                # 不是 consent 页面，可能只是加载慢，继续尝试提取结果
                time.sleep(2)

            # 提取搜索结果链接
            # 策略 1: 通过 a[jsname="UWckNb"]（Google 2024+ 有机结果链接）
            try:
                result_links = self.driver.find_elements(By.CSS_SELECTOR, 'a[jsname="UWckNb"]')
                for link in result_links:
                    href = link.get_attribute("href")
                    if href and href.startswith("http") and "google.com" not in href:
                        if href not in urls:
                            urls.append(href)
                        if len(urls) >= num:
                            return urls[:num]
            except Exception:
                pass

            # 策略 2: 从 h3 中提取链接（备选）
            if not urls:
                try:
                    for h3 in self.driver.find_elements(By.TAG_NAME, "h3"):
                        try:
                            parent_a = h3.find_element(By.XPATH, "./ancestor::a")
                            href = parent_a.get_attribute("href")
                            if href and href.startswith("http") and "google.com" not in href:
                                if href not in urls:
                                    urls.append(href)
                        except Exception:
                            continue
                        if len(urls) >= num:
                            return urls[:num]
                except Exception:
                    pass

            # 策略 3: 宽泛匹配所有外部链接
            if not urls:
                try:
                    all_links = self.driver.find_elements(By.CSS_SELECTOR, 'a[href^="http"]')
                    for link in all_links:
                        href = link.get_attribute("href")
                        if href and not any(d in href for d in [
                            "google.com", "googleadservices.com", "youtube.com",
                            "accounts.google", "policies.google", "support.google",
                        ]):
                            if href not in urls:
                                urls.append(href)
                            if len(urls) >= num:
                                return urls[:num]
                except Exception:
                    pass

            return urls[:num]

    # ── 工具方法 ───────────────────────────────────────
    def _col_to_index(self, col_letter: str) -> int:
        """A=0, B=1, ..."""
        col_letter = col_letter.strip().upper()
        result = 0
        for c in col_letter:
            result = result * 26 + (ord(c) - ord('A') + 1)
        return result - 1

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

    # ── 搜索与导出 ─────────────────────────────────────
    def do_search(self):
        excel_path = self.excel_path_var.get().strip()
        if not excel_path or not Path(excel_path).exists():
            messagebox.showerror("错误", "请先选择一个有效的 Excel 文件")
            return

        sheet_name = self.sheet_name_var.get().strip() or "Sheet1"
        query_col = self.query_col_var.get().strip() or "A"
        filter_col = self.filter_col_var.get().strip() or ""
        num_results = self.num_results_var.get()
        delay = self.delay_var.get()

        try:
            q_idx = self._col_to_index(query_col)
            f_idx = self._col_to_index(filter_col) if filter_col else None
        except Exception:
            messagebox.showerror("错误", f"列号格式错误: {query_col} / {filter_col}")
            return

        # 读取 Excel
        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True)
            if sheet_name not in wb.sheetnames:
                messagebox.showerror("错误", f"Sheet '{sheet_name}' 不存在。可用: {', '.join(wb.sheetnames)}")
                return
            ws = wb[sheet_name]
        except Exception as e:
            messagebox.showerror("错误", f"无法读取 Excel: {e}")
            return

        # 解析搜索词和过滤关键词
        rows_data = []
        for row in ws.iter_rows(min_row=2, values_only=True):  # 跳过标题行
            if not row or len(row) <= q_idx:
                continue
            query = str(row[q_idx]).strip() if row[q_idx] else ""
            if not query:
                continue
            filter_kw = ""
            if f_idx is not None and len(row) > f_idx and row[f_idx]:
                filter_kw = str(row[f_idx]).strip()
            rows_data.append((query, filter_kw))
        wb.close()

        if not rows_data:
            messagebox.showinfo("提示", "Excel 中没有找到搜索词（从第2行开始读取）")
            return

        self._log(f"📖 读取到 {len(rows_data)} 条搜索词")
        self._log(f"⚙️ 每条 {num_results} 个结果 | 间隔 {delay}s | 浏览器模式")

        # 清理旧结果
        self.all_results.clear()
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 开始搜索
        self.is_running = True
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self.export_btn.configure(state=tk.DISABLED)
        self.progress.configure(maximum=len(rows_data), value=0)

        def _run():
            total = len(rows_data)
            for i, (query, filter_kw) in enumerate(rows_data):
                if not self.is_running:
                    break

                self.progress_label.configure(text=f"{i+1}/{total}")
                self.progress.configure(value=i + 1)
                self.root.update_idletasks()

                self._log(f"🔍 [{i+1}/{total}] 搜索: {query}")

                # 解析过滤关键词
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
                        url_lower = url.lower()
                        is_kept = any(kw in url_lower for kw in include_kw)

                    self.all_results.append({
                        "query": query,
                        "title": title,
                        "url": url,
                        "kept": is_kept,
                    })
                    if is_kept:
                        kept += 1

                    # 实时更新列表
                    status = "✅" if is_kept else "⏭️"
                    self.root.after(0, lambda q=query, t=title, u=url, s=status:
                                    self.result_tree.insert("", tk.END, values=(q, t, u, s)))

                msg = f"    → {len(results)} 条结果"
                if include_kw:
                    msg += f", 保留 {kept} 条"
                self._log(msg)

                time.sleep(delay)

            # 完成
            self._cleanup_driver()
            kept_total = sum(1 for r in self.all_results if r["kept"])
            self._log(f"✅ 搜索完成 — 共 {len(self.all_results)} 条结果, 保留 {kept_total} 条")
            self.progress_label.configure(
                text=f"完成 {len(self.all_results)} 条" if self.is_running else "已停止"
            )
            self.start_btn.configure(state=tk.NORMAL)
            self.stop_btn.configure(state=tk.DISABLED)
            self.export_btn.configure(state=tk.NORMAL if self.all_results else tk.DISABLED)
            self.is_running = False

        threading.Thread(target=_run, daemon=True).start()

    def do_export(self):
        kept = [r for r in self.all_results if r["kept"]]
        if not kept:
            messagebox.showinfo("提示", "没有可导出的结果")
            return

        out_path = filedialog.asksaveasfilename(
            title="保存结果",
            defaultextension=".xlsx",
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

        # 也导出全部结果到另一个 sheet
        ws2 = wb.create_sheet("全部结果")
        ws2.append(["搜索词", "结果链接", "标题", "是否保留"])
        for r in self.all_results:
            ws2.append([r["query"], r["url"], r["title"], "✅" if r["kept"] else "⏭️"])

        wb.save(out_path)
        self._log(f"📥 已导出: {out_path}")
        messagebox.showinfo("导出成功", f"结果已保存到:\n{out_path}\n\n包含 {len(kept)} 条保留结果 ({len(self.all_results)} 条全部结果)")


def main():
    root = tk.Tk()
    GoogleBatchSearcher(root)
    root.mainloop()


if __name__ == "__main__":
    main()
