#!/usr/bin/env python3
"""
Google 批量搜索工具 — 从 Excel 读取搜索词，批量搜索，关键词过滤，导出结果
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
from googlesearch import search


class GoogleBatchSearcher:
    def __init__(self, root):
        self.root = root
        self.root.title("Google 批量搜索工具")
        self.root.geometry("820x680")
        self.root.resizable(True, True)
        self.root.minsize(700, 550)

        self.is_running = False
        self.all_results = []  # [{query, title, url, desc, kept}]

        self._build_ui()

    def _build_ui(self):
        # 顶部标题
        header = tk.Frame(self.root, bg="#fafafa", height=36)
        header.pack(fill=tk.X, padx=16, pady=(10, 0))
        tk.Label(header, text="Google 批量搜索工具", font=("Microsoft YaHei", 15, "bold"),
                 fg="#1d1d1f", bg="#fafafa").pack(side=tk.LEFT)

        # ---- 设置区域 ----
        card = tk.Frame(self.root, bg="white", highlightbackground="#e0e0e0", highlightthickness=1)
        card.pack(fill=tk.X, padx=16, pady=10)

        ttk.Label(card, text="📂 输入配置", font=("Microsoft YaHei", 10, "bold"), background="white").pack(anchor=tk.W, padx=12, pady=(10, 6))

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
        ttk.Spinbox(row3, from_=1, to=50, textvariable=self.num_results_var, width=5).pack(side=tk.LEFT)
        ttk.Label(row3, text="条结果", background="white").pack(side=tk.LEFT)
        ttk.Label(row3, text="搜索间隔:", background="white").pack(side=tk.LEFT, padx=(15, 0))
        self.delay_var = tk.DoubleVar(value=2.0)
        ttk.Spinbox(row3, from_=0.5, to=10, increment=0.5, textvariable=self.delay_var, width=5).pack(side=tk.LEFT)
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
        self.result_tree.column("url", width=280, minwidth=150)
        self.result_tree.column("kept", width=50, anchor=tk.CENTER, minwidth=40)

        sy = ttk.Scrollbar(result_card, orient=tk.VERTICAL, command=self.result_tree.yview)
        sx = ttk.Scrollbar(result_card, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=12)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx.pack(side=tk.BOTTOM, fill=tk.X, padx=12)

        # 双击复制链接
        self.result_tree.bind("<Double-1>", self._copy_url)

        # ---- 日志 ----
        log_frame = tk.Frame(self.root, bg="white", highlightbackground="#e0e0e0", highlightthickness=1)
        log_frame.pack(fill=tk.X, padx=16, pady=(0, 12))
        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=4, wrap=tk.WORD, font=("Consolas", 9), state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

    def _log(self, msg):
        self.log_text.configure(state=tk.NORMAL)
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.excel_path_var.set(path)

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
        self._log(f"⚙️ 每条 {num_results} 个结果 | 间隔 {delay}s")

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
                include_kw = [k.strip().lower() for k in filter_kw.replace("，", ",").split(",") if k.strip()] if filter_kw else []

                try:
                    results = list(search(query, num_results=num_results, lang="zh", sleep_interval=0))
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

                self._log(f"    → {len(results)} 条结果, 保留 {kept} 条" if include_kw else f"    → {len(results)} 条结果")

                time.sleep(delay)

            # 完成
            kept_total = sum(1 for r in self.all_results if r["kept"])
            self._log(f"✅ 搜索完成 — 共 {len(self.all_results)} 条结果, 保留 {kept_total} 条")
            self.progress_label.configure(text=f"完成 {len(self.all_results)} 条" if self.is_running else "已停止")
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
