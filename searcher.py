#!/usr/bin/env python3
"""
Google 批量搜索工具 — Glassmorphism UI (PySide6 + QWebEngine + Tailwind)
v6: 毛玻璃 Web UI + undetected-chromedriver 真实浏览器搜索
"""

import sys, json, time, threading, urllib.parse
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QFileDialog, QMessageBox
from PySide6.QtCore import Qt, QUrl, QObject, Signal, Slot, QThread, Property
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebChannel import QWebChannel

import openpyxl
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ── HTML 模板 ──────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Google Batch Searcher</title>
<script src="https://cdn.tailwindcss.com"></script>
<script>tailwind.config={darkMode:'class'}</script>
<script src="qrc:///qtwebchannel/qwebchannel.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
*{font-family:'Inter','Microsoft YaHei',sans-serif;margin:0;padding:0;box-sizing:border-box}
body{background:linear-gradient(135deg,#0f0c29,#302b63,#24243e);min-height:100vh;overflow-x:hidden;color:#e2e8f0}
body::before{content:'';position:fixed;top:-50%;left:-50%;width:200%;height:200%;
 background:radial-gradient(circle at 60% 0%,rgba(139,92,246,0.15) 0%,transparent 50%),
  radial-gradient(circle at 0% 80%,rgba(59,130,246,0.1) 0%,transparent 50%),
  radial-gradient(circle at 80% 100%,rgba(236,72,153,0.1) 0%,transparent 50%);
 animation:bg-shift 20s ease infinite;z-index:0}

@keyframes bg-shift{0%,100%{transform:translate(0,0)}33%{transform:translate(2%,1%)}66%{transform:translate(-1%,2%)}}

.glass{background:rgba(30,27,75,0.45);backdrop-filter:blur(20px) saturate(180%);-webkit-backdrop-filter:blur(20px) saturate(180%);
 border:1px solid rgba(255,255,255,0.08);border-radius:16px;box-shadow:0 8px 32px rgba(0,0,0,0.3);transition:all .3s}
.glass-light{background:rgba(45,40,95,0.5);backdrop-filter:blur(16px) saturate(150%);-webkit-backdrop-filter:blur(16px) saturate(150%);
 border:1px solid rgba(255,255,255,0.06);border-radius:12px}

.btn{display:inline-flex;align-items:center;gap:6px;padding:10px 22px;border-radius:10px;font-weight:600;font-size:14px;
 border:none;cursor:pointer;transition:all .2s;position:relative;overflow:hidden}
.btn::after{content:'';position:absolute;inset:0;background:linear-gradient(135deg,rgba(255,255,255,0.1),transparent);opacity:0;transition:opacity .2s}
.btn:hover::after{opacity:1}
.btn:active{transform:scale(.97)}

.btn-primary{background:linear-gradient(135deg,#8b5cf6,#6366f1);color:#fff;box-shadow:0 4px 15px rgba(139,92,246,0.4)}
.btn-primary:hover{box-shadow:0 6px 20px rgba(139,92,246,0.6);transform:translateY(-2px)}
.btn-danger{background:linear-gradient(135deg,#ef4444,#dc2626);color:#fff;box-shadow:0 4px 15px rgba(239,68,68,0.3)}
.btn-success{background:linear-gradient(135deg,#22c55e,#16a34a);color:#fff;box-shadow:0 4px 15px rgba(34,197,94,0.3)}
.btn-outline{background:rgba(255,255,255,0.05);color:#cbd5e1;border:1px solid rgba(255,255,255,0.1)}
.btn-outline:hover{background:rgba(255,255,255,0.1)}
.btn:disabled{opacity:.4;pointer-events:none;box-shadow:none}

.input-field{background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.1);border-radius:10px;padding:9px 14px;
 color:#e2e8f0;font-size:13px;outline:none;transition:all .2s;width:100%}
.input-field:focus{border-color:#8b5cf6;box-shadow:0 0 0 3px rgba(139,92,246,0.15);background:rgba(255,255,255,0.06)}
.input-field::placeholder{color:#64748b}
.input-sm{width:60px;text-align:center}
.input-md{width:100px}

.progress-bar{height:10px;background:rgba(255,255,255,0.06);border-radius:10px;overflow:hidden;width:100%;min-width:150px}
.progress-fill{height:100%;background:linear-gradient(90deg,#8b5cf6,#6366f1,#ec4899);border-radius:10px;transition:width .4s ease;
 background-size:200% 100%;animation:progress-shimmer 2s linear infinite}
@keyframes progress-shimmer{0%{background-position:200% 0}100%{background-position:-200% 0}}

.result-table{width:100%;border-collapse:separate;border-spacing:0;font-size:12px}
.result-table th{text-align:left;padding:10px 14px;background:rgba(139,92,246,0.15);color:#a5b4fc;font-weight:600;
 border-bottom:1px solid rgba(255,255,255,0.06);position:sticky;top:0;z-index:2}
.result-table td{padding:8px 14px;border-bottom:1px solid rgba(255,255,255,0.03);color:#cbd5e1;
 max-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.result-table tr:hover td{background:rgba(139,92,246,0.08)}
.result-table tr:nth-child(even) td{background:rgba(255,255,255,0.02)}
.result-table tr:nth-child(even):hover td{background:rgba(139,92,246,0.1)}
.url-cell{cursor:pointer;color:#93c5fd}
.url-cell:hover{color:#8b5cf6;text-decoration:underline}

.log-box{font-family:'Cascadia Code','Consolas',monospace;font-size:11px;color:#94a3b8;line-height:1.6;
 max-height:110px;overflow-y:auto;scrollbar-width:thin;scrollbar-color:rgba(255,255,255,0.1) transparent}
.log-box::-webkit-scrollbar{width:5px}
.log-box::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.1);border-radius:10px}

.status-dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px}
.status-kept{background:#22c55e;box-shadow:0 0 8px rgba(34,197,94,0.4)}
.status-skip{background:#64748b}

select.input-field{appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' fill='%2394a3b8' viewBox='0 0 16 16'%3E%3Cpath d='M8 11L3 6h10z'/%3E%3C/svg%3E");
 background-repeat:no-repeat;background-position:right 10px center;padding-right:28px}
option{background:#1e1b4b;color:#e2e8f0}

.tab-bar{display:flex;gap:4px;background:rgba(255,255,255,0.03);border-radius:12px;padding:3px}
.tab{flex:1;text-align:center;padding:7px 16px;border-radius:10px;font-size:12px;font-weight:500;
 cursor:pointer;transition:all .2s;color:#94a3b8}
.tab.active{background:rgba(139,92,246,0.2);color:#c4b5fd;box-shadow:0 2px 8px rgba(0,0,0,0.2)}
.tab:hover:not(.active){color:#e2e8f0}

.label{font-size:12px;color:#94a3b8;font-weight:500;margin-bottom:4px;display:block}

@keyframes fadeIn{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}
.animate-in{animation:fadeIn .3s ease}

.empty-state{text-align:center;padding:40px 20px;color:#64748b}
.empty-state svg{margin:0 auto 12px;opacity:.3}
</style>
</head>
<body>
<div class="relative z-10 max-w-5xl mx-auto p-5 space-y-4">
  <!-- Header -->
  <div class="glass p-5 flex items-center justify-between animate-in">
    <div class="flex items-center gap-3">
      <div class="w-10 h-10 rounded-xl flex items-center justify-center" style="background:linear-gradient(135deg,#8b5cf6,#6366f1)">
        <span class="text-xl">🔍</span>
      </div>
      <div>
        <h1 class="text-lg font-bold text-white">Google Batch Searcher</h1>
        <p class="text-xs text-slate-400">Excel 驱动 · 真实浏览器搜索 · 毛玻璃 UI</p>
      </div>
    </div>
    <div class="tab-bar" id="themeTabs">
      <div class="tab active" onclick="setTheme('dark')">🌙 Dark</div>
      <div class="tab" onclick="setTheme('light')">☀️ Light</div>
    </div>
  </div>

  <!-- Config Card -->
  <div class="glass p-5 animate-in" style="animation-delay:.05s">
    <h3 class="text-sm font-semibold text-violet-300 mb-4">📂 输入配置</h3>
    <div class="grid grid-cols-2 gap-4">
      <div class="col-span-2 flex gap-2">
        <div class="flex-1">
          <label class="label">Excel 文件</label>
          <input id="excelPath" class="input-field" placeholder="选择 Excel 文件...">
        </div>
        <button onclick="browseExcel()" class="btn btn-outline self-end">浏览</button>
      </div>
      <div>
        <label class="label">Sheet 名</label>
        <input id="sheetName" class="input-field input-md" value="Sheet1">
      </div>
      <div>
        <label class="label">搜索词列</label>
        <input id="queryCol" class="input-field input-sm" value="A">
      </div>
      <div>
        <label class="label">过滤列 (可选)</label>
        <input id="filterCol" class="input-field input-sm" value="B">
      </div>
      <div class="flex gap-3 items-end">
        <div>
          <label class="label">结果数</label>
          <input id="numResults" type="number" class="input-field input-sm" value="10" min="1" max="30">
        </div>
        <div>
          <label class="label">间隔(秒)</label>
          <input id="delaySec" type="number" class="input-field input-sm" value="3" min="1" max="15" step="0.5">
        </div>
      </div>
    </div>
  </div>

  <!-- Action Bar -->
  <div class="glass p-4 flex items-center gap-3 animate-in" style="animation-delay:.1s">
    <button id="startBtn" onclick="startSearch()" class="btn btn-primary">🔍 开始批量搜索</button>
    <button id="stopBtn" onclick="stopSearch()" class="btn btn-danger" disabled>⏹ 停止</button>
    <button id="exportBtn" onclick="exportResults()" class="btn btn-success" disabled>📥 导出 Excel</button>
    <div class="flex-1"></div>
    <div class="flex items-center gap-3">
      <span id="progressLabel" class="text-xs text-slate-400">就绪</span>
      <div class="progress-bar">
        <div id="progressFill" class="progress-fill" style="width:0%"></div>
      </div>
    </div>
  </div>

  <!-- Results -->
  <div class="glass p-5 animate-in flex-1" style="animation-delay:.15s;min-height:280px">
    <h3 class="text-sm font-semibold text-violet-300 mb-3">📋 搜索结果</h3>
    <div id="resultsContainer" style="max-height:320px;overflow-y:auto;scrollbar-width:thin;scrollbar-color:rgba(255,255,255,0.1) transparent">
      <table class="result-table" id="resultTable">
        <thead><tr><th>搜索词</th><th>标题</th><th class="w-3/5">链接</th><th class="text-center">保留</th></tr></thead>
        <tbody id="resultBody"></tbody>
      </table>
      <div id="emptyState" class="empty-state">
        <svg width="48" height="48" fill="none" stroke="#64748b" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
        <p class="text-sm">暂无搜索结果</p>
      </div>
    </div>
  </div>

  <!-- Log -->
  <div class="glass p-4 animate-in" style="animation-delay:.2s">
    <h3 class="text-sm font-semibold text-violet-300 mb-2">📜 运行日志</h3>
    <div class="log-box" id="logBox"></div>
  </div>
</div>

<script>
let allResults = [];
let totalItems = 0;
let bridge = null;

// ── QWebChannel bridge ──
new QWebChannel(qt.webChannelTransport, function(channel){
  bridge = channel.objects.bridge;
  bridge.logMsg.connect(function(msg){addLog(msg)});
  bridge.progressUpdate.connect(function(pct,text){
    document.getElementById('progressFill').style.width=pct+'%';
    document.getElementById('progressLabel').textContent=text;
  });
  bridge.resultFound.connect(function(query,title,url,kept){
    addResult(query,title,url,kept);
  });
  bridge.searchDone.connect(function(total,kept){
    document.getElementById('progressLabel').textContent='完成 '+total+' 条';
    setButtons(false);
  });
  bridge.enableButtons.connect(function(){
    setButtons(false);
    let ex = document.getElementById('exportBtn');
    ex.disabled = (allResults.filter(r=>r.kept).length===0);
  });
  bridge.setExcelPath.connect(function(p){document.getElementById('excelPath').value=p});
});

function setButtons(running){
  document.getElementById('startBtn').disabled=running;
  document.getElementById('stopBtn').disabled=!running;
  document.getElementById('exportBtn').disabled=running;
}

function addLog(msg){let b=document.getElementById('logBox');b.innerHTML+=msg+'<br>';b.scrollTop=b.scrollHeight}

function addResult(query,title,url,kept){
  document.getElementById('emptyState').style.display='none';
  allResults.push({query,title,url,kept});
  let tr=document.createElement('tr');tr.className='animate-in';
  tr.innerHTML='<td title="'+query+'">'+query+'</td><td>'+title+'</td>'+
    '<td><span class="url-cell" onclick="copyUrl(\''+url.replace(/'/g,"\\'")+'\')" title="'+url+'">'+url+'</span></td>'+
    '<td class="text-center"><span class="status-dot '+(kept?'status-kept':'status-skip')+'"></span>'+(kept?'✅':'⏭️')+'</td>';
  document.getElementById('resultBody').appendChild(tr);
}

function copyUrl(url){navigator.clipboard.writeText(url);addLog('📋 已复制: '+url)}

function browseExcel(){bridge.browseExcel()}
function startSearch(){
  if(bridge){
    let q=document.getElementById('excelPath').value.trim();
    let s=document.getElementById('sheetName').value.trim()||'Sheet1';
    let qc=document.getElementById('queryCol').value.trim()||'A';
    let fc=document.getElementById('filterCol').value.trim()||'';
    let n=parseInt(document.getElementById('numResults').value)||10;
    let d=parseFloat(document.getElementById('delaySec').value)||3;
    document.getElementById('resultBody').innerHTML='';
    document.getElementById('emptyState').style.display='block';
    allResults=[];
    setButtons(true);
    bridge.startSearch(q,s,qc,fc,n,d);
  }
}
function stopSearch(){if(bridge)bridge.stopSearch();setButtons(false)}
function exportResults(){if(bridge){let kept=allResults.filter(r=>r.kept);
  bridge.exportResults(JSON.stringify(kept),JSON.stringify(allResults))}}

function setTheme(t){
  document.querySelectorAll('#themeTabs .tab').forEach(e=>e.classList.remove('active'));
  event.target.classList.add('active');
  if(t==='light'){
    document.body.style.background='linear-gradient(135deg,#e0e7ff,#f0e6ff,#fce7f3)';
    document.body.style.color='#1e293b';
    document.querySelectorAll('.glass').forEach(g=>{
      g.style.background='rgba(255,255,255,0.55)';
      g.style.border='1px solid rgba(255,255,255,0.6)';
      g.style.boxShadow='0 8px 32px rgba(0,0,0,0.06)';
    });
    document.querySelectorAll('.input-field').forEach(i=>{
      i.style.background='rgba(255,255,255,0.7)';
      i.style.color='#1e293b';
      i.style.border='1px solid rgba(0,0,0,0.08)';
    });
    document.querySelectorAll('.result-table th').forEach(h=>{
      h.style.background='rgba(139,92,246,0.1)';
      h.style.color='#6d28d9';
    });
    document.querySelectorAll('.result-table td').forEach(d=>{
      d.style.color='#334155';
    });
    document.querySelector('.log-box').style.color='#475569';
  }else{
    document.body.style.background='linear-gradient(135deg,#0f0c29,#302b63,#24243e)';
    document.body.style.color='#e2e8f0';
    document.querySelectorAll('.glass').forEach(g=>{
      g.style.background='rgba(30,27,75,0.45)';
      g.style.border='1px solid rgba(255,255,255,0.08)';
      g.style.boxShadow='0 8px 32px rgba(0,0,0,0.3)';
    });
    document.querySelectorAll('.input-field').forEach(i=>{
      i.style.background='rgba(255,255,255,0.04)';
      i.style.color='#e2e8f0';
      i.style.border='1px solid rgba(255,255,255,0.1)';
    });
    document.querySelectorAll('.result-table th').forEach(h=>{
      h.style.background='rgba(139,92,246,0.15)';
      h.style.color='#a5b4fc';
    });
    document.querySelectorAll('.result-table td').forEach(d=>{
      d.style.color='#cbd5e1';
    });
    document.querySelector('.log-box').style.color='#94a3b8';
  }
}
</script>
</body>
</html>"""


# ── Python 桥接对象 ──────────────────────────────────────
class Bridge(QObject):
    logMsg = Signal(str)
    progressUpdate = Signal(int, str)
    resultFound = Signal(str, str, str, bool)
    searchDone = Signal(int, int)
    enableButtons = Signal()
    setExcelPath = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._worker = None
        self._main_window = parent

    @Slot()
    def browseExcel(self):
        path, _ = QFileDialog.getOpenFileName(
            self._main_window, "选择 Excel 文件", "",
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if path:
            self.setExcelPath.emit(path)

    @Slot(str, str, str, str, int, float)
    def startSearch(self, excel_path, sheet_name, query_col, filter_col, num_results, delay):
        self._worker = SearchWorker(excel_path, sheet_name, query_col, filter_col, num_results, delay)
        self._worker.log_msg.connect(self.logMsg)
        self._worker.progress_update.connect(self.progressUpdate)
        self._worker.result_found.connect(self.resultFound)
        self._worker.search_done.connect(self.searchDone)
        self._worker.enable_buttons.connect(self.enableButtons)
        self._worker.start()

    @Slot()
    def stopSearch(self):
        if self._worker:
            self._worker.stop()
        self.logMsg.emit("⏹ 用户停止了搜索")

    @Slot(str, str)
    def exportResults(self, kept_json, all_json):
        try:
            kept_items = json.loads(kept_json)
            all_items = json.loads(all_json)
        except Exception:
            self.logMsg.emit("⚠️ 导出数据解析失败")
            return

        if not kept_items:
            self.logMsg.emit("⚠️ 没有可导出的结果")
            return

        out_path, _ = QFileDialog.getSaveFileName(
            self._main_window, "保存结果",
            f"google_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel 文件 (*.xlsx)"
        )
        if not out_path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "搜索结果"
        ws.append(["搜索词", "结果链接", "标题", "是否保留"])
        for r in kept_items:
            ws.append([r.get("query",""), r.get("url",""), r.get("title",""),
                       "✅" if r.get("kept") else "⏭️"])

        ws2 = wb.create_sheet("全部结果")
        ws2.append(["搜索词", "结果链接", "标题", "是否保留"])
        for r in all_items:
            ws2.append([r.get("query",""), r.get("url",""), r.get("title",""),
                       "✅" if r.get("kept") else "⏭️"])

        wb.save(out_path)
        self.logMsg.emit(f"📥 已导出: {out_path}")


# ── 搜索工作线程 ───────────────────────────────────────
class SearchWorker(QThread):
    log_msg = Signal(str)
    progress_update = Signal(int, str)
    result_found = Signal(str, str, str, bool)
    search_done = Signal(int, int)
    enable_buttons = Signal()

    def __init__(self, excel_path, sheet_name, query_col, filter_col, num_results, delay):
        super().__init__()
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.query_col = query_col
        self.filter_col = filter_col
        self.num_results = num_results
        self.delay = delay
        self.is_running = True
        self.all_results = []

    def _col_to_index(self, col_letter):
        col_letter = col_letter.strip().upper()
        r = 0
        for c in col_letter:
            r = r * 26 + (ord(c) - ord('A') + 1)
        return r - 1

    def run(self):
        try:
            q_idx = self._col_to_index(self.query_col)
            f_idx = self._col_to_index(self.filter_col) if self.filter_col else None
        except Exception:
            self.log_msg.emit("⚠️ 列号格式错误")
            self.enable_buttons.emit()
            return

        try:
            wb = openpyxl.load_workbook(self.excel_path, read_only=True)
            if self.sheet_name not in wb.sheetnames:
                self.log_msg.emit(f"⚠️ Sheet '{self.sheet_name}' 不存在: {', '.join(wb.sheetnames)}")
                self.enable_buttons.emit()
                return
            ws = wb[self.sheet_name]
        except Exception as e:
            self.log_msg.emit(f"⚠️ 无法读取 Excel: {e}")
            self.enable_buttons.emit()
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
            self.log_msg.emit("⚠️ Excel 中没有找到搜索词")
            self.enable_buttons.emit()
            return

        self.log_msg.emit(f"📖 读取到 {len(rows_data)} 条搜索词")
        self.log_msg.emit(f"⚙️ 每条 {self.num_results} 个结果 | 间隔 {self.delay}s | 浏览器模式")

        # 启动浏览器
        try:
            driver = self._init_driver()
        except Exception as e:
            self.log_msg.emit(f"❌ 启动浏览器失败: {e}")
            self.enable_buttons.emit()
            return

        total = len(rows_data)
        for i, (query, filter_kw) in enumerate(rows_data):
            if not self.is_running:
                break

            pct = int((i + 1) / total * 100)
            self.progress_update.emit(pct, f"{i+1}/{total}")
            self.log_msg.emit(f"🔍 [{i+1}/{total}] 搜索: {query}")

            include_kw = (
                [k.strip().lower() for k in filter_kw.replace("，", ",").split(",") if k.strip()]
                if filter_kw else []
            )

            try:
                results = self._google_search(driver, query, self.num_results)
            except Exception as e:
                self.log_msg.emit(f"⚠️ 搜索失败: {e}")
                time.sleep(self.delay)
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
                self.result_found.emit(query, title, url, is_kept)

            msg = f"    → {len(results)} 条结果"
            if include_kw:
                msg += f", 保留 {kept} 条"
            self.log_msg.emit(msg)
            time.sleep(self.delay)

        try:
            driver.quit()
        except Exception:
            pass

        kept_total = sum(1 for r in self.all_results if r["kept"])
        self.log_msg.emit(f"✅ 搜索完成 — 共 {len(self.all_results)} 条, 保留 {kept_total} 条")
        self.search_done.emit(len(self.all_results), kept_total)
        self.enable_buttons.emit()

    def stop(self):
        self.is_running = False

    def _init_driver(self):
        self.log_msg.emit("🚀 正在启动 Chrome 浏览器...")
        options = uc.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-gpu")
        driver = uc.Chrome(options=options, version_main=None)
        driver.set_page_load_timeout(30)
        self.log_msg.emit("✅ Chrome 浏览器已启动")
        return driver

    def _google_search(self, driver, query, num, lang="zh"):
        urls = []
        params = {"q": query, "num": min(num, 30), "hl": lang}
        search_url = f"https://www.google.com/search?{urllib.parse.urlencode(params)}"
        try:
            driver.get(search_url)
        except Exception as e:
            raise Exception(f"页面加载失败: {e}")
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#search"))
            )
        except Exception:
            pt = driver.page_source[:2000].lower()
            if "captcha" in pt or "unusual traffic" in pt:
                raise Exception("Google 触发 CAPTCHA，请手动在浏览器中完成验证")
            if "consent.google" in driver.current_url:
                raise Exception("Google 弹出同意页，请手动点击 Accept all")
            time.sleep(2)
        try:
            for link in driver.find_elements(By.CSS_SELECTOR, 'a[jsname="UWckNb"]'):
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
                for h3 in driver.find_elements(By.TAG_NAME, "h3"):
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
                for link in driver.find_elements(By.CSS_SELECTOR, 'a[href^="http"]'):
                    href = link.get_attribute("href")
                    if href and not any(d in href for d in [
                        "google.com", "youtube.com", "accounts.google", "policies.google",
                    ]):
                        if href not in urls:
                            urls.append(href)
                        if len(urls) >= num:
                            return urls
            except Exception:
                pass
        return urls


# ── 主窗口 ────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Google Batch Searcher")
        self.resize(980, 820)
        self.setMinimumSize(800, 680)
        self._center()

        self.webview = QWebEngineView()
        self.channel = QWebChannel()
        self.bridge = Bridge(self)
        self.channel.registerObject("bridge", self.bridge)
        self.webview.page().setWebChannel(self.channel)
        self.webview.setHtml(HTML)

        self.setCentralWidget(self.webview)

    def _center(self):
        screen = QApplication.primaryScreen().availableGeometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Google Batch Searcher")
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
