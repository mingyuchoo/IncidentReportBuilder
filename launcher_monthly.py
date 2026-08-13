# -*- coding: utf-8 -*-
"""
장애 월간보고 생성기 - 로컬 웹 런처 (launcher_monthly.py)
=========================================================
주간 런처(launcher.py)의 헬퍼(_native_browse / _run_stream / _child_cmd)를 재사용하는
월간보고 전용 웹 UI. 기존 launcher.py / 주간 흐름은 건드리지 않는다 (import 만).

화면 입력:
  · 기준일
  · 원본 엑셀            (update_excel 로 갱신 + make_monthly_ppt 가 읽음)
  · 지난달 장애현황 PPT   (--prev : 템플릿 겸 기본 비교본, 결과도 이 폴더에 생성)
  · 주간 결과물 PPT       (--weekly : 별첨2 상세표 소스 = YYYYMMDD_장애공유.pptx)
  · (선택) 비교 파일      (--compare : 신규확정 N건 비교 기준. 미지정 시 지난달 PPT 사용)

버튼:  ① 엑셀 업데이트 / ② 월간 PPT 생성 / 전체 실행(①→②)
결과 파일:  지난달 PPT 폴더에  'YY년 M월 장애현황_기준일(YYYYMMDD).pptx'

실행:  월간보고_실행.bat   (또는  <venv>\python.exe launcher_monthly.py)
"""

import datetime as dt
import json
import os
import sys
import threading
import time
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import appconfig
import launcher as L          # 주간 런처의 헬퍼 재사용 (import 시 부작용 없음: main 은 __main__ 가드)
import make_monthly_ppt       # noqa: --child monthly 모드에서 직접 호출
import update_excel           # noqa

FROZEN = getattr(sys, "frozen", False)
BASE_DIR = os.path.dirname(sys.executable) if FROZEN else os.path.dirname(os.path.abspath(__file__))

_run_lock = threading.Lock()


def _child_cmd(mode):
    """월간 하위 작업 실행 커맨드(자기 자신 재실행). exe면 exe 자신을 --child 로."""
    if FROZEN:
        return [sys.executable, "--child", mode]
    return [sys.executable, os.path.abspath(__file__), "--child", mode]


DEFAULTS = {
    "date": dt.date.today().strftime("%Y-%m-%d"),
    "src": os.path.join(BASE_DIR, "data", appconfig.EXCEL_NAME),
    "prev": os.path.join(BASE_DIR, appconfig.TEMPLATE_NAME),
    "weekly": "",
    "compare": "",
}

PAGE = r"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8"><title>장애 월간보고 생성기</title>
<style>
  :root{--card:#fff;--line:#e2e8f0;--ink:#0f172a;--sub:#64748b;--brand:#6d28d9;
        --brand2:#7c3aed;--accent:#f1f5f9;}
  *{box-sizing:border-box} body{margin:0;font-family:'Segoe UI','Malgun Gothic',sans-serif;
    background:#f8fafc;color:var(--ink);}
  .wrap{max-width:900px;margin:0 auto;padding:28px 20px 60px;}
  h1{font-size:22px;margin:0 0 4px;} .desc{color:var(--sub);font-size:13px;margin:0 0 20px;}
  .card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:22px;
        box-shadow:0 1px 3px rgba(0,0,0,.05);margin-bottom:18px;}
  label{display:block;font-size:13px;font-weight:600;margin:14px 0 6px;}
  .row{display:flex;gap:8px;}
  input[type=text],input[type=date]{flex:1;padding:9px 11px;border:1px solid var(--line);
    border-radius:9px;font-size:13px;font-family:inherit;color:var(--ink);}
  input:focus{outline:2px solid #ddd6fe;border-color:var(--brand2);}
  .browse{padding:9px 12px;border:1px solid var(--line);background:var(--accent);
    border-radius:9px;font-size:12px;cursor:pointer;white-space:nowrap;} .browse:hover{background:#e2e8f0;}
  .btns{display:flex;gap:10px;flex-wrap:wrap;margin-top:22px;}
  button.run{border:0;border-radius:10px;padding:12px 18px;font-size:14px;font-weight:700;
    cursor:pointer;color:#fff;} .b1{background:var(--brand);} .b2{background:var(--brand2);} .ball{background:#0e7490;}
  button.run:hover{filter:brightness(1.08);} button.run:disabled{opacity:.5;cursor:not-allowed;}
  .ghost{background:#fff;color:var(--sub);border:1px solid var(--line);border-radius:10px;
    padding:10px 16px;font-size:13px;cursor:pointer;} .ghost:hover{background:#f1f5f9;}
  #log{background:#0b1020;color:#d1e0ff;font-family:Consolas,monospace;font-size:12.5px;line-height:1.55;
    border-radius:12px;padding:16px;height:340px;overflow:auto;white-space:pre-wrap;border:1px solid #1e293b;}
  #log .err{color:#fca5a5;} #log .ok{color:#86efac;} #log .hd{color:#93c5fd;}
  .status{font-size:13px;margin:10px 0;min-height:20px;}
  .results{display:flex;gap:10px;flex-wrap:wrap;margin-top:10px;}
  .results a{font-size:13px;background:#ecfdf5;border:1px solid #a7f3d0;color:#065f46;padding:9px 13px;
    border-radius:9px;text-decoration:none;cursor:pointer;}
  .hint{font-size:12px;color:var(--sub);margin-top:6px;}
  .topbar{display:flex;justify-content:space-between;align-items:center;}
</style></head>
<body><div class="wrap">
  <div class="topbar">
    <div><h1>🗓️ 장애 월간보고 생성기</h1>
    <p class="desc">기준일과 파일을 지정하고 실행하세요. 내 PC에서만 동작합니다.</p></div>
    <button class="ghost" onclick="quitApp()">■ 종료</button>
  </div>

  <div class="card">
    <label>기준일 — 이 날짜가 속한 달의 1일~기준일까지를 집계합니다</label>
    <div class="row"><input type="date" id="date"></div>

    <label>원본 엑셀 (데이터 시트) — 갱신 시 같은 폴더에 '_날짜' 백업 후 원본을 덮어씁니다</label>
    <div class="row"><input type="text" id="src"><button class="browse" onclick="browse('src','file','excel')">찾아보기</button></div>

    <label>지난달 장애현황 PPT — 템플릿 겸 기본 비교본. 결과 PPT는 이 파일 폴더에 생성됩니다</label>
    <div class="row"><input type="text" id="prev"><button class="browse" onclick="browse('prev','file','pptx')">찾아보기</button></div>

    <label>주간 결과물 PPT (YYYYMMDD_장애공유.pptx) — 별첨2 상세표를 사업부별로 가져옵니다</label>
    <div class="row"><input type="text" id="weekly"><button class="browse" onclick="browse('weekly','file','pptx')">찾아보기</button></div>

    <label>(선택) 비교 파일 — '전월 대비 신규확정 N건' 산정 기준. 비우면 위 지난달 PPT를 사용</label>
    <div class="row"><input type="text" id="compare"><button class="browse" onclick="browse('compare','file','pptx')">찾아보기</button></div>

    <div class="btns">
      <button class="run b1" id="btnExcel" onclick="run('excel')">① 엑셀 업데이트</button>
      <button class="run b2" id="btnPpt" onclick="run('ppt')">② 월간 PPT 생성</button>
      <button class="run ball" id="btnBoth" onclick="run('both')">전체 실행 (①→②)</button>
    </div>
    <div class="hint">전체 실행: 엑셀을 갱신한 뒤 그 결과로 월간 PPT까지 한 번에 생성합니다.</div>
  </div>

  <div class="card">
    <div class="status" id="status">대기 중…</div>
    <div id="log"></div>
    <div class="results" id="results"></div>
  </div>
</div>

<script>
const $=id=>document.getElementById(id);
async function loadDefaults(){
  const d=await (await fetch('/defaults')).json();
  for(const k of ['date','src','prev','weekly','compare']) $(k).value=d[k]||'';
}
async function browse(target,kind,flt){
  const cur=encodeURIComponent($(target).value||'');
  try{
    const r=await (await fetch(`/browse?kind=${kind}&flt=${flt}&cur=${cur}`)).json();
    if(r.path) $(target).value=r.path;
  }catch(e){ alert('파일 선택 창을 열 수 없습니다. 경로를 직접 입력하세요.'); }
}
function setBusy(b){ for(const id of ['btnExcel','btnPpt','btnBoth']) $(id).disabled=b; }
let es=null;
function run(mode){
  const p=new URLSearchParams({mode,date:$('date').value,src:$('src').value,
    prev:$('prev').value,weekly:$('weekly').value,compare:$('compare').value});
  $('log').innerHTML=''; $('results').innerHTML=''; setBusy(true); $('status').textContent='실행 중…';
  es=new EventSource('/run?'+p.toString());
  es.onmessage=e=>addLog(e.data);
  es.addEventListener('done',e=>{
    const r=JSON.parse(e.data); es.close(); setBusy(false);
    if(r.status==='ok'){ $('status').innerHTML='✅ 완료되었습니다.';
      const res=$('results'); res.innerHTML='';
      if(r.excel) res.appendChild(link('📗 엑셀 열기',r.excel));
      if(r.ppt) res.appendChild(link('📕 PPT 열기',r.ppt));
      if(r.ppt_dir) res.appendChild(link('📁 PPT 폴더',r.ppt_dir));
    } else $('status').innerHTML='❌ 오류가 발생했습니다. 로그를 확인하세요.';
  });
  es.onerror=()=>{ if(es) es.close(); setBusy(false); };
}
function addLog(line){
  const d=document.createElement('div');
  if(/오류|ERROR|Traceback|Permission/i.test(line)) d.className='err';
  else if(/완료|저장 완료|✅/.test(line)) d.className='ok';
  else if(/────|===|\[슬라이드|\[별첨|\[비교|\[주간|\[전체/.test(line)) d.className='hd';
  d.textContent=line; const log=$('log'); log.appendChild(d); log.scrollTop=log.scrollHeight;
}
function link(text,path){ const a=document.createElement('a'); a.textContent=text;
  a.onclick=()=>fetch('/open?path='+encodeURIComponent(path)); return a; }
function quitApp(){ if(confirm('런처를 종료할까요?')){ fetch('/quit');
  document.body.innerHTML='<div class="wrap"><h1>종료되었습니다.</h1><p class="desc">이 탭을 닫으셔도 됩니다.</p></div>'; } }
loadDefaults();
</script>
</body></html>"""


def _cmd_paths(qs):
    date = qs.get("date", DEFAULTS["date"])
    src = os.path.abspath(qs.get("src", DEFAULTS["src"]))
    prev = os.path.abspath(qs.get("prev", DEFAULTS["prev"]))
    weekly = qs.get("weekly", "").strip()
    compare = qs.get("compare", "").strip()
    weekly = os.path.abspath(weekly) if weekly else ""
    compare = os.path.abspath(compare) if compare else ""
    d = dt.datetime.strptime(date, "%Y-%m-%d")
    stamp = date.replace("-", "")
    prev_dir = os.path.dirname(prev)
    src_dir = os.path.dirname(src)
    out_ppt = os.path.join(prev_dir, f"{d:%y}년 {d.month}월 장애현황_기준일({stamp}).pptx")

    excel_cmd = L._child_cmd("excel") + ["--src", src, "--date", date]   # 주간 런처의 엑셀 갱신 재사용
    ppt_cmd = _child_cmd("monthly") + ["--xlsx", src, "--prev", prev,
                                       "--date", date, "--out", out_ppt]
    if weekly:
        ppt_cmd += ["--weekly", weekly]
    if compare:
        ppt_cmd += ["--compare", compare]
    return {"date": date, "src": src, "prev": prev, "weekly": weekly, "compare": compare,
            "out_ppt": out_ppt, "src_dir": src_dir, "prev_dir": prev_dir,
            "excel_cmd": excel_cmd, "ppt_cmd": ppt_cmd}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, ctype, body):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.wfile.write(body)

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        qs = {k: v[0] for k, v in urllib.parse.parse_qs(u.query).items()}
        if u.path == "/":
            self._send(200, "text/html; charset=utf-8", PAGE)
        elif u.path == "/defaults":
            self._send(200, "application/json", json.dumps(DEFAULTS))
        elif u.path == "/browse":
            path = L._native_browse(qs.get("kind", "file"), qs.get("flt", ""), qs.get("cur", ""))
            self._send(200, "application/json", json.dumps({"path": path}))
        elif u.path == "/open":
            p = qs.get("path", "")
            try:
                if p and os.path.exists(p):
                    os.startfile(p)  # noqa (Windows)
                    self._send(200, "application/json", json.dumps({"ok": True}))
                else:
                    self._send(404, "application/json", json.dumps({"ok": False}))
            except Exception as e:
                self._send(500, "application/json", json.dumps({"ok": False, "err": str(e)}))
        elif u.path == "/run":
            self._run(qs)
        elif u.path == "/quit":
            self._send(200, "application/json", json.dumps({"ok": True}))
            threading.Timer(0.4, lambda: os._exit(0)).start()
        else:
            self._send(404, "text/plain", "not found")

    def _run(self, qs):
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()

        def emit(line):
            self.wfile.write(f"data: {line}\n\n".encode("utf-8"))
            self.wfile.flush()

        c = _cmd_paths(qs)
        mode = qs.get("mode", "both")
        result = {"status": "error", "excel": None, "ppt": None, "ppt_dir": c["prev_dir"]}

        if not _run_lock.acquire(blocking=False):
            try:
                emit("[알림] 이미 실행 중입니다. 완료 후 다시 시도하세요.")
                self.wfile.write(f"event: done\ndata: {json.dumps(result)}\n\n".encode("utf-8"))
                self.wfile.flush()
            except Exception:
                pass
            return
        try:
            ok = True
            if not os.path.exists(c["src"]):
                emit(f"[오류] 원본 엑셀을 찾을 수 없습니다: {c['src']}"); ok = False
            if ok and mode in ("ppt", "both"):
                if not os.path.exists(c["prev"]):
                    emit(f"[오류] 지난달 PPT를 찾을 수 없습니다: {c['prev']}"); ok = False
                elif not c["weekly"]:
                    emit("[오류] 주간 결과물 PPT(YYYYMMDD_장애공유.pptx)를 지정하세요 (별첨2 소스)."); ok = False
                elif not os.path.exists(c["weekly"]):
                    emit(f"[오류] 주간 결과물 PPT를 찾을 수 없습니다: {c['weekly']}"); ok = False

            if ok and mode in ("excel", "both"):
                emit("──────── ① 엑셀 업데이트 시작 ────────")
                ok = L._run_stream(c["excel_cmd"], emit)
                if ok:
                    result["excel"] = c["src"]

            if ok and mode in ("ppt", "both"):
                emit("──────── ② 월간 PPT 생성 시작 ────────")
                ok = L._run_stream(c["ppt_cmd"], emit)
                if ok:
                    result["ppt"] = c["out_ppt"]

            result["status"] = "ok" if ok else "error"
        except Exception as e:
            try:
                emit(f"[오류] {e}")
            except Exception:
                pass
        finally:
            _run_lock.release()
        try:
            self.wfile.write(f"event: done\ndata: {json.dumps(result)}\n\n".encode("utf-8"))
            self.wfile.flush()
        except Exception:
            pass


def main():
    # 하위 작업 모드: 자기 자신을 재실행해 make_monthly_ppt / update_excel 를 직접 수행
    if len(sys.argv) >= 2 and sys.argv[1] == "--child":
        for _s in (sys.stdout, sys.stderr):
            try:
                _s.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
        mode = sys.argv[2] if len(sys.argv) >= 3 else ""
        rest = sys.argv[3:]
        if mode == "monthly":
            sys.argv = ["make_monthly_ppt"] + rest
            make_monthly_ppt.main()
        elif mode == "excel":
            sys.argv = ["update_excel"] + rest
            update_excel.main()
        return

    no_ui = "--no-ui" in sys.argv
    port_arg = 0
    if "--port" in sys.argv:
        try:
            port_arg = int(sys.argv[sys.argv.index("--port") + 1])
        except Exception:
            port_arg = 0
    httpd = ThreadingHTTPServer(("127.0.0.1", port_arg), Handler)
    url = f"http://127.0.0.1:{httpd.server_address[1]}/"
    print("=" * 60)
    print("  장애 월간보고 생성기 실행 중")
    print(f"  브라우저: {url}")
    print("  (이 창을 닫거나 웹페이지에서 '종료'를 누르면 끝납니다)")
    print("=" * 60)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    if not no_ui:
        time.sleep(0.4)
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Event().wait()


if __name__ == "__main__":
    main()
