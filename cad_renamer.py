"""
CAD 도면 파일명 일괄 변환기
PyInstaller로 단일 EXE 빌드 → Python 설치 불필요
DWG / DXF → 레이어 텍스트 추출 → 원본 파일명 직접 변경
"""
import http.server, json, os, re, shutil, subprocess
import tempfile, threading, time, webbrowser
from pathlib import Path

PORT = 19877

# ── DWG 변환기 자동 탐색 ─────────────────────────────────
ODA_PATHS = [
    r"C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe",
    r"C:\Program Files (x86)\ODA\ODAFileConverter\ODAFileConverter.exe",
    r"C:\Program Files\ODA File Converter\ODAFileConverter.exe",
]
LO_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]

def find_converter():
    for p in ODA_PATHS:
        if os.path.exists(p): return ("oda", p)
    for p in LO_PATHS:
        if os.path.exists(p): return ("lo", p)
    return (None, None)

CONV_TYPE, CONV_PATH = find_converter()

# ── DXF 파싱 ─────────────────────────────────────────────
def _read_dxf(path):
    for enc in ("utf-8", "cp949", "euc-kr", "latin-1"):
        try:
            return open(path, encoding=enc, errors="strict").read()
        except (UnicodeDecodeError, LookupError):
            continue
    return open(path, encoding="latin-1", errors="replace").read()

def _strip_mtext(raw):
    t = re.sub(r"\\[A-Za-z][^;]*;", "", raw)
    t = re.sub(r"[{}]", "", t)
    return re.sub(r"\s+", " ", t).strip()

def get_layers(dxf_path):
    layers = set()
    try:
        lines = _read_dxf(dxf_path).splitlines()
        in_e = False; i = 0
        while i < len(lines) - 1:
            try: code = int(lines[i].strip())
            except ValueError: i += 2; continue
            val = lines[i+1].strip(); i += 2
            if code == 0:
                in_e = val.upper() in ("TEXT","MTEXT","ATTRIB","ATTDEF")
            if in_e and code == 8:
                layers.add(val)
    except Exception:
        pass
    return sorted(layers)

def extract_texts(dxf_path, layer):
    results = []
    try:
        lines = _read_dxf(dxf_path).splitlines()
        i = 0; ent = None; ed = {}

        def flush():
            if ent in ("TEXT","ATTRIB","ATTDEF") and "1" in ed:
                if ed.get("8","").upper() == layer.upper():
                    results.append((ed["1"].strip(), float(ed.get("10",0)), float(ed.get("20",0))))
            elif ent == "MTEXT" and "1" in ed:
                if ed.get("8","").upper() == layer.upper():
                    results.append((_strip_mtext(ed["1"]), float(ed.get("10",0)), float(ed.get("20",0))))

        while i < len(lines) - 1:
            try: code = int(lines[i].strip())
            except ValueError: i += 2; continue
            val = lines[i+1].strip(); i += 2
            if code == 0: flush(); ent = val.upper(); ed = {}
            elif ent: ed[str(code)] = val
        flush()
    except Exception:
        pass
    return results

def pick_title(texts):
    cands = [(t,x,y) for t,x,y in texts
             if 2 <= len(t) <= 100
             and not re.fullmatch(r"[\d\s.\-+/\\,()]+", t)]
    if not cands: cands = [(t,x,y) for t,x,y in texts if t.strip()]
    if not cands: return None
    cands.sort(key=lambda v: v[1])
    return cands[0][0]

def sanitize(name):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name.strip())
    name = re.sub(r"\s+", " ", name).strip("_. ")
    return name[:120] or "unnamed"

# ── DWG → DXF 변환 ───────────────────────────────────────
def dwg_to_dxf(dwg_path, out_dir):
    stem = Path(dwg_path).stem
    dxf  = Path(out_dir) / (stem + ".dxf")
    if CONV_TYPE == "oda":
        try:
            subprocess.run([CONV_PATH, str(Path(dwg_path).parent), out_dir,
                            "ACAD2018", "DXF", "0", "1", stem],
                           capture_output=True, timeout=60)
            if dxf.exists(): return str(dxf), None
        except Exception as e:
            return None, str(e)
    if CONV_TYPE == "lo":
        try:
            r = subprocess.run([CONV_PATH, "--headless", "--convert-to", "dxf",
                                "--outdir", out_dir, str(dwg_path)],
                               capture_output=True, text=True, timeout=90)
            if dxf.exists(): return str(dxf), None
            return None, r.stderr or "LibreOffice 변환 실패"
        except subprocess.TimeoutExpired:
            return None, "변환 시간 초과"
        except Exception as e:
            return None, str(e)
    return None, "DWG 변환기 없음 (ODA File Converter 또는 LibreOffice 설치 필요)"

# ── HTTP 핸들러 ───────────────────────────────────────────
class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self._respond(200, "text/html; charset=utf-8", HTML.encode())
        elif self.path == "/api/info":
            self._json({"conv_type": CONV_TYPE, "conv_path": CONV_PATH})
        else:
            self._respond(404, "text/plain", b"not found")

    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(n))
        {"/api/scan": self._scan,
         "/api/preview": self._preview,
         "/api/rename": self._rename
        }.get(self.path, lambda b: self._json({"ok":False,"error":"not found"}, 404))(body)

    def _scan(self, body):
        folder = body.get("folder","").strip()
        if not folder or not os.path.isdir(folder):
            return self._json({"ok":False,"error":"폴더를 찾을 수 없습니다."})
        files = [{"name":f.name,"path":str(f),"ext":f.suffix.lower()[1:].upper()}
                 for f in sorted(Path(folder).iterdir())
                 if f.suffix.lower() in (".dxf",".dwg")]
        if not files:
            return self._json({"ok":False,"error":"DWG / DXF 파일이 없습니다."})
        layers = []
        tmpdir = tempfile.mkdtemp()
        try:
            for fi in files:
                dxf = fi["path"]
                if fi["ext"] == "DWG":
                    dxf, _ = dwg_to_dxf(fi["path"], tmpdir)
                if dxf and os.path.exists(dxf):
                    layers = get_layers(dxf)
                    if layers: break
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        self._json({"ok":True,"files":files,"layers":layers})

    def _preview(self, body):
        files = body.get("files",[]); layer = body.get("layer","").strip()
        if not files or not layer:
            return self._json({"ok":False,"error":"파일 또는 레이어 없음"})
        tmpdir = tempfile.mkdtemp(); results = []
        try:
            for fi in files:
                dxf = fi["path"]; err = None
                if fi["ext"] == "DWG":
                    dxf, err = dwg_to_dxf(fi["path"], tmpdir)
                if not dxf or not os.path.exists(dxf):
                    results.append({**fi,"ok":False,"error":err or "변환 실패","title":"","texts":[]}); continue
                texts = extract_texts(dxf, layer)
                if not texts:
                    results.append({**fi,"ok":False,"error":f"레이어 '{layer}'에 텍스트 없음","title":"","texts":[]}); continue
                title = pick_title(texts)
                if not title:
                    results.append({**fi,"ok":False,"error":"유효한 제목 없음","title":"","texts":[t for t,_,_ in texts]}); continue
                results.append({**fi,"ok":True,"title":sanitize(title),"texts":[t for t,_,_ in texts[:6]],"error":""})
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        self._json({"ok":True,"results":results})

    def _rename(self, body):
        items = body.get("items",[]); done, fail = [], []
        for item in items:
            src = Path(item["path"]); new = sanitize(item["new_name"])
            if not new:
                fail.append({"name":src.name,"error":"새 이름 없음"}); continue
            dst = src.parent / (new + src.suffix); c = 1
            while dst.exists() and dst != src:
                dst = src.parent / f"{new}_{c}{src.suffix}"; c += 1
            try:
                src.rename(dst); done.append({"old":src.name,"new":dst.name})
            except Exception as e:
                fail.append({"name":src.name,"error":str(e)})
        self._json({"ok":True,"done":done,"fail":fail})

    def _json(self, data, status=200):
        p = json.dumps(data, ensure_ascii=False).encode()
        self._respond(status, "application/json; charset=utf-8", p)

    def _respond(self, status, ctype, body):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

# ── HTML UI ───────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>CAD 도면 파일명 변환기</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&family=Share+Tech+Mono&display=swap" rel="stylesheet">
<style>
:root{--bg:#0e1117;--bg2:#161b24;--bg3:#1d2433;--cyan:#00e5ff;--cyan2:#00b8d9;--green:#00e676;--red:#ff5252;--yellow:#ffd740;--orange:#ff9800;--gray:#5c6b80;--light:#c9d6e3;--white:#eef2f7;--b1:rgba(0,229,255,.1);--b2:rgba(255,255,255,.06);}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--light);font-family:'Noto Sans KR',sans-serif;min-height:100vh;padding-bottom:80px}
.hdr{background:linear-gradient(160deg,#080d14,#0c1825 60%,#0a1520);border-bottom:1px solid var(--b1);padding:26px 40px;position:relative;overflow:hidden;}
.hdr::before{content:'';position:absolute;inset:0;background:repeating-linear-gradient(90deg,transparent,transparent 60px,rgba(0,229,255,.015) 60px,rgba(0,229,255,.015) 61px);pointer-events:none;}
.hdr::after{content:'CAD';font-family:'Share Tech Mono',monospace;font-size:180px;color:rgba(0,229,255,.03);position:absolute;right:-5px;bottom:-25px;line-height:1;pointer-events:none;user-select:none;}
.hi{position:relative;z-index:1}
.hbadge{font-family:'Share Tech Mono',monospace;font-size:10px;letter-spacing:3px;color:var(--cyan);background:rgba(0,229,255,.08);border:1px solid rgba(0,229,255,.2);display:inline-block;padding:3px 12px;border-radius:2px;margin-bottom:10px;}
.hdr h1{font-size:24px;font-weight:700;color:var(--white);line-height:1.3}
.hdr h1 span{color:var(--cyan)}
.hsub{font-size:12px;color:var(--gray);margin-top:5px}
.pills{display:flex;gap:7px;margin-top:12px;flex-wrap:wrap}
.pill{font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px;border:1px solid}
.pc{color:var(--cyan);border-color:rgba(0,229,255,.25);background:rgba(0,229,255,.07)}
.pg{color:var(--green);border-color:rgba(0,230,118,.25);background:rgba(0,230,118,.07)}
.py{color:var(--yellow);border-color:rgba(255,215,64,.25);background:rgba(255,215,64,.07)}
.po{color:var(--orange);border-color:rgba(255,152,0,.25);background:rgba(255,152,0,.07)}
.cbar{padding:8px 40px;font-size:11px;border-bottom:1px solid var(--b2)}
.cbar.ok{background:rgba(0,230,118,.06);color:var(--green)}
.cbar.warn{background:rgba(255,215,64,.06);color:var(--yellow)}
.wrap{max-width:1020px;margin:0 auto;padding:26px 18px}
.step{background:var(--bg2);border:1px solid var(--b2);border-radius:5px;padding:22px;margin-bottom:14px;position:relative;}
.snum{position:absolute;top:-1px;left:20px;background:var(--cyan);color:#000;font-family:'Share Tech Mono',monospace;font-size:10px;font-weight:700;padding:2px 12px;border-radius:0 0 4px 4px;letter-spacing:1px;}
.stitle{font-size:11px;font-weight:700;color:var(--cyan);letter-spacing:2px;text-transform:uppercase;margin:10px 0 12px;}
.row{display:flex;gap:10px;align-items:stretch}
input[type=text]{flex:1;background:var(--bg3);border:1px solid var(--b2);border-radius:3px;padding:10px 14px;color:var(--light);font-family:'Noto Sans KR',sans-serif;font-size:13px;outline:none;transition:border-color .2s;}
input[type=text]:focus{border-color:var(--cyan)}
input::placeholder{color:#2e3f52}
.btn{border:none;border-radius:3px;cursor:pointer;font-family:'Noto Sans KR',sans-serif;font-size:13px;font-weight:700;padding:10px 20px;display:inline-flex;align-items:center;gap:7px;transition:.2s;white-space:nowrap;}
.bc{background:var(--cyan);color:#000}.bc:hover{background:var(--cyan2)}.bc:disabled{background:#1a2e3a;color:#2a4555;cursor:not-allowed}
.bg_{background:var(--green);color:#000}.bg_:hover{filter:brightness(1.1)}.bg_:disabled{background:#102018;color:#1a4030;cursor:not-allowed}
.bgh{background:transparent;color:var(--gray);border:1px solid var(--b2)}.bgh:hover{border-color:var(--cyan);color:var(--cyan)}
.st{font-size:12px;margin-top:10px;padding:8px 14px;border-radius:3px;display:none;line-height:1.6}
.st.s{display:block}
.ld{background:rgba(0,229,255,.06);color:var(--cyan);border:1px solid rgba(0,229,255,.2)}
.er{background:rgba(255,82,82,.07);color:var(--red);border:1px solid rgba(255,82,82,.2)}
.ok{background:rgba(0,230,118,.06);color:var(--green);border:1px solid rgba(0,230,118,.2)}
.spin{display:inline-block;width:12px;height:12px;border:2px solid rgba(0,229,255,.2);border-top-color:var(--cyan);border-radius:50%;animation:spin .7s linear infinite;vertical-align:middle;margin-right:6px}
@keyframes spin{to{transform:rotate(360deg)}}
.lgrid{display:flex;flex-wrap:wrap;gap:7px;margin-top:6px}
.chip{background:var(--bg3);border:1px solid var(--b2);border-radius:3px;padding:6px 13px;font-size:11px;cursor:pointer;font-family:'Share Tech Mono',monospace;color:var(--gray);transition:.15s;user-select:none;}
.chip:hover{border-color:var(--cyan);color:var(--cyan)}.chip.sel{background:rgba(0,229,255,.12);border-color:var(--cyan);color:var(--cyan)}
.crec{font-size:9px;color:var(--yellow);margin-left:3px}
.twrap{overflow-x:auto;border-radius:4px;border:1px solid var(--b2)}
table{width:100%;border-collapse:collapse;font-size:12px}
thead th{background:var(--bg3);padding:9px 12px;text-align:left;font-family:'Share Tech Mono',monospace;font-size:9px;letter-spacing:2px;color:var(--cyan);border-bottom:1px solid var(--b2);white-space:nowrap;}
tbody tr{border-bottom:1px solid rgba(255,255,255,.03);transition:background .1s}
tbody tr:hover{background:rgba(0,229,255,.025)}tbody tr:last-child{border-bottom:none}
td{padding:9px 12px;vertical-align:middle}
.ext{font-family:'Share Tech Mono',monospace;font-size:10px;padding:2px 7px;border-radius:2px;border:1px solid rgba(0,229,255,.2);background:rgba(0,229,255,.07);color:var(--cyan)}
.ext.dwg{border-color:rgba(255,215,64,.2);background:rgba(255,215,64,.07);color:var(--yellow)}
.old{color:var(--gray);font-family:'Share Tech Mono',monospace;font-size:11px}
.arr{color:#253040;padding:0 5px}
.nw{color:var(--white);font-weight:500}
.rok td:first-child{border-left:2px solid var(--green)}.rer td:first-child{border-left:2px solid var(--red)}.rdn td:first-child{border-left:2px solid var(--cyan)}
.tok{color:var(--green);font-size:11px}.ter{color:var(--red);font-size:11px}.tdn{color:var(--cyan);font-size:11px}
.ed{background:transparent;border:none;border-bottom:1px dashed #2e3f52;color:var(--white);font-family:'Noto Sans KR',sans-serif;font-size:12px;outline:none;width:100%;padding:2px 4px;min-width:160px;}
.ed:focus{border-bottom-color:var(--cyan)}
.rbar{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px}
.rc{background:var(--bg3);border:1px solid var(--b2);border-radius:4px;padding:14px;text-align:center}
.rv{font-family:'Share Tech Mono',monospace;font-size:26px;font-weight:700;line-height:1}
.rl{font-size:9px;letter-spacing:2px;color:var(--gray);margin-top:3px}
.hint{font-size:11px;color:var(--gray);margin-top:8px;line-height:1.7}.hint b{color:var(--yellow)}
#s2,#s3,#s4{display:none}
</style>
</head>
<body>
<div class="hdr"><div class="hi">
  <div class="hbadge">CAD FILE RENAMER v2.0</div>
  <h1>도면 파일명 <span>일괄 변환기</span></h1>
  <p class="hsub">DWG / DXF 레이어 텍스트 추출 → 원본 파일명 직접 변경</p>
  <div class="pills">
    <span class="pill pc">✓ Python 불필요</span>
    <span class="pill pg">✓ DWG 결과물 유지</span>
    <span class="pill py">✓ 레이어 자동 감지</span>
    <span class="pill po">✓ 변경 전 미리보기</span>
  </div>
</div></div>
<div id="cb" class="cbar warn">⏳ DWG 변환기 확인 중...</div>
<div class="wrap">
  <div class="step">
    <div class="snum">STEP 01</div>
    <div class="stitle">📁 CAD 파일 폴더 경로 입력</div>
    <div class="row">
      <input type="text" id="fp" placeholder="예: C:\도면\2024프로젝트"/>
      <button class="btn bc" onclick="scan()">🔍 폴더 스캔</button>
    </div>
    <p class="hint">💡 폴더 안의 <b>모든 DWG / DXF</b> 파일을 불러옵니다. DWG는 <b>ODA File Converter</b>(무료) 또는 <b>LibreOffice</b> 필요.</p>
    <div class="st" id="st1"></div>
  </div>
  <div class="step" id="s2">
    <div class="snum">STEP 02</div>
    <div class="stitle">🗂 도면 제목 레이어 선택</div>
    <p style="font-size:12px;color:var(--gray);margin-bottom:10px">TEXT / MTEXT 레이어 목록입니다. 도면 표제란 레이어를 선택하세요.</p>
    <div class="lgrid" id="lg"></div>
    <p class="hint">💡 <b>★추천</b> 레이어를 먼저 시도해보세요.</p>
    <div style="margin-top:12px"><button class="btn bc" id="pb" onclick="preview()" disabled>👁 미리보기 생성</button></div>
    <div class="st" id="st2"></div>
  </div>
  <div class="step" id="s3">
    <div class="snum">STEP 03</div>
    <div class="stitle">✏️ 변경될 파일명 확인 / 수정</div>
    <p style="font-size:12px;color:var(--gray);margin-bottom:12px">확인 후 <b style="color:var(--cyan)">새 파일명 셀을 클릭</b>해 수정할 수 있습니다.</p>
    <div class="twrap"><table>
      <thead><tr><th>형식</th><th>현재 파일명</th><th></th><th>새 파일명 (클릭 수정)</th><th>추출된 텍스트</th><th>상태</th></tr></thead>
      <tbody id="pb2"></tbody>
    </table></div>
    <div style="margin-top:14px;display:flex;gap:10px;flex-wrap:wrap;align-items:center">
      <button class="btn bg_" id="rb" onclick="doRename()">🚀 일괄 파일명 변경</button>
      <button class="btn bgh" onclick="reset()">↩ 처음부터</button>
      <span style="font-size:11px;color:var(--yellow);margin-left:6px">⚠️ 원본 파일명이 직접 변경됩니다.</span>
    </div>
    <div class="st" id="st3"></div>
  </div>
  <div class="step" id="s4">
    <div class="snum">STEP 04</div>
    <div class="stitle">✅ 변경 완료</div>
    <div class="rbar" id="rb2"></div>
    <div class="twrap"><table>
      <thead><tr><th>결과</th><th>이전 파일명</th><th></th><th>변경된 파일명</th></tr></thead>
      <tbody id="rb3"></tbody>
    </table></div>
    <div style="margin-top:14px"><button class="btn bc" onclick="reset()">🔄 다시 변환하기</button></div>
  </div>
</div>
<script>
const KW=["제목","title","text","표제","도면명","name","글자","문자","annotation","drawing","titleblock"];
let SF=[],SL="",PD=[];
fetch("/api/info").then(r=>r.json()).then(i=>{
  const b=document.getElementById("cb");
  if(i.conv_type==="oda"){b.className="cbar ok";b.textContent="✅ ODA File Converter 감지됨 — DWG 처리 가능";}
  else if(i.conv_type==="lo"){b.className="cbar ok";b.textContent="✅ LibreOffice 감지됨 — DWG 처리 가능";}
  else{b.className="cbar warn";b.textContent="⚠️ DWG 변환기 없음 — DXF만 처리 가능 | ODA File Converter(무료) 또는 LibreOffice 설치 권장";}
});
function ss(id,m,t){const e=document.getElementById(id);e.className="st s "+t;e.innerHTML=t==="ld"?`<span class="spin"></span>${m}`:m;}
function cs(id){document.getElementById(id).className="st";}
async function scan(){
  const f=document.getElementById("fp").value.trim();
  if(!f){ss("st1","❌ 폴더 경로를 입력해주세요.","er");return;}
  ss("st1",`${f} 스캔 중...`,"ld");
  ["s2","s3","s4"].forEach(id=>document.getElementById(id).style.display="none");
  const r=await p("/api/scan",{folder:f});
  if(!r.ok){ss("st1","❌ "+r.error,"er");return;}
  SF=r.files;
  ss("st1",`✅ ${r.files.length}개 발견 — DWG: ${r.files.filter(x=>x.ext==="DWG").length} / DXF: ${r.files.filter(x=>x.ext==="DXF").length}`,"ok");
  buildL(r.layers);
  document.getElementById("s2").style.display="block";
  document.getElementById("s2").scrollIntoView({behavior:"smooth",block:"start"});
}
function buildL(layers){
  const g=document.getElementById("lg");g.innerHTML="";SL="";document.getElementById("pb").disabled=true;
  if(!layers.length){g.innerHTML=`<p style="color:var(--red);font-size:12px">텍스트 레이어 없음</p>`;return;}
  const sl=[...layers].sort((a,b)=>(KW.some(k=>a.toLowerCase().includes(k))?0:1)-(KW.some(k=>b.toLowerCase().includes(k))?0:1));
  sl.forEach((l,i)=>{
    const rec=KW.some(k=>l.toLowerCase().includes(k));
    const c=document.createElement("div");c.className="chip"+(i===0?" sel":"");
    c.innerHTML=l+(rec?`<span class="crec">★추천</span>`:"");
    c.onclick=()=>{document.querySelectorAll(".chip").forEach(x=>x.classList.remove("sel"));c.classList.add("sel");SL=l;document.getElementById("pb").disabled=false;};
    g.appendChild(c);
    if(i===0){SL=l;document.getElementById("pb").disabled=false;}
  });
}
async function preview(){
  if(!SL){ss("st2","❌ 레이어를 선택해주세요.","er");return;}
  ss("st2",`'${SL}' 레이어 텍스트 추출 중...`,"ld");
  document.getElementById("s3").style.display="none";document.getElementById("s4").style.display="none";
  const r=await p("/api/preview",{files:SF,layer:SL});
  if(!r.ok){ss("st2","❌ "+r.error,"er");return;}
  PD=r.results;
  ss("st2",`✅ 완료 — 성공: ${r.results.filter(x=>x.ok).length} / 실패: ${r.results.filter(x=>!x.ok).length}`,"ok");
  buildP(r.results);
  document.getElementById("s3").style.display="block";
  document.getElementById("s3").scrollIntoView({behavior:"smooth",block:"start"});
}
function buildP(results){
  const tb=document.getElementById("pb2");tb.innerHTML="";
  results.forEach((r,i)=>{
    const tr=document.createElement("tr");tr.className=r.ok?"rok":"rer";
    const nv=r.ok?r.title:r.name.replace(/\.[^.]+$/,"");
    const h=r.texts&&r.texts.length?r.texts.slice(0,4).join(" / "):"-";
    tr.innerHTML=`<td><span class="ext ${r.ext==="DWG"?"dwg":""}">${r.ext}</span></td>
      <td class="old">${e(r.name)}</td><td class="arr">→</td>
      <td><input class="ed" id="n${i}" value="${e(nv)}"></td>
      <td style="color:var(--gray);font-size:11px;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${e(h)}">${e(h)}</td>
      <td>${r.ok?`<span class="tok">✅ 성공</span>`:`<span class="ter">❌ ${e(r.error)}</span>`}</td>`;
    tb.appendChild(tr);
  });
}
async function doRename(){
  const items=PD.map((r,i)=>({path:r.path,new_name:(document.getElementById(`n${i}`)?.value||"").trim()})).filter(x=>x.new_name);
  if(!items.length){ss("st3","❌ 변경할 항목이 없습니다.","er");return;}
  document.getElementById("rb").disabled=true;
  ss("st3",`${items.length}개 변경 중...`,"ld");
  const r=await p("/api/rename",{items});cs("st3");
  buildR(r.done,r.fail);
  document.getElementById("s4").style.display="block";
  document.getElementById("s4").scrollIntoView({behavior:"smooth",block:"start"});
  document.getElementById("rb").disabled=false;
}
function buildR(done,fail){
  document.getElementById("rb2").innerHTML=`
    <div class="rc"><div class="rv" style="color:var(--cyan)">${done.length+fail.length}</div><div class="rl">전체</div></div>
    <div class="rc"><div class="rv" style="color:var(--green)">${done.length}</div><div class="rl">성공</div></div>
    <div class="rc"><div class="rv" style="color:var(--red)">${fail.length}</div><div class="rl">실패</div></div>`;
  const tb=document.getElementById("rb3");tb.innerHTML="";
  done.forEach(d=>{const tr=document.createElement("tr");tr.className="rdn";tr.innerHTML=`<td><span class="tdn">✅ 완료</span></td><td class="old">${e(d.old)}</td><td class="arr">→</td><td class="nw">${e(d.new)}</td>`;tb.appendChild(tr);});
  fail.forEach(f=>{const tr=document.createElement("tr");tr.className="rer";tr.innerHTML=`<td><span class="ter">❌ 실패</span></td><td class="old">${e(f.name)}</td><td></td><td style="color:var(--red);font-size:11px">${e(f.error)}</td>`;tb.appendChild(tr);});
}
function reset(){SF=[];SL="";PD=[];document.getElementById("fp").value="";["s2","s3","s4"].forEach(id=>document.getElementById(id).style.display="none");["st1","st2","st3"].forEach(id=>document.getElementById(id).className="st");document.getElementById("rb").disabled=false;window.scrollTo({top:0,behavior:"smooth"});}
function e(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}
async function p(url,body){const r=await fetch(url,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(body)});return r.json();}
</script>
</body></html>"""

# ── 실행 ─────────────────────────────────────────────────
if __name__ == "__main__":
    server = http.server.HTTPServer(("127.0.0.1", PORT), Handler)
    print(f"\n  🗂  CAD 도면 파일명 변환기\n  브라우저: http://localhost:{PORT}\n  종료: 창 닫기\n")
    threading.Thread(target=lambda: (time.sleep(0.9), webbrowser.open(f"http://localhost:{PORT}")), daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
