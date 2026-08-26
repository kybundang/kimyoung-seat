from fastapi import FastAPI, Form, UploadFile, File, Request
from fastapi.responses import HTMLResponse
import sqlite3
import pandas as pd
import io
import re
import json
import hashlib
from collections import Counter
from datetime import datetime

app = FastAPI()

SECRET_KEY = "kimyoung_seat_secret_key"

def get_today_str():
    return datetime.now().strftime("%Y%m%d")

def generate_daily_token(date_str: str):
    return hashlib.sha256(f"{date_str}_{SECRET_KEY}".encode()).hexdigest()[:12]

# DB 연결 헬퍼 (WAL 모드 & 10초 타임아웃 적용으로 동시성 극대화)
def get_db_connection():
    conn = sqlite3.connect("seat_system.db", timeout=10.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA busy_timeout=5000;")
    return conn

# 1. DB 및 인덱스 초기화
def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            username TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            class_name TEXT DEFAULT '-',
            password TEXT DEFAULT '1234',
            status TEXT DEFAULT 'active'
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS rooms (
            room_name TEXT PRIMARY KEY,
            title TEXT,
            rows_count INTEGER,
            cols_count INTEGER,
            grid_json TEXT
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS class_configs (
            class_name TEXT PRIMARY KEY,
            room_name TEXT DEFAULT '301호',
            rows_count INTEGER DEFAULT 0,
            cols_count INTEGER DEFAULT 0,
            reset_datetime TEXT DEFAULT '',
            last_reset_at TEXT DEFAULT ''
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS seat_reservations (
            class_name TEXT,
            seat_id TEXT,
            username TEXT,
            name TEXT,
            PRIMARY KEY (class_name, seat_id)
        )
    """)
    
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_students_class ON students(class_name);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_reservations_class ON seat_reservations(class_name);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_reservations_user ON seat_reservations(username);")
    
    conn.commit()
    conn.close()

init_db()

# 예약된 리셋 일시 도달 시 좌석 자동 삭제 체크 (매일 반복 DAILY:HH:MM 지원)
def check_and_apply_resets(conn):
    cursor = conn.cursor()
    now = datetime.now()
    now_str = now.strftime("%Y-%m-%dT%H:%M")
    today_date_str = now.strftime("%Y-%m-%d")
    current_hm = now.strftime("%H:%M")
    
    cursor.execute("SELECT class_name, reset_datetime, last_reset_at FROM class_configs")
    configs = cursor.fetchall()
    
    for c_name, r_dt, l_reset in configs:
        if not r_dt:
            continue
            
        # 매일 반복 형태인 경우 (DAILY:HH:MM)
        if r_dt.startswith("DAILY:"):
            target_hm = r_dt.replace("DAILY:", "").strip()
            # 오늘 설정 시간 이후이고, 오늘 아직 리셋되지 않았다면 리셋
            if current_hm >= target_hm and (not l_reset or l_reset < today_date_str):
                cursor.execute("DELETE FROM seat_reservations WHERE class_name = ?", (c_name,))
                cursor.execute("UPDATE class_configs SET last_reset_at = ? WHERE class_name = ?", (today_date_str, c_name))
        else:
            # 1회성 특정 날짜 리셋인 경우 (YYYY-MM-DDTHH:MM)
            if r_dt <= now_str and (not l_reset or l_reset < r_dt):
                cursor.execute("DELETE FROM seat_reservations WHERE class_name = ?", (c_name,))
                cursor.execute("UPDATE class_configs SET last_reset_at = ? WHERE class_name = ?", (r_dt, c_name))
                
    conn.commit()

# 2. 학생 명단 파일 파서
def parse_student_file(contents: bytes) -> pd.DataFrame:
    try:
        df = pd.read_excel(io.BytesIO(contents))
        if '아이디' in df.columns and '이름' in df.columns: return df
    except Exception: pass

    text = ""
    for enc in ['euc-kr', 'cp949', 'utf-8-sig', 'utf-8']:
        try:
            text = contents.decode(enc)
            break
        except Exception: continue
    if not text: text = contents.decode('euc-kr', errors='ignore')

    if "<table" in text.lower() or "<tr" in text.lower():
        try:
            rows = []
            tr_matches = re.findall(r'<tr[^>]*>(.*?)</tr>', text, re.DOTALL | re.IGNORECASE)
            for tr in tr_matches:
                cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', tr, re.DOTALL | re.IGNORECASE)
                clean = [re.sub(r'<[^>]+>', '', c).replace('&nbsp;', ' ').strip() for c in cells]
                if any(clean): rows.append(clean)
            if rows:
                header, data = rows[0], rows[1:]
                max_l = max(len(r) for r in rows)
                header += [f"Col_{i}" for i in range(len(header), max_l)]
                data = [r + [''] * (max_l - len(r)) for r in data]
                df = pd.DataFrame(data, columns=header)
                if '아이디' in df.columns and '이름' in df.columns: return df
        except Exception: pass

    for sep in [',', '\t', '|']:
        try:
            df = pd.read_csv(io.StringIO(text), sep=sep)
            if '아이디' in df.columns and '이름' in df.columns: return df
        except Exception: pass

    return None

# 3. 강의실 좌석표 엑셀 파서
def parse_seat_excel(contents: bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
    seat_pattern = re.compile(r'^[A-Z]\d{1,2}$')
    
    parsed_rooms = []
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        max_r, max_c = ws.max_row, ws.max_column
        
        title = sheetname
        for r in range(1, min(max_r + 1, 5)):
            for c in range(1, max_c + 1):
                val = ws.cell(r, c).value
                if val and ('[' in str(val) or '호' in str(val)):
                    title = str(val).strip()
                    break
        
        seat_rows = []
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                val = str(ws.cell(r, c).value or '').strip()
                if seat_pattern.match(val):
                    seat_rows.append(r)
                    break
                    
        if not seat_rows: continue

        min_col, max_col = 999, 0
        for r in seat_rows:
            for c in range(1, max_c + 1):
                val = str(ws.cell(r, c).value or '').strip()
                if seat_pattern.match(val) or val.isdigit():
                    if c < min_col: min_col = c
                    if c > max_col: max_col = c
                    
        grid = []
        for r in seat_rows:
            row_cells = []
            for c in range(min_col, max_col + 1):
                val = str(ws.cell(r, c).value or '').strip()
                if seat_pattern.match(val):
                    row_cells.append({"type": "seat", "id": val})
                elif val.isdigit():
                    row_cells.append({"type": "aisle", "val": val})
                else:
                    row_cells.append({"type": "empty"})
            grid.append(row_cells)
            
        parsed_rooms.append({
            "room_name": sheetname.strip(),
            "title": title,
            "rows_count": len(grid),
            "cols_count": max_col - min_col + 1,
            "grid": grid
        })
    return parsed_rooms

# ==============================================================================
# 1. 학생용 접속 게이트웨이 및 신청 화면
# ==============================================================================
@app.get("/", response_class=HTMLResponse)
def student_view(date: str = "", token: str = ""):
    today_str = get_today_str()
    valid_token = generate_daily_token(today_str)
    
    is_expired = False
    if date or token:
        if date != today_str or token != valid_token:
            is_expired = True

    if is_expired:
        return """
        <!DOCTYPE html>
        <html lang="ko">
        <head>
            <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>QR 만료 안내 - 김영편입 좌석표 신청</title><script src="https://cdn.tailwindcss.com"></script>
        </head>
        <body class="bg-slate-100 flex justify-center p-4 min-h-screen items-center text-center">
            <div class="max-w-md bg-white p-8 rounded-2xl shadow-md border">
                <h2 class="text-lg font-bold text-slate-900 mb-2">QR코드 유효기간 만료</h2>
                <p class="text-sm text-slate-600 mb-6">해당 QR코드는 유효기간(매일 자정 24:00 만료)이 지났습니다.<br>데스크의 <b>오늘자 최신 QR코드</b>를 다시 스캔해 주세요.</p>
                <div class="text-xs text-slate-400">매일 자정 보안을 위해 QR이 자동 갱신됩니다.</div>
            </div>
        </body>
        </html>
        """

    return """
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>김영편입 좌석표 신청</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <style>
            .x-box {
                background: linear-gradient(to top right, transparent calc(50% - 1px), #cbd5e1, transparent calc(50% + 1px)),
                            linear-gradient(to bottom right, transparent calc(50% - 1px), #cbd5e1, transparent calc(50% + 1px));
            }
        </style>
    </head>
    <body class="bg-slate-100 flex justify-center p-4 min-h-screen items-center">
        <div id="main-card" class="w-full max-w-md bg-white rounded-2xl p-6 shadow-md border border-slate-200 transition-all duration-300">
            
            <div id="step-login">
                <div class="text-center mb-6">
                    <span class="inline-block bg-blue-100 text-blue-700 text-xs px-2.5 py-1 rounded-full font-bold mb-2">오늘의 좌석신청</span>
                    <h2 class="text-xl font-bold text-slate-900">김영편입 좌석표 신청</h2>
                </div>
                <div class="space-y-4">
                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">아이디</label>
                        <input type="text" id="username" placeholder="아이디 입력" class="w-full border rounded-lg p-2.5 text-sm outline-none focus:border-blue-500 font-medium">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">비밀번호</label>
                        <input type="password" id="password" placeholder="초기 비밀번호: 1234" class="w-full border rounded-lg p-2.5 text-sm outline-none focus:border-blue-500 font-medium">
                    </div>
                    <button onclick="login()" class="w-full bg-blue-600 text-white font-bold py-3 rounded-lg hover:bg-blue-700 transition shadow-sm">로그인 하기</button>
                </div>
            </div>

            <div id="step-class" class="hidden">
                <div class="flex justify-between items-center bg-blue-50 p-3 rounded-xl mb-4">
                    <div>
                        <span id="user-info-text" class="text-sm font-bold text-blue-900"></span>
                        <p class="text-xs text-blue-600 font-medium">신청할 반을 선택하세요.</p>
                    </div>
                    <div class="flex gap-1.5">
                        <button onclick="openPwModal()" class="text-xs bg-white border border-slate-300 text-slate-700 px-2 py-1.5 rounded-lg hover:bg-slate-50 font-bold">비번변경</button>
                        <button onclick="logout()" class="text-xs bg-red-500 text-white px-2 py-1.5 rounded-lg hover:bg-red-600 font-bold">로그아웃</button>
                    </div>
                </div>
                <div class="text-xs font-bold text-slate-500 mb-2">수강 중인 반 목록</div>
                <div id="class-list-container" class="space-y-2 my-3"></div>
            </div>

            <div id="step-seat" class="hidden">
                <div class="flex justify-between items-center bg-slate-50 border p-3 rounded-xl mb-4">
                    <div>
                        <div id="selected-class-title" class="text-sm font-bold text-slate-800"></div>
                        <div id="room-info-badge" class="text-xs text-blue-600 font-bold mt-0.5"></div>
                    </div>
                    <button onclick="backToClassSelect()" class="text-xs bg-slate-200 text-slate-700 px-3 py-1.5 rounded-lg hover:bg-slate-300 font-bold">◀ 반 목록</button>
                </div>

                <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                    <div class="lg:col-span-5 bg-white border border-slate-300 rounded-xl overflow-hidden shadow-sm">
                        <div class="bg-amber-400 p-3 text-white font-extrabold text-base border-b border-amber-500">
                            <span id="excel-blueprint-title">[강의실]</span>
                        </div>
                        <div class="flex bg-slate-200 text-xs font-bold text-slate-700 border-b border-slate-300">
                            <div class="flex-1 py-1.5 text-center tracking-widest">칠 판</div>
                            <div class="bg-slate-500 text-white px-3 py-1.5 font-bold">출입문</div>
                        </div>
                        <div id="excel-blueprint-grid" class="p-3 max-h-[50vh] overflow-auto bg-white"></div>
                        <div class="bg-slate-50 p-2.5 border-t border-slate-200 text-[11px] text-slate-600 flex justify-around font-bold">
                            <span><b class="text-emerald-600">■</b> 내 좌석</span>
                            <span><b class="text-red-500">■</b> 신청 완료</span>
                            <span><b class="text-slate-400">■</b> 통로/빈자리</span>
                        </div>
                    </div>

                    <div class="lg:col-span-7">
                        <div class="w-full bg-slate-700 text-white py-1.5 rounded-lg text-xs font-bold text-center mb-3 tracking-widest shadow-sm">칠 판 / 교 탁 (앞면)</div>
                        <div id="seats-grid-container" class="grid gap-1.5 max-h-[55vh] overflow-auto p-2 bg-slate-50 rounded-xl border border-slate-200"></div>
                    </div>
                </div>
            </div>

        </div>

        <div id="pw-modal" class="fixed inset-0 bg-slate-900/50 hidden items-center justify-center z-50 p-4">
            <div class="bg-white rounded-2xl p-6 w-full max-w-sm shadow-xl">
                <h3 class="text-base font-bold text-slate-900 mb-4">비밀번호 변경</h3>
                <div class="space-y-3">
                    <input type="password" id="cur-pw" placeholder="현재 비밀번호" class="w-full border rounded-lg p-2 text-sm font-medium">
                    <input type="password" id="new-pw" placeholder="새 비밀번호" class="w-full border rounded-lg p-2 text-sm font-medium">
                    <input type="password" id="new-pw-confirm" placeholder="새 비밀번호 확인" class="w-full border rounded-lg p-2 text-sm font-medium">
                </div>
                <div class="flex justify-end gap-2 mt-5">
                    <button onclick="closePwModal()" class="px-3 py-2 border rounded-lg text-xs text-slate-600 font-bold">취소</button>
                    <button onclick="changePassword()" class="px-4 py-2 bg-blue-600 text-white rounded-lg text-xs font-bold shadow-sm">변경 완료</button>
                </div>
            </div>
        </div>

        <script>
            let currentUsername = "";
            let currentStudentName = "";
            let currentSelectedClass = "";

            async function login() {
                const uName = document.getElementById("username").value;
                const pw = document.getElementById("password").value;
                const res = await fetch("/api/login", {
                    method: "POST",
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: `username=${encodeURIComponent(uName)}&password=${encodeURIComponent(pw)}`
                });
                const data = await res.json();
                if (data.success) {
                    currentUsername = data.username;
                    currentStudentName = data.name;
                    document.getElementById("user-info-text").innerText = `${currentStudentName} 님`;

                    const listContainer = document.getElementById("class-list-container");
                    listContainer.innerHTML = "";
                    data.class_list.forEach(c => {
                        listContainer.innerHTML += `
                            <button onclick="selectClass('${c}')" class="w-full text-left p-3.5 border border-slate-200 rounded-xl hover:border-blue-500 hover:bg-blue-50/50 transition flex justify-between items-center group">
                                <span class="text-sm font-bold text-slate-800 group-hover:text-blue-700">${c}</span>
                                <span class="text-xs bg-blue-100 text-blue-700 px-2.5 py-1 rounded-full font-bold">좌석 선택 ➔</span>
                            </button>
                        `;
                    });

                    document.getElementById("step-login").classList.add("hidden");
                    document.getElementById("step-class").classList.remove("hidden");
                } else {
                    alert(data.message);
                }
            }

            function selectClass(className) {
                currentSelectedClass = className;
                document.getElementById("selected-class-title").innerText = className;
                document.getElementById("main-card").className = "w-full max-w-5xl bg-white rounded-2xl p-6 shadow-md border border-slate-200 transition-all duration-300";
                document.getElementById("step-class").classList.add("hidden");
                document.getElementById("step-seat").classList.remove("hidden");
                loadSeats();
            }

            function backToClassSelect() {
                document.getElementById("main-card").className = "w-full max-w-lg bg-white rounded-2xl p-6 shadow-md border border-slate-200 transition-all duration-300";
                document.getElementById("step-seat").classList.add("hidden");
                document.getElementById("step-class").classList.remove("hidden");
            }

            function logout() {
                currentUsername = "";
                document.getElementById("main-card").className = "w-full max-w-md bg-white rounded-2xl p-6 shadow-md border border-slate-200 transition-all duration-300";
                document.getElementById("step-class").classList.add("hidden");
                document.getElementById("step-seat").classList.add("hidden");
                document.getElementById("step-login").classList.remove("hidden");
            }

            async function loadSeats() {
                const res = await fetch(`/api/seats?username=${encodeURIComponent(currentUsername)}&class_name=${encodeURIComponent(currentSelectedClass)}`);
                const data = await res.json();

                document.getElementById("room-info-badge").innerText = `강의실: ${data.room_name}`;
                document.getElementById("excel-blueprint-title").innerText = data.room_title || `[${data.room_name}]`;

                const bpGrid = document.getElementById("excel-blueprint-grid");
                bpGrid.innerHTML = "";
                const table = document.createElement("table");
                table.className = "w-full border-collapse text-center text-xs";
                
                data.grid.forEach(row => {
                    const tr = document.createElement("tr");
                    row.forEach(cell => {
                        const td = document.createElement("td");
                        if (cell.type === 'seat') {
                            let bg = "bg-white text-slate-800 border-slate-400";
                            if (cell.status_class === 'mine') bg = "bg-emerald-100 text-emerald-800 border-emerald-500 font-bold";
                            if (cell.status_class === 'occupied') bg = "bg-red-100 text-red-700 border-red-400";
                            
                            td.className = `border p-1 ${bg}`;
                            td.innerHTML = `<div class="font-bold">${cell.id}</div><div class="text-[9px] text-slate-400 font-normal">| &nbsp; |</div>`;
                        } else if (cell.type === 'aisle') {
                            td.className = "border-x border-dashed bg-slate-100/70 text-slate-500 font-bold px-1";
                            td.innerText = cell.val;
                        } else {
                            td.className = "border border-slate-200 x-box p-1 text-transparent";
                            td.innerText = "X";
                        }
                        tr.appendChild(td);
                    });
                    table.appendChild(tr);
                });
                bpGrid.appendChild(table);

                const container = document.getElementById("seats-grid-container");
                container.style.gridTemplateColumns = `repeat(${data.cols}, minmax(0, 1fr))`;
                container.innerHTML = "";

                data.grid.forEach(row => {
                    row.forEach(cell => {
                        const div = document.createElement("div");
                        if (cell.type === 'seat') {
                            let color = "bg-blue-100 text-blue-700 border-blue-200 hover:bg-blue-200 cursor-pointer";
                            if (cell.status_class === 'mine') color = "bg-emerald-500 text-white border-emerald-600 font-bold shadow-sm";
                            if (cell.status_class === 'occupied') color = "bg-red-100 text-red-400 border-red-200 cursor-not-allowed";

                            div.className = `p-2 rounded-lg border text-center text-xs font-bold transition ${color}`;
                            div.innerText = cell.id;
                            if (cell.status_class !== 'occupied') {
                                div.onclick = () => reserveSeat(cell.id);
                            }
                        } else if (cell.type === 'aisle') {
                            div.className = "flex items-center justify-center text-[11px] font-bold text-slate-400 bg-slate-200/50 rounded";
                            div.innerText = cell.val;
                        } else {
                            div.className = "p-2 opacity-0";
                        }
                        container.appendChild(div);
                    });
                });
            }

            async function reserveSeat(seatId) {
                if (!confirm(`[${currentSelectedClass}] ${seatId}번 좌석으로 신청(또는 변경)하시겠습니까?`)) return;
                const res = await fetch("/api/reserve", {
                    method: "POST",
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: `username=${encodeURIComponent(currentUsername)}&name=${encodeURIComponent(currentStudentName)}&class_name=${encodeURIComponent(currentSelectedClass)}&seat_id=${encodeURIComponent(seatId)}`
                });
                const data = await res.json();
                if (data.success) loadSeats();
                else alert(data.message);
            }

            function openPwModal() { document.getElementById("pw-modal").classList.replace("hidden", "flex"); }
            function closePwModal() { document.getElementById("pw-modal").classList.replace("flex", "hidden"); }

            async function changePassword() {
                const curPw = document.getElementById("cur-pw").value;
                const newPw = document.getElementById("new-pw").value;
                const confirmPw = document.getElementById("new-pw-confirm").value;

                if (newPw !== confirmPw) { alert("새 비밀번호 확인이 일치하지 않습니다."); return; }

                const res = await fetch("/api/change-password", {
                    method: "POST",
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: `username=${encodeURIComponent(currentUsername)}&current_password=${encodeURIComponent(curPw)}&new_password=${encodeURIComponent(newPw)}`
                });
                const data = await res.json();
                alert(data.message);
                if (data.success) closePwModal();
            }
        </script>
    </body>
    </html>
    """

# ==============================================================================
# 2. 관리자 화면 (매일 리셋 옵션 추가)
# ==============================================================================
@app.get("/admin/students", response_class=HTMLResponse)
def admin_students_view(request: Request):
    return """
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>김영편입 좌석표 신청 - 관리자</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"></script>
    </head>
    <body class="bg-slate-50 flex h-screen text-slate-800 antialiased">
        <aside class="w-64 bg-white border-r border-slate-200 flex flex-col hidden md:flex p-5 shadow-sm">
            <h1 class="text-lg font-bold text-blue-700 mb-6 tracking-tight">김영편입 좌석표 신청</h1>
            <nav class="space-y-1.5">
                <button onclick="switchTab('students')" id="tab-btn-students" class="w-full text-left px-3.5 py-2.5 bg-blue-50 text-blue-700 font-bold rounded-xl transition">학생명단</button>
                <button onclick="switchTab('classes')" id="tab-btn-classes" class="w-full text-left px-3.5 py-2.5 text-slate-600 hover:bg-slate-50 font-bold rounded-xl transition">좌석표 설정 및 인쇄</button>
                <button onclick="switchTab('qr')" id="tab-btn-qr" class="w-full text-left px-3.5 py-2.5 text-slate-600 hover:bg-slate-50 font-bold rounded-xl transition">로그인 QR</button>
                <a href="/" target="_blank" class="block px-3.5 py-2.5 text-slate-600 hover:bg-slate-50 font-bold rounded-xl mt-4 border-t pt-4 transition">좌석표 신청(학생용) ↗</a>
            </nav>
        </aside>

        <main class="flex-1 overflow-y-auto p-8">
            <!-- 탭 1: 학생명단 -->
            <div id="tab-students">
                <div class="flex justify-between items-center mb-4">
                    <div>
                        <div class="flex items-center gap-3">
                            <h2 class="text-2xl font-bold text-slate-900">학생명단</h2>
                            <button onclick="toggleClassFilterPanel()" id="btn-class-filter-toggle" class="bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1.5 shadow-xs">
                                📋 반별보기 <span id="filter-active-count" class="bg-blue-600 text-white px-1.5 py-0.5 rounded-full text-[10px] hidden font-bold">0</span>
                            </button>
                        </div>
                        <p id="summary-text" class="text-sm text-slate-500 mt-1 font-medium">불러오는 중...</p>
                    </div>
                    <div class="flex gap-2">
                        <button onclick="openAddStudentModal()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm font-bold shadow-sm transition">
                            ➕ 학생 수동 추가
                        </button>
                        <label class="cursor-pointer bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-bold shadow-sm transition">
                            수강생리스트 파일 업로드 (.xls / .xlsx / .csv)
                            <input type="file" id="upload-student-file" accept=".xls,.xlsx,.csv,.htm,.html" class="hidden" onchange="uploadStudentFile()">
                        </label>
                    </div>
                </div>

                <!-- 반별 보기 체크박스 패널 -->
                <div id="class-filter-panel" class="hidden bg-white border border-slate-200 rounded-xl p-4 shadow-md mb-6">
                    <div class="flex justify-between items-center pb-2 border-b border-slate-100 mb-3">
                        <span class="text-xs font-bold text-slate-700">반 선택 필터</span>
                        <div class="flex gap-2 text-xs">
                            <button onclick="selectAllClassFilters(true)" class="text-blue-600 font-bold hover:underline">전체 선택</button>
                            <span class="text-slate-300">|</span>
                            <button onclick="selectAllClassFilters(false)" class="text-slate-500 font-bold hover:underline">전체 해제</button>
                        </div>
                    </div>
                    <div id="class-filter-checkboxes" class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2.5 max-h-48 overflow-y-auto">
                        <!-- JS로 동적 생성 -->
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl border border-slate-200 mb-6 shadow-sm">
                    <input type="text" id="search-keyword" oninput="fetchStudents()" placeholder="아이디, 이름 또는 반명 검색..." class="w-full border rounded-lg px-3 py-2 text-sm outline-none focus:border-blue-500 font-medium">
                </div>
                <div class="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
                    <table class="w-full text-left text-sm">
                        <thead class="bg-slate-50 border-b text-xs font-bold text-slate-500 select-none">
                            <tr>
                                <th onclick="toggleSort('username')" class="py-3.5 px-5 w-48 cursor-pointer hover:bg-slate-100 transition">
                                    <div class="flex items-center gap-1">
                                        아이디 <span id="sort-icon-username" class="text-slate-400">↕</span>
                                    </div>
                                </th>
                                <th onclick="toggleSort('name')" class="py-3.5 px-5 w-36 cursor-pointer hover:bg-slate-100 transition">
                                    <div class="flex items-center gap-1">
                                        이름 <span id="sort-icon-name" class="text-slate-400">↕</span>
                                    </div>
                                </th>
                                <th onclick="toggleSort('class_name')" class="py-3.5 px-5 cursor-pointer hover:bg-slate-100 transition">
                                    <div class="flex items-center gap-1">
                                        수강 반 목록 <span id="sort-icon-class_name" class="text-slate-400">↕</span>
                                    </div>
                                </th>
                                <th class="py-3.5 px-5 w-32 text-center">비밀번호</th>
                                <th class="py-3.5 px-5 w-24 text-center">관리</th>
                            </tr>
                        </thead>
                        <tbody id="student-tbody" class="divide-y divide-slate-100"></tbody>
                    </table>
                </div>
            </div>

            <!-- 탭 2: 좌석표 설정 및 인쇄 -->
            <div id="tab-classes" class="hidden">
                <div class="flex justify-between items-center mb-6">
                    <div>
                        <h2 class="text-2xl font-bold text-slate-900">좌석표 설정 및 인쇄</h2>
                        <p class="text-sm text-slate-500 mt-1 font-medium">반별 좌석 리셋 날짜/시간 또는 매일 반복 주기를 설정한 후 [저장]할 수 있습니다.</p>
                    </div>
                    <label class="cursor-pointer bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm font-bold shadow-sm transition">
                        강의실 좌석표 엑셀 업로드 (.xlsx)
                        <input type="file" id="upload-room-file" accept=".xlsx" class="hidden" onchange="uploadRoomFile()">
                    </label>
                </div>

                <div class="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
                    <table class="w-full text-left text-sm">
                        <thead class="bg-slate-50 border-b text-xs font-bold text-slate-500">
                            <tr>
                                <th class="py-3.5 px-4">반(강좌)명</th>
                                <th class="py-3.5 px-3 text-center w-16">총원</th>
                                <th class="py-3.5 px-3 text-center w-16">신청</th>
                                <th class="py-3.5 px-3 text-center w-20">미신청</th>
                                <th class="py-3.5 px-2 text-center w-16">행</th>
                                <th class="py-3.5 px-2 text-center w-16">열</th>
                                <th class="py-3.5 px-3 w-32">배정 강의실</th>
                                <th class="py-3.5 px-3 w-48 text-center">자동 리셋 일시 예약</th>
                                <th class="py-3.5 px-2 text-center w-16">저장</th>
                                <th class="py-3.5 px-2 text-center w-20">즉시리셋</th>
                                <th class="py-3.5 px-2 text-center w-16">인쇄</th>
                            </tr>
                        </thead>
                        <tbody id="class-config-tbody" class="divide-y divide-slate-100"></tbody>
                    </table>
                </div>
            </div>

            <!-- 탭 3: 로그인 QR -->
            <div id="tab-qr" class="hidden">
                <div class="mb-6 flex justify-between items-center">
                    <div>
                        <h2 class="text-2xl font-bold text-slate-900">로그인 QR</h2>
                        <p class="text-sm text-slate-500 mt-1 font-medium">이 QR코드는 보안을 위해 <b>오늘 밤 24:00(자정)에 자동 만료</b>되며 매일 새로 갱신됩니다.</p>
                    </div>
                    <button onclick="window.print()" class="bg-slate-800 text-white px-4 py-2 rounded-lg text-sm font-bold hover:bg-slate-900 transition shadow-sm">포스터 인쇄 / PDF 저장</button>
                </div>

                <div class="flex justify-center my-6">
                    <div class="bg-white border-2 border-blue-600 rounded-3xl p-8 max-w-sm w-full shadow-lg text-center">
                        <div class="text-2xl font-extrabold text-slate-900 mb-4" id="qr-date-title"></div>
                        
                        <div class="bg-slate-50 p-4 rounded-2xl inline-block border border-slate-200 mb-4">
                            <div id="qrcode" class="flex justify-center"></div>
                        </div>

                        <p class="text-xs font-bold text-blue-600 mb-1">카메라로 스캔 ➔ 로그인 ➔ 좌석 선택</p>
                        <p class="text-[11px] text-slate-400 font-medium">초기 비밀번호는 1234 입니다.</p>
                        
                        <div class="mt-4 pt-3 border-t border-slate-100 text-[11px] text-red-500 font-bold">
                            매일 자정(24:00) 만료 및 자동 갱신
                        </div>
                    </div>
                </div>
            </div>
        </main>

        <!-- 학생 수동 추가 모달 -->
        <div id="add-student-modal" class="fixed inset-0 bg-slate-900/50 hidden items-center justify-center z-50 p-4">
            <div class="bg-white rounded-2xl p-6 w-full max-w-md shadow-2xl flex flex-col">
                <div class="flex justify-between items-center pb-3 border-b border-slate-100 mb-4">
                    <h3 class="text-base font-bold text-slate-900">➕ 학생 수동 추가</h3>
                    <button onclick="closeAddStudentModal()" class="text-slate-400 hover:text-slate-600 font-bold text-xl">&times;</button>
                </div>

                <div class="space-y-4 mb-6">
                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">아이디</label>
                        <input type="text" id="add-username" placeholder="예: kimyoung123" class="w-full border border-slate-300 rounded-lg p-2.5 text-sm font-bold outline-none focus:border-blue-500">
                    </div>

                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">이름</label>
                        <input type="text" id="add-name" placeholder="예: 홍길동" class="w-full border border-slate-300 rounded-lg p-2.5 text-sm font-bold outline-none focus:border-blue-500">
                    </div>

                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">수강 반 선택 (복수 선택 가능)</label>
                        <div id="add-class-options" class="border border-slate-200 rounded-lg p-3 max-h-40 overflow-y-auto space-y-1.5 bg-slate-50">
                            <!-- JS로 동적 생성 -->
                        </div>
                    </div>
                </div>

                <div class="flex justify-end gap-2 pt-3 border-t border-slate-100">
                    <button onclick="closeAddStudentModal()" class="px-4 py-2 bg-slate-200 text-slate-700 text-xs font-bold rounded-lg hover:bg-slate-300">취소</button>
                    <button onclick="submitAddStudent()" class="px-5 py-2 bg-emerald-600 text-white text-xs font-bold rounded-lg hover:bg-emerald-700 shadow-sm">추가 등록</button>
                </div>
            </div>
        </div>

        <!-- 미신청 학생 명단 모달 -->
        <div id="unreserved-modal" class="fixed inset-0 bg-slate-900/50 hidden items-center justify-center z-50 p-4">
            <div class="bg-white rounded-2xl p-6 w-full max-w-md shadow-2xl flex flex-col max-h-[80vh]">
                <div class="flex justify-between items-center pb-3 border-b border-slate-100">
                    <div>
                        <h3 class="text-base font-bold text-slate-900" id="unreserved-modal-title">미신청자 명단</h3>
                        <p class="text-xs text-rose-500 font-bold mt-0.5" id="unreserved-modal-count"></p>
                    </div>
                    <button onclick="closeUnreservedModal()" class="text-slate-400 hover:text-slate-600 font-bold text-xl">&times;</button>
                </div>

                <div class="overflow-y-auto py-3 space-y-2 flex-1 my-2" id="unreserved-student-list"></div>

                <div class="pt-3 border-t border-slate-100 flex justify-end">
                    <button onclick="closeUnreservedModal()" class="px-4 py-2 bg-slate-800 text-white text-xs font-bold rounded-lg hover:bg-slate-900">닫기</button>
                </div>
            </div>
        </div>

        <!-- 커스텀 날짜/시간/매일 선택 [확인] 모달 -->
        <div id="datetime-picker-modal" class="fixed inset-0 bg-slate-900/50 hidden items-center justify-center z-50 p-4">
            <div class="bg-white rounded-2xl p-6 w-full max-w-sm shadow-2xl flex flex-col">
                <div class="flex justify-between items-center pb-3 border-b border-slate-100 mb-4">
                    <h3 class="text-base font-bold text-slate-900" id="dt-modal-title">자동 리셋 일시 설정</h3>
                    <button onclick="closeDateTimePickerModal()" class="text-slate-400 hover:text-slate-600 font-bold text-xl">&times;</button>
                </div>

                <div class="space-y-4 mb-6">
                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">리셋 주기 선택</label>
                        <div class="grid grid-cols-2 gap-2">
                            <label class="flex items-center justify-center gap-1.5 p-2 border rounded-lg cursor-pointer hover:bg-slate-50 font-bold text-xs">
                                <input type="radio" name="dt-modal-type" value="once" checked onchange="onResetTypeChange()" class="text-blue-600">
                                <span>특정 날짜 (1회)</span>
                            </label>
                            <label class="flex items-center justify-center gap-1.5 p-2 border rounded-lg cursor-pointer hover:bg-slate-50 font-bold text-xs">
                                <input type="radio" name="dt-modal-type" value="daily" onchange="onResetTypeChange()" class="text-blue-600">
                                <span>매일 반복</span>
                            </label>
                        </div>
                    </div>

                    <div id="dt-modal-date-wrapper">
                        <label class="block text-xs font-bold text-slate-600 mb-1">날짜 선택</label>
                        <input type="date" id="dt-modal-date" class="w-full border border-slate-300 rounded-lg p-2.5 text-sm font-bold text-slate-800 outline-none focus:border-blue-500">
                    </div>

                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">시간 선택</label>
                        <div class="grid grid-cols-3 gap-2">
                            <select id="dt-modal-ampm" class="border border-slate-300 rounded-lg p-2 text-sm font-bold text-slate-800 bg-white cursor-pointer">
                                <option value="AM">오전</option>
                                <option value="PM" selected>오후</option>
                            </select>
                            <select id="dt-modal-hour" class="border border-slate-300 rounded-lg p-2 text-sm font-bold text-slate-800 bg-white cursor-pointer"></select>
                            <select id="dt-modal-minute" class="border border-slate-300 rounded-lg p-2 text-sm font-bold text-slate-800 bg-white cursor-pointer"></select>
                        </div>
                    </div>
                </div>

                <div class="flex justify-between items-center pt-3 border-t border-slate-100">
                    <button onclick="clearDateTimePickerModal()" class="px-3 py-2 bg-rose-50 text-rose-600 text-xs font-bold rounded-lg hover:bg-rose-100">예약 해제</button>
                    <div class="flex gap-2">
                        <button onclick="closeDateTimePickerModal()" class="px-3 py-2 bg-slate-200 text-slate-700 text-xs font-bold rounded-lg hover:bg-slate-300">취소</button>
                        <button onclick="confirmDateTimePickerModal()" class="px-5 py-2 bg-blue-600 text-white text-xs font-bold rounded-lg hover:bg-blue-700 shadow-sm">확인</button>
                    </div>
                </div>
            </div>
        </div>

        <script>
            let availableRooms = [];
            let targetConfigIdx = -1;
            
            let currentSortColumn = "username";
            let currentSortOrder = "asc";
            let selectedClassFilters = [];

            function switchTab(tab) {
                document.getElementById("tab-students").classList.add("hidden");
                document.getElementById("tab-classes").classList.add("hidden");
                document.getElementById("tab-qr").classList.add("hidden");

                document.getElementById("tab-btn-students").className = "w-full text-left px-3.5 py-2.5 text-slate-600 hover:bg-slate-50 font-bold rounded-xl transition";
                document.getElementById("tab-btn-classes").className = "w-full text-left px-3.5 py-2.5 text-slate-600 hover:bg-slate-50 font-bold rounded-xl transition";
                document.getElementById("tab-btn-qr").className = "w-full text-left px-3.5 py-2.5 text-slate-600 hover:bg-slate-50 font-bold rounded-xl transition";

                if (tab === 'students') {
                    document.getElementById("tab-students").classList.remove("hidden");
                    document.getElementById("tab-btn-students").className = "w-full text-left px-3.5 py-2.5 bg-blue-50 text-blue-700 font-bold rounded-xl transition";
                    fetchStudents();
                } else if (tab === 'classes') {
                    document.getElementById("tab-classes").classList.remove("hidden");
                    document.getElementById("tab-btn-classes").className = "w-full text-left px-3.5 py-2.5 bg-blue-50 text-blue-700 font-bold rounded-xl transition";
                    fetchClassConfigs();
                } else if (tab === 'qr') {
                    document.getElementById("tab-qr").classList.remove("hidden");
                    document.getElementById("tab-btn-qr").className = "w-full text-left px-3.5 py-2.5 bg-blue-50 text-blue-700 font-bold rounded-xl transition";
                    generateTodayQR();
                }
            }

            function toggleSort(column) {
                if (currentSortColumn === column) {
                    currentSortOrder = (currentSortOrder === "asc") ? "desc" : "asc";
                } else {
                    currentSortColumn = column;
                    currentSortOrder = "asc";
                }
                
                updateSortIcons();
                fetchStudents();
            }

            function updateSortIcons() {
                ["username", "name", "class_name"].forEach(col => {
                    const iconSpan = document.getElementById(`sort-icon-${col}`);
                    if (iconSpan) {
                        if (currentSortColumn === col) {
                            iconSpan.innerText = currentSortOrder === "asc" ? "▲" : "▼";
                            iconSpan.className = "text-blue-600 font-bold";
                        } else {
                            iconSpan.innerText = "↕";
                            iconSpan.className = "text-slate-400 font-normal";
                        }
                    }
                });
            }

            async function toggleClassFilterPanel() {
                const panel = document.getElementById("class-filter-panel");
                if (panel.classList.contains("hidden")) {
                    panel.classList.remove("hidden");
                    await loadClassFilterList();
                } else {
                    panel.classList.add("hidden");
                }
            }

            async function loadClassFilterList() {
                const res = await fetch("/api/admin/classes");
                const data = await res.json();
                const container = document.getElementById("class-filter-checkboxes");
                container.innerHTML = "";
                
                if (data.classes.length === 0) {
                    container.innerHTML = `<div class="text-xs text-slate-400 col-span-full py-2 text-center font-bold">등록된 수강 반이 없습니다.</div>`;
                    return;
                }
                
                data.classes.forEach(c => {
                    const isChecked = selectedClassFilters.includes(c.class_name);
                    container.innerHTML += `
                        <label class="flex items-center gap-2 cursor-pointer p-1.5 hover:bg-slate-50 rounded-lg border border-slate-100 transition">
                            <input type="checkbox" value="${c.class_name}" ${isChecked ? 'checked' : ''} onchange="onClassFilterChange()" class="class-filter-cb w-4 h-4 text-blue-600 rounded">
                            <span class="text-xs font-bold text-slate-800 truncate" title="${c.class_name}">${c.class_name}</span>
                        </label>
                    `;
                });
            }

            function onClassFilterChange() {
                const checkboxes = document.querySelectorAll('.class-filter-cb:checked');
                selectedClassFilters = Array.from(checkboxes).map(cb => cb.value);
                
                const activeBadge = document.getElementById("filter-active-count");
                if (selectedClassFilters.length > 0) {
                    activeBadge.innerText = selectedClassFilters.length;
                    activeBadge.classList.remove("hidden");
                } else {
                    activeBadge.classList.add("hidden");
                }
                
                fetchStudents();
            }

            function selectAllClassFilters(select) {
                const checkboxes = document.querySelectorAll('.class-filter-cb');
                checkboxes.forEach(cb => cb.checked = select);
                onClassFilterChange();
            }

            async function fetchStudents() {
                const keyword = document.getElementById("search-keyword").value;
                const classFilterStr = selectedClassFilters.join(",");
                const res = await fetch(`/api/admin/students?keyword=${encodeURIComponent(keyword)}&sort_by=${currentSortColumn}&order=${currentSortOrder}&class_filter=${encodeURIComponent(classFilterStr)}`);
                const data = await res.json();
                
                if (selectedClassFilters.length > 0) {
                    document.getElementById("summary-text").innerText = `총 등록 학생: ${data.total}명 | 선택한 반 학생: ${data.filtered_total}명`;
                } else {
                    document.getElementById("summary-text").innerText = `총 등록 학생: ${data.total}명`;
                }

                const tbody = document.getElementById("student-tbody");
                tbody.innerHTML = "";
                data.students.forEach(s => {
                    tbody.innerHTML += `
                        <tr class="hover:bg-slate-50 font-medium">
                            <td class="py-3 px-5 font-mono text-slate-700 font-bold">${s.username}</td>
                            <td class="py-3 px-5 font-bold text-slate-900">${s.name}</td>
                            <td class="py-3 px-5 text-slate-600 font-medium">${s.class_name}</td>
                            <td class="py-3 px-5 text-center">
                                <button onclick="resetPassword('${s.username}')" class="text-xs bg-slate-100 hover:bg-slate-200 px-2 py-1 rounded border font-bold">초기화</button>
                            </td>
                            <td class="py-3 px-5 text-center">
                                <button onclick="deleteStudent('${s.username}')" class="text-xs text-red-500 hover:text-red-700 font-bold">삭제</button>
                            </td>
                        </tr>
                    `;
                });
            }

            async function openAddStudentModal() {
                document.getElementById("add-username").value = "";
                document.getElementById("add-name").value = "";
                
                const res = await fetch("/api/admin/classes");
                const data = await res.json();
                
                const classContainer = document.getElementById("add-class-options");
                classContainer.innerHTML = "";
                
                if (data.classes.length === 0) {
                    classContainer.innerHTML = `<div class="text-xs text-slate-400 py-2 text-center font-bold">등록된 수강 반이 없습니다. 파일 업로드를 먼저 해주세요.</div>`;
                } else {
                    data.classes.forEach((c, idx) => {
                        classContainer.innerHTML += `
                            <label class="flex items-center gap-2 cursor-pointer p-1 hover:bg-white rounded transition">
                                <input type="checkbox" name="add-student-class" value="${c.class_name}" class="w-4 h-4 text-blue-600 rounded">
                                <span class="text-xs font-bold text-slate-800">${c.class_name}</span>
                            </label>
                        `;
                    });
                }

                document.getElementById("add-student-modal").classList.replace("hidden", "flex");
            }

            function closeAddStudentModal() {
                document.getElementById("add-student-modal").classList.replace("flex", "hidden");
            }

            async function submitAddStudent() {
                const username = document.getElementById("add-username").value.trim();
                const name = document.getElementById("add-name").value.trim();
                
                const checkboxes = document.querySelectorAll('input[name="add-student-class"]:checked');
                const selectedClasses = Array.from(checkboxes).map(cb => cb.value);

                if (!username || !name) {
                    alert("아이디와 이름을 모두 입력해 주세요.");
                    return;
                }

                if (selectedClasses.length === 0) {
                    alert("최소 하나 이상의 수강 반을 선택해 주세요.");
                    return;
                }

                const res = await fetch("/api/admin/students/add-manual", {
                    method: "POST",
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: `username=${encodeURIComponent(username)}&name=${encodeURIComponent(name)}&class_names=${encodeURIComponent(selectedClasses.join(', '))}`
                });

                const data = await res.json();
                alert(data.message);
                if (data.success) {
                    closeAddStudentModal();
                    fetchStudents();
                }
            }

            function formatResetDisplay(dtStr) {
                if (!dtStr) return "일시 미설정";
                
                if (dtStr.startsWith("DAILY:")) {
                    const tStr = dtStr.replace("DAILY:", "");
                    let [h, m] = tStr.split(":");
                    let hNum = parseInt(h, 10);
                    const ampm = hNum >= 12 ? "오후" : "오전";
                    if (hNum > 12) hNum -= 12;
                    if (hNum === 0) hNum = 12;
                    const hDisplay = hNum < 10 ? '0' + hNum : '' + hNum;
                    return `매일 ${ampm} ${hDisplay}:${m}`;
                }

                if (dtStr.includes("T")) {
                    const [dStr, tStr] = dtStr.split("T");
                    let [h, m] = tStr.split(":");
                    let hNum = parseInt(h, 10);
                    const ampm = hNum >= 12 ? "오후" : "오전";
                    if (hNum > 12) hNum -= 12;
                    if (hNum === 0) hNum = 12;
                    const hDisplay = hNum < 10 ? '0' + hNum : '' + hNum;
                    return `${dStr} ${ampm} ${hDisplay}:${m}`;
                }

                return "일시 미설정";
            }

            async function fetchClassConfigs() {
                const res = await fetch("/api/admin/classes");
                const data = await res.json();
                availableRooms = data.rooms;

                const tbody = document.getElementById("class-config-tbody");
                tbody.innerHTML = "";

                data.classes.forEach((c, idx) => {
                    let options = availableRooms.map(r => `<option value="${r.room_name}" ${r.room_name === c.room_name ? 'selected' : ''}>${r.room_name}</option>`).join("");
                    if (!options) options = `<option value="${c.room_name}">${c.room_name}</option>`;

                    const unreservedBadge = c.unreserved_count > 0 
                        ? `<button onclick="showUnreserved('${c.class_name}')" class="px-2 py-1 bg-rose-50 text-rose-600 border border-rose-200 rounded-full text-xs font-extrabold hover:bg-rose-100 transition shadow-sm">${c.unreserved_count}명 🔍</button>`
                        : `<span class="px-2 py-0.5 bg-slate-100 text-slate-400 rounded-full text-xs font-bold">0명</span>`;

                    const displayDtText = formatResetDisplay(c.reset_datetime);

                    tbody.innerHTML += `
                        <tr class="hover:bg-slate-50 font-medium" id="row-${idx}">
                            <td class="py-3.5 px-4 font-bold text-slate-800 text-xs">${c.class_name}</td>
                            <td class="py-3.5 px-1 text-center font-bold text-slate-700 text-xs">${c.total_students}명</td>
                            <td class="py-3.5 px-1 text-center"><span class="px-1.5 py-0.5 bg-blue-50 text-blue-700 border border-blue-200 rounded-full text-xs font-bold">${c.reserved_count}명</span></td>
                            <td class="py-3.5 px-1 text-center">${unreservedBadge}</td>
                            <td class="py-3.5 px-1 text-center">
                                <input type="number" id="rows-${idx}" value="${c.rows_count}" min="1" max="30" class="border rounded px-1 py-1 text-xs w-12 text-center font-bold text-blue-700 bg-blue-50/30">
                            </td>
                            <td class="py-3.5 px-1 text-center">
                                <input type="number" id="cols-${idx}" value="${c.cols_count}" min="1" max="30" class="border rounded px-1 py-1 text-xs w-12 text-center font-bold text-blue-700 bg-blue-50/30">
                            </td>
                            <td class="py-3.5 px-2">
                                <select id="room-${idx}" onchange="onRoomChange(${idx})" class="border rounded px-1.5 py-1 text-xs w-full bg-white font-bold cursor-pointer">
                                    ${options}
                                </select>
                            </td>
                            <td class="py-3.5 px-2 text-center">
                                <input type="hidden" id="reset-dt-${idx}" value="${c.reset_datetime || ''}">
                                <input type="text" id="reset-dt-display-${idx}" value="${displayDtText}" readonly onclick="openDateTimePickerModal(${idx}, '${c.class_name}')" class="border border-slate-300 rounded px-2 py-1 text-xs font-bold text-slate-700 bg-white hover:bg-blue-50/50 cursor-pointer w-44 text-center transition shadow-xs" title="클릭하여 리셋 일시 설정">
                            </td>
                            <td class="py-3.5 px-1 text-center">
                                <button onclick="saveClassConfig('${c.class_name}', ${idx})" id="save-btn-${idx}" class="text-xs bg-blue-600 text-white px-2.5 py-1 rounded-lg hover:bg-blue-700 font-bold shadow-sm transition">저장</button>
                            </td>
                            <td class="py-3.5 px-1 text-center">
                                <button onclick="resetClassNow('${c.class_name}')" class="text-xs bg-rose-500 text-white px-2 py-1 rounded-lg hover:bg-rose-600 font-bold transition shadow-sm">초기화</button>
                            </td>
                            <td class="py-3.5 px-1 text-center">
                                <a href="/admin/print-seating-chart?class_name=${encodeURIComponent(c.class_name)}" target="_blank" class="inline-block text-xs bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-2 py-1 rounded-lg shadow-sm transition">
                                    인쇄
                                </a>
                            </td>
                        </tr>
                    `;
                });
            }

            function initDateTimeOptions() {
                const hourSelect = document.getElementById("dt-modal-hour");
                hourSelect.innerHTML = "";
                for(let h=1; h<=12; h++) {
                    const val = h < 10 ? '0' + h : '' + h;
                    hourSelect.innerHTML += `<option value="${val}">${val}시</option>`;
                }
                
                const minSelect = document.getElementById("dt-modal-minute");
                minSelect.innerHTML = "";
                for(let m=0; m<60; m+=5) {
                    const val = m < 10 ? '0' + m : '' + m;
                    minSelect.innerHTML += `<option value="${val}">${val}분</option>`;
                }
            }

            function onResetTypeChange() {
                const selectedType = document.querySelector('input[name="dt-modal-type"]:checked').value;
                const dateWrapper = document.getElementById("dt-modal-date-wrapper");
                if (selectedType === 'daily') {
                    dateWrapper.style.display = 'none';
                } else {
                    dateWrapper.style.display = 'block';
                }
            }

            function openDateTimePickerModal(idx, className) {
                targetConfigIdx = idx;
                initDateTimeOptions();
                
                document.getElementById("dt-modal-title").innerText = `[${className}] 자동 리셋 일시 설정`;
                
                const currentVal = document.getElementById(`reset-dt-${idx}`).value;
                const dateInput = document.getElementById("dt-modal-date");
                const ampmSelect = document.getElementById("dt-modal-ampm");
                const hourSelect = document.getElementById("dt-modal-hour");
                const minSelect = document.getElementById("dt-modal-minute");
                const typeRadios = document.getElementsByName("dt-modal-type");
                
                if (currentVal && currentVal.startsWith("DAILY:")) {
                    typeRadios[1].checked = true; // 매일
                    const tStr = currentVal.replace("DAILY:", "");
                    let [h, m] = tStr.split(":").map(Number);
                    if (h >= 12) {
                        ampmSelect.value = "PM";
                        if (h > 12) h -= 12;
                    } else {
                        ampmSelect.value = "AM";
                        if (h === 0) h = 12;
                    }
                    const hStr = h < 10 ? '0' + h : '' + h;
                    const mRounded = Math.floor(m / 5) * 5;
                    const mStr = mRounded < 10 ? '0' + mRounded : '' + mRounded;
                    hourSelect.value = hStr;
                    minSelect.value = mStr;
                } else if (currentVal && currentVal.includes("T")) {
                    typeRadios[0].checked = true; // 특정 날짜
                    const [dStr, tStr] = currentVal.split("T");
                    dateInput.value = dStr;
                    
                    let [h, m] = tStr.split(":").map(Number);
                    if (h >= 12) {
                        ampmSelect.value = "PM";
                        if (h > 12) h -= 12;
                    } else {
                        ampmSelect.value = "AM";
                        if (h === 0) h = 12;
                    }
                    
                    const hStr = h < 10 ? '0' + h : '' + h;
                    const mRounded = Math.floor(m / 5) * 5;
                    const mStr = mRounded < 10 ? '0' + mRounded : '' + mRounded;
                    
                    hourSelect.value = hStr;
                    minSelect.value = mStr;
                } else {
                    typeRadios[0].checked = true;
                    const now = new Date();
                    const yyyy = now.getFullYear();
                    const mm = String(now.getMonth() + 1).padStart(2, '0');
                    const dd = String(now.getDate()).padStart(2, '0');
                    dateInput.value = `${yyyy}-${mm}-${dd}`;
                    ampmSelect.value = "PM";
                    hourSelect.value = "12";
                    minSelect.value = "00";
                }
                
                onResetTypeChange();
                document.getElementById("datetime-picker-modal").classList.replace("hidden", "flex");
            }

            function closeDateTimePickerModal() {
                document.getElementById("datetime-picker-modal").classList.replace("flex", "hidden");
            }

            function clearDateTimePickerModal() {
                if (targetConfigIdx >= 0) {
                    document.getElementById(`reset-dt-${targetConfigIdx}`).value = "";
                    document.getElementById(`reset-dt-display-${targetConfigIdx}`).value = "일시 미설정";
                }
                closeDateTimePickerModal();
            }

            function confirmDateTimePickerModal() {
                if (targetConfigIdx < 0) return;
                
                const selectedType = document.querySelector('input[name="dt-modal-type"]:checked').value;
                const ampm = document.getElementById("dt-modal-ampm").value;
                let h = parseInt(document.getElementById("dt-modal-hour").value, 10);
                const m = document.getElementById("dt-modal-minute").value;
                
                if (ampm === "PM" && h < 12) h += 12;
                if (ampm === "AM" && h === 12) h = 0;
                
                const hStr = h < 10 ? '0' + h : '' + h;
                const displayAmpm = ampm === "PM" ? "오후" : "오전";
                const displayHour = document.getElementById("dt-modal-hour").value;

                if (selectedType === 'daily') {
                    const saveVal = `DAILY:${hStr}:${m}`;
                    document.getElementById(`reset-dt-${targetConfigIdx}`).value = saveVal;
                    document.getElementById(`reset-dt-display-${targetConfigIdx}`).value = `매일 ${displayAmpm} ${displayHour}:${m}`;
                } else {
                    const dVal = document.getElementById("dt-modal-date").value;
                    if (!dVal) {
                        alert("날짜를 선택해 주세요.");
                        return;
                    }
                    const saveVal = `${dVal}T${hStr}:${m}`;
                    document.getElementById(`reset-dt-${targetConfigIdx}`).value = saveVal;
                    document.getElementById(`reset-dt-display-${targetConfigIdx}`).value = `${dVal} ${displayAmpm} ${displayHour}:${m}`;
                }
                
                closeDateTimePickerModal();
            }

            function onRoomChange(idx) {
                const selectedRoomName = document.getElementById(`room-${idx}`).value;
                const targetRoom = availableRooms.find(r => r.room_name === selectedRoomName);
                if (targetRoom) {
                    document.getElementById(`rows-${idx}`).value = targetRoom.rows_count || 0;
                    document.getElementById(`cols-${idx}`).value = targetRoom.cols_count || 0;
                }
            }

            async function showUnreserved(className) {
                const res = await fetch(`/api/admin/unreserved-students?class_name=${encodeURIComponent(className)}`);
                const data = await res.json();

                document.getElementById("unreserved-modal-title").innerText = `[${className}] 미신청 학생`;
                document.getElementById("unreserved-modal-count").innerText = `총 ${data.unreserved.length}명이 아직 좌석을 선택하지 않았습니다.`;

                const listContainer = document.getElementById("unreserved-student-list");
                listContainer.innerHTML = "";

                if (data.unreserved.length === 0) {
                    listContainer.innerHTML = `<div class="text-center py-6 text-slate-400 text-sm font-bold">모든 학생이 좌석 신청을 완료했습니다!</div>`;
                } else {
                    data.unreserved.forEach(s => {
                        const idDisplay = s.is_duplicate_name 
                            ? `<span class="font-mono text-xs text-rose-600 bg-rose-50 px-2 py-0.5 rounded border border-rose-200 font-bold">${s.username}</span>`
                            : ``;

                        listContainer.innerHTML += `
                            <div class="flex justify-between items-center p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm">
                                <span class="font-bold text-slate-800">${s.name}</span>
                                ${idDisplay}
                            </div>
                        `;
                    });
                }

                document.getElementById("unreserved-modal").classList.replace("hidden", "flex");
            }

            function closeUnreservedModal() {
                document.getElementById("unreserved-modal").classList.replace("flex", "hidden");
            }

            async function saveClassConfig(className, idx) {
                const saveBtn = document.getElementById(`save-btn-${idx}`);
                const origText = saveBtn.innerText;
                saveBtn.disabled = true;
                saveBtn.innerText = "저장중...";

                const room = document.getElementById(`room-${idx}`).value;
                const rows = document.getElementById(`rows-${idx}`).value || 0;
                const cols = document.getElementById(`cols-${idx}`).value || 0;
                const resetDt = document.getElementById(`reset-dt-${idx}`).value || "";

                try {
                    const res = await fetch("/api/admin/classes/update", {
                        method: "POST",
                        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                        body: `class_name=${encodeURIComponent(className)}&room_name=${encodeURIComponent(room)}&rows_count=${rows}&cols_count=${cols}&reset_datetime=${encodeURIComponent(resetDt)}`
                    });
                    const result = await res.json();

                    if (result.success) {
                        saveBtn.innerText = "✓ 완료";
                        saveBtn.className = "text-xs bg-emerald-600 text-white px-2.5 py-1 rounded-lg font-bold shadow-sm transition";
                        setTimeout(() => {
                            saveBtn.innerText = "저장";
                            saveBtn.className = "text-xs bg-blue-600 text-white px-2.5 py-1 rounded-lg hover:bg-blue-700 font-bold shadow-sm transition";
                            saveBtn.disabled = false;
                        }, 1200);
                    } else {
                        alert(result.message);
                        saveBtn.innerText = origText;
                        saveBtn.disabled = false;
                    }
                } catch (e) {
                    alert("저장 중 오류가 발생했습니다.");
                    saveBtn.innerText = origText;
                    saveBtn.disabled = false;
                }
            }

            async function resetClassNow(className) {
                if (!confirm(`[${className}] 반의 현재 모든 좌석 신청 내역을 즉시 초기화하시겠습니까?`)) return;
                const res = await fetch("/api/admin/classes/reset-now", {
                    method: "POST",
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: `class_name=${encodeURIComponent(className)}`
                });
                const result = await res.json();
                alert(result.message);
                fetchClassConfigs();
            }

            async function uploadStudentFile() {
                const fileInput = document.getElementById("upload-student-file");
                if (!fileInput.files[0]) return;
                const formData = new FormData();
                formData.append("file", fileInput.files[0]);

                document.getElementById("summary-text").innerText = "명단 분석 및 등록 중...";
                const res = await fetch("/api/admin/students/upload", { method: "POST", body: formData });
                const result = await res.json();
                alert(result.message);
                fileInput.value = "";
                fetchStudents();
            }

            async function uploadRoomFile() {
                const fileInput = document.getElementById("upload-room-file");
                if (!fileInput.files[0]) return;
                const formData = new FormData();
                formData.append("file", fileInput.files[0]);

                const res = await fetch("/api/admin/rooms/upload", { method: "POST", body: formData });
                const result = await res.json();
                alert(result.message);
                fileInput.value = "";
                fetchClassConfigs();
            }

            async function generateTodayQR() {
                const res = await fetch("/api/admin/qr-info");
                const data = await res.json();

                document.getElementById("qr-date-title").innerText = `좌석표 신청 (${data.formatted_date})`;
                
                const currentHost = window.location.origin;
                const targetUrl = `${currentHost}/?date=${data.date}&token=${data.token}`;
                
                const qrContainer = document.getElementById("qrcode");
                qrContainer.innerHTML = "";
                new QRCode(qrContainer, {
                    text: targetUrl,
                    width: 180,
                    height: 180,
                    colorDark : "#0f172a",
                    colorLight : "#ffffff",
                    correctLevel : QRCode.CorrectLevel.H
                });
            }

            async function resetPassword(username) {
                if (!confirm(`${username} 학생의 비밀번호를 '1234'로 초기화하시겠습니까?`)) return;
                const res = await fetch(`/api/admin/students/${encodeURIComponent(username)}/reset-pw`, { method: 'POST' });
                const result = await res.json();
                alert(result.message);
            }

            async function deleteStudent(username) {
                if (!confirm(`${username} 학생을 삭제하시겠습니까?`)) return;
                await fetch(`/api/admin/students/${encodeURIComponent(username)}`, { method: 'DELETE' });
                fetchStudents();
            }

            window.onload = function() {
                updateSortIcons();
                fetchStudents();
            };
        </script>
    </body>
    </html>
    """

# ==============================================================================
# 3. 당일 신청 좌석표 인쇄 / PDF 다운로드 전용 뷰어 화면
# ==============================================================================
@app.get("/admin/print-seating-chart", response_class=HTMLResponse)
def print_seating_chart(class_name: str = ""):
    conn = get_db_connection()
    check_and_apply_resets(conn)
    cursor = conn.cursor()
    
    cursor.execute("SELECT room_name, rows_count, cols_count FROM class_configs WHERE class_name = ?", (class_name,))
    config = cursor.fetchone()
    room_name = config[0] if config else "301호"
    custom_rows = config[1] if config else 0
    custom_cols = config[2] if config else 0
    
    cursor.execute("SELECT title, rows_count, cols_count, grid_json FROM rooms WHERE room_name = ?", (room_name,))
    room_data = cursor.fetchone()
    
    if room_data:
        title, orig_rows, orig_cols, grid = room_data[0], room_data[1], room_data[2], json.loads(room_data[3])
    else:
        title = room_name
        grid = [[{"type": "seat", "id": f"{chr(65+r)}{c+1}"} for c in range(4)] for r in range(5)]
        
    if custom_rows > 0: grid = grid[:custom_rows]
    if custom_cols > 0: grid = [row[:custom_cols] for row in grid]

    cursor.execute("SELECT seat_id, name, username FROM seat_reservations WHERE class_name = ?", (class_name,))
    res_rows = cursor.fetchall()
    seat_to_name = {r[0]: r[1] for r in res_rows}
    reserved_user_set = set([r[2] for r in res_rows])

    cursor.execute("SELECT username, name, class_name FROM students")
    all_students = cursor.fetchall()
    conn.close()

    total_class_students = []
    unreserved_students = []
    for u_id, u_name, c_str in all_students:
        c_list = [c.strip() for c in c_str.split(',') if c.strip()]
        if class_name in c_list:
            total_class_students.append((u_id, u_name))
            if u_id not in reserved_user_set:
                unreserved_students.append((u_id, u_name))

    name_counts = Counter([u[1] for u in total_class_students])

    total_count = len(total_class_students)
    reserved_count = len(reserved_user_set)
    unreserved_count = len(unreserved_students)

    now = datetime.now()
    formatted_date = f"{now.year}년 {now.month}월 {now.day}일"

    table_rows_html = ""
    for row in grid:
        table_rows_html += "<tr>"
        for cell in row:
            if cell.get("type") == "seat":
                s_id = cell["id"]
                s_name = seat_to_name.get(s_id, "")
                
                name_display = f'<div style="font-size: 15px; font-weight: 800; color: #1e3a8a; margin-top: 4px;">{s_name}</div>' if s_name else '<div style="font-size: 11px; color: #94a3b8; margin-top: 6px;">| &nbsp;&nbsp;&nbsp; |</div>'
                bg_color = "background-color: #eff6ff;" if s_name else "background-color: #ffffff;"
                
                table_rows_html += f'''
                    <td style="border: 1px solid #475569; width: 85px; height: 58px; text-align: center; vertical-align: top; padding: 4px; {bg_color}">
                        <div style="font-size: 12px; font-weight: 700; color: #334155;">{s_id}</div>
                        {name_display}
                    </td>
                '''
            elif cell.get("type") == "aisle":
                table_rows_html += f'''
                    <td style="border-left: 1px dashed #94a3b8; border-right: 1px dashed #94a3b8; background-color: #f8fafc; font-size: 13px; font-weight: 800; color: #64748b; text-align: center; width: 35px;">
                        {cell["val"]}
                    </td>
                '''
            else:
                table_rows_html += '''
                    <td style="border: 1px solid #cbd5e1; background: linear-gradient(to top right, transparent calc(50% - 1px), #cbd5e1, transparent calc(50% + 1px)), linear-gradient(to bottom right, transparent calc(50% - 1px), #cbd5e1, transparent calc(50% + 1px)); width: 85px; height: 58px;"></td>
                '''
        table_rows_html += "</tr>"

    if unreserved_students:
        badges = []
        for u_id, u_name in unreserved_students:
            display_text = f"{u_name} ({u_id})" if name_counts[u_name] > 1 else u_name
            badges.append(f'<span style="display: inline-block; background: #fff; border: 1px solid #fecdd3; color: #e11d48; font-size: 11px; font-weight: bold; padding: 3px 8px; border-radius: 6px; margin: 2px 4px;">{display_text}</span>')
        unres_badges = "".join(badges)
    else:
        unres_badges = '<span style="color: #10b981; font-size: 12px; font-weight: bold;">모든 학생이 좌석 신청을 완료했습니다.</span>'

    return f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <title>[김영편입 좌석표] {class_name} ({formatted_date})</title>
        <style>
            @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
            * {{ box-sizing: border-box; font-family: 'Pretendard', sans-serif; }}
            body {{ margin: 0; padding: 20px; background-color: #f1f5f9; display: flex; flex-direction: column; align-items: center; }}
            .paper {{ background: white; width: 210mm; min-height: 297mm; padding: 15mm; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); border-radius: 8px; display: flex; flex-direction: column; justify-content: space-between; }}
            @media print {{
                body {{ background: white; padding: 0; }}
                .paper {{ width: 100%; box-shadow: none; padding: 0; min-height: 100vh; }}
                .no-print {{ display: none !important; }}
            }}
        </style>
    </head>
    <body>
        <div class="no-print" style="margin-bottom: 20px; display: flex; gap: 10px;">
            <button onclick="window.print()" style="background-color: #2563eb; color: white; border: none; padding: 10px 24px; border-radius: 8px; font-weight: 700; cursor: pointer; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">A4 인쇄 / PDF 다운로드</button>
            <button onclick="window.close()" style="background-color: #64748b; color: white; border: none; padding: 10px 18px; border-radius: 8px; font-weight: 600; cursor: pointer; font-size: 14px;">닫기</button>
        </div>

        <div class="paper">
            <div>
                <div style="background-color: #f59e0b; color: white; padding: 16px 20px; border-radius: 6px 6px 0 0; display: flex; justify-content: space-between; align-items: center;">
                    <div style="font-size: 22px; font-weight: 900;">[{class_name}] 좌석 배치표</div>
                    <div style="font-size: 15px; font-weight: 700; background: rgba(0,0,0,0.15); padding: 4px 10px; border-radius: 4px;">{room_name} ({formatted_date})</div>
                </div>

                <div style="display: flex; background-color: #cbd5e1; border: 1px solid #94a3b8; border-top: none; font-weight: 800; font-size: 13px;">
                    <div style="flex: 1; text-align: center; padding: 8px; letter-spacing: 12px; color: #1e293b;">칠 &nbsp;&nbsp; 판</div>
                    <div style="background-color: #475569; color: white; padding: 8px 16px;">출입문</div>
                </div>

                <div style="margin-top: 15px; display: flex; justify-content: center;">
                    <table style="border-collapse: collapse; margin: 0 auto;">
                        {table_rows_html}
                    </table>
                </div>
            </div>

            <div style="margin-top: 20px;">
                <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 16px; display: flex; justify-content: space-between; align-items: center; font-size: 13px; font-weight: bold;">
                    <span style="color: #475569;">김영편입 좌석표 신청</span>
                    <div style="display: flex; gap: 16px;">
                        <span style="color: #334155;">총원: <b>{total_count}명</b></span>
                        <span style="color: #2563eb;">신청: <b>{reserved_count}명</b></span>
                        <span style="color: #e11d48;">미신청: <b>{unreserved_count}명</b></span>
                    </div>
                </div>

                <div style="margin-top: 8px; background-color: #fff1f2; border: 1px solid #ffe4e6; border-radius: 8px; padding: 10px 14px;">
                    <div style="font-size: 12px; font-weight: 800; color: #be123c; margin-bottom: 6px;">미신청 학생 명단 ({unreserved_count}명)</div>
                    <div style="line-height: 1.8;">
                        {unres_badges}
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """

# ==============================================================================
# 4. 백엔드 API
# ==============================================================================
@app.get("/api/admin/qr-info")
def get_qr_info():
    today_str = get_today_str()
    token = generate_daily_token(today_str)
    now = datetime.now()
    formatted = f"{now.month}월 {now.day}일"
    return {
        "date": today_str,
        "token": token,
        "formatted_date": formatted
    }

@app.post("/api/login")
def api_login(username: str = Form(...), password: str = Form(...)):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name, class_name FROM students WHERE username = ? AND password = ?", (username, password))
    user = cursor.fetchone()
    conn.close()
    if user:
        raw_classes = [c.strip() for c in user[1].split(",") if c.strip() and c.strip() != "-"]
        return {"success": True, "username": username, "name": user[0], "class_list": raw_classes if raw_classes else ["기본반"]}
    return {"success": False, "message": "아이디 또는 비밀번호가 일치하지 않습니다."}

@app.post("/api/change-password")
def api_change_password(username: str = Form(...), current_password: str = Form(...), new_password: str = Form(...)):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT username FROM students WHERE username = ? AND password = ?", (username, current_password))
    user = cursor.fetchone()
    if not user:
        conn.close()
        return {"success": False, "message": "현재 비밀번호가 일치하지 않습니다."}
        
    cursor.execute("UPDATE students SET password = ? WHERE username = ?", (new_password, username))
    conn.commit()
    conn.close()
    return {"success": True, "message": "비밀번호가 성공적으로 변경되었습니다."}

@app.post("/api/admin/students/{username}/reset-pw")
def api_reset_password(username: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE students SET password = '1234' WHERE username = ?", (username,))
    conn.commit()
    conn.close()
    return {"success": True, "message": f"{username} 학생의 비밀번호가 1234로 초기화되었습니다."}

@app.get("/api/seats")
def api_seats(username: str, class_name: str):
    conn = get_db_connection()
    check_and_apply_resets(conn)
    cursor = conn.cursor()
    
    cursor.execute("SELECT room_name, rows_count, cols_count FROM class_configs WHERE class_name = ?", (class_name,))
    config = cursor.fetchone()
    if config:
        room_name, custom_rows, custom_cols = config[0], config[1], config[2]
    else:
        room_name, custom_rows, custom_cols = "301호", 0, 0
    
    cursor.execute("SELECT title, rows_count, cols_count, grid_json FROM rooms WHERE room_name = ?", (room_name,))
    room_data = cursor.fetchone()
    
    if room_data:
        title, orig_rows, orig_cols, grid = room_data[0], room_data[1], room_data[2], json.loads(room_data[3])
        cols = orig_cols
    else:
        title, cols = room_name, 4
        grid = [[{"type": "seat", "id": f"{chr(65+r)}{c+1}"} for c in range(4)] for r in range(5)]
        
    if custom_rows > 0: grid = grid[:custom_rows]
    if custom_cols > 0:
        cols = custom_cols
        grid = [row[:custom_cols] for row in grid]

    cursor.execute("SELECT seat_id, username FROM seat_reservations WHERE class_name = ?", (class_name,))
    reservations = dict(cursor.fetchall())
    conn.close()

    total_seats = 0
    for r in grid:
        for cell in r:
            if cell.get("type") == "seat":
                total_seats += 1
                assigned_user = reservations.get(cell["id"])
                if assigned_user == username:
                    cell["status_class"] = "mine"
                elif assigned_user:
                    cell["status_class"] = "occupied"
                else:
                    cell["status_class"] = "empty"
                    
    return {"room_name": room_name, "room_title": title, "cols": cols, "total_seats": total_seats, "grid": grid}

@app.post("/api/reserve")
def api_reserve(username: str = Form(...), name: str = Form(...), class_name: str = Form(...), seat_id: str = Form(...)):
    conn = get_db_connection()
    check_and_apply_resets(conn)
    cursor = conn.cursor()
    
    cursor.execute("SELECT username FROM seat_reservations WHERE class_name = ? AND seat_id = ?", (class_name, seat_id))
    target = cursor.fetchone()
    if target and target[0] and target[0] != username:
        conn.close()
        return {"success": False, "message": "이미 다른 학생이 선택한 좌석입니다!"}
        
    cursor.execute("DELETE FROM seat_reservations WHERE class_name = ? AND username = ?", (class_name, username))
    cursor.execute("INSERT INTO seat_reservations (class_name, seat_id, username, name) VALUES (?, ?, ?, ?)", (class_name, seat_id, username, name))
    conn.commit()
    conn.close()
    return {"success": True}

@app.get("/api/admin/students")
def get_students(keyword: str = "", sort_by: str = "username", order: str = "asc", class_filter: str = ""):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    valid_cols = {"username": "username", "name": "name", "class_name": "class_name"}
    col_sql = valid_cols.get(sort_by, "username")
    order_sql = "DESC" if order.lower() == "desc" else "ASC"

    query = f"SELECT username, name, class_name FROM students WHERE 1=1"
    params = []
    if keyword:
        query += " AND (username LIKE ? OR name LIKE ? OR class_name LIKE ?)"
        params.extend([f"%{keyword}%", f"%{keyword}%", f"%{keyword}%"])
        
    query += f" ORDER BY {col_sql} {order_sql}"

    cursor.execute(query, params)
    rows = cursor.fetchall()
    
    cursor.execute("SELECT COUNT(*) FROM students")
    total = cursor.fetchone()[0]
    conn.close()

    filter_classes = [c.strip() for c in class_filter.split(",") if c.strip()]

    student_list = []
    for r in rows:
        u_id, u_name, c_str = r[0], r[1], r[2]
        student_c_list = [c.strip() for c in c_str.split(",") if c.strip()]
        
        if filter_classes:
            if not any(fc in student_c_list for fc in filter_classes):
                continue
                
        student_list.append({"username": u_id, "name": u_name, "class_name": c_str})

    filtered_total = len(student_list)

    return {
        "total": total, 
        "filtered_total": filtered_total, 
        "students": student_list
    }

@app.post("/api/admin/students/add-manual")
def add_student_manual(username: str = Form(...), name: str = Form(...), class_names: str = Form(...)):
    username = username.strip()
    name = name.strip()
    class_names = class_names.strip()

    if not username or not name:
        return {"success": False, "message": "아이디와 이름을 모두 입력해 주세요."}

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT username FROM students WHERE username = ?", (username,))
    if cursor.fetchone():
        conn.close()
        return {"success": False, "message": f"이미 존재하는 아이디({username})입니다."}

    cursor.execute("""
        INSERT INTO students (username, name, class_name, password, status)
        VALUES (?, ?, ?, '1234', 'active')
    """, (username, name, class_names))

    for c in class_names.split(','):
        c_clean = c.strip()
        if c_clean and c_clean != '-':
            cursor.execute("INSERT OR IGNORE INTO class_configs (class_name, room_name, rows_count, cols_count) VALUES (?, '301호', 0, 0)", (c_clean,))

    conn.commit()
    conn.close()

    return {"success": True, "message": f"학생 [{name}({username})]이 성공적으로 추가되었습니다."}

@app.delete("/api/admin/students/{username}")
def delete_student(username: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM students WHERE username = ?", (username,))
    cursor.execute("DELETE FROM seat_reservations WHERE username = ?", (username,))
    conn.commit()
    conn.close()
    return {"success": True}

@app.get("/api/admin/classes")
def get_classes():
    conn = get_db_connection()
    check_and_apply_resets(conn)
    cursor = conn.cursor()
    
    cursor.execute("SELECT room_name, title, rows_count, cols_count FROM rooms ORDER BY room_name")
    rooms = cursor.fetchall()
    rooms_dict = {r[0]: {"rows": r[2], "cols": r[3]} for r in rooms}
    
    cursor.execute("SELECT class_name, room_name, rows_count, cols_count, reset_datetime FROM class_configs ORDER BY class_name")
    classes = cursor.fetchall()
    
    cursor.execute("SELECT username, name, class_name FROM students")
    all_students = cursor.fetchall()
    
    cursor.execute("SELECT class_name, username FROM seat_reservations")
    all_reservations = cursor.fetchall()
    
    conn.close()

    class_students_map = {}
    for u_id, u_name, c_str in all_students:
        for c in c_str.split(','):
            c_clean = c.strip()
            if c_clean and c_clean != '-':
                if c_clean not in class_students_map:
                    class_students_map[c_clean] = []
                class_students_map[c_clean].append((u_id, u_name))
                
    class_reserved_map = {}
    for c_name, u_id in all_reservations:
        if c_name not in class_reserved_map:
            class_reserved_map[c_name] = set()
        class_reserved_map[c_name].add(u_id)

    class_list = []
    for c in classes:
        c_name, r_name, r_cnt, c_cnt, reset_dt = c[0], c[1], c[2], c[3], c[4]
        if (r_cnt == 0 or c_cnt == 0) and r_name in rooms_dict:
            r_cnt = rooms_dict[r_name]["rows"]
            c_cnt = rooms_dict[r_name]["cols"]
            
        st_list = class_students_map.get(c_name, [])
        total_st = len(st_list)
        reserved_set = class_reserved_map.get(c_name, set())
        reserved_st = len([s for s in st_list if s[0] in reserved_set])
        unreserved_st = total_st - reserved_st
        
        class_list.append({
            "class_name": c_name,
            "room_name": r_name,
            "rows_count": r_cnt or 0,
            "cols_count": c_cnt or 0,
            "reset_datetime": reset_dt or "",
            "total_students": total_st,
            "reserved_count": reserved_st,
            "unreserved_count": max(0, unreserved_st)
        })
        
    return {
        "classes": class_list,
        "rooms": [{"room_name": r[0], "title": r[1], "rows_count": r[2], "cols_count": r[3]} for r in rooms]
    }

@app.get("/api/admin/unreserved-students")
def get_unreserved_students(class_name: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT username, name, class_name FROM students")
    all_students = cursor.fetchall()
    
    cursor.execute("SELECT username FROM seat_reservations WHERE class_name = ?", (class_name,))
    reserved_users = set([r[0] for r in cursor.fetchall()])
    conn.close()

    total_class_students = []
    unreserved_list = []
    for u_id, u_name, c_str in all_students:
        c_list = [c.strip() for c in c_str.split(',') if c.strip()]
        if class_name in c_list:
            total_class_students.append(u_name)
            if u_id not in reserved_users:
                unreserved_list.append({"username": u_id, "name": u_name})
                
    name_counts = Counter(total_class_students)
    for item in unreserved_list:
        item["is_duplicate_name"] = name_counts[item["name"]] > 1
            
    return {"class_name": class_name, "unreserved": unreserved_list}

@app.post("/api/admin/classes/update")
def update_class_config(
    class_name: str = Form(...), 
    room_name: str = Form(...), 
    rows_count: int = Form(0), 
    cols_count: int = Form(0),
    reset_datetime: str = Form("")
):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO class_configs (class_name, room_name, rows_count, cols_count, reset_datetime)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(class_name) DO UPDATE SET 
            room_name = excluded.room_name,
            rows_count = excluded.rows_count,
            cols_count = excluded.cols_count,
            reset_datetime = excluded.reset_datetime
    """, (class_name, room_name, rows_count, cols_count, reset_datetime))
    conn.commit()
    check_and_apply_resets(conn)
    conn.close()
    return {"success": True, "message": f"[{class_name}] 설정이 성공적으로 저장되었습니다."}

@app.post("/api/admin/classes/reset-now")
def reset_class_now(class_name: str = Form(...)):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM seat_reservations WHERE class_name = ?", (class_name,))
    conn.commit()
    conn.close()
    return {"success": True, "message": f"[{class_name}] 반의 모든 좌석 신청 내역이 즉시 초기화되었습니다."}

@app.post("/api/admin/students/upload")
async def upload_students_file(file: UploadFile = File(...)):
    contents = await file.read()
    df = parse_student_file(contents)

    if df is None or '아이디' not in df.columns or '이름' not in df.columns:
        return {"message": "파일에서 '아이디'와 '이름' 컬럼을 찾을 수 없습니다."}

    class_col = '강좌명' if '강좌명' in df.columns else ('반명' if '반명' in df.columns else None)

    if class_col:
        grouped = df.groupby(['아이디', '이름'])[class_col].apply(
            lambda s: ', '.join(dict.fromkeys(s.dropna().astype(str)))
        ).reset_index()
    else:
        grouped = df[['아이디', '이름']].drop_duplicates()
        grouped['class_name'] = '-'

    conn = get_db_connection()
    cursor = conn.cursor()
    
    unique_classes = set()
    for _, row in grouped.iterrows():
        u_id = str(row['아이디']).strip()
        u_name = str(row['이름']).strip()
        c_name = str(row[class_col if class_col else 'class_name']).strip()
        
        for c in c_name.split(','):
            if c.strip() and c.strip() != '-':
                unique_classes.add(c.strip())

        cursor.execute("""
            INSERT INTO students (username, name, class_name, password, status)
            VALUES (?, ?, ?, '1234', 'active')
            ON CONFLICT(username) DO UPDATE SET
                name = excluded.name,
                class_name = excluded.class_name
        """, (u_id, u_name, c_name))
        
    for c in unique_classes:
        cursor.execute("INSERT OR IGNORE INTO class_configs (class_name, room_name, rows_count, cols_count) VALUES (?, '301호', 0, 0)", (c,))

    conn.commit()
    conn.close()

    return {"message": f"총 {len(grouped)}명의 학생과 {len(unique_classes)}개 반 목록이 등록되었습니다!"}

@app.post("/api/admin/rooms/upload")
async def upload_room_excel(file: UploadFile = File(...)):
    contents = await file.read()
    parsed_rooms = parse_seat_excel(contents)
    
    if not parsed_rooms:
        return {"message": "좌석표 엑셀 파일에서 유효한 강의실 시트를 찾을 수 없습니다."}
        
    conn = get_db_connection()
    cursor = conn.cursor()
    
    for r in parsed_rooms:
        cursor.execute("""
            INSERT INTO rooms (room_name, title, rows_count, cols_count, grid_json)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(room_name) DO UPDATE SET
                title = excluded.title,
                rows_count = excluded.rows_count,
                cols_count = excluded.cols_count,
                grid_json = excluded.grid_json
        """, (r["room_name"], r["title"], r["rows_count"], r["cols_count"], json.dumps(r["grid"])))
        
    conn.commit()
    conn.close()
    
    room_names = ", ".join([r["room_name"] for r in parsed_rooms])
    return {"message": f"총 {len(parsed_rooms)}개 강의실({room_names}) 좌석표가 등록되었습니다!"}

if __name__ == "__main__":
    import os
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port)