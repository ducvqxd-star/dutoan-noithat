import streamlit as st
import pandas as pd
import json
import io
import os
import re
import time
import math
import sqlite3
import hashlib
from datetime import datetime, timezone, timedelta
from PIL import Image
from pydantic import BaseModel
from typing import List, Optional
from google import genai
from google.genai import types

# -------------------------------------------------------------
# 1. CẤU HÌNH GIAO DIỆN & TỰ ĐỘNG NẠP LOGO SPACE
# -------------------------------------------------------------
LOGO_FILE = None
for fname in ["logo.png", "logo.jpg", "logo.jpeg"]:
    if os.path.exists(fname):
        LOGO_FILE = fname
        break

if LOGO_FILE:
    try:
        app_icon = Image.open(LOGO_FILE)
    except Exception:
        app_icon = "🏗️"
else:
    app_icon = "🏗️"

st.set_page_config(
    page_title="Hệ Thống Dự Toán Nội Thất - SPACE", 
    page_icon=app_icon,
    layout="wide"
)

st.markdown("""
<style>
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
        opacity: 1 !important;
        filter: none !important;
        width: 100% !important;
    }
    .stDataFrame td, .stDataFrame th {
        font-weight: 500 !important;
        font-size: 13px !important;
        white-space: nowrap !important;
    }
    .stApp {
        opacity: 1 !important;
    }
</style>
""", unsafe_allow_html=True)

DB_FILE = "qs_multiuser.db"
JSON_STORAGE_FILE = "saved_projects_backup.json"
IMAGE_DIR = "uploaded_drawings"
os.makedirs(IMAGE_DIR, exist_ok=True)

try:
    CLOUD_API_KEY = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    CLOUD_API_KEY = ""

def get_vietnam_time():
    tz_vn = timezone(timedelta(hours=7))
    return datetime.now(tz_vn).strftime("%Y-%m-%d %H:%M:%S")

def parse_dimensions(dim_str: str):
    nums = re.findall(r'\d+', str(dim_str).replace('.', '').replace(',', ''))
    if len(nums) >= 3:
        return float(nums[0]), float(nums[1]), float(nums[2])
    elif len(nums) == 2:
        return float(nums[0]), float(nums[1]), 100.0
    return 1000.0, 1000.0, 100.0

# -------------------------------------------------------------
# 2. CƠ CHẾ SAO LƯU & KHÔI PHỤC DỮ LIỆU CHỐNG MẤT
# -------------------------------------------------------------
def sync_projects_to_json():
    try:
        conn = sqlite3.connect(DB_FILE)
        p_df = pd.read_sql_query("SELECT * FROM projects", conn)
        b_df = pd.read_sql_query("SELECT * FROM boq_items", conn)
        conn.close()
        
        backup_data = {
            "projects": p_df.to_dict(orient="records"),
            "boq_items": b_df.to_dict(orient="records")
        }
        with open(JSON_STORAGE_FILE, "w", encoding="utf-8") as f:
            json.dump(backup_data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def restore_projects_from_json():
    if not os.path.exists(JSON_STORAGE_FILE):
        return
    try:
        with open(JSON_STORAGE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM projects")
        if cur.fetchone()[0] == 0:
            for p in data.get("projects", []):
                cur.execute('''
                    INSERT OR REPLACE INTO projects 
                    (id, user_id, project_name, item_code, item_name, dimensions, specs, direct_cost, overhead_cost, tender_price, technical_risks, image_paths, created_at, review_status, review_note)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    p.get("id"), p.get("user_id"), p.get("project_name"), p.get("item_code"), p.get("item_name"),
                    p.get("dimensions"), p.get("specs"), p.get("direct_cost"), p.get("overhead_cost"), p.get("tender_price"),
                    p.get("technical_risks"), p.get("image_paths"), p.get("created_at"), p.get("review_status"), p.get("review_note")
                ))
                
            for b in data.get("boq_items", []):
                cur.execute('''
                    INSERT OR REPLACE INTO boq_items 
                    (id, project_id, section_code, material_code, description, location, dai_m, rong_m, cao_m, so_luong, unit, net_qty, waste_percent, waste_qty, gross_qty, unit_cost, amount, meaning, formula)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    b.get("id"), b.get("project_id"), b.get("section_code"), b.get("material_code"), b.get("description"),
                    b.get("location"), b.get("dai_m"), b.get("rong_m"), b.get("cao_m"), b.get("so_luong"),
                    b.get("unit"), b.get("net_qty"), b.get("waste_percent"), b.get("waste_qty"), b.get("gross_qty"),
                    b.get("unit_cost"), b.get("amount"), b.get("meaning"), b.get("formula")
                ))
            conn.commit()
        conn.close()
    except Exception:
        pass

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password_hash TEXT, full_name TEXT, role TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS projects (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, project_name TEXT, item_code TEXT, item_name TEXT, dimensions TEXT, specs TEXT, direct_cost REAL, overhead_cost REAL, tender_price REAL, technical_risks TEXT, image_paths TEXT, created_at TIMESTAMP, review_status TEXT DEFAULT 'Chờ duyệt', review_note TEXT DEFAULT '', FOREIGN KEY (user_id) REFERENCES users(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS boq_items (id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER, section_code TEXT, material_code TEXT, description TEXT, location TEXT, dai_m REAL DEFAULT 0.0, rong_m REAL DEFAULT 0.0, cao_m REAL DEFAULT 0.0, so_luong REAL DEFAULT 1.0, unit TEXT, net_qty REAL, waste_percent REAL, waste_qty REAL, gross_qty REAL, unit_cost REAL, amount REAL, meaning TEXT, formula TEXT, FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE)''')
    admin_pw = hashlib.sha256("admin123".encode()).hexdigest()
    qs1_pw = hashlib.sha256("qs123".encode()).hexdigest()
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, full_name, role) VALUES ('admin', ?, 'Quản Trị Viên', 'admin')", (admin_pw,))
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, full_name, role) VALUES ('qs01', ?, 'Kỹ Sư QS 01', 'qs')", (qs1_pw,))
    conn.commit()
    conn.close()

init_db()
restore_projects_from_json()

# -------------------------------------------------------------
# 3. THUẬT TOÁN ĐỊNH MỨC NHÂN CÔNG
# -------------------------------------------------------------
def calculate_dynamic_labor(length_m: float, depth_m: float, height_m: float, item_name: str, specs_text: str, is_night_shift: bool, is_high_access: bool):
    name_low = str(item_name).lower()
    spec_low = str(specs_text).lower()

    is_ceiling = any(k in name_low for k in ["trần", "tran", "la phông"])
    is_wall_panel = any(k in name_low for k in ["vách ốp", "vach op", "ốp tường", "lam sóng"])
    is_door = any(k in name_low for k in ["cửa", "cua"])
    is_table = any(k in name_low for k in ["bàn", "ban"])
    is_bed = any(k in name_low for k in ["giường", "giuong"])

    if is_ceiling:
        area = max(length_m * depth_m, 1.0) if depth_m > 0 else max(length_m * height_m, 1.0)
        base_sx = area * 0.25
        base_ld = area * 0.35
    elif is_wall_panel:
        area = max(length_m * height_m, 1.0)
        base_sx = area * 0.20
        base_ld = area * 0.25
    elif is_door:
        area = max(length_m * height_m, 1.0)
        base_sx = area * 0.30
        base_ld = max(math.ceil(area * 0.20), 1)
    elif is_table or is_bed:
        base_sx = 2.0
        base_ld = 1.0
    else:
        area = max(length_m * height_m, 1.0)
        base_sx = area * 0.45
        base_ld = area * 0.50

    if any(k in spec_low for k in ["sơn pu", "sơn màu", "sơn hiệu ứng"]):
        base_sx += 1.0
    if any(k in spec_low for k in ["khung sắt", "sắt hộp", "thép"]):
        base_sx += 0.5
    if any(k in spec_low for k in ["đá", "nung kết"]) and not is_ceiling:
        base_ld += 0.8
    if any(k in spec_low for k in ["da", "nệm"]) and not is_ceiling:
        base_sx += 0.5

    if is_high_access:
        base_ld = base_ld / 0.7 
    if is_night_shift:
        base_ld = base_ld * 1.5

    prod_days = max(math.ceil(base_sx), 1)
    inst_days = max(math.ceil(base_ld), 1)
    return prod_days, inst_days

# -------------------------------------------------------------
# 4. BỘ NẠP MASTER EXCEL
# -------------------------------------------------------------
@st.cache_data(ttl=3600)
def load_price_master():
    candidate_files = ["BIEU-MAU-DU-TOAN_MAPPED_MASTER.xlsx", "BIEU-MAU-DU-TOAN_VERSION-2.xlsx", "Gia cost_updated.xlsx", "Gia cost.xlsx", "Gia_cost_CapNhat_AI.xlsx"]
    target_file = next((f for f in candidate_files if os.path.exists(f)), None)
    if not target_file:
        all_xlsx = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~$")]
        target_file = all_xlsx[0] if all_xlsx else None
        if not target_file: return {}, "Không tìm thấy file Excel"

    try:
        xls = pd.ExcelFile(target_file)
        target_sheet = next((s for s in xls.sheet_names if any(k in s.strip().lower() for k in ["dmvt", "đơn giá vt", "don gia vt", "dongia", "vt"])), xls.sheet_names[0])
        df_raw = pd.read_excel(target_file, sheet_name=target_sheet, header=None)
        
        col_code, col_name, col_unit, col_price = 1, 2, 3, 14
        start_row = 3

        for r_idx in range(min(12, len(df_raw))):
            row_vals = [str(x).strip().lower() for x in df_raw.iloc[r_idx] if pd.notna(x)]
            if any("đơn giá" in v or "thành tiền" in v for v in row_vals):
                for c_idx, val in enumerate(df_raw.iloc[r_idx]):
                    val_str = str(val).strip().lower()
                    if "mã" in val_str and "vật tư" in val_str: col_code = c_idx
                    elif "tên vật tư" in val_str or "hạng mục" in val_str: col_name = c_idx
                    elif val_str == "đvt" or "đơn vị" in val_str: col_unit = c_idx
                    elif val_str == "đơn giá" or val_str == "đơn giá tính": col_price = c_idx
                start_row = r_idx + 1
                break

        price_dict = {}
        current_cat = "CHUNG"

        for r in range(start_row, len(df_raw)):
            code = df_raw.iloc[r, col_code] if len(df_raw.columns) > col_code else None
            name = df_raw.iloc[r, col_name] if len(df_raw.columns) > col_name else None
            unit = df_raw.iloc[r, col_unit] if len(df_raw.columns) > col_unit else 'cái'
            price = df_raw.iloc[r, col_price] if len(df_raw.columns) > col_price else 0

            if str(code).strip().lower() in ["mã vật tư", "code", "stt", "nan", ""]: continue
            if pd.isna(code) and pd.notna(name):
                current_cat = str(name).strip()
            elif pd.notna(code) and pd.notna(name):
                try: p_val = float(price) if isinstance(price, (int, float)) else float(str(price).replace(",", "").replace(".", "").strip())
                except: p_val = 0.0
                if p_val > 0:
                    price_dict[str(code).strip()] = {"category": current_cat, "name": str(name).strip().replace('\n', ' '), "unit": str(unit).strip() if pd.notna(unit) else 'cái', "unit_cost": p_val}
        
        special_items = {
            "NC-SX": {"category": "NHÂN CÔNG", "name": "NC-SX: Nhân công sản xuất xưởng", "unit": "Công", "unit_cost": 500000.0},
            "NC-LD": {"category": "NHÂN CÔNG", "name": "NC-LD: Nhân công lắp đặt công trình", "unit": "Công", "unit_cost": 550000.0},
            "KS.V30": {"category": "KIM LOẠI", "name": "Khung thép hộp 20x40x1.2mm", "unit": "md", "unit_cost": 55000.0},
            "PU69": {"category": "SƠN", "name": "Sơn PU bóng", "unit": "m2", "unit_cost": 270000.0},
            "PP": {"category": "PHỤ LIỆU", "name": "Phụ liệu keo vít", "unit": "Bộ", "unit_cost": 500000.0}
        }
        for k, v in special_items.items():
            if k not in price_dict: price_dict[k] = v
        return price_dict, f"{target_file} (Sheet: {target_sheet})"
    except Exception as e:
        return {}, f"Lỗi: {str(e)}"

PRICE_MASTER, SOURCE_INFO = load_price_master()

def get_approved_benchmark(item_code: str, item_name: str):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, project_name, item_name, dimensions, direct_cost FROM projects WHERE review_status = '✅ Đúng giá' AND (item_code = ? OR item_name = ?) ORDER BY id DESC LIMIT 1", (item_code, item_name))
    row = c.fetchone()
    if not row:
        conn.close()
        return None, ""
    bench_id, p_name, itm_name, dims, d_cost = row
    items_df = pd.read_sql_query("SELECT section_code, material_code, description, location, dai_m, rong_m, cao_m, so_luong, unit, waste_percent FROM boq_items WHERE project_id = ?", conn, params=(bench_id,))
    conn.close()
    if items_df.empty: return None, ""
    bench_text = f"HỒ SƠ MẪU CHUẨN CÙNG LOẠI [{itm_name} - Kích thước: {dims}]:\n"
    for _, itm in items_df.iterrows():
        bench_text += f"- [{itm['section_code']}] {itm['material_code']} | {itm['description']} ({itm['location']}) -> Dài:{itm['dai_m']}m x Rộng:{itm['rong_m']}m x Cao:{itm['cao_m']}m x SL:{itm['so_luong']} {itm['unit']} (Hao hụt:{itm['waste_percent']}%)\n"
    return bench_id, bench_text

# -------------------------------------------------------------
# 5. ĐỘNG CƠ QUY ĐỔI KHỐI LƯỢNG M3 -> M2 / MD THÔNG MINH
# -------------------------------------------------------------
VALID_SECTIONS = ["I. VẬT TƯ THẦU PHỤ", "II. VẬT TƯ CHÍNH", "III. HOÀN THIỆN (KIM LOẠI, PU...)", "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ", "V. BAO BÌ", "VI. PHỤ KIỆN", "VII. SOURCE NGOÀI"]

def resolve_material_master(section_code: str, raw_code: str, description: str, location: str, unit_hint: str):
    sec = str(section_code).strip().upper()
    loc_clean = str(location).strip().lower()
    desc_clean = str(description).strip().lower()
    full_text = loc_clean + " " + desc_clean
    raw_c = str(raw_code).strip()

    if any(k in full_text for k in ["lắp đặt", "nc-ld", "công trình"]): return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ", "NC-LD", PRICE_MASTER.get("NC-LD")
    if any(k in full_text for k in ["sản xuất", "nc-sx", "thợ mộc", "xưởng"]): return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ", "NC-SX", PRICE_MASTER.get("NC-SX")
    if "sơn pu" in full_text or "sơn màu" in full_text: return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ", "PU69", PRICE_MASTER.get("PU69")
    if any(k in full_text for k in ["keo", "vít", "phụ liệu", "đinh"]): return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ", "PP", PRICE_MASTER.get("PP")
    if any(k in full_text for k in ["bao bì", "màng pe", "carton"]): return "V. BAO BÌ", "B.BBTN.1", PRICE_MASTER.get("B.BBTN.1")
    if any(k in full_text for k in ["sắt hộp", "khung sắt", "thép hộp"]): return "II. VẬT TƯ CHÍNH", "KS.V30", PRICE_MASTER.get("KS.V30")

    if bool(re.search(r'(?:\b|\s|^)đá(?:\s|$|[,\.\-_])', full_text)) and "vân đá" not in full_text and "đáy" not in loc_clean:
        return "VII. SOURCE NGOÀI", "ST.01", PRICE_MASTER.get("ST.01")
    if any(k in full_text for k in ["bọc da", "nệm bọc da", "leather"]): return "VII. SOURCE NGOÀI", "FB LE-01", PRICE_MASTER.get("FB LE-01")
    if any(k in full_text for k in ["cánh kính", "kính cường lực", "gương"]): return "VII. SOURCE NGOÀI", "KI.1.04", PRICE_MASTER.get("KI.1.04")

    if raw_c in PRICE_MASTER: return section_code, raw_c, PRICE_MASTER[raw_c]

    words = [w for w in re.split(r'[,;\-\s_/\(\)\+]+', desc_clean) if len(w) >= 2]
    best_code, best_data, best_score = None, None, 0
    for p_code, p_data in PRICE_MASTER.items():
        p_name_low = p_data["name"].lower()
        score = sum(3 for w in words if w in p_name_low)
        if desc_clean in p_name_low: score += 5
        if score > best_score:
            best_score, best_code, best_data = score, p_code, p_data

    if best_score >= 3 and best_data: return section_code, best_code, best_data
    return section_code, "CHƯA_CÓ_MÃ", None

def smart_unit_conversion(master_unit: str, master_price: float, master_name: str, ai_unit: str, r_m: float, c_m: float):
    m_u = str(master_unit).lower().strip()
    a_u = str(ai_unit).lower().strip()

    if "m3" in m_u and a_u in ["m2", "md", "m"]:
        if a_u == "m2":
            thick = c_m if c_m > 0 else 0.02
            new_price = master_price * thick * 1.35
            return "m2", new_price, f"{master_name} (Đã quy đổi m2 dày {thick*1000:.0f}mm)"
        else:
            width = r_m if r_m > 0 else 0.05
            thick = c_m if c_m > 0 else 0.05
            new_price = master_price * (width * thick) * 1.45
            return "md", new_price, f"{master_name} (Đã quy đổi md {width*1000:.0f}x{thick*1000:.0f}mm)"
    
    return master_unit, master_price, master_name

def calculate_clean_net_quantity(d_m: float, r_m: float, c_m: float, sl: float, unit_str: str):
    unit_lower = str(unit_str).lower().strip()
    if "m3" in unit_lower: return round(d_m * (r_m if r_m>0 else 0.05) * (c_m if c_m>0 else 0.05) * sl, 4)
    elif "m2" in unit_lower: return round(d_m * r_m * sl, 3) if r_m > 0 else round(d_m * c_m * sl, 3)
    elif "md" in unit_lower or unit_lower == "m": return round(d_m * sl, 3)
    else: return round(sl, 2)

def clean_section_row(row):
    val = row.get('Mục chiết tính (I -> VII)')
    if val in VALID_SECTIONS: return val
    code = str(row.get('Code', '')).strip().upper()
    full = str(row.get('Vị trí', '')).strip().lower() + " " + str(row.get('Vật tư / Hạng mục', '')).strip().lower()

    if code.startswith("NC") or any(k in full for k in ["nhân công", "keo", "vít"]): return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ"
    if any(k in full for k in ["thân tủ", "hông tủ", "đáy tủ", "khung sắt"]): return "II. VẬT TƯ CHÍNH"
    if code.startswith("DA.") or code.startswith("ST.") or any(k in full for k in ["kính", "da tam tính", "mặt kệ đá"]): return "VII. SOURCE NGOÀI"
    if code.startswith("PK") or any(k in full for k in ["bản lề", "ray", "led"]): return "VI. PHỤ KIỆN"
    if code.startswith("BB") or any(k in full for k in ["bao bì", "carton"]): return "V. BAO BÌ"
    if code.startswith("CC") or code.startswith("IN") or any(k in full for k in ["chỉ cạnh", "nẹp", "inox", "sơn"]): return "III. HOÀN THIỆN (KIM LOẠI, PU...)"
    return "II. VẬT TƯ CHÍNH"

# -------------------------------------------------------------
# 6. GIAO DIỆN CHÍNH
# -------------------------------------------------------------
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
    st.session_state['user_info'] = None

def login(username, password):
    pw_hash = hashlib.sha256(password.encode()).hexdigest()
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, username, full_name, role FROM users WHERE username = ? AND password_hash = ?", (username, pw_hash))
    user = c.fetchone()
    conn.close()
    if user:
        st.session_state['logged_in'] = True
        st.session_state['user_info'] = {"id": user[0], "username": user[1], "full_name": user[2], "role": user[3]}
        st.rerun()
    else: st.error("Sai tên đăng nhập hoặc mật khẩu!")

def logout():
    st.session_state['logged_in'] = False
    st.session_state['user_info'] = None
    st.rerun()

if not st.session_state['logged_in']:
    st.title("🔐 Hệ Thống Báo Giá & Dự Toán Nội Thất (SPACE AI QS)")
    col_l1, col_l2, _ = st.columns([1, 1, 1])
    with col_l1:
        u_in = st.text_input("Tên đăng nhập:")
        p_in = st.text_input("Mật khẩu:", type="password")
        if st.button("Đăng nhập", type="primary", use_container_width=True): login(u_in, p_in)
    st.stop()

class ChiTietVatTu(BaseModel):
    section_code: str
    material_code: str
    description: str
    location: str
    dai_m: float
    rong_m: float
    cao_m: float
    so_luong: float
    unit: str
    waste_factor_percent: float

class TakeoffResult(BaseModel):
    item_code: str
    item_name: str
    breakdown_items: List[ChiTietVatTu]
    technical_risks_and_notes: List[str]

user = st.session_state['user_info']

with st.sidebar:
    if LOGO_FILE:
        st.image(LOGO_FILE, use_container_width=True)
        st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"👤 **{user['full_name']}** (`{user['role']}`)")
    if st.button("Đăng xuất", use_container_width=True): logout()
    st.markdown("---")
    
    st.header("⚙️ Các hệ số rủi ro & Chi phí mềm")
    api_key = CLOUD_API_KEY if CLOUD_API_KEY else st.text_input("Gemini API Key:", type="password")
        
    is_night_shift = st.checkbox("👷 Thi công đêm (TTTM/VP)")
    is_high_access = st.checkbox("🪜 Thi công trên cao >3.5m / Cầu thang")
    pm_fee_rate = st.slider("Phí QLDA %:", 0.0, 10.0, 3.0)
    dlp_rate = st.slider("Quỹ bảo hành %:", 0.0, 5.0, 2.0)
    consumables_rate = st.slider("Vật tư tiêu hao %:", 0.0, 5.0, 2.0)
    overhead_rate = st.slider("Chi phí chung Xưởng %:", 0.0, 25.0, 8.0)
    profit_rate = st.slider("Biên Lợi nhuận mục tiêu %:", 0.0, 30.0, 12.0)
    
    st.caption(f"Nguồn dữ liệu: `{SOURCE_INFO}`")
    st.markdown("---")
    if os.path.exists(JSON_STORAGE_FILE):
        with open(JSON_STORAGE_FILE, "rb") as fp_j:
            st.download_button("📥 Tải file JSON sao lưu hồ sơ", data=fp_j.read(), file_name="saved_projects_backup.json", mime="application/json", use_container_width=True)

tab1, tab2, tab3 = st.tabs(["📝 Bóc Tách Dự Toán Mới", "📂 Thư Viện Hồ Sơ & Chỉnh Sửa Chiết Tính", "📊 Quản Lý Đơn Giá Master"])

# TAB 1: BÓC TÁCH MỚI ĐA LUỒNG
with tab1:
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("1. Dữ liệu bản vẽ & Thông số kỹ thuật")
        project_name = st.text_input("Tên dự án / Công trình:", value="Biệt thự Park City")
        item_code = st.text_input("Mã sản phẩm / Hạng mục:", value="SP12")
        item_name = st.text_input("Tên sản phẩm:", value="Tủ rượu gỗ tự nhiên")
        dimensions = st.text_input("Kích thước bao phủ bì (DxRxC mm):", value="4095 x 450 x 2600 mm")
        
        col_up_2d, col_up_3d = st.columns(2)
        with col_up_2d:
            uploaded_2d = st.file_uploader("Upload Bản vẽ 2D (Kích thước chuẩn):", type=["jpg", "png", "pdf", "jpeg"], accept_multiple_files=True)
        with col_up_3d:
            uploaded_3d = st.file_uploader("Upload Ảnh 3D (Thẩm mỹ & Cấu kiện):", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        
        all_uploaded_files = []
        if uploaded_2d: all_uploaded_files.extend(uploaded_2d)
        if uploaded_3d: all_uploaded_files.extend(uploaded_3d)

        if all_uploaded_files:
            c_preview = st.columns(min(len(all_uploaded_files), 3))
            for idx, f in enumerate(all_uploaded_files):
                with c_preview[idx % 3]: st.image(f, use_container_width=True)

        specs = st.text_area(
            "Quy cách kỹ thuật & Vật liệu chỉ định:",
            height=140, 
            value="- Gỗ sồi tự nhiên nhập khẩu sơn PU hoàn thiện (Bóc chi tiết từng đợt m2, nẹp md).\n- HẬU TỦ: MDF Veneer Sồi chống ẩm 9mm.\n- KHUNG GIA CỐ: Thép hộp 20x40x1.2mm.\n- CÁNH TỦ: Cánh kính cường lực 5mm khung nhôm, phụ kiện giảm chấn."
        )

        submit_btn = st.button("🚀 Bóc tách chéo (2D+3D) với 15 Rule QS", type="primary", use_container_width=True)

    with c2:
        st.subheader("2. Chi phí tổng quan & Đánh giá Rủi ro")
        if submit_btn:
            if not api_key: st.error("Vui lòng cấu hình Gemini API Key.")
            elif not all_uploaded_files: st.warning("Vui lòng tải lên ít nhất 1 bản vẽ 2D hoặc 3D.")
            else:
                with st.spinner(f"AI QS đang đối chiếu 2D-3D và phân bổ quy tắc QS cho {item_name}..."):
                    try:
                        client = genai.Client(api_key=api_key)
                        images = []
                        saved_paths = []
                        for idx, f in enumerate(all_uploaded_files):
                            img = Image.open(f)
                            img.thumbnail((1600, 1600))
                            images.append(img)
                            spath = os.path.join(IMAGE_DIR, f"{user['username']}_{item_code}_{idx}_{f.name}")
                            img.save(spath)
                            saved_paths.append(spath)

                        l_mm, d_mm, h_mm = parse_dimensions(dimensions)
                        l_m = l_mm / 1000.0; d_m = d_mm / 1000.0; h_m = h_mm / 1000.0

                        prod_days_calc, inst_days_calc = calculate_dynamic_labor(l_m, d_m, h_m, item_name, specs, is_night_shift, is_high_access)
                        bench_id, benchmark_context = get_approved_benchmark(item_code, item_name)

                        # ĐỊNH TUYẾN CẤU TRÚC ĐỂ AI KHÔNG BỊ "LƯỜI"
                        name_low = str(item_name).lower()
                        if any(k in name_low for k in ["tủ", "kệ", "cabinet"]):
                            structural_hint = "- BẮT BUỘC bóc tách chi tiết 15-30 dòng cấu kiện (Hông, đợt, vách chia, nóc, đáy, cánh...). KHÔNG ĐƯỢC GỘP TẮT."
                        else: structural_hint = ""

                        # PROMPT TÍCH HỢP ĐA LUỒNG (2D+3D) & 15 RULE QS
                        prompt = f"""
                        BẠN LÀ SENIOR QS NỘI THẤT. BẠN ĐANG NHẬN ĐƯỢC BẢN VẼ KỸ THUẬT (2D) VÀ/HOẶC ẢNH RENDER (3D).
                        - Tên sản phẩm: {item_name} (Mã: {item_code})
                        - Kích thước phủ bì tổng: Dài {l_m:.3f}m x Rộng {d_m:.3f}m x Cao {h_m:.3f}m.
                        - Specs: {specs}
                        
                        {structural_hint}

                        QUY TẮC BÓC TÁCH KÉP (DUAL-INPUT RULE) BẮT BUỘC ÁP DỤNG:
                        1. KÍCH THƯỚC: Chỉ nội suy chia tỉ lệ các tấm ván, hộc kéo, khoảng cách đợt dựa vào con số Dim trên bản vẽ 2D (nếu có).
                        2. THẨM MỸ & VẬT TƯ NGẦM: Quét ảnh 3D để nhận diện:
                           - Có thanh nhôm LED hắt sáng không? -> Bóc LED (md).
                           - Cánh kính hay cánh gỗ? Mặt lưng bọc da hay ốp đá?
                        3. ĐỐI CHIẾU XUNG ĐỘT: Nếu hình dáng, số lượng đợt kệ trên 2D và 3D lệch nhau, ưu tiên bóc theo 2D, ghi chú vào `technical_risks_and_notes`.

                        15 QUY TẮC QS NÂNG CAO BẮT BUỘC:
                        - Đối với Gỗ tự nhiên (Sồi, óc chó...): TẤM PHẲNG BẮT BUỘC ĐVT 'm2' (Ghi chiều Dày tại cao_m). THANH DÀI (Phào nẹp) BẮT BUỘC ĐVT 'md'. Cấm bóc m3 cho chi tiết rời.
                        - Khẩu độ đợt tủ >800mm -> bóc 2 lớp ván (kẹp đôi). Bếp/Lavabo -> Cốt chống ẩm/Picomat.
                        - Ép Laminate/Veneer -> Bóc lớp Balance mặt sau. Sơn PU -> m2 khai triển x 1.15 hoặc 2.15.
                        - L_ray = Floor((Sâu - 50)/50)*50. Bản lề: <1200(2 cái), 1200-1600(3 cái), 1600-2200(4 cái). 
                        - Kính/Gương: Mọi tấm < 0.3 m2 -> Ép làm tròn 0.3 m2. 
                        - Đá lật vân (Bookmatch) hao hụt 50%.
                        - Khung sắt tính md, Ván/Đá/Kính tính m2, Phào nẹp len chân tính md.

                        ĐỊNH MỨC NHÂN CÔNG:
                        - Dòng 1: material_code='NC-SX', description='NC-SX: Nhân công sản xuất xưởng', ĐVT='Công', SL={prod_days_calc}, hao hụt=0.
                        - Dòng 2: material_code='NC-LD', description='NC-LD: Nhân công lắp đặt công trình', ĐVT='Công', SL={inst_days_calc}, hao hụt=0.
                        """

                        if benchmark_context: prompt += f"\nTHAM CHIẾU HỒ SƠ CHUẨN ĐÃ DUYỆT:\n{benchmark_context}"

                        MODELS_TO_TRY = ['gemini-3.8-flash', 'gemini-3.5-flash-lite']
                        response = None
                        for current_model in MODELS_TO_TRY:
                            for attempt in range(3):
                                try:
                                    response = client.models.generate_content(
                                        model=current_model, contents=[*images, prompt],
                                        config=types.GenerateContentConfig(
                                            system_instruction="BẠN LÀ SENIOR QS NỘI THẤT. Ưu tiên 2D, nhận dạng 3D. Tuân thủ 15 quy tắc QS, bóc m2/md cho gỗ tấm.",
                                            response_mime_type="application/json", response_schema=TakeoffResult, temperature=0.0
                                        )
                                    )
                                    break
                                except Exception as err:
                                    if "RESOURCE_EXHAUSTED" in str(err): break
                                    time.sleep(2)
                            if response: break

                        data = json.loads(response.text)
                        
                        total_direct = 0
                        boq_records = []
                        for comp in data.get("breakdown_items", []):
                            s_code = comp.get("section_code", "II. VẬT TƯ CHÍNH")
                            raw_code = str(comp.get("material_code", "")).strip()
                            ai_desc = comp.get("description", "").strip()
                            loc = comp.get("location", "").strip()
                            ai_unit_str = str(comp.get("unit", "cái")).strip()
                            
                            d_m = float(comp.get("dai_m", 0.0))
                            r_m = float(comp.get("rong_m", 0.0))
                            c_m = float(comp.get("cao_m", 0.0))
                            sl = float(comp.get("so_luong", 1.0))
                            w_pct = float(comp.get("waste_factor_percent", 0.0))

                            final_sec, matched_code, item_info = resolve_material_master(s_code, raw_code, ai_desc, loc, ai_unit_str)

                            if item_info:
                                master_u_cost = item_info["unit_cost"]
                                master_name = item_info["name"]
                                master_unit = item_info["unit"]
                                
                                final_unit, final_cost, final_desc = smart_unit_conversion(
                                    master_unit, master_u_cost, master_name, ai_unit_str, r_m, c_m
                                )
                            else:
                                matched_code = "CHƯA_CÓ_MÃ"
                                final_cost = 0.0
                                final_desc = ai_desc
                                final_unit = ai_unit_str

                            if matched_code == "KS.V30" or "thép hộp" in ai_desc.lower():
                                final_unit = "md"
                                if r_m > 0 and d_m < 20.0: d_m = round(d_m * r_m * 5.0, 1); r_m = 0.0

                            net_q = calculate_clean_net_quantity(d_m, r_m, c_m, sl, final_unit)

                            # Ứng dụng Rule Kính < 0.3m2
                            if "kính" in ai_desc.lower() or matched_code == "KI.1.04":
                                if net_q < 0.3 and sl == 1.0: net_q = 0.3
                                elif net_q / max(sl, 1) < 0.3: net_q = 0.3 * sl

                            is_discrete = any(u in final_unit.lower() for u in ['cai', 'cái', 'bộ', 'bo', 'con', 'chiếc', 'hệ', 'he', 'công'])
                            if is_discrete:
                                waste_q = math.ceil(net_q * (w_pct / 100.0)) if w_pct > 0 else 0
                                gross_q = net_q + waste_q
                            else:
                                waste_q = round(net_q * (w_pct / 100.0), 4 if "m3" in final_unit.lower() else 3)
                                gross_q = round(net_q + waste_q, 4 if "m3" in final_unit.lower() else 3)

                            amt = round(gross_q * final_cost, 0)
                            total_direct += amt
                            
                            boq_records.append({
                                "section_code": final_sec, "material_code": matched_code, "description": final_desc,
                                "location": loc, "dai_m": d_m, "rong_m": r_m, "cao_m": c_m, "so_luong": sl,
                                "unit": final_unit, "net_qty": net_q, "waste_percent": w_pct, "waste_qty": waste_q,
                                "gross_qty": gross_q, "unit_cost": final_cost, "amount": amt,
                                "meaning": f"{loc}: {d_m:.2f}x{r_m:.2f}x{sl:.0f}", "formula": f"{gross_q:.3f} x {final_cost:,.0f} đ"
                            })

                        # Tính chi phí mềm (Soft Costs - Rule 11 & 15)
                        consumables_amt = total_direct * (consumables_rate / 100.0)
                        pm_fee_amt = total_direct * (pm_fee_rate / 100.0)
                        dlp_amt = total_direct * (dlp_rate / 100.0)
                        overhead_amt = total_direct * (overhead_rate / 100.0)
                        
                        total_cost_with_soft = total_direct + consumables_amt + pm_fee_amt + dlp_amt + overhead_amt
                        tender_amt = total_cost_with_soft * (1 + profit_rate / 100.0)
                        
                        total_overhead_pct_db = overhead_rate + pm_fee_rate + dlp_rate + consumables_rate
                        overhead_amt_db = total_direct * (total_overhead_pct_db / 100.0)

                        created_time = get_vietnam_time()

                        conn = sqlite3.connect(DB_FILE)
                        cur = conn.cursor()
                        cur.execute('''INSERT INTO projects (user_id, project_name, item_code, item_name, dimensions, specs, direct_cost, overhead_cost, tender_price, technical_risks, image_paths, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                            (user["id"], project_name, item_code, item_name, dimensions, specs, total_direct, overhead_amt_db, tender_amt, json.dumps(data.get("technical_risks_and_notes", [])), json.dumps(saved_paths), created_time))
                        proj_id = cur.lastrowid
                        
                        for r in boq_records:
                            cur.execute('''INSERT INTO boq_items (project_id, section_code, material_code, description, location, dai_m, rong_m, cao_m, so_luong, unit, net_qty, waste_percent, waste_qty, gross_qty, unit_cost, amount, meaning, formula) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                                (proj_id, r["section_code"], r["material_code"], r["description"], r["location"], r["dai_m"], r["rong_m"], r["cao_m"], r["so_luong"], r["unit"], r["net_qty"], r["waste_percent"], r["waste_qty"], r["gross_qty"], r["unit_cost"], r["amount"], r["meaning"], r["formula"]))
                        conn.commit()
                        conn.close()

                        sync_projects_to_json()
                        st.success(f"✅ Bóc tách chiết tính thành công!")
                        
                        st.markdown(f"**💡 Cảnh báo rủi ro 2D/3D (Từ AI QS):**")
                        for note in data.get("technical_risks_and_notes", []):
                            st.write(f"- 🔴 {note}")

                        st.markdown("---")
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Tổng giá Cost xưởng (Direct)", f"{total_direct:,.0f} đ")
                        m2.metric(f"Chi phí mềm + Overhead ({total_overhead_pct_db}%)", f"{overhead_amt_db:,.0f} đ")
                        m3.metric(f"Giá chào thầu ({profit_rate}%)", f"{tender_amt:,.0f} đ")

                    except Exception as e:
                        st.error(f"Lỗi xử lý: {str(e)}")

# TAB 2: THƯ VIỆN HỒ SƠ & BẢNG CHỈNH SỬA CHIẾT TÍNH ĐẦY ĐỦ
with tab2:
    st.subheader(f"📂 Thư Viện Hồ Sơ & Chỉnh Sửa Chiết Tính Chi Tiết")
    conn = sqlite3.connect(DB_FILE)
    df_saved = pd.read_sql_query("SELECT p.id, u.full_name AS nguoi_tao, p.project_name, p.item_code, p.item_name, p.dimensions, p.direct_cost, p.tender_price, p.review_status, p.review_note, p.created_at, p.user_id FROM projects p LEFT JOIN users u ON p.user_id = u.id ORDER BY p.id DESC", conn)
    
    if df_saved.empty:
        st.info("Chưa có hồ sơ dự toán nào được lưu.")
    else:
        st.dataframe(df_saved[["id", "nguoi_tao", "item_name", "dimensions", "direct_cost", "tender_price", "review_status"]].style.format({"direct_cost": "{:,.0f} đ", "tender_price": "{:,.0f} đ"}), use_container_width=True)
        
        st.markdown("### 🛠️ Quản Trị & Đào Tạo Bộ Não AI")
        sel_col1, sel_col2, sel_col3 = st.columns([1.5, 2.5, 1])
        with sel_col1:
            sel_id = st.selectbox("Chọn ID hồ sơ cần duyệt / xem:", df_saved["id"].tolist())
            selected_row = df_saved[df_saved["id"] == sel_id].iloc[0]
        with sel_col2:
            if user['role'] == 'admin':
                new_status = st.radio("Đánh giá Admin:", ['Chờ duyệt', '✅ Đúng giá', '❌ Sai giá'], index=['Chờ duyệt', '✅ Đúng giá', '❌ Sai giá'].index(selected_row['review_status'] if pd.notna(selected_row['review_status']) else 'Chờ duyệt'), horizontal=True)
                new_note = st.text_input("Ghi chú Admin:", value=selected_row['review_note'] if pd.notna(selected_row['review_note']) else '')
                if st.button("💾 Lưu phê duyệt", type="primary"):
                    conn.cursor().execute("UPDATE projects SET review_status = ?, review_note = ? WHERE id = ?", (new_status, new_note, sel_id))
                    conn.commit()
                    sync_projects_to_json()
                    st.success("Đã lưu kết quả kiểm duyệt!")
                    time.sleep(0.5); st.rerun()
            else: st.info(f"Trạng thái: {selected_row['review_status']}")

        with sel_col3:
            st.write("")
            if (user['role'] == 'admin' or selected_row['user_id'] == user['id']) and st.button(f"🗑️ Xóa hồ sơ #{sel_id}", type="secondary", use_container_width=True):
                conn.cursor().execute("DELETE FROM boq_items WHERE project_id = ?", (sel_id,))
                conn.cursor().execute("DELETE FROM projects WHERE id = ?", (sel_id,))
                conn.commit()
                sync_projects_to_json()
                st.warning(f"Đã xóa hoàn toàn hồ sơ #{sel_id}."); time.sleep(0.5); st.rerun()

        st.markdown("---")
        state_key = f"df_edit_data_{sel_id}"
        if state_key not in st.session_state:
            df_init = pd.read_sql_query("SELECT id AS item_db_id, section_code AS 'Mục chiết tính (I -> VII)', material_code AS 'Code', description AS 'Vật tư / Hạng mục', location AS 'Vị trí', dai_m AS 'Dài (m)', rong_m AS 'Rộng (m)', cao_m AS 'Cao (m)', so_luong AS 'SL', unit AS 'ĐVT', net_qty AS 'KL Tịnh (Net)', waste_percent AS '% Hao hụt', gross_qty AS 'KL Tổng (Gross)', unit_cost AS 'Đơn giá', amount AS 'Thành tiền' FROM boq_items WHERE project_id = ?", conn, params=(sel_id,))
            df_init['Mục chiết tính (I -> VII)'] = df_init.apply(clean_section_row, axis=1)
            st.session_state[state_key] = df_init

        edited_df = st.data_editor(st.session_state[state_key], use_container_width=True, num_rows="dynamic", height=540, key=f"editor_table_{sel_id}", column_config={
                "item_db_id": None, "Mục chiết tính (I -> VII)": st.column_config.SelectboxColumn("Mục chiết tính", options=VALID_SECTIONS, required=True),
                "Code": st.column_config.TextColumn("Code"), "Vật tư / Hạng mục": st.column_config.TextColumn("Vật tư", required=True),
                "Dài (m)": st.column_config.NumberColumn("Dài", format="%.3f"), "Rộng (m)": st.column_config.NumberColumn("Rộng", format="%.3f"), "Cao (m)": st.column_config.NumberColumn("Cao", format="%.3f"),
                "KL Tịnh (Net)": st.column_config.NumberColumn("Net", disabled=True, format="%.4f"), "KL Tổng (Gross)": st.column_config.NumberColumn("Gross", disabled=True, format="%.4f"),
                "Thành tiền": st.column_config.NumberColumn("Thành tiền", disabled=True, format="%d đ")
            })

        c_tool1, c_tool2 = st.columns([2.5, 2.5])
        with c_tool1:
            if st.button("🔍 Đồng Bộ Giá & Quy đổi Tự Động", use_container_width=True):
                updated_df = edited_df.copy()
                for idx in updated_df.index:
                    desc_val = str(updated_df.at[idx, 'Vật tư / Hạng mục']).strip()
                    code_val = str(updated_df.at[idx, 'Code']).strip()
                    loc_val = str(updated_df.at[idx, 'Vị trí']).strip()
                    unit_v = str(updated_df.at[idx, 'ĐVT']).strip()

                    d_m = float(updated_df.at[idx, 'Dài (m)'] or 0.0)
                    r_m = float(updated_df.at[idx, 'Rộng (m)'] or 0.0)
                    c_m = float(updated_df.at[idx, 'Cao (m)'] or 0.0)
                    sl = float(updated_df.at[idx, 'SL'] or 1.0)
                    w_pct = float(updated_df.at[idx, '% Hao hụt'] or 0.0)

                    target_sec, found_code, master_item = resolve_material_master(str(updated_df.at[idx, 'Mục chiết tính (I -> VII)']), code_val, desc_val, loc_val, unit_v)
                    
                    cost_val = float(updated_df.at[idx, 'Đơn giá'] or 0.0)
                    if master_item:
                        m_unit, m_cost, m_name = smart_unit_conversion(master_item['unit'], master_item['unit_cost'], master_item['name'], unit_v, r_m, c_m)
                        code_val = found_code; desc_val = m_name; unit_v = m_unit; cost_val = m_cost

                    net_q = calculate_clean_net_quantity(d_m, r_m, c_m, sl, unit_v)
                    waste_q = round(net_q * (w_pct / 100.0), 4 if "m3" in unit_v.lower() else 3)
                    gross_q = round(net_q + waste_q, 4 if "m3" in unit_v.lower() else 3)

                    updated_df.at[idx, 'Mục chiết tính (I -> VII)'] = target_sec
                    updated_df.at[idx, 'Code'] = code_val
                    updated_df.at[idx, 'Vật tư / Hạng mục'] = desc_val
                    updated_df.at[idx, 'ĐVT'] = unit_v
                    updated_df.at[idx, 'Đơn giá'] = cost_val
                    updated_df.at[idx, 'KL Tịnh (Net)'] = net_q
                    updated_df.at[idx, 'KL Tổng (Gross)'] = gross_q
                    updated_df.at[idx, 'Thành tiền'] = round(gross_q * cost_val, 0)
                
                updated_df['Mục chiết tính (I -> VII)'] = updated_df.apply(clean_section_row, axis=1)
                st.session_state[state_key] = updated_df
                st.success(f"✅ Đã chuẩn hóa quy đổi m3 -> m2/md thành công!"); st.rerun()

        with c_tool2:
            if st.button("💾 Lưu thay đổi bảng chiết xuất", type="primary", use_container_width=True):
                cur_save = conn.cursor()
                cur_save.execute("DELETE FROM boq_items WHERE project_id = ?", (sel_id,))
                new_total_direct = 0.0
                
                for _, row in edited_df.iterrows():
                    desc = str(row['Vật tư / Hạng mục']).strip(); m_code = str(row['Code']).strip(); loc = str(row['Vị trí'])
                    d_m = float(row['Dài (m)'] or 0.0); r_m = float(row['Rộng (m)'] or 0.0); c_m = float(row['Cao (m)'] or 0.0)
                    sl = float(row['SL'] or 1.0); unit = str(row['ĐVT']).strip()
                    s_code = clean_section_row(row)
                    
                    _, f_code, f_item = resolve_material_master(s_code, m_code, desc, loc, unit)
                    u_cost = float(row['Đơn giá'] or 0.0)
                    if f_item and u_cost == 0:
                        unit, u_cost, desc = smart_unit_conversion(f_item['unit'], f_item['unit_cost'], f_item['name'], unit, r_m, c_m)
                        m_code = f_code

                    net_q = calculate_clean_net_quantity(d_m, r_m, c_m, sl, unit)
                    w_pct = float(row['% Hao hụt'] or 0.0)
                    waste_q = round(net_q * (w_pct / 100.0), 4 if "m3" in unit.lower() else 3)
                    gross_q = round(net_q + waste_q, 4 if "m3" in unit.lower() else 3)
                    amt = round(gross_q * u_cost, 0)
                    new_total_direct += amt
                    
                    cur_save.execute('''INSERT INTO boq_items (project_id, section_code, material_code, description, location, dai_m, rong_m, cao_m, so_luong, unit, net_qty, waste_percent, waste_qty, gross_qty, unit_cost, amount, meaning, formula) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                        (sel_id, s_code, m_code, desc, loc, d_m, r_m, c_m, sl, unit, net_q, w_pct, waste_q, gross_q, u_cost, amt, f"{loc}: {d_m:.2f}x{r_m:.2f}x{sl:.0f}", f"{gross_q:.3f} x {u_cost:,.0f} đ"))

                total_overhead_pct_db = overhead_rate + pm_fee_rate + dlp_rate + consumables_rate
                new_overhead = new_total_direct * (total_overhead_pct_db / 100.0)
                new_tender = (new_total_direct + new_overhead) * (1 + profit_rate / 100.0)
                cur_save.execute("UPDATE projects SET direct_cost = ?, overhead_cost = ?, tender_price = ? WHERE id = ?", (new_total_direct, new_overhead, new_tender, sel_id))
                
                conn.commit()
                sync_projects_to_json()
                st.session_state.pop(state_key, None)
                st.success(f"✅ Đã lưu thành công! Giá vốn mới: {new_total_direct:,.0f} đ"); time.sleep(0.8); st.rerun()
    conn.close()

# -------------------------------------------------------------
# TAB 3: QUẢN LÝ ĐƠN GIÁ MASTER
# -------------------------------------------------------------
with tab3:
    st.subheader("📊 Quản Lý & Chỉnh Sửa Đơn Giá Master")
    if PRICE_MASTER:
        c_search, c_page = st.columns([3, 1])
        with c_search: search_kw = st.text_input("🔍 Tìm kiếm vật tư nhanh:")
        
        items_list = [{"MÃ CODE": k, "NHÓM": v.get("category", "CHUNG"), "TÊN VẬT TƯ / NHÂN CÔNG": v["name"], "ĐVT": v["unit"], "ĐƠN GIÁ XƯỞNG": v["unit_cost"]} for k, v in PRICE_MASTER.items() if not search_kw or search_kw.lower() in v["name"].lower() or search_kw.lower() in k.lower()]
        total_found = len(items_list)
        
        page_size = 50
        with c_page: current_page = st.number_input("Trang số:", min_value=1, max_value=max(math.ceil(total_found/page_size), 1), value=1)
        
        df_page = pd.DataFrame(items_list[(current_page-1)*page_size : min((current_page-1)*page_size + page_size, total_found)])
        edited_page_df = st.data_editor(df_page, use_container_width=True, num_rows="fixed", height=480, key=f"master_editor_{current_page}")

        if st.button(f"💾 Lưu cập nhật giá của Trang {current_page}", type="primary"):
            for _, r in edited_page_df.iterrows():
                if r["MÃ CODE"] in PRICE_MASTER: PRICE_MASTER[r["MÃ CODE"]].update({"name": r["TÊN VẬT TƯ / NHÂN CÔNG"], "unit": r["ĐVT"], "unit_cost": float(r["ĐƠN GIÁ XƯỞNG"])})
            pd.DataFrame([{"MÃ CODE": k, "NHÓM": v.get("category", "CHUNG"), "TÊN VẬT TƯ / NHÂN CÔNG": v["name"], "ĐVT": v["unit"], "ĐƠN GIÁ XƯỞNG": v["unit_cost"]} for k, v in PRICE_MASTER.items()]).to_excel("BIEU-MAU-DU-TOAN_MAPPED_MASTER.xlsx", index=False, sheet_name='DMVT')
            st.cache_data.clear()
            st.success("✅ Đã lưu giá mới vào Master!"); time.sleep(1); st.rerun()
