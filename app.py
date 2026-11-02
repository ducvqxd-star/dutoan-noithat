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
        return float(nums[0]), 0.0, float(nums[1])
    return 1000.0, 1000.0, 1000.0

# -------------------------------------------------------------
# 2. CƠ CHẾ SAO LƯU & KHÔI PHỤC DỮ LIỆU
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
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            password_hash TEXT,
            full_name TEXT,
            role TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            project_name TEXT,
            item_code TEXT,
            item_name TEXT,
            dimensions TEXT,
            specs TEXT,
            direct_cost REAL,
            overhead_cost REAL,
            tender_price REAL,
            technical_risks TEXT,
            image_paths TEXT,
            created_at TIMESTAMP,
            review_status TEXT DEFAULT 'Chờ duyệt',
            review_note TEXT DEFAULT '',
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS boq_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            section_code TEXT,
            material_code TEXT,
            description TEXT,
            location TEXT,
            dai_m REAL DEFAULT 0.0,
            rong_m REAL DEFAULT 0.0,
            cao_m REAL DEFAULT 0.0,
            so_luong REAL DEFAULT 1.0,
            unit TEXT,
            net_qty REAL,
            waste_percent REAL,
            waste_qty REAL,
            gross_qty REAL,
            unit_cost REAL,
            amount REAL,
            meaning TEXT,
            formula TEXT,
            FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
        )
    ''')
    admin_pw = hashlib.sha256("admin123".encode()).hexdigest()
    qs1_pw = hashlib.sha256("qs123".encode()).hexdigest()
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, full_name, role) VALUES ('admin', ?, 'Quản Trị Viên', 'admin')", (admin_pw,))
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, full_name, role) VALUES ('qs01', ?, 'Kỹ Sư QS 01', 'qs')", (qs1_pw,))
    conn.commit()
    conn.close()

init_db()
restore_projects_from_json()

# -------------------------------------------------------------
# 3. THUẬT TOÁN ĐỊNH MỨC NHÂN CÔNG TỔNG QUÁT THEO LOẠI SẢN PHẨM
# -------------------------------------------------------------
def calculate_dynamic_labor(length_m: float, depth_m: float, height_m: float, item_name: str, specs_text: str):
    """
    Tính định mức nhân công linh hoạt dựa trên LOẠI SẢN PHẨM THỰC TẾ:
    - Trần gỗ / Vách ốp: Tính theo diện tích mặt phẳng (Dài x Rộng hoặc Dài x Cao).
    - Tủ / Kệ: Tính theo diện tích mặt đứng kết cấu khối.
    - Bàn / Giường: Tính theo cái/bộ.
    """
    name_low = str(item_name).lower()
    spec_low = str(specs_text).lower()

    is_ceiling = any(k in name_low for k in ["trần", "tran", "la phông"])
    is_wall_panel = any(k in name_low for k in ["vách ốp", "vach op", "ốp tường", "lam sóng"])
    is_door = any(k in name_low for k in ["cửa", "cua"])
    is_table = any(k in name_low for k in ["bàn", "ban"])
    is_bed = any(k in name_low for k in ["giường", "giuong"])

    if is_ceiling:
        # Trần gỗ: Diện tích mặt bằng = Dài x Rộng
        area = max(length_m * depth_m, 1.0) if depth_m > 0 else max(length_m * height_m, 1.0)
        # Định mức thi công xưởng: gia công khung xương, cắt nan gỗ, sơn lót PU
        base_sx = area * 0.25
        # Định mức lắp đặt trần: dựng giàn giáo, bắn ty treo, cân laser, ốp nan trần trên cao
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
        # Hệ tủ khối (tủ áo, tủ bếp, tủ rượu)
        area = max(length_m * height_m, 1.0)
        base_sx = area * 0.35
        base_ld = area * 0.40

    # Phụ trội vật liệu nếu có chỉ định rõ ràng trong specs
    if any(k in spec_low for k in ["sơn pu", "sơn màu"]):
        base_sx += 0.8
    if any(k in spec_low for k in ["khung sắt", "sắt hộp", "thép"]):
        base_sx += 0.5
    if any(k in spec_low for k in ["đá", "nung kết"]) and not is_ceiling:
        base_ld += 0.8
    if any(k in spec_low for k in ["da", "nệm"]) and not is_ceiling:
        base_sx += 0.5

    prod_days = max(math.ceil(base_sx), 1)
    inst_days = max(math.ceil(base_ld), 1)
    return prod_days, inst_days

# -------------------------------------------------------------
# 4. BỘ NẠP MASTER EXCEL
# -------------------------------------------------------------
@st.cache_data(ttl=3600)
def load_price_master():
    candidate_files = [
        "BIEU-MAU-DU-TOAN_MAPPED_MASTER.xlsx",
        "BIEU-MAU-DU-TOAN_VERSION-2.xlsx",
        "Gia cost_updated.xlsx",
        "Gia cost.xlsx",
        "Gia_cost_CapNhat_AI.xlsx"
    ]
    target_file = None
    for fname in candidate_files:
        if os.path.exists(fname):
            target_file = fname
            break
            
    if not target_file:
        all_xlsx = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~$")]
        if all_xlsx:
            target_file = all_xlsx[0]
        else:
            return {}, "Không tìm thấy file Excel đơn giá"

    try:
        xls = pd.ExcelFile(target_file)
        target_sheet = None
        for s in xls.sheet_names:
            s_clean = s.strip().lower()
            if any(k in s_clean for k in ["dmvt", "đơn giá vt", "don gia vt", "dongia", "vt"]):
                target_sheet = s
                break
        if not target_sheet:
            target_sheet = xls.sheet_names[0]

        df_raw = pd.read_excel(target_file, sheet_name=target_sheet, header=None)
        
        col_code = 1
        col_name = 2
        col_unit = 3
        col_price = 14
        start_row = 3

        for r_idx in range(min(12, len(df_raw))):
            row_vals = [str(x).strip().lower() for x in df_raw.iloc[r_idx] if pd.notna(x)]
            if any("đơn giá" in v or "thành tiền" in v for v in row_vals):
                for c_idx, val in enumerate(df_raw.iloc[r_idx]):
                    val_str = str(val).strip().lower()
                    if "mã" in val_str and "vật tư" in val_str:
                        col_code = c_idx
                    elif "tên vật tư" in val_str or "hạng mục" in val_str:
                        col_name = c_idx
                    elif val_str == "đvt" or "đơn vị" in val_str:
                        col_unit = c_idx
                    elif val_str == "đơn giá" or val_str == "đơn giá tính":
                        col_price = c_idx
                start_row = r_idx + 1
                break

        price_dict = {}
        current_cat = "CHUNG"

        for r in range(start_row, len(df_raw)):
            code = df_raw.iloc[r, col_code] if len(df_raw.columns) > col_code else None
            name = df_raw.iloc[r, col_name] if len(df_raw.columns) > col_name else None
            unit = df_raw.iloc[r, col_unit] if len(df_raw.columns) > col_unit else 'cái'
            price = df_raw.iloc[r, col_price] if len(df_raw.columns) > col_price else 0

            if str(code).strip().lower() in ["mã vật tư", "code", "stt", "nan", ""]:
                continue

            if pd.isna(code) and pd.notna(name):
                current_cat = str(name).strip()
            elif pd.notna(code) and pd.notna(name):
                try:
                    if isinstance(price, (int, float)):
                        p_val = float(price)
                    else:
                        p_val = float(str(price).replace(",", "").replace(".", "").strip())
                except Exception:
                    p_val = 0.0
                
                if p_val > 0:
                    price_dict[str(code).strip()] = {
                        "category": current_cat,
                        "name": str(name).strip().replace('\n', ' '),
                        "unit": str(unit).strip() if pd.notna(unit) else 'cái',
                        "unit_cost": p_val
                    }
        
        # Bổ sung các mã vật tư định mức cơ bản nếu thiếu
        special_items = {
            "NC-SX": {"category": "NHÂN CÔNG", "name": "NC-SX: Nhân công sản xuất xưởng (Định mức)", "unit": "Công", "unit_cost": 500000.0},
            "NC-LD": {"category": "NHÂN CÔNG", "name": "NC-LD: Nhân công lắp đặt công trình (Định mức)", "unit": "Công", "unit_cost": 550000.0},
            "KS.V30": {"category": "KIM LOẠI", "name": "Khung thép hộp 20x40x1.2mm gia cố", "unit": "md", "unit_cost": 55000.0},
            "PU69": {"category": "SƠN HOÀN THIỆN", "name": "Sơn PU bóng (60-90%)", "unit": "m2", "unit_cost": 270000.0},
            "PP": {"category": "PHỤ LIỆU", "name": "Phụ liệu (Keo, vít, đinh, ke góc)", "unit": "Bộ", "unit_cost": 500000.0},
            "B.BBTN.1": {"category": "BAO BÌ", "name": "B.BBTN.1: Bao bì sản phẩm (Màng PE, carton)", "unit": "m2", "unit_cost": 25353.0},
            "ANC.MEL.3.03": {"category": "VÁN", "name": "MDF thường E2 melamine nhóm A - 17mm", "unit": "m2", "unit_cost": 165984.0},
            "ANC.MEL.3.01": {"category": "VÁN", "name": "MDF thường E2 melamine nhóm A - 9mm", "unit": "m2", "unit_cost": 120000.0},
            "ST.01": {"category": "ĐÁ", "name": "ST-01: Đá tam tính", "unit": "m2", "unit_cost": 3000000.0},
            "FB LE-01": {"category": "SOURCE NGOÀI", "name": "FB LE-01: Da tam tính (Bọc da trang trí)", "unit": "m2", "unit_cost": 500000.0},
            "KI.1.04": {"category": "SOURCE NGOÀI", "name": "Kính trong cường lực - 5 mm", "unit": "m2", "unit_cost": 1000000.0}
        }
        for k, v in special_items.items():
            if k not in price_dict:
                price_dict[k] = v

        return price_dict, f"{target_file} (Sheet: {target_sheet})"
    except Exception as e:
        return {}, f"Lỗi đọc file: {str(e)}"

PRICE_MASTER, SOURCE_INFO = load_price_master()

# -------------------------------------------------------------
# 5. TRUY VẤN HỒ SƠ ĐỐI SOÁT CÙNG CHỦNG LOẠI SẢN PHẨM
# -------------------------------------------------------------
def get_approved_benchmark(item_code: str, item_name: str):
    """CHỈ tìm kiếm hồ sơ cùng loại sản phẩm đã duyệt '✅ Đúng giá' để học tập"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Tìm chính xác theo tên sản phẩm hoặc mã hiệu
    c.execute("""
        SELECT id, project_name, item_name, dimensions, direct_cost 
        FROM projects 
        WHERE review_status = '✅ Đúng giá' AND (item_code = ? OR item_name = ?)
        ORDER BY id DESC LIMIT 1
    """, (item_code, item_name))
    row = c.fetchone()
    
    if not row:
        conn.close()
        return None, ""
        
    bench_id, p_name, itm_name, dims, d_cost = row
    items_df = pd.read_sql_query("""
        SELECT section_code, material_code, description, location, dai_m, rong_m, cao_m, so_luong, unit, waste_percent 
        FROM boq_items 
        WHERE project_id = ?
    """, conn, params=(bench_id,))
    conn.close()

    if items_df.empty:
        return None, ""

    bench_text = f"HỒ SƠ MẪU CHUẨN CÙNG LOẠI [{itm_name} - Kích thước: {dims} | Chi phí gốc: {d_cost:,.0f} đ]:\n"
    for _, itm in items_df.iterrows():
        bench_text += f"- [{itm['section_code']}] {itm['material_code']} | {itm['description']} ({itm['location']}) -> Dài:{itm['dai_m']}m x Rộng:{itm['rong_m']}m x Cao:{itm['cao_m']}m x SL:{itm['so_luong']} {itm['unit']} (Hao hụt:{itm['waste_percent']}%)\n"
    
    return bench_id, bench_text

# -------------------------------------------------------------
# 6. THUẬT TOÁN ĐỊNH DANH VẬT TƯ THỰC TẾ (KHÔNG ÉP BUỘC ĐÁ/DA)
# -------------------------------------------------------------
VALID_SECTIONS = [
    "I. VẬT TƯ THẦU PHỤ",
    "II. VẬT TƯ CHÍNH",
    "III. HOÀN THIỆN (KIM LOẠI, PU...)",
    "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ",
    "V. BAO BÌ",
    "VI. PHỤ KIỆN",
    "VII. SOURCE NGOÀI"
]

def resolve_material_master(section_code: str, raw_code: str, description: str, location: str, unit_hint: str):
    """Tra cứu và bảo toàn đúng chủng loại vật tư thực tế, không tự tiện đổi thành đá"""
    sec = str(section_code).strip().upper()
    loc_clean = str(location).strip().lower()
    desc_clean = str(description).strip().lower()
    full_text = loc_clean + " " + desc_clean
    raw_c = str(raw_code).strip()

    # 1. Nhân công
    if any(k in full_text for k in ["lắp đặt", "nc-ld", "công trình"]):
        return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ", "NC-LD", PRICE_MASTER.get("NC-LD")
    if any(k in full_text for k in ["sản xuất", "nc-sx", "thợ mộc", "xưởng"]):
        return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ", "NC-SX", PRICE_MASTER.get("NC-SX")
    if "sơn pu" in full_text or "sơn màu" in full_text:
        return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ", "PU69", PRICE_MASTER.get("PU69")
    if any(k in full_text for k in ["keo", "vít", "phụ liệu", "đinh"]):
        return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ", "PP", PRICE_MASTER.get("PP")

    # 2. Bao bì
    if any(k in full_text for k in ["bao bì", "màng pe", "carton", "bọc hàng"]):
        return "V. BAO BÌ", "B.BBTN.1", PRICE_MASTER.get("B.BBTN.1")

    # 3. Khung thép / sắt hộp
    if any(k in full_text for k in ["sắt hộp", "khung sắt", "khung thép", "xương sắt", "thép hộp"]):
        return "II. VẬT TƯ CHÍNH", "KS.V30", PRICE_MASTER.get("KS.V30")

    # 4. Đá thật: BẮT BUỘC có chữ 'đá' độc lập và KHÔNG nằm trong từ 'đáy', KHÔNG phải 'vân đá'
    has_real_stone = bool(re.search(r'(?:\b|\s|^)đá(?:\s|$|[,\.\-_])', full_text)) and "vân đá" not in full_text and "đáy" not in loc_clean
    if has_real_stone or "đá tam tính" in full_text:
        return "VII. SOURCE NGOÀI", "ST.01", PRICE_MASTER.get("ST.01")

    # 5. Da thật / nệm bọc
    if any(k in full_text for k in ["bọc da", "da tam tính", "nệm bọc da", "leather"]):
        return "VII. SOURCE NGOÀI", "FB LE-01", PRICE_MASTER.get("FB LE-01")

    # 6. Kính
    if any(k in full_text for k in ["cánh kính", "kính cường lực", "gương", "kính trong"]):
        return "VII. SOURCE NGOÀI", "KI.1.04", PRICE_MASTER.get("KI.1.04")

    # 7. Khớp trực tiếp mã Master nếu có
    if raw_c in PRICE_MASTER:
        return section_code, raw_c, PRICE_MASTER[raw_c]

    # 8. Tìm kiếm theo từ khóa trong Master
    words = [w for w in re.split(r'[,;\-\s_/\(\)\+]+', desc_clean) if len(w) >= 2]
    best_code, best_data, best_score = None, None, 0
    for p_code, p_data in PRICE_MASTER.items():
        p_name_low = p_data["name"].lower()
        score = sum(3 for w in words if w in p_name_low)
        if desc_clean in p_name_low:
            score += 5
        if score > best_score:
            best_score = score
            best_code = p_code
            best_data = p_data

    if best_score >= 3 and best_data:
        return section_code, best_code, best_data

    return section_code, "CHƯA_CÓ_MÃ", None

def clean_section_row(row):
    val = row.get('Mục chiết tính (I -> VII)')
    if val in VALID_SECTIONS:
        return val

    code = str(row.get('Code', '')).strip().upper()
    loc = str(row.get('Vị trí', '')).strip().lower()
    name = str(row.get('Vật tư / Hạng mục', '')).strip().lower()
    full = loc + " " + name

    if code.startswith("NC") or any(k in full for k in ["nhân công", "vật liệu phụ", "keo", "vít", "đinh"]):
        return "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ"
    if code.startswith("DA.") or code.startswith("ST.") or code.startswith("FB") or code.startswith("KI.") or any(k in full for k in ["kính", "da tam tính", "mặt kệ đá", "vách đá"]):
        return "VII. SOURCE NGOÀI"
    if code.startswith("PK") or any(k in full for k in ["bản lề", "ray", "tay nắm", "khóa", "led", "đèn"]):
        return "VI. PHỤ KIỆN"
    if code.startswith("BB") or any(k in full for k in ["bao bì", "pe", "carton"]):
        return "V. BAO BÌ"
    if code.startswith("CC") or code.startswith("IN") or any(k in full for k in ["chỉ cạnh", "nẹp", "inox", "sơn", "pu"]):
        return "III. HOÀN THIỆN (KIM LOẠI, PU...)"

    return "II. VẬT TƯ CHÍNH"

# -------------------------------------------------------------
# 7. QUẢN LÝ ĐĂNG NHẬP
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
    else:
        st.error("Sai tên đăng nhập hoặc mật khẩu!")

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
        if st.button("Đăng nhập", type="primary", use_container_width=True):
            login(u_in, p_in)
        st.caption("Tài khoản: `admin` / `admin123`")
    st.stop()

# -------------------------------------------------------------
# 8. CẤU TRÚC DỮ LIỆU ĐỊNH NGHĨA PHÂN RÃ CẤU KIỆN
# -------------------------------------------------------------
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

# -------------------------------------------------------------
# 9. GIAO DIỆN CHÍNH
# -------------------------------------------------------------
user = st.session_state['user_info']

with st.sidebar:
    if LOGO_FILE:
        st.image(LOGO_FILE, use_container_width=True)
        st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(f"👤 **{user['full_name']}** (`{user['role']}`)")
    if st.button("Đăng xuất", use_container_width=True):
        logout()
    st.markdown("---")
    st.header("⚙️ Cấu hình hệ thống")
    
    if CLOUD_API_KEY:
        api_key = CLOUD_API_KEY
        st.success("✅ Đã kết nối Gemini API từ Cloud")
    else:
        api_key = st.text_input("Gemini API Key:", type="password")
        
    overhead_rate = st.slider("Chi phí chung (Overhead) %:", 0, 25, 8)
    profit_rate = st.slider("Lợi nhuận mục tiêu (Margin) %:", 0, 30, 12)
    
    st.caption(f"Nguồn dữ liệu: `{SOURCE_INFO}`")
    st.caption(f"Đã nạp **{len(PRICE_MASTER)}** mã đơn giá Master.")

    st.markdown("---")
    st.subheader("💾 Sao lưu & Khôi phục")
    if os.path.exists(JSON_STORAGE_FILE):
        with open(JSON_STORAGE_FILE, "rb") as fp_j:
            st.download_button(
                "📥 Tải file JSON sao lưu hồ sơ",
                data=fp_j.read(),
                file_name="saved_projects_backup.json",
                mime="application/json",
                use_container_width=True
            )

tab1, tab2, tab3 = st.tabs(["📝 Bóc Tách Dự Toán Mới", "📂 Thư Viện Hồ Sơ & Chỉnh Sửa Chiết Tính", "📊 Quản Lý Đơn Giá Master"])

# TAB 1: BÓC TÁCH MỚI (DYNAMIC THEO BẢN VẼ)
with tab1:
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("1. Dữ liệu bản vẽ & Thông số kỹ thuật")
        project_name = st.text_input("Tên dự án / Công trình:", value="Biệt thự Park City")
        item_code = st.text_input("Mã sản phẩm / Hạng mục:", value="TG-01")
        item_name = st.text_input("Tên sản phẩm:", value="Trần gỗ tiêu âm phòng khách")
        dimensions = st.text_input("Kích thước bao (DxRxC mm):", value="5000 x 4000 x 150 mm")
        
        uploaded_files = st.file_uploader("Upload ảnh 3D hoặc bản vẽ kỹ thuật:", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        if uploaded_files:
            c_preview = st.columns(min(len(uploaded_files), 3))
            for idx, f in enumerate(uploaded_files):
                with c_preview[idx % 3]:
                    st.image(f, use_container_width=True)

        specs = st.text_area(
            "Quy cách kỹ thuật & Vật liệu chỉ định:",
            height=140, 
            value="- Khung xương sắt hộp mạ kẽm 20x40x1.2mm bắt ty treo chống rung.\n- Tấm ốp trần MDF chống ẩm phủ Melamine/Veneer sồi xẻ rãnh tiêu âm.\n- Phào cổ trần và nẹp chỉ viền gỗ sồi tự nhiên sơn PU hoàn thiện.\n- Đèn rãnh LED hắt trần viền xung quanh."
        )

        submit_btn = st.button("🚀 Bắt đầu bóc tách dự toán (Chính xác theo bản vẽ)", type="primary", use_container_width=True)

    with c2:
        st.subheader("2. Chi phí tổng quan")
        if submit_btn:
            if not api_key:
                st.error("Vui lòng cấu hình Gemini API Key.")
            elif not uploaded_files:
                st.warning("Vui lòng tải lên ít nhất 1 ảnh bản vẽ kỹ thuật.")
            else:
                with st.spinner(f"AI QS đang phân tích bản vẽ {item_name} (chỉ bóc vật tư có thực)..."):
                    try:
                        client = genai.Client(api_key=api_key)
                        images = []
                        saved_paths = []
                        for idx, f in enumerate(uploaded_files):
                            img = Image.open(f)
                            img.thumbnail((1600, 1600))
                            images.append(img)
                            spath = os.path.join(IMAGE_DIR, f"{user['username']}_{item_code}_{idx}_{f.name}")
                            img.save(spath)
                            saved_paths.append(spath)

                        l_mm, d_mm, h_mm = parse_dimensions(dimensions)
                        l_m = l_mm / 1000.0
                        d_m = d_mm / 1000.0
                        h_m = h_mm / 1000.0

                        prod_days_calc, inst_days_calc = calculate_dynamic_labor(l_m, d_m, h_m, item_name, specs)
                        bench_id, benchmark_context = get_approved_benchmark(item_code, item_name)

                        # PROMPT TỔNG QUÁT: TUYỆT ĐỐI KHÔNG ÉP CỨNG ĐÁ HOẶC DA
                        prompt = f"""
                        BẠN LÀ SENIOR QS NỘI THẤT. BÓC TÁCH CHI TIẾT TỪNG TẤM (CUTTING LIST / BOM) DỰA THEO BẢN VẼ VÀ QUY CÁCH:
                        - Tên sản phẩm: {item_name}
                        - Mã sản phẩm: {item_code}
                        - Kích thước phủ bì: Dài {l_m:.3f}m x Rộng {d_m:.3f}m x Cao {h_m:.3f}m ({l_mm:.0f}x{d_mm:.0f}x{h_mm:.0f}mm).
                        - Quy cách kỹ thuật:
                        {specs}

                        NGUYÊN TẮC BÓC TÁCH SỐ 1: CHỈ BÓC TÁCH NHỮNG VẬT TƯ THỰC SỰ CÓ TRONG BẢN VẼ VÀ MÔ TẢ KỸ THUẬT.
                        - TUYỆT ĐỐI KHÔNG TỰ BỊA ĐÁ, DA, NỆM, HOẶC KÍNH NẾU SẢN PHẨM KHÔNG CÓ (Ví dụ: Trần gỗ thì KHÔNG ĐƯỢC bóc đá hay da).
                        - Phân loại chính xác từng dòng vào 7 nhóm (I -> VII):
                          I. VẬT TƯ THẦU PHỤ (Gia công ngoài đặc thù nếu có)
                          II. VẬT TƯ CHÍNH (Gỗ xẻ, ván MDF/HDF, tấm tiêu âm, khung thép hộp gia cố, ty treo...)
                          III. HOÀN THIỆN (Chỉ nẹp viền, phào cổ trần, nẹp kim loại Inox/nhôm, chỉ dán cạnh...)
                          IV. NHÂN CÔNG & PHỤ LIỆU:
                              * BẮT BUỘC có 2 dòng nhân công định mức:
                                + 'NC-SX: Nhân công sản xuất xưởng (Định mức)' -> đúng {prod_days_calc} công.
                                + 'NC-LD: Nhân công lắp đặt công trình (Định mức)' -> đúng {inst_days_calc} công.
                              * Sơn PU bóng hoàn thiện bề mặt (m2).
                              * Phụ liệu keo vít, đinh, bát treo (Bộ).
                          V. BAO BÌ (Màng bọc PE, carton bảo vệ)
                          VI. PHỤ KIỆN (Bản lề, ray, đèn LED hắt nếu có trong bản vẽ)
                          VII. SOURCE NGOÀI (CHỈ BÓC NẾU CÓ TRONG THIẾT KẾ: Đá mặt bàn/kệ, Kính, Gương, Da bọc nệm. NẾU KHÔNG CÓ THÌ BỎ QUA).
                        """

                        if benchmark_context:
                            prompt += f"""
                            ----------------------------------------------------------------------
                            THAM CHIẾU HỒ SƠ CHUẨN CÙNG LOẠI ĐÃ ĐƯỢC ADMIN DUYỆT:
                            {benchmark_context}
                            Hãy kế thừa danh mục và tỷ lệ hao hụt của hồ sơ chuẩn trên.
                            ----------------------------------------------------------------------
                            """

                        MODELS_TO_TRY = ['gemini-3.8-flash', 'gemini-3.5-flash-lite']
                        response = None
                        last_err = None

                        for current_model in MODELS_TO_TRY:
                            for attempt in range(3):
                                try:
                                    response = client.models.generate_content(
                                        model=current_model,
                                        contents=[*images, prompt],
                                        config=types.GenerateContentConfig(
                                            system_instruction="BẠN LÀ SENIOR QS NỘI THẤT. Chỉ bóc tách vật tư có thực trên bản vẽ, tuân thủ nguyên tắc trung thực kỹ thuật, không tự bịa vật tư.",
                                            response_mime_type="application/json",
                                            response_schema=TakeoffResult,
                                            temperature=0.0
                                        )
                                    )
                                    break
                                except Exception as err:
                                    last_err = err
                                    err_str = str(err)
                                    if "RESOURCE_EXHAUSTED" in err_str or "quota" in err_str.lower():
                                        break
                                    elif any(k in err_str for k in ["503", "UNAVAILABLE"]):
                                        time.sleep(2 * (attempt + 1))
                                    else:
                                        break
                            if response:
                                break

                        if not response:
                            raise last_err

                        data = json.loads(response.text)
                        
                        total_direct = 0
                        boq_records = []
                        for comp in data.get("breakdown_items", []):
                            s_code = comp.get("section_code", "II. VẬT TƯ CHÍNH")
                            raw_code = str(comp.get("material_code", "")).strip()
                            ai_desc = comp.get("description", "").strip()
                            loc = comp.get("location", "").strip()
                            unit_str = str(comp.get("unit", "cái")).strip()
                            
                            d_m = float(comp.get("dai_m", 0.0))
                            r_m = float(comp.get("rong_m", 0.0))
                            c_m = float(comp.get("cao_m", 0.0))
                            sl = float(comp.get("so_luong", 1.0))
                            w_pct = float(comp.get("waste_factor_percent", 0.0))

                            unit_lower = unit_str.lower()
                            if "m3" in unit_lower:
                                net_q = round(d_m * r_m * c_m * sl, 3)
                            elif "m2" in unit_lower:
                                net_q = round(d_m * r_m * sl, 3) if r_m > 0 else round(d_m * c_m * sl, 3)
                            elif "md" in unit_lower or "m" == unit_lower:
                                net_q = round(d_m * sl, 3)
                            else:
                                net_q = round(sl, 2)

                            final_sec, matched_code, item_info = resolve_material_master(s_code, raw_code, ai_desc, loc, unit_str)

                            if item_info:
                                u_cost = item_info["unit_cost"]
                                final_desc = item_info["name"]
                                final_unit = item_info["unit"]
                            else:
                                matched_code = "CHƯA_CÓ_MÃ"
                                u_cost = 0.0
                                final_desc = ai_desc
                                final_unit = unit_str

                            is_discrete = any(u in final_unit.lower() for u in ['cai', 'cái', 'bộ', 'bo', 'con', 'chiếc', 'chiec', 'hệ', 'he', 'công', 'cong'])
                            if is_discrete:
                                waste_q = math.ceil(net_q * (w_pct / 100.0)) if w_pct > 0 else 0
                                gross_q = net_q + waste_q
                            else:
                                waste_q = round(net_q * (w_pct / 100.0), 3)
                                gross_q = round(net_q + waste_q, 3)

                            amt = gross_q * u_cost
                            total_direct += amt
                            
                            meaning_text = f"{loc}: {d_m:.2f}x{r_m:.2f}x{sl:.0f}"
                            formula_text = f"{gross_q:.2f} x {u_cost:,.0f} đ"

                            boq_records.append({
                                "section_code": final_sec,
                                "material_code": matched_code,
                                "description": final_desc,
                                "location": loc,
                                "dai_m": d_m,
                                "rong_m": r_m,
                                "cao_m": c_m,
                                "so_luong": sl,
                                "unit": final_unit,
                                "net_qty": net_q,
                                "waste_percent": w_pct,
                                "waste_qty": waste_q,
                                "gross_qty": gross_q,
                                "unit_cost": u_cost,
                                "amount": amt,
                                "meaning": meaning_text,
                                "formula": formula_text
                            })

                        overhead_amt = total_direct * (overhead_rate / 100.0)
                        tender_amt = (total_direct + overhead_amt) * (1 + profit_rate / 100.0)
                        created_time_vn = get_vietnam_time()

                        conn = sqlite3.connect(DB_FILE)
                        cur = conn.cursor()
                        cur.execute('''
                            INSERT INTO projects 
                            (user_id, project_name, item_code, item_name, dimensions, specs, direct_cost, overhead_cost, tender_price, technical_risks, image_paths, created_at, review_status, review_note)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Chờ duyệt', '')
                        ''', (
                            user["id"], project_name, item_code, item_name, dimensions, specs,
                            total_direct, overhead_amt, tender_amt,
                            json.dumps(data.get("technical_risks_and_notes", []), ensure_ascii=False),
                            json.dumps(saved_paths), created_time_vn
                        ))
                        proj_id = cur.lastrowid
                        
                        for r in boq_records:
                            cur.execute('''
                                INSERT INTO boq_items 
                                (project_id, section_code, material_code, description, location, dai_m, rong_m, cao_m, so_luong, unit, net_qty, waste_percent, waste_qty, gross_qty, unit_cost, amount, meaning, formula)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', (
                                proj_id, r["section_code"], r["material_code"], r["description"], r["location"],
                                r["dai_m"], r["rong_m"], r["cao_m"], r["so_luong"],
                                r["unit"], r["net_qty"], r["waste_percent"], r["waste_qty"], r["gross_qty"], 
                                r["unit_cost"], r["amount"], r["meaning"], r["formula"]
                            ))
                        conn.commit()
                        conn.close()

                        sync_projects_to_json()

                        st.success(f"✅ Bóc tách chiết tính thành công lúc {created_time_vn} (Giờ VN)!")
                        
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Tổng giá Cost xưởng (Direct)", f"{total_direct:,.0f} đ")
                        m2.metric(f"Chi phí chung ({overhead_rate}%)", f"{overhead_amt:,.0f} đ")
                        m3.metric(f"Giá chào thầu ({profit_rate}%)", f"{tender_amt:,.0f} đ")

                        st.info(f"👷 **Định mức nhân công {item_name}:** Sản xuất xưởng: **{prod_days_calc} công** (NC-SX: 500k/công) | Lắp đặt: **{inst_days_calc} công** (NC-LD: 550k/công) — Làm tròn lớn hơn.")

                    except Exception as e:
                        st.error(f"Lỗi xử lý: {str(e)}")

# TAB 2: THƯ VIỆN HỒ SƠ & BẢNG CHỈNH SỬA CHIẾT TÍNH ĐẦY ĐỦ
with tab2:
    st.subheader(f"📂 Thư Viện Hồ Sơ & Chỉnh Sửa Chiết Tính Chi Tiết")
    conn = sqlite3.connect(DB_FILE)
    
    query = """
        SELECT 
            p.id, 
            u.full_name AS nguoi_tao,
            p.project_name, 
            p.item_code, 
            p.item_name, 
            p.dimensions, 
            p.direct_cost, 
            p.tender_price, 
            p.review_status,
            p.review_note,
            p.created_at,
            p.user_id
        FROM projects p
        LEFT JOIN users u ON p.user_id = u.id
    """
    params = ()
    if user['role'] != 'admin':
        query += " WHERE p.user_id = ?"
        params = (user['id'],)
    query += " ORDER BY p.id DESC"
    
    df_saved = pd.read_sql_query(query, conn, params=params)
    
    if df_saved.empty:
        st.info("Chưa có hồ sơ dự toán nào được lưu. Bạn có thể sang Tab 1 để bóc tách mới sản phẩm bất kỳ!")
    else:
        df_display = df_saved.rename(columns={
            "id": "ID",
            "nguoi_tao": "Người tạo",
            "project_name": "Dự án",
            "item_code": "Mã hiệu",
            "item_name": "Tên sản phẩm",
            "dimensions": "Kích thước (DxRxC)",
            "direct_cost": "Chi phí gốc",
            "tender_price": "Giá chào thầu",
            "review_status": "Đánh giá Admin",
            "review_note": "Ghi chú kiểm duyệt",
            "created_at": "Thời gian tạo (Giờ VN)"
        })

        st.dataframe(
            df_display[["ID", "Người tạo", "Dự án", "Mã hiệu", "Tên sản phẩm", "Kích thước (DxRxC)", "Chi phí gốc", "Giá chào thầu", "Đánh giá Admin", "Ghi chú kiểm duyệt", "Thời gian tạo (Giờ VN)"]].style.format({
                "Chi phí gốc": "{:,.0f} đ",
                "Giá chào thầu": "{:,.0f} đ"
            }),
            use_container_width=True
        )
        
        st.markdown("---")
        
        st.markdown("### 🛠️ Quản Trị & Đào Tạo Bộ Não AI")
        sel_col1, sel_col2, sel_col3 = st.columns([1.5, 2.5, 1])
        
        with sel_col1:
            sel_id = st.selectbox("Chọn ID hồ sơ cần duyệt / xem:", df_saved["id"].tolist())
            selected_row = df_saved[df_saved["id"] == sel_id].iloc[0]
            st.caption(f"Hạng mục: **{selected_row['item_name']}** | Tạo bởi: **{selected_row['nguoi_tao']}**")

        with sel_col2:
            if user['role'] == 'admin':
                cur_status = selected_row['review_status'] if pd.notna(selected_row['review_status']) else 'Chờ duyệt'
                status_options = ['Chờ duyệt', '✅ Đúng giá', '❌ Sai giá']
                status_idx = status_options.index(cur_status) if cur_status in status_options else 0
                
                c_rev1, c_rev2 = st.columns([1.5, 2.5])
                with c_rev1:
                    new_status = st.radio("Đánh giá Admin:", status_options, index=status_idx, horizontal=True)
                with c_rev2:
                    new_note = st.text_input("Ghi chú Admin:", value=selected_row['review_note'] if pd.notna(selected_row['review_note']) else '')
                
                if st.button("💾 Lưu phê duyệt & Đào tạo AI", type="primary"):
                    cur_u = conn.cursor()
                    cur_u.execute("UPDATE projects SET review_status = ?, review_note = ? WHERE id = ?", (new_status, new_note, sel_id))
                    conn.commit()
                    sync_projects_to_json()
                    st.success("Đã lưu kết quả kiểm duyệt!")
                    time.sleep(0.5)
                    st.rerun()
            else:
                st.info(f"Trạng thái: **{selected_row['review_status']}** | Ghi chú: {selected_row['review_note'] if selected_row['review_note'] else 'Không có'}")

        with sel_col3:
            st.write("")
            st.write("")
            if user['role'] == 'admin' or selected_row['user_id'] == user['id']:
                if st.button(f"🗑️ Xóa hồ sơ #{sel_id}", type="secondary", use_container_width=True):
                    cur_d = conn.cursor()
                    cur_d.execute("DELETE FROM boq_items WHERE project_id = ?", (sel_id,))
                    cur_d.execute("DELETE FROM projects WHERE id = ?", (sel_id,))
                    conn.commit()
                    sync_projects_to_json()
                    st.warning(f"Đã xóa hoàn toàn hồ sơ #{sel_id}.")
                    time.sleep(0.5)
                    st.rerun()

        st.markdown("---")
        st.markdown(f"### ✏️ Bảng Chi Tiết Chiết Tính & Kích Thước Hình Học (Hồ Sơ #{sel_id})")

        state_key = f"df_edit_data_{sel_id}"
        if state_key not in st.session_state:
            df_init = pd.read_sql_query(
                """SELECT 
                    id AS item_db_id,
                    section_code AS 'Mục chiết tính (I -> VII)', 
                    material_code AS 'Code', 
                    description AS 'Vật tư / Hạng mục', 
                    location AS 'Vị trí',
                    dai_m AS 'Dài (m)',
                    rong_m AS 'Rộng (m)',
                    cao_m AS 'Cao (m)',
                    so_luong AS 'SL',
                    unit AS 'ĐVT', 
                    net_qty AS 'KL Tịnh (Net)', 
                    waste_percent AS '% Hao hụt', 
                    gross_qty AS 'KL Tổng (Gross)', 
                    unit_cost AS 'Đơn giá', 
                    amount AS 'Thành tiền'
                FROM boq_items WHERE project_id = ?""",
                conn, params=(sel_id,)
            )
            df_init['Mục chiết tính (I -> VII)'] = df_init.apply(clean_section_row, axis=1)
            st.session_state[state_key] = df_init

        # Bảng Data Editor chỉnh sửa trực tiếp hồ sơ
        edited_df = st.data_editor(
            st.session_state[state_key],
            use_container_width=True,
            num_rows="dynamic",
            height=540,
            key=f"editor_table_{sel_id}",
            column_config={
                "item_db_id": None,
                "Mục chiết tính (I -> VII)": st.column_config.SelectboxColumn(
                    "Mục chiết tính (I -> VII)",
                    options=VALID_SECTIONS,
                    required=True
                ),
                "Code": st.column_config.TextColumn("Code (Mã Master)"),
                "Vật tư / Hạng mục": st.column_config.TextColumn("Vật tư / Hạng mục", required=True),
                "Vị trí": st.column_config.TextColumn("Vị trí"),
                "Dài (m)": st.column_config.NumberColumn("Dài (m)", min_value=0.0, format="%.3f"),
                "Rộng (m)": st.column_config.NumberColumn("Rộng (m)", min_value=0.0, format="%.3f"),
                "Cao (m)": st.column_config.NumberColumn("Cao (m)", min_value=0.0, format="%.3f"),
                "SL": st.column_config.NumberColumn("SL", min_value=0.0, format="%.1f"),
                "ĐVT": st.column_config.TextColumn("ĐVT", required=True),
                "KL Tịnh (Net)": st.column_config.NumberColumn("KL Tịnh (Net)", min_value=0.0, format="%.3f"),
                "% Hao hụt": st.column_config.NumberColumn("% Hao hụt", min_value=0.0, max_value=100.0, format="%.1f%%"),
                "KL Tổng (Gross)": st.column_config.NumberColumn("KL Tổng (Gross)", disabled=True, format="%.3f"),
                "Đơn giá": st.column_config.NumberColumn("Đơn giá", min_value=0.0, format="%d đ"),
                "Thành tiền": st.column_config.NumberColumn("Thành tiền", disabled=True, format="%d đ")
            }
        )

        c_tool1, c_tool2, c_tool3 = st.columns([2.4, 1.8, 1.8])
        with c_tool1:
            if st.button("🔍 Tra Cứu & Đồng Bộ Giá Tự Động Theo Tên", use_container_width=True):
                updated_df = edited_df.copy()
                match_count = 0
                for idx in updated_df.index:
                    desc_val = str(updated_df.at[idx, 'Vật tư / Hạng mục']).strip()
                    code_val = str(updated_df.at[idx, 'Code']).strip()
                    loc_val = str(updated_df.at[idx, 'Vị trí']).strip()
                    sec_val = str(updated_df.at[idx, 'Mục chiết tính (I -> VII)']).strip()
                    cost_val = float(updated_df.at[idx, 'Đơn giá'] or 0.0)

                    target_sec, found_code, master_item = resolve_material_master(sec_val, code_val, desc_val, loc_val, str(updated_df.at[idx, 'ĐVT']))
                    
                    if master_item and (cost_val == 0.0 or code_val != found_code or sec_val != target_sec):
                        updated_df.at[idx, 'Mục chiết tính (I -> VII)'] = target_sec
                        updated_df.at[idx, 'Code'] = found_code
                        updated_df.at[idx, 'Vật tư / Hạng mục'] = master_item['name']
                        updated_df.at[idx, 'ĐVT'] = master_item['unit']
                        updated_df.at[idx, 'Đơn giá'] = master_item['unit_cost']
                        
                        net_q = float(updated_df.at[idx, 'KL Tịnh (Net)'] or 1.0)
                        w_pct = float(updated_df.at[idx, '% Hao hụt'] or 0.0)
                        gross_q = round(net_q * (1 + w_pct / 100.0), 3)
                        updated_df.at[idx, 'KL Tổng (Gross)'] = gross_q
                        updated_df.at[idx, 'Thành tiền'] = round(gross_q * master_item['unit_cost'], 0)
                        match_count += 1
                
                updated_df['Mục chiết tính (I -> VII)'] = updated_df.apply(clean_section_row, axis=1)
                st.session_state[state_key] = updated_df
                st.success(f"✅ Đã đồng bộ giá chuẩn xác cho {match_count} cấu kiện!")
                st.rerun()

        with c_tool2:
            if st.button("💾 Lưu thay đổi bảng chiết xuất", type="primary", use_container_width=True):
                cur_save = conn.cursor()
                cur_save.execute("DELETE FROM boq_items WHERE project_id = ?", (sel_id,))
                
                new_total_direct = 0.0
                
                for _, row in edited_df.iterrows():
                    desc = str(row['Vật tư / Hạng mục']).strip() if pd.notna(row['Vật tư / Hạng mục']) else ""
                    m_code = str(row['Code']).strip() if pd.notna(row['Code']) else "CHƯA_CÓ_MÃ"
                    s_code = clean_section_row(row)
                    loc = str(row['Vị trí']) if pd.notna(row['Vị trí']) else ""
                    
                    d_m = float(row['Dài (m)']) if pd.notna(row['Dài (m)']) else 0.0
                    r_m = float(row['Rộng (m)']) if pd.notna(row['Rộng (m)']) else 0.0
                    c_m = float(row['Cao (m)']) if pd.notna(row['Cao (m)']) else 0.0
                    sl = float(row['SL']) if pd.notna(row['SL']) else 1.0
                    unit = str(row['ĐVT']) if pd.notna(row['ĐVT']) else "cái"
                    
                    net_q = float(row['KL Tịnh (Net)']) if pd.notna(row['KL Tịnh (Net)']) else 0.0
                    if net_q == 0.0:
                        u_low = unit.lower()
                        if "m3" in u_low:
                            net_q = round(d_m * r_m * c_m * sl, 3)
                        elif "m2" in u_low:
                            net_q = round(d_m * r_m * sl, 3) if r_m > 0 else round(d_m * c_m * sl, 3)
                        elif "md" in u_low:
                            net_q = round(d_m * sl, 3)
                        else:
                            net_q = sl

                    w_pct = float(row['% Hao hụt']) if pd.notna(row['% Hao hụt']) else 0.0
                    u_cost = float(row['Đơn giá']) if pd.notna(row['Đơn giá']) else 0.0
                    
                    if u_cost == 0.0 or m_code in ["CHƯA_CÓ_MÃ", "None", "nan", ""]:
                        _, f_code, f_item = resolve_material_master(s_code, m_code, desc, loc, unit)
                        if f_item:
                            m_code = f_code
                            desc = f_item['name']
                            unit = f_item['unit']
                            u_cost = f_item['unit_cost']

                    waste_q = round(net_q * (w_pct / 100.0), 3)
                    gross_q = round(net_q + waste_q, 3)
                    amt = round(gross_q * u_cost, 0)
                    new_total_direct += amt
                    
                    meaning = f"{loc}: {d_m:.2f}x{r_m:.2f}x{sl:.0f}"
                    formula = f"{gross_q:.2f} x {u_cost:,.0f} đ"
                    
                    cur_save.execute('''
                        INSERT INTO boq_items 
                        (project_id, section_code, material_code, description, location, dai_m, rong_m, cao_m, so_luong, unit, net_qty, waste_percent, waste_qty, gross_qty, unit_cost, amount, meaning, formula)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        sel_id, s_code, m_code, desc, loc, d_m, r_m, c_m, sl,
                        unit, net_q, w_pct, waste_q, gross_q, u_cost, amt, meaning, formula
                    ))

                new_overhead = new_total_direct * (overhead_rate / 100.0)
                new_tender = (new_total_direct + new_overhead) * (1 + profit_rate / 100.0)
                
                cur_save.execute('''
                    UPDATE projects 
                    SET direct_cost = ?, overhead_cost = ?, tender_price = ? 
                    WHERE id = ?
                ''', (new_total_direct, new_overhead, new_tender, sel_id))
                
                conn.commit()
                sync_projects_to_json()
                st.session_state.pop(state_key, None)
                st.success(f"✅ Đã lưu thành công! Tổng chi phí mới: {new_total_direct:,.0f} đ | Giá chào thầu: {new_tender:,.0f} đ")
                time.sleep(0.8)
                st.rerun()

        with c_tool3:
            buf_re = io.BytesIO()
            with pd.ExcelWriter(buf_re, engine='openpyxl') as wr:
                df_to_export = edited_df.drop(columns=['item_db_id'], errors='ignore')
                df_to_export.to_excel(wr, index=False, sheet_name='ChiTiet_BOM')
            st.download_button(
                f"📥 Tải file Excel chiết tính (#{sel_id})", 
                data=buf_re.getvalue(), 
                file_name=f"BOM_Chiết_Tính_ID_{sel_id}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    conn.close()

# -------------------------------------------------------------
# TAB 3: QUẢN LÝ ĐƠN GIÁ MASTER (PHÂN TRANG TIẾT KIỆM RAM)
# -------------------------------------------------------------
with tab3:
    st.subheader("📊 Quản Lý & Chỉnh Sửa Đơn Giá Master")
    st.caption(f"Nguồn dữ liệu hiện hành: `{SOURCE_INFO}` | Tổng số mã: **{len(PRICE_MASTER)}**")

    with st.expander("📤 **Tải lên file Excel Đơn giá mới (Ghi đè file hiện tại)**", expanded=False):
        uploaded_master = st.file_uploader(
            "Chọn file Excel đơn giá mới (.xlsx):", 
            type=["xlsx"], 
            key="uploader_master_file"
        )
        if uploaded_master is not None:
            if st.button("Xác nhận nạp file Master này", type="primary"):
                save_master_name = "BIEU-MAU-DU-TOAN_MAPPED_MASTER.xlsx"
                with open(save_master_name, "wb") as f:
                    f.write(uploaded_master.getbuffer())
                
                st.cache_data.clear()
                st.success("✅ Đã cập nhật file Master thành công! Đang tải lại dữ liệu...")
                time.sleep(1)
                st.rerun()

    st.markdown("---")

    if PRICE_MASTER:
        c_search, c_page = st.columns([3, 1])
        with c_search:
            search_kw = st.text_input("🔍 Tìm kiếm vật tư nhanh (gõ mã hoặc tên):", placeholder="Ví dụ: gỗ sồi, mdf, nẹp, phào, ty treo...")

        items_list = []
        q_kw = search_kw.strip().lower()
        for k, v in PRICE_MASTER.items():
            if not q_kw or q_kw in v["name"].lower() or q_kw in k.lower():
                items_list.append({
                    "MÃ CODE": k,
                    "NHÓM": v.get("category", "CHUNG"),
                    "TÊN VẬT TƯ / NHÂN CÔNG": v["name"],
                    "ĐVT": v["unit"],
                    "ĐƠN GIÁ XƯỞNG": v["unit_cost"]
                })

        total_found = len(items_list)
        st.caption(f"Tìm thấy **{total_found}** vật tư phù hợp.")

        page_size = 50
        total_pages = max(math.ceil(total_found / page_size), 1)
        
        with c_page:
            current_page = st.number_input("Trang số:", min_value=1, max_value=total_pages, value=1, step=1)

        start_idx = (current_page - 1) * page_size
        end_idx = min(start_idx + page_size, total_found)
        page_items = items_list[start_idx:end_idx]

        df_page = pd.DataFrame(page_items)

        edited_page_df = st.data_editor(
            df_page,
            use_container_width=True,
            num_rows="fixed",
            height=480,
            key=f"master_page_editor_{current_page}",
            column_config={
                "MÃ CODE": st.column_config.TextColumn("Mã Code", disabled=True),
                "NHÓM": st.column_config.TextColumn("Nhóm", disabled=True),
                "TÊN VẬT TƯ / NHÂN CÔNG": st.column_config.TextColumn("Tên vật tư / Quy cách", required=True),
                "ĐVT": st.column_config.TextColumn("ĐVT", required=True),
                "ĐƠN GIÁ XƯỞNG": st.column_config.NumberColumn("Đơn giá xưởng (VNĐ)", min_value=0.0, format="%d đ", required=True)
            }
        )

        c_save_m, c_dl_m = st.columns([2, 2])
        with c_save_m:
            if st.button(f"💾 Lưu cập nhật giá của Trang {current_page}", type="primary", use_container_width=True):
                for _, r in edited_page_df.iterrows():
                    code_key = r["MÃ CODE"]
                    if code_key in PRICE_MASTER:
                        PRICE_MASTER[code_key]["name"] = r["TÊN VẬT TƯ / NHÂN CÔNG"]
                        PRICE_MASTER[code_key]["unit"] = r["ĐVT"]
                        PRICE_MASTER[code_key]["unit_cost"] = float(r["ĐƠN GIÁ XƯỞNG"])

                full_export_list = [
                    {"MÃ CODE": k, "NHÓM": v.get("category", "CHUNG"), "TÊN VẬT TƯ / NHÂN CÔNG": v["name"], "ĐVT": v["unit"], "ĐƠN GIÁ XƯỞNG": v["unit_cost"]}
                    for k, v in PRICE_MASTER.items()
                ]
                df_all_export = pd.DataFrame(full_export_list)
                with pd.ExcelWriter("BIEU-MAU-DU-TOAN_MAPPED_MASTER.xlsx", engine='openpyxl') as writer:
                    df_all_export.to_excel(writer, index=False, sheet_name='DMVT')

                st.cache_data.clear()
                st.success(f"✅ Đã lưu giá mới cho Trang {current_page} vào file Master!")
                time.sleep(1)
                st.rerun()

        with c_dl_m:
            buf_master_export = io.BytesIO()
            full_export_list = [
                {"MÃ CODE": k, "NHÓM": v.get("category", "CHUNG"), "TÊN VẬT TƯ / NHÂN CÔNG": v["name"], "ĐVT": v["unit"], "ĐƠN GIÁ XƯỞNG": v["unit_cost"]}
                for k, v in PRICE_MASTER.items()
            ]
            df_all_export = pd.DataFrame(full_export_list)
            with pd.ExcelWriter(buf_master_export, engine='openpyxl') as wr:
                df_all_export.to_excel(wr, index=False, sheet_name='DMVT')
            st.download_button(
                "📥 Tải toàn bộ file Master (.xlsx) về máy",
                data=buf_master_export.getvalue(),
                file_name="Don_Gia_Master_Toan_Bo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.warning("Chưa có danh mục đơn giá. Hãy tải lên file Excel ở trên.")
