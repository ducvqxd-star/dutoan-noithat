import streamlit as st
import pandas as pd
import json
import io
import os
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
# 1. CẤU HÌNH GIAO DIỆN & TĂNG TƯƠNG PHẢN (CHỐNG MỜ CHỮ)
# -------------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Dự Toán Nội Thất (AI QS)", 
    page_icon="🏗️",
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
IMAGE_DIR = "uploaded_drawings"
os.makedirs(IMAGE_DIR, exist_ok=True)

try:
    CLOUD_API_KEY = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    CLOUD_API_KEY = ""

def get_vietnam_time():
    tz_vn = timezone(timedelta(hours=7))
    return datetime.now(tz_vn).strftime("%Y-%m-%d %H:%M:%S")

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
    for col_def in [
        "review_status TEXT DEFAULT 'Chờ duyệt'",
        "review_note TEXT DEFAULT ''"
    ]:
        try:
            c.execute(f"ALTER TABLE projects ADD COLUMN {col_def}")
        except Exception:
            pass

    c.execute('''
        CREATE TABLE IF NOT EXISTS boq_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            section_code TEXT,
            material_code TEXT,
            description TEXT,
            location TEXT,
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
    for col_def in [
        "section_code TEXT DEFAULT ''",
        "location TEXT DEFAULT ''",
        "waste_qty REAL DEFAULT 0.0",
        "meaning TEXT DEFAULT ''",
        "formula TEXT DEFAULT ''"
    ]:
        try:
            c.execute(f"ALTER TABLE boq_items ADD COLUMN {col_def}")
        except Exception:
            pass

    admin_pw = hashlib.sha256("admin123".encode()).hexdigest()
    qs1_pw = hashlib.sha256("qs123".encode()).hexdigest()
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, full_name, role) VALUES ('admin', ?, 'Quản Trị Viên', 'admin')", (admin_pw,))
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, full_name, role) VALUES ('qs01', ?, 'Kỹ Sư QS 01', 'qs')", (qs1_pw,))
    conn.commit()
    conn.close()

init_db()

# -------------------------------------------------------------
# 2. BỘ NẠP MASTER EXCEL TỰ ĐỘNG ĐỊNH VỊ CỘT O VÀ CỘT D (ĐVT)
# -------------------------------------------------------------
@st.cache_data(ttl=600)
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
        return price_dict, f"{target_file} (Sheet: {target_sheet})"
    except Exception as e:
        return {}, f"Lỗi đọc file: {str(e)}"

PRICE_MASTER, SOURCE_INFO = load_price_master()

# -------------------------------------------------------------
# 3. QUẢN LÝ ĐĂNG NHẬP
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
    st.title("🔐 Hệ Thống Báo Giá & Dự Toán Nội Thất (AI QS)")
    st.markdown("Cổng làm việc nội bộ phòng Dự toán & QS.")
    col_l1, col_l2, _ = st.columns([1, 1, 1])
    with col_l1:
        u_in = st.text_input("Tên đăng nhập:")
        p_in = st.text_input("Mật khẩu:", type="password")
        if st.button("Đăng nhập", type="primary", use_container_width=True):
            login(u_in, p_in)
        st.caption("Tài khoản mẫu: `admin` / `admin123` hoặc `qs01` / `qs123`")
    st.stop()

# -------------------------------------------------------------
# 4. CẤU TRÚC DỮ LIỆU ĐỊNH NGHĨA CHO AI THEO CHUẨN 7 PHẦN (I -> VII)
# -------------------------------------------------------------
class ChiTietVatTu(BaseModel):
    material_code: str
    description: str
    location: str
    unit: str
    net_quantity: float
    waste_factor_percent: float

class NhomChiTiet(BaseModel):
    section_code: str
    items: List[ChiTietVatTu]

class TakeoffResult(BaseModel):
    item_code: str
    item_name: str
    breakdown_sections: List[NhomChiTiet]
    technical_risks_and_notes: List[str]

# -------------------------------------------------------------
# 5. GIAO DIỆN CHÍNH
# -------------------------------------------------------------
user = st.session_state['user_info']

with st.sidebar:
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

tab1, tab2, tab3 = st.tabs(["📝 Bóc Tách Dự Toán Mới", "📂 Thư Viện Hồ Sơ & Chỉnh Sửa Chiết Tính", "📊 Bảng Tra Đơn Giá Master"])

# TAB 1: BÓC TÁCH MỚI
with tab1:
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("1. Dữ liệu bản vẽ & Thông số kỹ thuật")
        project_name = st.text_input("Tên dự án / Công trình:", value="Biệt thự Park City")
        item_code = st.text_input("Mã sản phẩm / Hạng mục:", value="SP10")
        item_name = st.text_input("Tên sản phẩm:", value="Tủ pantry khung gỗ đợt gỗ hoàn thiện")
        dimensions = st.text_input("Kích thước bao (DxRxC mm):", value="2600 x 600 x 2250 mm")
        
        uploaded_files = st.file_uploader("Upload ảnh 3D hoặc bản vẽ kỹ thuật:", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        if uploaded_files:
            c_preview = st.columns(min(len(uploaded_files), 3))
            for idx, f in enumerate(uploaded_files):
                with c_preview[idx % 3]:
                    st.image(f, use_container_width=True)

        specs = st.text_area(
            "Quy cách kỹ thuật & Vật liệu chỉ định:",
            height=130, 
            value="- Khung gỗ tự nhiên hoàn thiện sơn màu chỉ định.\n- Vách và đợt tủ gỗ hoàn thiện theo mẫu.\n- Hệ tủ dưới: Thùng vách MDF chống ẩm dán Veneer.\n- Mặt tủ ốp đá tự nhiên/nung kết líp mòi.\n- Hệ tay nắm nhôm Hafele, bản lề giảm chấn và ray âm giảm chấn."
        )

        submit_btn = st.button("🚀 Bắt đầu bóc tách & Tính chiết tính", type="primary", use_container_width=True)

    with c2:
        st.subheader("2. Chi phí tổng quan")
        if submit_btn:
            if not api_key:
                st.error("Vui lòng cấu hình Gemini API Key.")
            elif not uploaded_files:
                st.warning("Vui lòng tải lên ít nhất 1 ảnh bản vẽ kỹ thuật.")
            else:
                with st.spinner("AI QS đang bóc tách chiết tính theo chuẩn 7 phần (I -> VII)..."):
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

                        prompt = f"""
                        BẠN LÀ SENIOR QS NỘI THẤT. BÓC TÁCH KHỐI LƯỢNG KỸ THUẬT VÀ CHIẾT TÍNH ĐÚNG CẤU TRÚC TỪ I ĐẾN VII:
                        - Mã hiệu sản phẩm: {item_code}
                        - Tên sản phẩm: {item_name}
                        - Kích thước bao (DxRxC mm): {dimensions}
                        - Quy cách vật liệu:
                        {specs}

                        PHÂN BỔ BẮT BUỘC ĐÚNG 7 NHÓM SAU (TÊN section_code BẮT BUỘC CHÍNH XÁC):
                        I. VẬT TƯ THẦU PHỤ: Hạng mục gia công ngoài / nhà thầu phụ (nếu có).
                        II. VẬT TƯ CHÍNH: Gỗ tự nhiên xẻ sấy, ván dán Plywood, ván cốt MDF/HDF chống ẩm, tấm Cemboard lót.
                        III. HOÀN THIỆN (KIM LOẠI, PU...): Nẹp chỉ kim loại, thanh Inox La mạ PVD, khung nhôm viền.
                        IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ: Nhân công thợ mộc, sơn PU, ép dán Veneer/Laminate, keo/vít.
                        V. BAO BÌ: Màng PE, bọc carton bảo vệ sản phẩm chống trầy.
                        VI. PHỤ KIỆN: Bản lề giảm chấn, ray trượt âm, tay nắm nhôm, khóa.
                        VII. SOURCE NGOÀI: Mặt bàn đá nung kết/đá tự nhiên líp mòi 45 độ, kính cánh tủ, màng xuyên sáng.

                        Ghi rõ vị trí cụ thể (location) cho từng chi tiết.
                        """

                        MODELS_TO_TRY = ['gemini-3.8-flash', 'gemini-3.1-pro']
                        response = None
                        last_err = None

                        for current_model in MODELS_TO_TRY:
                            for attempt in range(3):
                                try:
                                    response = client.models.generate_content(
                                        model=current_model,
                                        contents=[*images, prompt],
                                        config=types.GenerateContentConfig(
                                            system_instruction="BẠN LÀ CHUYÊN GIA SENIOR QS NỘI THẤT. Bóc tách chiết tính 7 phần I -> VII chuẩn xác.",
                                            response_mime_type="application/json",
                                            response_schema=TakeoffResult,
                                            temperature=0.0
                                        )
                                    )
                                    break
                                except Exception as err:
                                    last_err = err
                                    err_str = str(err)
                                    if any(k in err_str for k in ["503", "UNAVAILABLE", "429"]):
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
                        for section in data.get("breakdown_sections", []):
                            s_code = section.get("section_code", "CHUNG")
                            for comp in section.get("items", []):
                                code = str(comp.get("material_code", "")).strip()
                                desc = comp.get("description", "").lower().strip()
                                loc = comp.get("location", "").strip()
                                unit_str = str(comp.get("unit", "cái")).strip().lower()
                                
                                item_info = PRICE_MASTER.get(code)
                                if not item_info and desc:
                                    best_match = None
                                    best_score = 0
                                    clean_words = [w for w in desc.replace(",", " ").replace("-", " ").replace("/", " ").split() if len(w) > 2]
                                    for p_code, p_data in PRICE_MASTER.items():
                                        p_name_lower = p_data["name"].lower()
                                        score = sum(1 for w in clean_words if w in p_name_lower)
                                        if score > best_score:
                                            best_score = score
                                            best_match = p_data
                                            code = p_code
                                    if best_score >= 1:
                                        item_info = best_match

                                if not item_info:
                                    item_info = {"name": comp.get("description"), "unit": comp.get("unit"), "unit_cost": 0.0}

                                u_cost = item_info["unit_cost"]
                                net_q = float(comp.get("net_quantity", 0.0))
                                w_pct = float(comp.get("waste_factor_percent", 0.0))

                                is_discrete = any(u in unit_str for u in ['cai', 'cái', 'bộ', 'bo', 'con', 'chiếc', 'chiec'])
                                if is_discrete:
                                    waste_q = math.ceil(net_q * (w_pct / 100.0)) if w_pct > 0 else 0
                                    gross_q = net_q + waste_q
                                else:
                                    waste_q = round(net_q * (w_pct / 100.0), 3)
                                    gross_q = round(net_q + waste_q, 3)

                                amt = gross_q * u_cost
                                total_direct += amt
                                
                                meaning_text = f"Đo vẽ {loc}: KL Net + {w_pct:.0f}% Hao hụt = KL Gross xưởng"
                                formula_text = f"Gross = {net_q:.2f} + {waste_q:.2f} -> {gross_q:.2f} x {u_cost:,.0f} đ"

                                boq_records.append({
                                    "section_code": s_code,
                                    "material_code": code if code else "CHƯA_CÓ_MÃ",
                                    "description": comp.get("description"),
                                    "location": loc,
                                    "unit": item_info.get("unit", comp.get("unit")),
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
                                (project_id, section_code, material_code, description, location, unit, net_qty, waste_percent, waste_qty, gross_qty, unit_cost, amount, meaning, formula)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', (
                                proj_id, r["section_code"], r["material_code"], r["description"], r["location"],
                                r["unit"], r["net_qty"], r["waste_percent"], r["waste_qty"], r["gross_qty"], 
                                r["unit_cost"], r["amount"], r["meaning"], r["formula"]
                            ))
                        conn.commit()
                        conn.close()

                        st.success(f"✅ Bóc tách chiết tính 7 phần thành công lúc {created_time_vn} (Giờ VN)!")
                        
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Tổng giá Cost xưởng", f"{total_direct:,.0f} đ")
                        m2.metric(f"Chi phí chung ({overhead_rate}%)", f"{overhead_amt:,.0f} đ")
                        m3.metric(f"Giá chào thầu ({profit_rate}%)", f"{tender_amt:,.0f} đ")

                        st.markdown("⚠️ **Lưu ý kỹ thuật & Cảnh báo gia công:**")
                        for rsk in data.get("technical_risks_and_notes", []):
                            st.info(f"• {rsk}")

                        st.info("💡 Chuyển sang tab **'📂 Thư Viện Hồ Sơ & Chỉnh Sửa Chiết Tính'** để xem, đổi mã code tự động dò master hoặc chỉnh sửa số liệu.")

                    except Exception as e:
                        st.error(f"Lỗi xử lý: {str(e)}")

# TAB 2: THƯ VIỆN HỒ SƠ & BẢNG CHỈNH SỬA CHIẾT TÍNH (TỰ ĐỘNG DÒ CODE MASTER)
with tab2:
    st.subheader(f"📂 Thư Viện Hồ Sơ & Chỉnh Sửa Chiết Tính")
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
        st.info("Chưa có hồ sơ dự toán nào được lưu.")
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
        
        st.markdown("### 🛠️ Quản Trị & Thao Tác Hồ Sơ")
        sel_col1, sel_col2, sel_col3 = st.columns([1.5, 2.5, 1])
        
        with sel_col1:
            sel_id = st.selectbox("Chọn ID hồ sơ cần xem / chỉnh sửa:", df_saved["id"].tolist())
            selected_row = df_saved[df_saved["id"] == sel_id].iloc[0]
            st.caption(f"Hạng mục: **{selected_row['item_name']}** | Tạo bởi: **{selected_row['nguoi_tao']}**")

        with sel_col2:
            if user['role'] == 'admin':
                cur_status = selected_row['review_status'] if pd.notna(selected_row['review_status']) else 'Chờ duyệt'
                status_options = ['Chờ duyệt', '✅ Đúng giá', '❌ Sai giá']
                status_idx = status_options.index(cur_status) if cur_status in status_options else 0
                
                c_rev1, c_rev2 = st.columns([1.5, 2.5])
                with c_rev1:
                    new_status = st.radio("Đánh giá của Admin:", status_options, index=status_idx, horizontal=True)
                with c_rev2:
                    new_note = st.text_input("Ghi chú Admin:", value=selected_row['review_note'] if pd.notna(selected_row['review_note']) else '')
                
                if st.button("💾 Cập nhật kiểm duyệt", type="primary"):
                    cur_u = conn.cursor()
                    cur_u.execute("UPDATE projects SET review_status = ?, review_note = ? WHERE id = ?", (new_status, new_note, sel_id))
                    conn.commit()
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
                    st.warning(f"Đã xóa hoàn toàn hồ sơ #{sel_id}.")
                    time.sleep(0.5)
                    st.rerun()

        st.markdown("---")
        st.markdown(f"### ✏️ Bảng Chi Tiết Chiết Tính (Hồ Sơ #{sel_id})")
        st.caption("💡 **Cơ chế dò mã Master:** Bạn có thể nhập/sửa trực tiếp mã ở cột **Code**. Hệ thống sẽ tự động tra cứu trong kho Master để cập nhật đúng **Tên vật tư**, **ĐVT** và **Đơn giá**.")

        # Lấy dữ liệu hồ sơ vào session state để quản lý chỉnh sửa
        state_key = f"df_edit_data_{sel_id}"
        if state_key not in st.session_state:
            df_init = pd.read_sql_query(
                """SELECT 
                    id AS item_db_id,
                    section_code AS 'Mục chiết tính (I -> VII)', 
                    material_code AS 'Code', 
                    description AS 'Vật tư / Hạng mục', 
                    location AS 'Vị trí',
                    unit AS 'ĐVT', 
                    net_qty AS 'KL Tịnh (Net)', 
                    waste_percent AS '% Hao hụt', 
                    gross_qty AS 'KL Tổng (Gross)', 
                    unit_cost AS 'Đơn giá', 
                    amount AS 'Thành tiền',
                    meaning AS 'Ý nghĩa'
                FROM boq_items WHERE project_id = ?""",
                conn, params=(sel_id,)
            )
            df_init['Mục chiết tính (I -> VII)'] = df_init['Mục chiết tính (I -> VII)'].fillna("II. VẬT TƯ CHÍNH")
            df_init['Vị trí'] = df_init['Vị trí'].fillna("Theo bản vẽ")
            df_init['Ý nghĩa'] = df_init['Ý nghĩa'].fillna("KL Net + Hao hụt = KL Gross xưởng")
            st.session_state[state_key] = df_init

        # Nút tiện ích tự động quét đồng bộ lại tên và giá theo mã code
        c_tool1, c_tool2 = st.columns([2, 3])
        with c_tool1:
            if st.button("🔍 Tự Động Tra Lại Tên & Giá Theo Mã Code", help="Tra cứu toàn bộ cột Code để lấy đúng tên vật tư và đơn giá Master"):
                for idx in st.session_state[state_key].index:
                    code_val = str(st.session_state[state_key].at[idx, 'Code']).strip()
                    if code_val in PRICE_MASTER:
                        master_item = PRICE_MASTER[code_val]
                        st.session_state[state_key].at[idx, 'Vật tư / Hạng mục'] = master_item['name']
                        st.session_state[state_key].at[idx, 'ĐVT'] = master_item['unit']
                        st.session_state[state_key].at[idx, 'Đơn giá'] = master_item['unit_cost']
                        # Tính lại Gross & Thành tiền
                        net_q = float(st.session_state[state_key].at[idx, 'KL Tịnh (Net)'] or 0.0)
                        w_pct = float(st.session_state[state_key].at[idx, '% Hao hụt'] or 0.0)
                        gross_q = round(net_q * (1 + w_pct / 100.0), 3)
                        st.session_state[state_key].at[idx, 'KL Tổng (Gross)'] = gross_q
                        st.session_state[state_key].at[idx, 'Thành tiền'] = round(gross_q * master_item['unit_cost'], 0)
                st.success("Đã đồng bộ lại toàn bộ vật tư theo mã Code!")
                st.rerun()

        # Bảng dữ liệu có thể chỉnh sửa
        edited_df = st.data_editor(
            st.session_state[state_key],
            use_container_width=True,
            num_rows="dynamic",
            height=520,
            key=f"editor_table_{sel_id}",
            column_config={
                "item_db_id": None,
                "Mục chiết tính (I -> VII)": st.column_config.SelectboxColumn(
                    "Mục chiết tính (I -> VII)",
                    options=[
                        "I. VẬT TƯ THẦU PHỤ",
                        "II. VẬT TƯ CHÍNH",
                        "III. HOÀN THIỆN (KIM LOẠI, PU...)",
                        "IV. NHÂN CÔNG & VẬT LIỆU HOÀN THIỆN PHỤ",
                        "V. BAO BÌ",
                        "VI. PHỤ KIỆN",
                        "VII. SOURCE NGOÀI"
                    ],
                    required=True
                ),
                "Code": st.column_config.TextColumn("Code (Mã Master)", required=True),
                "Vật tư / Hạng mục": st.column_config.TextColumn("Vật tư / Hạng mục", required=True),
                "Vị trí": st.column_config.TextColumn("Vị trí"),
                "ĐVT": st.column_config.TextColumn("ĐVT", required=True),
                "KL Tịnh (Net)": st.column_config.NumberColumn("KL Tịnh (Net)", min_value=0.0, format="%.2f"),
                "% Hao hụt": st.column_config.NumberColumn("% Hao hụt", min_value=0.0, max_value=100.0, format="%.1f%%"),
                "KL Tổng (Gross)": st.column_config.NumberColumn("KL Tổng (Gross)", disabled=True, format="%.2f"),
                "Đơn giá": st.column_config.NumberColumn("Đơn giá", min_value=0.0, format="%d đ"),
                "Thành tiền": st.column_config.NumberColumn("Thành tiền", disabled=True, format="%d đ"),
                "Ý nghĩa": st.column_config.TextColumn("Ý nghĩa")
            }
        )

        col_save_btn, col_exp_btn = st.columns([1.5, 3.5])
        with col_save_btn:
            if st.button("💾 Lưu thay đổi bảng chiết tính", type="primary", use_container_width=True):
                cur_save = conn.cursor()
                cur_save.execute("DELETE FROM boq_items WHERE project_id = ?", (sel_id,))
                
                new_total_direct = 0.0
                
                for _, row in edited_df.iterrows():
                    s_code = str(row['Mục chiết tính (I -> VII)'])
                    m_code = str(row['Code']).strip() if pd.notna(row['Code']) else "CHƯA_CÓ_MÃ"
                    desc = str(row['Vật tư / Hạng mục']) if pd.notna(row['Vật tư / Hạng mục']) else ""
                    loc = str(row['Vị trí']) if pd.notna(row['Vị trí']) else ""
                    unit = str(row['ĐVT']) if pd.notna(row['ĐVT']) else "cái"
                    
                    net_q = float(row['KL Tịnh (Net)']) if pd.notna(row['KL Tịnh (Net)']) else 0.0
                    w_pct = float(row['% Hao hụt']) if pd.notna(row['% Hao hụt']) else 0.0
                    u_cost = float(row['Đơn giá']) if pd.notna(row['Đơn giá']) else 0.0
                    
                    # Tự động đồng bộ từ Master nếu mã Code trùng khớp với kho
                    if m_code in PRICE_MASTER:
                        master_data = PRICE_MASTER[m_code]
                        # Nếu người dùng chưa gõ tên, tự động lấy tên Master
                        if not desc or desc == "None":
                            desc = master_data['name']
                        if not unit or unit == "cái":
                            unit = master_data['unit']
                        if u_cost == 0.0:
                            u_cost = master_data['unit_cost']

                    # Tính toán lại các giá trị phụ thuộc
                    waste_q = round(net_q * (w_pct / 100.0), 3)
                    gross_q = round(net_q + waste_q, 3)
                    amt = round(gross_q * u_cost, 0)
                    new_total_direct += amt
                    
                    meaning = str(row['Ý nghĩa']) if pd.notna(row['Ý nghĩa']) else f"Đo vẽ {loc}"
                    formula = f"Gross = {net_q:.2f} + {waste_q:.2f} -> {gross_q:.2f} x {u_cost:,.0f} đ"
                    
                    cur_save.execute('''
                        INSERT INTO boq_items 
                        (project_id, section_code, material_code, description, location, unit, net_qty, waste_percent, waste_qty, gross_qty, unit_cost, amount, meaning, formula)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        sel_id, s_code, m_code, desc, loc, unit,
                        net_q, w_pct, waste_q, gross_q, u_cost, amt, meaning, formula
                    ))

                new_overhead = new_total_direct * (overhead_rate / 100.0)
                new_tender = (new_total_direct + new_overhead) * (1 + profit_rate / 100.0)
                
                cur_save.execute('''
                    UPDATE projects 
                    SET direct_cost = ?, overhead_cost = ?, tender_price = ? 
                    WHERE id = ?
                ''', (new_total_direct, new_overhead, new_tender, sel_id))
                
                conn.commit()
                # Xóa cache session state để nạp lại dữ liệu mới nhất
                st.session_state.pop(state_key, None)
                st.success(f"✅ Đã lưu thành công! Tổng chi phí mới: {new_total_direct:,.0f} đ | Giá chào thầu: {new_tender:,.0f} đ")
                time.sleep(0.8)
                st.rerun()

        with col_exp_btn:
            buf_re = io.BytesIO()
            with pd.ExcelWriter(buf_re, engine='openpyxl') as wr:
                df_to_export = edited_df.drop(columns=['item_db_id'], errors='ignore')
                df_to_export.to_excel(wr, index=False, sheet_name='ChietTinh_SauChinhSua')
            st.download_button(
                f"📥 Tải file Excel chiết tính (đã cập nhật) của hồ sơ #{sel_id}", 
                data=buf_re.getvalue(), 
                file_name=f"Chiết_Tính_Cập_Nhật_ID_{sel_id}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    conn.close()

# TAB 3: BẢNG TRA ĐƠN GIÁ MASTER
with tab3:
    st.subheader("📊 Bảng Tra Cứu Đơn Giá Master")
    st.caption(f"Nguồn dữ liệu hiện hành: `{SOURCE_INFO}`")
    if PRICE_MASTER:
        df_pview = pd.DataFrame.from_dict(PRICE_MASTER, orient='index').reset_index()
        df_pview.columns = ["MÃ CODE", "NHÓM", "TÊN VẬT TƯ / NHÂN CÔNG", "ĐVT", "ĐƠN GIÁ XƯỞNG"]
        
        search_kw = st.text_input("🔍 Tìm kiếm vật tư nhanh (ví dụ: 'đá', 'melamine', 'inox', 'lắp đặt'):")
        if search_kw:
            df_pview = df_pview[df_pview['TÊN VẬT TƯ / NHÂN CÔNG'].str.contains(search_kw, case=False, na=False) | df_pview['MÃ CODE'].str.contains(search_kw, case=False, na=False)]
            
        st.dataframe(df_pview.style.format({"ĐƠN GIÁ XƯỞNG": "{:,.0f} đ"}), use_container_width=True, height=500)
    else:
        st.warning("Chưa nạp được danh mục đơn giá. Vui lòng kiểm tra lại file Excel.")
