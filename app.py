import streamlit as st
import pandas as pd
import json
import io
import os
import time
import sqlite3
import hashlib
from datetime import datetime, timezone, timedelta
from PIL import Image
from pydantic import BaseModel
from typing import List, Optional
from google import genai
from google.genai import types

# -------------------------------------------------------------
# 1. CẤU HÌNH GIAO DIỆN & TĂNG ĐỘ TƯƠNG PHẢN (CHỐNG MỜ CHỮ)
# -------------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Dự Toán Nội Thất (AI QS)", 
    page_icon="🏗️",
    layout="wide"
)

# Chống mờ chữ trên màn hình sáng/tối
st.markdown("""
<style>
    div[data-testid="stDataFrame"] {
        opacity: 1 !important;
        filter: none !important;
    }
    .stDataFrame td, .stDataFrame th {
        font-weight: 500 !important;
        font-size: 14px !important;
    }
    .stApp {
        opacity: 1 !important;
    }
</style>
""", unsafe_allow_html=True)

DB_FILE = "qs_multiuser.db"
IMAGE_DIR = "uploaded_drawings"
os.makedirs(IMAGE_DIR, exist_ok=True)

# Lấy API Key từ Secrets của Streamlit Cloud (bảo vệ tránh crash khi chạy local)
try:
    CLOUD_API_KEY = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    CLOUD_API_KEY = ""

def get_vietnam_time():
    """Lấy thời gian chuẩn theo múi giờ Việt Nam (GMT+7)"""
    tz_vn = timezone(timedelta(hours=7))
    return datetime.now(tz_vn).strftime("%Y-%m-%d %H:%M:%S")

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Bảng người dùng
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            password_hash TEXT,
            full_name TEXT,
            role TEXT
        )
    ''')
    # Bảng lưu hồ sơ dự toán
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
    try:
        c.execute("ALTER TABLE projects ADD COLUMN review_status TEXT DEFAULT 'Chờ duyệt'")
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE projects ADD COLUMN review_note TEXT DEFAULT ''")
    except Exception:
        pass

    # Bảng lưu chi tiết từng dòng BOQ
    c.execute('''
        CREATE TABLE IF NOT EXISTS boq_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            module_name TEXT,
            material_code TEXT,
            description TEXT,
            unit TEXT,
            net_qty REAL,
            waste_percent REAL,
            gross_qty REAL,
            unit_cost REAL,
            amount REAL,
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
        col_unit = 3   # Cột D (ĐVT)
        col_price = 14 # Cột O (Đơn giá tính)
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
# 3. QUẢN LÝ ĐĂNG NHẬP (MULTI-USER)
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
# 4. CẤU TRÚC DỮ LIỆU ĐỊNH NGHĨA CHO AI (PYDANTIC SCHEMA)
# -------------------------------------------------------------
class ComponentItem(BaseModel):
    component_id: str
    description: str
    material_master_code: str
    unit: str
    net_quantity: float
    waste_factor_percent: float
    gross_quantity: float
    technical_spec: str

class SubAssembly(BaseModel):
    sub_assembly_id: str
    sub_assembly_name: str
    components: List[ComponentItem]

class TakeoffResult(BaseModel):
    item_code: str
    item_name: str
    structural_breakdown: List[SubAssembly]
    technical_risks_and_notes: List[str]

# -------------------------------------------------------------
# 5. GIAO DIỆN CHÍNH CỦA ỨNG DỤNG
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

tab1, tab2, tab3 = st.tabs(["📝 Bóc Tách Dự Toán Mới", "📂 Thư Viện Hồ Sơ Đã Lưu", "📊 Bảng Tra Đơn Giá Master"])

# TAB 1: BÓC TÁCH MỚI
with tab1:
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("1. Dữ liệu bản vẽ & Thông số kỹ thuật")
        project_name = st.text_input("Tên dự án / Công trình:", value="Biệt thự Park City")
        item_code = st.text_input("Mã sản phẩm / Hạng mục:", value="TR-01")
        item_name = st.text_input("Tên sản phẩm:", value="Vách Tivi + tủ quần áo p ngủ bà")
        dimensions = st.text_input("Kích thước bao (DxRxC mm):", value="4095 x 450 x 2600 mm")
        
        uploaded_files = st.file_uploader("Upload ảnh 3D hoặc bản vẽ kỹ thuật:", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        if uploaded_files:
            c_preview = st.columns(min(len(uploaded_files), 3))
            for idx, f in enumerate(uploaded_files):
                with c_preview[idx % 3]:
                    st.image(f, use_container_width=True)

        specs = st.text_area(
            "Quy cách kỹ thuật & Vật liệu chỉ định:",
            height=130, 
            value="- Hệ vách tivi: Khung xương gỗ ghép thanh kết hợp cốt MDF chống ẩm phủ Melamine.\n- Tủ áo: Thùng cánh MDF chống ẩm, bản lề giảm chấn, ray trượt âm.\n- Kích thước gia công và phụ kiện theo đúng bản vẽ đính kèm."
        )

        submit_btn = st.button("🚀 Bắt đầu bóc tách & Tính dự toán", type="primary", use_container_width=True)

    with c2:
        st.subheader("2. Bảng BOQ & Chiết tính chi phí xưởng")
        if submit_btn:
            if not api_key:
                st.error("Vui lòng cấu hình Gemini API Key.")
            elif not uploaded_files:
                st.warning("Vui lòng tải lên ít nhất 1 ảnh bản vẽ kỹ thuật.")
            else:
                with st.spinner("AI QS đang bóc tách hình học và tra cứu đơn giá master..."):
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
                        BÓC TÁCH KỸ THUẬT NỘI THẤT (ĐỘ CHÍNH XÁC CAO - QUY TẮC BẮT BUỘC):
                        - Hạng mục: {item_name} | Mã hiệu: {item_code}
                        - Kích thước phủ bì cố định: {dimensions}
                        - Quy cách vật liệu thực tế:
                        {specs}

                        QUY TẮC HÌNH HỌC & ĐỊNH MỨC BẮT BUỘC:
                        1. TÍNH TOÁN THEO KÍCH THƯỚC PHỦ BÌ:
                           - Tổng diện tích mặt đứng không được vượt quá: (Dài x Cao) / 1.000.000 (m2).
                           - Bóc tách đúng các module: thùng, hậu, vách, cánh, đợt, phụ kiện.
                        2. TUÂN THỦ NGHIÊM NGẶT SPECS:
                           - Chỉ bóc tách những gì thực sự có trong bản vẽ và được nêu trong phần quy cách.
                        3. HỆ SỐ HAO HỤT (waste_factor_percent) CỐ ĐỊNH:
                           - Ván công nghiệp (MDF, MFC, Plywood, Gỗ ghép): cố định đúng 8.0%
                           - Đá tự nhiên / Đá nung kết: cố định đúng 10.0%
                           - Nẹp / Inox / Phụ kiện: cố định đúng 5.0%
                        4. ÁNH XẠ MÃ VẬT TƯ (material_master_code): Khớp mã trong CSDL nếu có, hoặc ghi rõ tên chủng loại chính xác.
                        """

                        # Áp dụng danh sách model chỉ định: gemini-3.8-flash, gemini-3.1-pro
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
                                            system_instruction="BẠN LÀ CHUYÊN GIA SENIOR QS NỘI THẤT. Tính toán khối lượng tất định, không suy diễn ngẫu nhiên, mô tả chi tiết vật tư.",
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
                        
                        # Khớp đơn giá từ Master Database với thuật toán tìm kiếm thông minh
                        total_direct = 0
                        boq_records = []
                        for sub in data.get("structural_breakdown", []):
                            m_name = sub.get("sub_assembly_name")
                            for comp in sub.get("components", []):
                                code = str(comp.get("material_master_code", "")).strip()
                                desc = comp.get("description", "").lower().strip()
                                
                                # 1. Khớp chính xác theo mã code
                                item_info = PRICE_MASTER.get(code)
                                
                                # 2. Nếu không khớp mã code, tự động tìm kiếm tương đối theo từ khóa tên vật tư
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

                                # 3. Nếu không tìm thấy trong Master, không gán 250k mà để 0 để cảnh báo
                                if not item_info:
                                    item_info = {
                                        "name": comp.get("description"),
                                        "unit": comp.get("unit"),
                                        "unit_cost": 0.0
                                    }

                                u_cost = item_info["unit_cost"]
                                gross_q = comp.get("gross_quantity", 0.0)
                                amt = gross_q * u_cost
                                total_direct += amt
                                
                                boq_records.append({
                                    "module_name": m_name,
                                    "material_code": code if code else "CHƯA_CÓ_MÃ",
                                    "description": comp.get("description"),
                                    "technical_spec": comp.get("technical_spec"),
                                    "unit": item_info.get("unit", comp.get("unit")),
                                    "net_qty": comp.get("net_quantity"),
                                    "waste_percent": comp.get("waste_factor_percent"),
                                    "gross_qty": gross_q,
                                    "unit_cost": u_cost,
                                    "amount": amt
                                })

                        overhead_amt = total_direct * (overhead_rate / 100.0)
                        tender_amt = (total_direct + overhead_amt) * (1 + profit_rate / 100.0)
                        created_time_vn = get_vietnam_time()

                        # Lưu vào CSDL
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
                                (project_id, module_name, material_code, description, unit, net_qty, waste_percent, gross_qty, unit_cost, amount)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', (
                                proj_id, r["module_name"], r["material_code"], r["description"],
                                r["unit"], r["net_qty"], r["waste_percent"], r["gross_qty"], r["unit_cost"], r["amount"]
                            ))
                        conn.commit()
                        conn.close()

                        st.success(f"✅ Bóc tách thành công lúc {created_time_vn} (Giờ VN)!")
                        
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Chi phí trực tiếp (Direct Cost)", f"{total_direct:,.0f} đ")
                        m2.metric(f"Chi phí chung ({overhead_rate}%)", f"{overhead_amt:,.0f} đ")
                        m3.metric(f"Giá chào thầu ({profit_rate}%)", f"{tender_amt:,.0f} đ")

                        df_view = pd.DataFrame(boq_records)
                        st.dataframe(df_view[["module_name", "material_code", "description", "unit", "gross_qty", "unit_cost", "amount"]].style.format({
                            "gross_qty": "{:.2f}",
                            "unit_cost": "{:,.0f} đ",
                            "amount": "{:,.0f} đ"
                        }), use_container_width=True)

                        st.markdown("⚠️ **Lưu ý kỹ thuật & Cảnh báo thi công:**")
                        for rsk in data.get("technical_risks_and_notes", []):
                            st.info(f"• {rsk}")

                        buf = io.BytesIO()
                        with pd.ExcelWriter(buf, engine='openpyxl') as wr:
                            df_view.to_excel(wr, index=False, sheet_name='BOQ_ChiTiet')
                        st.download_button(
                            "📥 Tải file BOQ Excel (.xlsx)", 
                            data=buf.getvalue(), 
                            file_name=f"BOQ_{item_code}.xlsx", 
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                    except Exception as e:
                        st.error(f"Lỗi xử lý: {str(e)}")

# TAB 2: THƯ VIỆN HỒ SƠ ĐÃ LƯU & QUẢN TRỊ KIỂM DUYỆT / XÓA
with tab2:
    st.subheader(f"📂 Thư Viện Hồ Sơ Dự Toán")
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
        
        col_act1, col_act2 = st.columns([1, 1])
        
        with col_act1:
            st.markdown("### 📥 Chi tiết BOQ & Xuất Excel")
            sel_id = st.selectbox("Chọn ID hồ sơ:", df_saved["id"].tolist())
            if sel_id:
                df_item_detail = pd.read_sql_query(
                    "SELECT module_name, material_code, description, unit, net_qty, waste_percent, gross_qty, unit_cost, amount FROM boq_items WHERE project_id = ?",
                    conn, params=(sel_id,)
                )
                st.dataframe(df_item_detail.style.format({
                    "gross_qty": "{:.2f}",
                    "unit_cost": "{:,.0f} đ",
                    "amount": "{:,.0f} đ"
                }), use_container_width=True)
                
                buf_re = io.BytesIO()
                with pd.ExcelWriter(buf_re, engine='openpyxl') as wr:
                    df_item_detail.to_excel(wr, index=False, sheet_name='BOQ_ChiTiet')
                st.download_button(
                    f"📥 Tải file Excel của hồ sơ #{sel_id}", 
                    data=buf_re.getvalue(), 
                    file_name=f"BOQ_ID_{sel_id}.xlsx",
                    use_container_width=True
                )

        with col_act2:
            st.markdown("### 🛡️ Thao tác Hồ sơ & Kiểm duyệt")
            selected_row = df_saved[df_saved["id"] == sel_id].iloc[0]
            st.info(f"Đang chọn: **#{sel_id} - {selected_row['item_name']}** (Người tạo: {selected_row['nguoi_tao']})")
            
            # Quản trị viên đánh giá Đúng/Sai giá
            if user['role'] == 'admin':
                st.markdown("#### 1. Đánh giá tính chính xác của Đơn giá:")
                cur_status = selected_row['review_status'] if pd.notna(selected_row['review_status']) else 'Chờ duyệt'
                status_options = ['Chờ duyệt', '✅ Đúng giá', '❌ Sai giá']
                status_idx = status_options.index(cur_status) if cur_status in status_options else 0
                
                new_status = st.radio("Kết quả kiểm duyệt:", status_options, index=status_idx, horizontal=True)
                new_note = st.text_input("Ghi chú của Quản trị viên:", value=selected_row['review_note'] if pd.notna(selected_row['review_note']) else '')
                
                if st.button("💾 Cập nhật đánh giá", type="primary", use_container_width=True):
                    cur_u = conn.cursor()
                    cur_u.execute("UPDATE projects SET review_status = ?, review_note = ? WHERE id = ?", (new_status, new_note, sel_id))
                    conn.commit()
                    st.success("Đã lưu kết quả kiểm duyệt đơn giá thành công!")
                    time.sleep(0.5)
                    st.rerun()
            else:
                st.caption(f"Trạng thái kiểm duyệt: **{selected_row['review_status']}**")
                if selected_row['review_note']:
                    st.caption(f"Ghi chú: {selected_row['review_note']}")

            st.markdown("---")
            # Quyền xóa hồ sơ: Admin hoặc chính chủ tạo hồ sơ
            st.markdown("#### 2. Xóa hồ sơ:")
            if user['role'] == 'admin' or selected_row['user_id'] == user['id']:
                if st.button(f"🗑️ Xóa vĩnh viễn hồ sơ #{sel_id}", type="secondary", use_container_width=True):
                    cur_d = conn.cursor()
                    cur_d.execute("DELETE FROM boq_items WHERE project_id = ?", (sel_id,))
                    cur_d.execute("DELETE FROM projects WHERE id = ?", (sel_id,))
                    conn.commit()
                    st.warning(f"Đã xóa hoàn toàn hồ sơ #{sel_id} khỏi hệ thống.")
                    time.sleep(0.5)
                    st.rerun()
            else:
                st.caption("🔒 Bạn chỉ có thể xóa các hồ sơ do chính bạn tạo.")

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
