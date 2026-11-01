import streamlit as st
import pandas as pd
import json
import io
import os
import sqlite3
import hashlib
from datetime import datetime
from PIL import Image
from pydantic import BaseModel
from typing import List, Optional
from google import genai
from google.genai import types

# -------------------------------------------------------------
# 1. CẤU HÌNH HỆ THỐNG & ĐỌC SECRETS AN TOÀN
# -------------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Dự Toán Nội Thất (AI QS)", 
    page_icon="🏗️",
    layout="wide"
)

DB_FILE = "qs_multiuser.db"
IMAGE_DIR = "uploaded_drawings"
os.makedirs(IMAGE_DIR, exist_ok=True)

# Lấy API Key từ Secrets của Streamlit Cloud một cách an toàn (tránh văng lỗi khi chạy Local)
try:
    CLOUD_API_KEY = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    CLOUD_API_KEY = ""

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
    # Bảng lưu hồ sơ dự toán theo từng User
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
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
    ''')
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
    # Khởi tạo sẵn tài khoản quản trị và kỹ sư QS mặc định
    admin_pw = hashlib.sha256("admin123".encode()).hexdigest()
    qs1_pw = hashlib.sha256("qs123".encode()).hexdigest()
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, full_name, role) VALUES ('admin', ?, 'Quản Trị Viên', 'admin')", (admin_pw,))
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, full_name, role) VALUES ('qs01', ?, 'Kỹ Sư QS 01', 'qs')", (qs1_pw,))
    conn.commit()
    conn.close()

init_db()

# -------------------------------------------------------------
# 2. HÀM NẠP CSDL MASTER TỰ ĐỘNG NHẬN DIỆN MỌI ĐỊNH DẠNG FILE
# -------------------------------------------------------------
@st.cache_data(ttl=600)
def load_price_master():
    # Danh sách các tên file ưu tiên kiểm tra
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
            
    # Nếu không trùng tên danh sách trên, tự động quét bất kỳ file .xlsx nào có trong thư mục
    if not target_file:
        all_xlsx = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~$")]
        if all_xlsx:
            target_file = all_xlsx[0]
        else:
            return {}, "Không tìm thấy file Excel đơn giá nào"

    try:
        xls = pd.ExcelFile(target_file)
        
        # Tự động tìm sheet chứa dữ liệu đơn giá
        target_sheet = None
        for s in xls.sheet_names:
            s_clean = s.strip().lower()
            if any(k in s_clean for k in ["đơn giá vt", "don gia vt", "dongia", "vt"]):
                target_sheet = s
                break
        if not target_sheet:
            target_sheet = xls.sheet_names[0]

        df_raw = pd.read_excel(target_file, sheet_name=target_sheet, header=None)
        
        # Nhận diện cấu trúc file theo tên sheet
        is_gia_cost = "dongia" in str(target_sheet).lower()
        col_code = 1
        col_name = 2
        col_unit = 4
        col_price = 7 if is_gia_cost else 5
        start_row = 6 if is_gia_cost else 2

        price_dict = {}
        current_cat = "CHUNG"

        for r in range(start_row, len(df_raw)):
            code = df_raw.iloc[r, col_code]
            name = df_raw.iloc[r, col_name]
            unit = df_raw.iloc[r, col_unit] if len(df_raw.columns) > col_unit else 'cái'
            price = df_raw.iloc[r, col_price] if len(df_raw.columns) > col_price else 0

            if pd.isna(code) and pd.notna(name):
                current_cat = str(name).strip()
            elif pd.notna(code) and pd.notna(name) and pd.notna(price):
                try:
                    p_val = float(price)
                except Exception:
                    p_val = 0.0
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
# 3. XỬ LÝ ĐĂNG NHẬP & PHÂN QUYỀN
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
# 4. SCHEMA DỮ LIỆU ĐỊNH NGHĨA CHO AI
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
# 5. GIAO DIỆN CHÍNH
# -------------------------------------------------------------
user = st.session_state['user_info']

with st.sidebar:
    st.markdown(f"👤 **{user['full_name']}** (`{user['role']}`)")
    if st.button("Đăng xuất", use_container_width=True):
        logout()
    st.markdown("---")
    st.header("⚙️ Cấu hình hệ thống")
    
    # Tự động nhận diện API Key từ Cloud Secrets nếu có
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
        item_name = st.text_input("Tên sản phẩm:", value="Tủ rượu quầy bar kết hợp vách xuyên sáng")
        dimensions = st.text_input("Kích thước bao (DxRxC mm):", value="4095 x 450 x 2600 mm")
        
        uploaded_files = st.file_uploader("Upload ảnh 3D hoặc bản vẽ kỹ thuật:", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        if uploaded_files:
            c_preview = st.columns(min(len(uploaded_files), 3))
            for idx, f in enumerate(uploaded_files):
                with c_preview[idx % 3]:
                    st.image(f, use_container_width=True)

        specs = st.text_area(
            "Quy cách kỹ thuật & Vật liệu chỉ định:",
            height=120, 
            value="- Kệ treo: Thanh Inox La 5x15mm mạ màu PVD, đợt gỗ sơn hoàn thiện rãnh LED.\n- Tủ dưới: Thùng cánh MDF chống ẩm phủ Melamine An Cường.\n- Tủ đứng: Cánh kính khung nhôm Slim, bản lề trục xoay âm nóc-đáy, đợt đỡ chai rượu nghiêng.\n- Mặt bàn bar: Đá nung kết Neolith líp mòi 45 độ, lót đệm ván Plywood.\n- Vách màng căng vi sinh xuyên sáng kèm hệ LED ziczac hộp đèn."
        )

        submit_btn = st.button("🚀 Bắt đầu bóc tách & Tính dự toán", type="primary", use_container_width=True)

    with c2:
        st.subheader("2. Bảng BOQ & Chiết tính chi phí xưởng")
        if submit_btn:
            if not api_key:
                st.error("Vui lòng cấu hình Gemini API Key ở thanh bên trái.")
            elif not uploaded_files:
                st.warning("Vui lòng tải lên ít nhất 1 ảnh phối cảnh hoặc bản vẽ kỹ thuật.")
            else:
                with st.spinner("AI QS đang đọc bản vẽ, bóc tách khối lượng và tra cứu đơn giá master..."):
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
                        BÓC TÁCH KHỐI LƯỢNG KỸ THUẬT NỘI THẤT:
                        - Mã hiệu: {item_code} | Tên: {item_name} | Kích thước: {dimensions}
                        - Quy cách vật liệu:
                        {specs}
                        - LƯU Ý BẮT BUỘC: Ánh xạ mã material_master_code chuẩn theo hệ mã phân tầng Master:
                          + Đá nung kết / gia công líp mòi: 10.03.01, 10.03.02, 10.03.03, 10.03.04
                          + Cánh kính Slim & Phụ kiện âm: 9.03.01, 15.02.01, 15.02.02
                          + Màng xuyên sáng & LED hộp đèn: 16.03.01, 16.03.02, 16.03.03, 16.03.04
                          + Inox La PVD & Kệ treo trần: 8.03.01, 8.03.02, 15.03.01, 17.01.07
                          + Da bọc vách / nệm: 12.02.01, 12.02.02, 12.02.03, 17.02.27
                          + Ván MDF chống ẩm Melamine: 3.01.04 (hoặc mã tương đương trong kho)
                          + Nhân công lắp đặt hiện trường: 17.02.01 đến 17.02.26
                        """

                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[*images, prompt],
                            config=types.GenerateContentConfig(
                                system_instruction="BẠN LÀ SENIOR QS NỘI THẤT. Bóc tách khối lượng net, tỷ lệ hao hụt waste%, mapping material_master_code phân tầng chuẩn xác.",
                                response_mime_type="application/json",
                                response_schema=TakeoffResult,
                                temperature=0.1,
                                thinking_config=types.ThinkingConfig(thinking_budget=0)
                            )
                        )

                        data = json.loads(response.text)
                        
                        # Khớp đơn giá từ Master Database
                        total_direct = 0
                        boq_records = []
                        for sub in data.get("structural_breakdown", []):
                            m_name = sub.get("sub_assembly_name")
                            for comp in sub.get("components", []):
                                code = comp.get("material_master_code", "DEFAULT")
                                item_info = PRICE_MASTER.get(code, {"name": comp.get("description"), "unit": comp.get("unit"), "unit_cost": 250000})
                                u_cost = item_info["unit_cost"]
                                gross_q = comp.get("gross_quantity", 0.0)
                                amt = gross_q * u_cost
                                total_direct += amt
                                
                                boq_records.append({
                                    "module_name": m_name,
                                    "material_code": code,
                                    "description": comp.get("description"),
                                    "technical_spec": comp.get("technical_spec"),
                                    "unit": comp.get("unit"),
                                    "net_qty": comp.get("net_quantity"),
                                    "waste_percent": comp.get("waste_factor_percent"),
                                    "gross_qty": gross_q,
                                    "unit_cost": u_cost,
                                    "amount": amt
                                })

                        overhead_amt = total_direct * (overhead_rate / 100.0)
                        tender_amt = (total_direct + overhead_amt) * (1 + profit_rate / 100.0)

                        # Lưu vào CSDL
                        conn = sqlite3.connect(DB_FILE)
                        cur = conn.cursor()
                        cur.execute('''
                            INSERT INTO projects 
                            (user_id, project_name, item_code, item_name, dimensions, specs, direct_cost, overhead_cost, tender_price, technical_risks, image_paths, created_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            user["id"], project_name, item_code, item_name, dimensions, specs,
                            total_direct, overhead_amt, tender_amt,
                            json.dumps(data.get("technical_risks_and_notes", []), ensure_ascii=False),
                            json.dumps(saved_paths), datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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

                        st.success(f"✅ Đã bóc tách thành công và lưu vào tài khoản {user['full_name']}!")
                        
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Chi phí trực tiếp (Direct Cost)", f"{total_direct:,.0f} đ")
                        m2.metric(f"Chi phí chung ({overhead_rate}%)", f"{overhead_amt:,.0f} đ")
                        m3.metric(f"Giá chào thầu ({profit_rate}%)", f"{tender_amt:,.0f} đ")

                        df_view = pd.DataFrame(boq_records)
                        st.dataframe(df_view[["module_name", "material_code", "description", "unit", "gross_qty", "unit_cost", "amount"]].style.format({
                            "gross_qty": "{:.2f}",
                            "unit_cost": "{:,.0f}",
                            "amount": "{:,.0f}"
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

# TAB 2: THƯ VIỆN HỒ SƠ ĐÃ LƯU
with tab2:
    st.subheader(f"📂 Hồ sơ dự toán của: {user['full_name']}")
    conn = sqlite3.connect(DB_FILE)
    query = "SELECT id, project_name, item_code, item_name, dimensions, direct_cost, tender_price, created_at FROM projects"
    params = ()
    if user['role'] != 'admin':
        query += " WHERE user_id = ?"
        params = (user['id'],)
    query += " ORDER BY id DESC"
    
    df_saved = pd.read_sql_query(query, conn, params=params)
    if df_saved.empty:
        st.info("Chưa có hồ sơ nào được tạo.")
    else:
        st.dataframe(df_saved.style.format({
            "direct_cost": "{:,.0f} đ",
            "tender_price": "{:,.0f} đ"
        }), use_container_width=True)
        
        st.markdown("---")
        sel_id = st.selectbox("Chọn ID hồ sơ muốn tải lại BOQ:", df_saved["id"].tolist())
        if sel_id:
            df_item_detail = pd.read_sql_query(
                "SELECT module_name, material_code, description, unit, net_qty, waste_percent, gross_qty, unit_cost, amount FROM boq_items WHERE project_id = ?",
                conn, params=(sel_id,)
            )
            st.dataframe(df_item_detail.style.format({
                "gross_qty": "{:.2f}",
                "unit_cost": "{:,.0f}",
                "amount": "{:,.0f}"
            }), use_container_width=True)
            
            buf_re = io.BytesIO()
            with pd.ExcelWriter(buf_re, engine='openpyxl') as wr:
                df_item_detail.to_excel(wr, index=False, sheet_name='BOQ_ChiTiet')
            st.download_button(f"📥 Tải lại BOQ của mục ID {sel_id}", data=buf_re.getvalue(), file_name=f"BOQ_ID_{sel_id}.xlsx")
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
