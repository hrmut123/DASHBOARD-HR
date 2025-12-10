import streamlit as st
import pandas as pd
import os
import io
import plotly.express as px
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from streamlit_option_menu import option_menu

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="HRIS Dark Pro",
    layout="wide",
    page_icon="🚀",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. CSS DARK MODE PREMIUM
# ==========================================
st.markdown("""
    <style>
    /* Background & Sidebar */
    .stApp { background-color: #0f172a; color: #f8fafc; }
    section[data-testid="stSidebar"] { background-color: #1e293b; border-right: 1px solid #334155; }
    
    /* Kartu Metric */
    div[data-testid="metric-container"] {
        background-color: #1e293b; border: 1px solid #334155;
        padding: 15px; border-radius: 10px; color: white;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3); transition: 0.3s;
    }
    div[data-testid="metric-container"]:hover { border-color: #3b82f6; transform: translateY(-5px); }
    div[data-testid="metric-container"] label { color: #94a3b8; }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #38bdf8; }

    /* Tabel */
    div[data-testid="stDataEditor"] { background-color: #1e293b; border: 1px solid #475569; border-radius: 10px; }
    th { background-color: #020617 !important; color: #38bdf8 !important; border-bottom: 2px solid #334155 !important; text-align: center !important; }
    td { color: #e2e8f0 !important; background-color: #1e293b !important; text-align: center !important; }

    /* Input */
    .stTextInput input, .stSelectbox, .stNumberInput input, .stDateInput input, .stTextArea textarea {
        background-color: #334155 !important; color: white !important; border: 1px solid #475569 !important; border-radius: 5px;
    }
    
    /* Tombol */
    .stButton button { background-color: #3b82f6; color: white; border-radius: 8px; font-weight: bold; border: none; }
    .stButton button:hover { background-color: #2563eb; }
    
    /* Tombol Hapus (Merah) */
    .delete-btn button { background-color: #ef4444 !important; }
    .delete-btn button:hover { background-color: #dc2626 !important; }

    /* Expander */
    .streamlit-expanderHeader { background-color: #1e293b !important; color: white !important; border: 1px solid #334155; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. BACKEND LOGIC
# ==========================================
FILE_EMP = 'data_karyawan.csv'
FILE_ATT = 'data_absensi.csv'
DEFAULT_COLS = ['PT', 'NIK', 'Nama', 'Jabatan', 'Departemen']

def load_data():
    if os.path.exists(FILE_EMP):
        try:
            df = pd.read_csv(FILE_EMP, dtype=str)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            # Bersihkan kolom helper jika ada sisa
            drop_cols = [c for c in df.columns if c in ['No', 'Ceklist', 'Pilih']]
            if drop_cols: df = df.drop(columns=drop_cols)
            if len(df.columns) == 0: df = pd.DataFrame(columns=DEFAULT_COLS)
        except: df = pd.DataFrame(columns=DEFAULT_COLS)
    else: df = pd.DataFrame(columns=DEFAULT_COLS)

    if os.path.exists(FILE_ATT):
        try:
            df_att = pd.read_csv(FILE_ATT, dtype=str)
            df_att = df_att.loc[:, ~df_att.columns.str.contains('^Unnamed')]
            if 'Pilih' in df_att.columns: df_att = df_att.drop(columns=['Pilih'])
        except: df_att = pd.DataFrame(columns=['Tanggal', 'NIK', 'Nama', 'Departemen', 'Jenis', 'Keterangan', 'Waktu_Input'])
    else: df_att = pd.DataFrame(columns=['Tanggal', 'NIK', 'Nama', 'Departemen', 'Jenis', 'Keterangan', 'Waktu_Input'])
    
    return df, df_att

def save_data(df, df_att):
    if 'Pilih' in df.columns: df = df.drop(columns=['Pilih'])
    if 'Pilih' in df_att.columns: df_att = df_att.drop(columns=['Pilih'])
    df.to_csv(FILE_EMP, index=False)
    df_att.to_csv(FILE_ATT, index=False)

def create_colorful_excel(df, title_text):
    output = io.BytesIO()
    clean_df = df.copy()
    for col in ['No', 'Ceklist', 'Pilih']:
        if col in clean_df.columns: clean_df = clean_df.drop(columns=[col])

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        clean_df.to_excel(writer, index=False, sheet_name='Laporan', startrow=3)
        ws = writer.sheets['Laporan']
        
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        row_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True, size=11)
        font_title = Font(color="1F4E78", bold=True, size=16)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center = Alignment(horizontal="center", vertical="center")

        ws['A1'] = title_text; ws['A1'].font = font_title
        ws['A2'] = f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')}"

        for cell in ws[4]:
            cell.fill = header_fill; cell.font = font_white; cell.alignment = center; cell.border = border
        
        for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = border; cell.alignment = center
                if cell.row % 2 == 0: cell.fill = row_fill
        
        for col in ws.columns:
            max_len = 0
            col_let = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if len(str(cell.value)) > max_len: max_len = len(str(cell.value))
                except: pass
            ws.column_dimensions[col_let].width = (max_len + 2) * 1.2
            
    output.seek(0)
    return output

df_employees, df_attendance = load_data()

# ==========================================
# 4. SIDEBAR MENU
# ==========================================
with st.sidebar:
    st.markdown("<h1 style='text-align: center; color: #38bdf8;'>⚡ HRIS PRO</h1>", unsafe_allow_html=True)
    selected = option_menu(
        menu_title=None,
        options=["Dashboard Karyawan", "Input Absensi", "Laporan Rekap"],
        icons=["people-fill", "clipboard-data", "file-earmark-bar-graph"],
        default_index=0,
        styles={
            "container": {"padding": "0!important", "background-color": "transparent"},
            "icon": {"color": "#38bdf8", "font-size": "18px"}, 
            "nav-link": {"font-size": "15px", "text-align": "left", "margin":"5px", "--hover-color": "#334155", "color": "#e2e8f0"},
            "nav-link-selected": {"background-color": "#3b82f6", "color": "white"},
        }
    )
    st.markdown("---")
    st.caption("Mode: Dark Premium")

# ==========================================
# 5. DASHBOARD KARYAWAN
# ==========================================
if selected == "Dashboard Karyawan":
    st.title("📂 Database Karyawan")
    st.markdown("---")
    
    if 'uploaded_template' not in st.session_state: st.session_state['uploaded_template'] = None
    if 'sheet_name_template' not in st.session_state: st.session_state['sheet_name_template'] = ""
    if 'header_row_template' not in st.session_state: st.session_state['header_row_template'] = 6

    if not df_employees.empty:
        # Metrics
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Pegawai", len(df_employees))
        dept_num = df_employees['Departemen'].nunique() if 'Departemen' in df_employees.columns else 0
        m2.metric("Departemen", dept_num)
        jab_num = df_employees['Jabatan'].nunique() if 'Jabatan' in df_employees.columns else 0
        m3.metric("Jabatan", jab_num)
        m4.metric("Status", "Active")
        st.write("")
        
        # Charts (Grafik)
        has_dept = 'Departemen' in df_employees.columns
        has_jab = 'Jabatan' in df_employees.columns
        if has_dept or has_jab:
            c1, c2 = st.columns(2)
            with c1:
                if has_dept:
                    # Ambil 10 Departemen terbanyak agar grafik tidak penuh sesak
                    d_cnt = df_employees['Departemen'].value_counts().head(10).reset_index()
                    d_cnt.columns = ['Departemen', 'Jumlah']
                    fig = px.bar(d_cnt, x='Departemen', y='Jumlah', color='Departemen', title="Top 10 Departemen", template='plotly_dark')
                    fig.update_layout(showlegend=False, height=320, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig, use_container_width=True)
            with c2:
                if has_jab:
                    # Ambil 10 Jabatan terbanyak
                    j_cnt = df_employees['Jabatan'].value_counts().head(10).reset_index()
                    j_cnt.columns = ['Jabatan', 'Jumlah']
                    fig = px.pie(j_cnt, names='Jabatan', values='Jumlah', title="Top 10 Jabatan", hole=0.5, template='plotly_dark')
                    fig.update_layout(height=320, paper_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig, use_container_width=True)
        st.divider()

    c_up, c_add = st.columns(2)
    with c_up:
        with st.expander("📥 Import Excel (SO MUT)", expanded=False):
            up_file = st.file_uploader("File .xlsx", type=['xlsx'])
            if up_file:
                st.session_state['uploaded_template'] = up_file
                try:
                    xls = pd.ExcelFile(up_file)
                    idx = 0
                    for i, n in enumerate(xls.sheet_names): 
                        if "DATABASE SESUAI SO".lower() in n.lower(): idx = i; break
                    sh = st.selectbox("Sheet:", xls.sheet_names, index=idx)
                    rw = st.number_input("Header Baris:", 1, 20, 6)
                    
                    if st.button("Load Data", type="primary"):
                        df = pd.read_excel(up_file, sheet_name=sh, header=rw-1, dtype=str)
                        
                        # --- SMART MAPPING V2 (SUPER LENGKAP) ---
                        # Kita bersihkan nama kolom dulu (hapus spasi, uppercase)
                        df.columns = [str(c).strip().upper() for c in df.columns]
                        
                        # Kamus terjemahan yang lebih luas
                        rename_map = {
                            "NO. INDUK": "NIK", "NIK/NRP": "NIK", "NO INDUK": "NIK",
                            "NAMA LENGKAP": "Nama", "NAMA KARYAWAN": "Nama", "NAMA": "Nama",
                            "JABATAN": "Jabatan", "POSISI": "Jabatan", "JABATAN BARU": "Jabatan",
                            "DEPARTEMEN": "Departemen", "DIVISI": "Departemen", "BAGIAN": "Departemen", "DEPARTEMEN BARU": "Departemen", "DEPT": "Departemen",
                            "PT": "PT", "PERUSAHAAN": "PT", "ENTITY": "PT", "PT BARU": "PT"
                        }
                        
                        # Proses Rename
                        df.rename(columns=rename_map, inplace=True)
                        
                        # Hapus kolom sampah
                        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                        
                        # Filter Data Valid (Wajib ada NIK)
                        if 'NIK' in df.columns: 
                            df = df[df['NIK'].notna()]
                            df = df[df['NIK'].astype(str).str.strip() != '']
                        
                        df_employees = df; save_data(df_employees, df_attendance)
                        st.session_state['sheet_name_template'] = sh; st.session_state['header_row_template'] = rw
                        st.success(f"Loaded {len(df)} rows."); st.rerun()
                except Exception as e: st.error(f"Error: {e}")

    with c_add:
        with st.expander("➕ Tambah Manual"):
            cols = [c for c in df_employees.columns if c not in ['No','Ceklist','Pilih']] or DEFAULT_COLS
            with st.form("add"):
                v = {}
                cg = st.columns(2)
                for i, col in enumerate(cols):
                    with cg[i%2]: v[col] = st.text_input(col)
                if st.form_submit_button("Simpan"):
                    if any(v.values()):
                        df_employees = pd.concat([df_employees, pd.DataFrame([v])], ignore_index=True)
                        save_data(df_employees, df_attendance); st.success("OK"); st.rerun()

    if not df_employees.empty:
        st.subheader("📝 Editor Data & Hapus")
        search = st.text_input("🔍 Filter Nama/NIK:")
        if search:
            df_show = df_employees[df_employees.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)].copy()
        else:
            df_show = df_employees.copy()
        
        df_show.insert(0, "Pilih", False)
        
        edited = st.data_editor(
            df_show, 
            num_rows="dynamic", 
            use_container_width=True, 
            hide_index=True, 
            height=450,
            column_config={"Pilih": st.column_config.CheckboxColumn("Hapus?", width="small")}
        )
        
        b1, b2, b3 = st.columns([1, 1, 3])
        
        with b1:
            if st.button("💾 Simpan Perubahan", type="primary"):
                if search: st.warning("Harap hapus kata kunci pencarian.")
                else: 
                    final_df = edited.drop(columns=['Pilih'])
                    save_data(final_df, df_attendance)
                    st.success("Tersimpan!"); st.rerun()
        
        with b2:
            if st.button("🗑️ Hapus Terpilih", type="secondary"):
                if search:
                    st.warning("Hapus filter dulu.")
                else:
                    rows_to_keep = edited[edited['Pilih'] == False]
                    final_df = rows_to_keep.drop(columns=['Pilih'])
                    deleted_count = len(df_employees) - len(final_df)
                    if deleted_count > 0:
                        save_data(final_df, df_attendance)
                        st.success(f"Dihapus: {deleted_count}!"); st.rerun()
                    else: st.info("Pilih data dulu.")

        with b3:
            if st.session_state['uploaded_template']:
                try:
                    out = io.BytesIO()
                    st.session_state['uploaded_template'].seek(0)
                    wb = load_workbook(st.session_state['uploaded_template'])
                    ws = wb[st.session_state['sheet_name_template']]
                    export_df = edited.drop(columns=['Pilih'])
                    rows = export_df.values.tolist()
                    start = st.session_state['header_row_template'] + 1
                    for i, r_data in enumerate(rows):
                        for j, val in enumerate(r_data):
                            ws.cell(row=i+start, column=j+1, value=val)
                    wb.save(out); out.seek(0)
                    st.download_button("📥 Download Excel Asli", out, "Update_Karyawan.xlsx")
                except: st.warning("Gagal format asli.")
            else:
                export_df = edited.drop(columns=['Pilih'])
                xl = create_colorful_excel(export_df, "DATABASE KARYAWAN")
                st.download_button("📥 Download Excel Baru", xl, "Database.xlsx")
    else: st.info("Database kosong.")

# ==========================================
# 6. INPUT ABSENSI (DENGAN FITUR HAPUS)
# ==========================================
elif selected == "Input Absensi":
    st.title("📝 Presensi Harian")
    st.markdown("---")
    if df_employees.empty: st.error("Database Kosong.")
    else:
        c1, c2 = st.columns([1, 2])
        with c1:
            with st.container(border=True):
                st.info("Form Input")
                cols = df_employees.columns
                cnik = next((c for c in cols if 'NIK' in c.upper()), cols[0])
                cnm = next((c for c in cols if 'NAMA' in c.upper()), cols[1])
                mst = df_employees[[cnik, cnm]].drop_duplicates().dropna()
                opts = [f"{r[cnik]} - {r[cnm]}" for _, r in mst.iterrows()]
                
                sel = st.selectbox("Karyawan:", opts)
                jenis = st.radio("Status:", ["Sakit", "Izin", "Alpha", "Cuti"], horizontal=True)
                tgl = st.date_input("Tanggal:", datetime.now())
                ket = st.text_area("Ket:", height=80)
                
                if st.button("Simpan", type="primary", use_container_width=True):
                    nik_val = sel.split(" - ")[0]
                    nm_val = sel.split(" - ")[1]
                    dpt_val = "-"
                    cdep = next((c for c in cols if 'DEP' in c.upper()), None)
                    if cdep:
                        tmp = df_employees[df_employees[cnik] == nik_val]
                        if not tmp.empty: dpt_val = tmp.iloc[0][cdep]
                    
                    new = {'Tanggal':tgl, 'NIK':nik_val, 'Nama':nm_val, 'Departemen':dpt_val, 
                           'Jenis':jenis, 'Keterangan':ket, 'Waktu_Input':datetime.now().strftime("%Y-%m-%d %H:%M")}
                    df_attendance = pd.concat([df_attendance, pd.DataFrame([new])], ignore_index=True)
                    save_data(df_employees, df_attendance); st.success("Masuk!"); st.rerun()

        with c2:
            st.subheader("Riwayat & Edit")
            if not df_attendance.empty:
                df_show = df_attendance.sort_values('Waktu_Input', ascending=False).reset_index(drop=True)
                df_show.insert(0, "Pilih", False)
                
                edited_history = st.data_editor(
                    df_show,
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        "Pilih": st.column_config.CheckboxColumn("Hapus?", width="small"),
                        "Waktu_Input": st.column_config.TextColumn("Waktu", disabled=True),
                        "Nama": st.column_config.TextColumn("Nama", disabled=True),
                        "Jenis": st.column_config.TextColumn("Jenis", disabled=True)
                    }
                )
                
                col_del, col_space = st.columns([1, 3])
                with col_del:
                    if st.button("🗑️ Hapus Data Terpilih", type="secondary"):
                        rows_to_keep = edited_history[edited_history['Pilih'] == False]
                        rows_to_keep = rows_to_keep.drop(columns=['Pilih'])
                        df_attendance = rows_to_keep
                        save_data(df_employees, df_attendance)
                        st.success("Dihapus!"); st.rerun()
            else: 
                st.info("Belum ada data absensi.")

# ==========================================
# 7. REKAP LAPORAN
# ==========================================
elif selected == "Laporan Rekap":
    st.title("📊 Laporan Bulanan")
    st.markdown("---")
    if df_employees.empty: st.warning("Database Kosong.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        bln = c1.selectbox("Bulan", range(1,13), index=datetime.now().month-1)
        thn = c2.number_input("Tahun", value=datetime.now().year)
        hk = c3.number_input("Hari Kerja", value=26)
        
        df_att_show = df_attendance.copy()
        if not df_att_show.empty:
            df_att_show['Tanggal'] = pd.to_datetime(df_att_show['Tanggal'])
            mask = (df_att_show['Tanggal'].dt.month == bln) & (df_att_show['Tanggal'].dt.year == thn)
            fil = df_att_show[mask]
            
            c4.metric("Total Absen", len(fil))
            
            abs_cnt = fil.groupby('NIK').size().reset_index(name='Total_Absen')
            cols = df_employees.columns
            cnik = next((c for c in cols if 'NIK' in c.upper()), cols[0])
            cnm = next((c for c in cols if 'NAMA' in c.upper()), cols[1])
            cdep = next((c for c in cols if 'DEP' in c.upper()), None)
            
            if cdep: mst = df_employees[[cnik, cnm, cdep]].drop_duplicates()
            else: mst = df_employees[[cnik, cnm]].drop_duplicates(); mst['Departemen'] = "-"
            
            mst.columns = ['NIK', 'Nama', 'Departemen']
            mst['NIK'] = mst['NIK'].astype(str); abs_cnt['NIK'] = abs_cnt['NIK'].astype(str)
            fin = pd.merge(mst, abs_cnt, on='NIK', how='left')
            fin['Total_Absen'] = fin['Total_Absen'].fillna(0).astype(int)
            fin['Total_Hadir'] = (hk - fin['Total_Absen']).clip(lower=0)
            fin['Persentase'] = ((fin['Total_Hadir']/hk)*100).round(1).astype(str) + '%'
            
            st.dataframe(fin, use_container_width=True, hide_index=True)
            
            xl = create_colorful_excel(fin, f"REKAP {bln}/{thn}")
            st.download_button("📥 Download Laporan (Excel)", xl, f"Rekap_{bln}_{thn}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            
            st.divider()
            if not fil.empty:
                g1, g2 = st.columns(2)
                with g1:
                    pc = fil['Jenis'].value_counts().reset_index(); pc.columns=['Jenis','Jml']
                    fig = px.pie(pc, names='Jenis', values='Jml', title="Proporsi Izin", hole=0.4, template="plotly_dark")
                    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig, use_container_width=True)
                with g2:
                    top = fin[fin['Total_Absen']>0].sort_values('Total_Absen', ascending=False).head(5)
                    fig = px.bar(top, x='Total_Absen', y='Nama', orientation='h', title="Top 5 Absen", template="plotly_dark")
                    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig, use_container_width=True)
        else: st.info("Belum ada data absensi.")