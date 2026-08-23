import streamlit as st
import os
import json
import pandas as pd
from datetime import datetime
import shutil

# --- TAMBAHAN UNTUK EXCEL & GAMBAR ---
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
from PIL import Image as PILImage
import io

st.set_page_config(page_title="Tracker Proyek RS", page_icon="🏗️", layout="wide")

# ==========================================
# 1. PERSIAPAN FOLDER & FILE DATABASE
# ==========================================
DIR_SAAT_INI = os.path.dirname(os.path.abspath(__file__))
FILE_DATABASE = os.path.join(DIR_SAAT_INI, "database_proyek.json")
FOLDER_FOTO = os.path.join(DIR_SAAT_INI, "foto_progres")

if not os.path.exists(FOLDER_FOTO):
    os.makedirs(FOLDER_FOTO)

def simpan_database():
    with open(FILE_DATABASE, 'w') as f:
        json.dump(st.session_state.database_tasks, f, indent=4)

def muat_database():
    if os.path.exists(FILE_DATABASE):
        try:
            with open(FILE_DATABASE, 'r') as f:
                return json.load(f)
        except:
            return {}
    else:
        return {}

if 'database_tasks' not in st.session_state:
    st.session_state.database_tasks = muat_database()
    simpan_database()


# ==========================================
# 2. SIDEBAR (PANEL INPUT & RESTORE)
# ==========================================
with st.sidebar:
    st.header("➕ Manajemen Data")
    
    with st.expander("1. Tambah Area Baru"):
        with st.form("form_area"):
            area_baru = st.text_input("Nama Area")
            if st.form_submit_button("Tambah Area") and area_baru:
                if area_baru not in st.session_state.database_tasks:
                    st.session_state.database_tasks[area_baru] = {}
                    simpan_database()
                    st.success("Berhasil! Refresh halaman.")

    with st.expander("2. Tambah Pekerjaan Utama"):
        with st.form("form_pekerjaan"):
            if st.session_state.database_tasks:
                pilih_area_1 = st.selectbox("Pilih Area", list(st.session_state.database_tasks.keys()))
                pekerjaan_baru = st.text_input("Nama Pekerjaan")
                if st.form_submit_button("Tambah Pekerjaan") and pekerjaan_baru:
                    if pekerjaan_baru not in st.session_state.database_tasks[pilih_area_1]:
                        st.session_state.database_tasks[pilih_area_1][pekerjaan_baru] = {}
                        simpan_database()
                        st.success("Berhasil! Refresh halaman.")
            else:
                st.write("Buat area dulu.")

    with st.expander("3. Tambah Printilan"):
        with st.form("form_printilan"):
            if st.session_state.database_tasks:
                pilih_area_2 = st.selectbox("Pilih Area", list(st.session_state.database_tasks.keys()))
                if st.session_state.database_tasks[pilih_area_2]:
                    pilih_pekerjaan = st.selectbox("Pilih Pekerjaan", list(st.session_state.database_tasks[pilih_area_2].keys()))
                    printilan_baru = st.text_input("Nama Printilan")
                    if st.form_submit_button("Tambah Printilan") and printilan_baru:
                        if printilan_baru not in st.session_state.database_tasks[pilih_area_2][pilih_pekerjaan]:
                            st.session_state.database_tasks[pilih_area_2][pilih_pekerjaan][printilan_baru] = False
                            simpan_database()
                            st.success("Berhasil! Refresh halaman.")
    st.divider()

    st.header("📂 Restore Data (Upload Excel)")
    file_excel_upload = st.file_uploader("Upload Excel Backup (*.xlsx)", type=["xlsx"])
    if file_excel_upload:
        if st.button("🔄 Pulihkan Data"):
            try:
                df_upload = pd.read_excel(file_excel_upload)
                database_baru = {}
                for index, row in df_upload.iterrows():
                    area = str(row['Area']).strip()
                    pekerjaan = str(row['Pekerjaan Utama']).strip()
                    printilan = str(row['Item Printilan']).strip()
                    status_text = str(row['Status']).strip()
                    
                    if area not in database_baru: database_baru[area] = {}
                    if pekerjaan not in database_baru[area]: database_baru[area][pekerjaan] = {}
                        
                    if printilan != "-" and printilan.lower() != "nan":
                        database_baru[area][pekerjaan][printilan] = (status_text.lower() == "selesai")
                
                st.session_state.database_tasks = database_baru
                simpan_database()
                st.success("✅ Berhasil dipulihkan!")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal memulihkan: {e}")

    st.divider()
    
    # ==========================================
    # FITUR EXPORT EXCEL DENGAN FOTO INSERT
    # ==========================================
    st.header("📊 Export Laporan (+ Foto)")
    st.write("Unduh Excel lengkap dengan fotonya.")
    
    if st.button("Siapkan File Excel Lengkap"):
        data_untuk_excel = []
        for area, dict_pekerjaan in st.session_state.database_tasks.items():
            for pekerjaan, dict_printilan in dict_pekerjaan.items():
                total = len(dict_printilan)
                selesai = sum(dict_printilan.values()) if total > 0 else 0
                progres_persen = (selesai / total) if total > 0 else 0
                
                # Masukkan baris judul pekerjaan (Di sini foto akan ditempel nanti)
                data_untuk_excel.append({"Area": area, "Pekerjaan Utama": pekerjaan, "Item Printilan": "-", "Status": "PROGRES KESELURUHAN", "Progres (%)": progres_persen})
                
                # Masukkan baris anak-anaknya (printilan)
                for printilan, status in dict_printilan.items():
                    data_untuk_excel.append({"Area": area, "Pekerjaan Utama": pekerjaan, "Item Printilan": printilan, "Status": "Selesai" if status else "Belum", "Progres (%)": 1.0 if status else 0.0})
        
        if data_untuk_excel:
            df = pd.DataFrame(data_untuk_excel)
            file_excel_temp = os.path.join(DIR_SAAT_INI, "Laporan_Temp.xlsx")
            df.to_excel(file_excel_temp, index=False)
            
            # --- PROSES MENEMPELKAN FOTO KE EXCEL ---
            try:
                wb = openpyxl.load_workbook(file_excel_temp)
                ws = wb.active
                
                # Tambah judul kolom untuk Foto
                ws.cell(row=1, column=6, value="Dokumentasi (Foto)")
                ws.column_dimensions['F'].width = 25 # Lebar kolom
                
                daftar_foto = os.listdir(FOLDER_FOTO)
                
                # Cek setiap baris dari baris ke-2
                for row_idx, row_data in enumerate(data_untuk_excel, start=2):
                    area = row_data["Area"]
                    pekerjaan = row_data["Pekerjaan Utama"]
                    item_printilan = row_data["Item Printilan"]
                    
                    # Kita hanya menaruh foto di baris "PROGRES KESELURUHAN"
                    if item_printilan == "-":
                        ws.row_dimensions[row_idx].height = 80 # Tinggikan baris
                        
                        # Cari foto yang nama file-nya diawali dengan area & pekerjaan tersebut
                        prefix = f"{area}_{pekerjaan}_".replace(" ", "_")
                        foto_ditemukan = None
                        
                        # Cari foto terbaru/terkait
                        for f in daftar_foto:
                            if f.startswith(prefix):
                                foto_ditemukan = os.path.join(FOLDER_FOTO, f)
                                break
                        
                        # Jika fotonya ada, perkecil dan tempelkan!
                        if foto_ditemukan:
                            try:
                                img_pil = PILImage.open(foto_ditemukan)
                                # Konversi ke RGB agar aman (mencegah error gambar transparansi/PNG)
                                if img_pil.mode != 'RGB':
                                    img_pil = img_pil.convert('RGB')
                                
                                # Resize gambar agar tidak kebesaran di Excel (Lebar 150px)
                                img_pil.thumbnail((150, 100)) 
                                
                                img_byte_arr = io.BytesIO()
                                img_pil.save(img_byte_arr, format='JPEG')
                                
                                img_xl = OpenPyxlImage(img_byte_arr)
                                ws.add_image(img_xl, f"F{row_idx}")
                            except Exception as e:
                                pass # Jika file rusak, abaikan
                
                # Simpan final
                file_excel_final = os.path.join(DIR_SAAT_INI, "Laporan_Progres_Final.xlsx")
                wb.save(file_excel_final)
                
                with open(file_excel_final, "rb") as f:
                    st.download_button(label="📥 Unduh Excel (+ Foto)", data=f, file_name=f"Laporan_Lengkap_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            except Exception as e:
                st.error(f"Gagal memproses foto ke Excel: {e}")

# ==========================================
# 3. HALAMAN UTAMA (TRACKER)
# ==========================================
st.title("🏗️ Aplikasi Tracker Progres Proyek")

if not st.session_state.database_tasks:
    st.info("👈 Data kosong. Tambah data di panel kiri.")
else:
    area_terpilih = st.selectbox("📍 Pilih Area Pekerjaan:", list(st.session_state.database_tasks.keys()))
    st.divider()

    pekerjaan_utama_dict = st.session_state.database_tasks[area_terpilih]

    if not pekerjaan_utama_dict:
        st.warning(f"Belum ada pekerjaan di area {area_terpilih}.")
    else:
        for main_task, sub_tasks in pekerjaan_utama_dict.items():
            with st.expander(f"🛠️ {main_task}", expanded=True):
                if not sub_tasks:
                    st.write("*Belum ada checklist.*")
                else:
                    total_printilan = len(sub_tasks)
                    printilan_selesai = sum(sub_tasks.values())
                    persentase = int((printilan_selesai / total_printilan) * 100) if total_printilan > 0 else 0
                    
                    st.progress(persentase / 100)
                    st.markdown(f"**Progres Pekerjaan: {persentase}%**  *( {printilan_selesai} / {total_printilan} Selesai )*")
                    st.write("---")
                    
                    for printilan, is_done in sub_tasks.items():
                        unik_key = f"{area_terpilih}_{main_task}_{printilan}"
                        def update_status(a=area_terpilih, p=main_task, pr=printilan, key=unik_key):
                            st.session_state.database_tasks[a][p][pr] = st.session_state[key]
                            simpan_database()
                        
                        st.checkbox(printilan, value=is_done, key=unik_key, on_change=update_status)
                
                st.write("---")
                with st.form(f"form_foto_{main_task}", clear_on_submit=True):
                    foto_file = st.file_uploader("📸 Upload Foto Dokumentasi (Khusus dari HP)", type=["jpg", "png", "jpeg"])
                    submit_foto = st.form_submit_button("Simpan Foto")
                    if submit_foto and foto_file:
                        # Penamaan foto distandarisasi agar mudah dicari oleh Excel nanti
                        nama_file = f"{area_terpilih}_{main_task}_{foto_file.name}".replace(" ", "_")
                        path_simpan = os.path.join(FOLDER_FOTO, nama_file)
                        with open(path_simpan, "wb") as f:
                            f.write(foto_file.getbuffer())
                        st.success("✅ Foto dokumentasi tersimpan dan terhubung dengan laporan Excel!")
