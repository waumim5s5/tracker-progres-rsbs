import streamlit as st
import os
import json
import pandas as pd
from datetime import datetime
import shutil

# --- MODUL UNTUK EXCEL & GAMBAR ---
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
        json.dump({
            "tasks": st.session_state.database_tasks,
            "workers": st.session_state.database_workers
        }, f, indent=4)

def muat_database():
    if os.path.exists(FILE_DATABASE):
        try:
            with open(FILE_DATABASE, 'r') as f:
                data = json.load(f)
                if "tasks" not in data:
                    return {"tasks": data, "workers": {}}
                return data
        except:
            return {"tasks": {}, "workers": {}}
    else:
        return {"tasks": {}, "workers": {}}

if 'db_loaded' not in st.session_state:
    db = muat_database()
    st.session_state.database_tasks = db["tasks"]
    st.session_state.database_workers = db["workers"]
    st.session_state.db_loaded = True
    simpan_database()


# ==========================================
# 2. SIDEBAR (PANEL INPUT, HAPUS, & RESTORE)
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
        if st.session_state.database_tasks:
            pilih_area_1 = st.selectbox("Pilih Area", list(st.session_state.database_tasks.keys()), key="sb_area_pek")
            with st.form("form_pekerjaan"):
                pekerjaan_baru = st.text_input("Nama Pekerjaan")
                if st.form_submit_button("Tambah Pekerjaan") and pekerjaan_baru:
                    if pekerjaan_baru not in st.session_state.database_tasks[pilih_area_1]:
                        st.session_state.database_tasks[pilih_area_1][pekerjaan_baru] = {}
                        simpan_database()
                        st.success("Berhasil! Refresh halaman.")
        else:
            st.write("Buat area dulu.")

    with st.expander("3. Tambah Printilan"):
        if st.session_state.database_tasks:
            pilih_area_2 = st.selectbox("Pilih Area", list(st.session_state.database_tasks.keys()), key="sb_area_prin")
            if st.session_state.database_tasks[pilih_area_2]:
                pilih_pekerjaan = st.selectbox("Pilih Pekerjaan", list(st.session_state.database_tasks[pilih_area_2].keys()), key="sb_pek_prin")
                with st.form("form_printilan"):
                    printilan_baru = st.text_input("Nama Printilan")
                    if st.form_submit_button("Tambah Printilan") and printilan_baru:
                        if printilan_baru not in st.session_state.database_tasks[pilih_area_2][pilih_pekerjaan]:
                            # DISINI DIUBAH: Default progress printilan adalah angka 0, bukan False
                            st.session_state.database_tasks[pilih_area_2][pilih_pekerjaan][printilan_baru] = 0
                            simpan_database()
                            st.success("Berhasil! Refresh halaman.")
            else:
                st.write("Buat pekerjaan dulu.")
        else:
            st.write("Buat area dulu.")
    
    st.divider()

    st.header("🗑️ Hapus Data")
    with st.expander("Panel Hapus Data (Hati-hati!)"):
        if st.session_state.database_tasks:
            jenis_hapus = st.radio("Apa yang ingin dihapus?", ["Area Seluruhnya", "Pekerjaan Utama", "Printilan (Checklist)"])
            
            if jenis_hapus == "Area Seluruhnya":
                area_hapus = st.selectbox("Pilih Area yang DIHAPUS", list(st.session_state.database_tasks.keys()), key="del_area")
                if st.button("🚨 HAPUS AREA INI"):
                    del st.session_state.database_tasks[area_hapus]
                    simpan_database()
                    st.success(f"Area {area_hapus} terhapus!")
                    st.rerun()

            elif jenis_hapus == "Pekerjaan Utama":
                area_pilih = st.selectbox("Dari Area Mana?", list(st.session_state.database_tasks.keys()), key="del_area_pek")
                if st.session_state.database_tasks[area_pilih]:
                    pek_hapus = st.selectbox("Pilih Pekerjaan", list(st.session_state.database_tasks[area_pilih].keys()), key="del_pek")
                    if st.button("🚨 HAPUS PEKERJAAN INI"):
                        del st.session_state.database_tasks[area_pilih][pek_hapus]
                        st.session_state.database_workers.pop(f"{area_pilih}_{pek_hapus}_tukang", None)
                        st.session_state.database_workers.pop(f"{area_pilih}_{pek_hapus}_peladen", None)
                        simpan_database()
                        st.success(f"Pekerjaan {pek_hapus} terhapus!")
                        st.rerun()
                else:
                    st.write("Tidak ada pekerjaan di area ini.")

            elif jenis_hapus == "Printilan (Checklist)":
                area_pilih2 = st.selectbox("Pilih Area", list(st.session_state.database_tasks.keys()), key="del_area_prin")
                if st.session_state.database_tasks[area_pilih2]:
                    pek_pilih = st.selectbox("Pilih Pekerjaan", list(st.session_state.database_tasks[area_pilih2].keys()), key="del_pek_prin")
                    if st.session_state.database_tasks[area_pilih2][pek_pilih]:
                        prin_hapus = st.selectbox("Pilih Printilan", list(st.session_state.database_tasks[area_pilih2][pek_pilih].keys()), key="del_prin")
                        if st.button("🚨 HAPUS PRINTILAN INI"):
                            del st.session_state.database_tasks[area_pilih2][pek_pilih][prin_hapus]
                            simpan_database()
                            st.success(f"Printilan {prin_hapus} terhapus!")
                            st.rerun()
                    else:
                        st.write("Tidak ada printilan.")
                else:
                    st.write("Tidak ada pekerjaan.")
        else:
            st.info("Database kosong.")

    st.divider()
    if st.button("🔄 Muat Ulang Halaman (Refresh)"):
        st.rerun()
    st.divider()

    # --- FITUR RESTORE DATA (UPLOAD EXCEL) ---
    st.header("📂 Restore Data (Upload Excel)")
    file_excel_upload = st.file_uploader("Upload Excel Backup (*.xlsx)", type=["xlsx"])
    if file_excel_upload:
        if st.button("🔄 Pulihkan Data"):
            try:
                df_upload = pd.read_excel(file_excel_upload)
                database_baru = {}
                workers_baru = {}
                
                for index, row in df_upload.iterrows():
                    area = str(row['Area']).strip()
                    pekerjaan = str(row['Pekerjaan Utama']).strip()
                    tukang = str(row.get('Nama Tukang', '')).strip()
                    peladen = str(row.get('Nama Peladen', '')).strip()
                    printilan = str(row['Item Printilan']).strip()
                    progres_val = float(row.get('Progres (%)', 0))
                    
                    if area not in database_baru: database_baru[area] = {}
                    if pekerjaan not in database_baru[area]: database_baru[area][pekerjaan] = {}
                    
                    if tukang and tukang.lower() != "nan" and tukang != "-":
                        workers_baru[f"{area}_{pekerjaan}_tukang"] = tukang
                    if peladen and peladen.lower() != "nan" and peladen != "-":
                        workers_baru[f"{area}_{pekerjaan}_peladen"] = peladen
                        
                    if printilan != "-" and printilan.lower() != "nan":
                        # Restore angka progres ke database
                        database_baru[area][pekerjaan][printilan] = int(progres_val * 100)
                
                st.session_state.database_tasks = database_baru
                st.session_state.database_workers = workers_baru
                simpan_database()
                st.success("✅ Berhasil dipulihkan!")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal memulihkan: {e}")

    st.divider()
    
    # --- FITUR EXPORT EXCEL (+ FOTO) ---
    st.header("📊 Export Laporan (+ Foto)")
    if st.button("Siapkan File Excel Lengkap"):
        data_untuk_excel = []
        for area, dict_pekerjaan in st.session_state.database_tasks.items():
            for pekerjaan, dict_printilan in dict_pekerjaan.items():
                
                # Perhitungan Progres Rata-rata dari Slider
                jumlah_printilan = len(dict_printilan)
                if jumlah_printilan > 0:
                    total_progres = 0
                    for val in dict_printilan.values():
                        # Keamanan: ubah True/False jadi angka 100/0 jika ada data lama
                        if isinstance(val, bool): val = 100 if val else 0
                        total_progres += val
                    progres_rata = (total_progres / (jumlah_printilan * 100))
                else:
                    progres_rata = 0
                
                nama_tukang = st.session_state.database_workers.get(f"{area}_{pekerjaan}_tukang", "-")
                nama_peladen = st.session_state.database_workers.get(f"{area}_{pekerjaan}_peladen", "-")
                
                data_untuk_excel.append({"Area": area, "Pekerjaan Utama": pekerjaan, "Nama Tukang": nama_tukang, "Nama Peladen": nama_peladen, "Item Printilan": "-", "Status": "PROGRES KESELURUHAN", "Progres (%)": progres_rata})
                
                for printilan, val in dict_printilan.items():
                    if isinstance(val, bool): val = 100 if val else 0
                    status_text = "Selesai" if val == 100 else f"Proses {val}%"
                    data_untuk_excel.append({"Area": area, "Pekerjaan Utama": pekerjaan, "Nama Tukang": nama_tukang, "Nama Peladen": nama_peladen, "Item Printilan": printilan, "Status": status_text, "Progres (%)": (val / 100)})
        
        if data_untuk_excel:
            df = pd.DataFrame(data_untuk_excel)
            file_excel_temp = os.path.join(DIR_SAAT_INI, "Laporan_Temp.xlsx")
            df.to_excel(file_excel_temp, index=False)
            
            try:
                wb = openpyxl.load_workbook(file_excel_temp)
                ws = wb.active
                
                ws.cell(row=1, column=8, value="Dokumentasi (Foto)")
                ws.column_dimensions['H'].width = 25 
                
                daftar_foto = os.listdir(FOLDER_FOTO)
                
                for row_idx, row_data in enumerate(data_untuk_excel, start=2):
                    area = row_data["Area"]
                    pekerjaan = row_data["Pekerjaan Utama"]
                    item_printilan = row_data["Item Printilan"]
                    
                    if item_printilan == "-":
                        ws.row_dimensions[row_idx].height = 80 
                        prefix = f"{area}_{pekerjaan}_".replace(" ", "_")
                        foto_ditemukan = None
                        
                        for f in daftar_foto:
                            if f.startswith(prefix):
                                foto_ditemukan = os.path.join(FOLDER_FOTO, f)
                                break
                        
                        if foto_ditemukan:
                            try:
                                img_pil = PILImage.open(foto_ditemukan)
                                if img_pil.mode != 'RGB':
                                    img_pil = img_pil.convert('RGB')
                                img_pil.thumbnail((150, 100)) 
                                img_byte_arr = io.BytesIO()
                                img_pil.save(img_byte_arr, format='JPEG')
                                img_xl = OpenPyxlImage(img_byte_arr)
                                ws.add_image(img_xl, f"H{row_idx}") 
                            except Exception as e:
                                pass 
                
                file_excel_final = os.path.join(DIR_SAAT_INI, "Laporan_Progres_Final.xlsx")
                wb.save(file_excel_final)
                
                with open(file_excel_final, "rb") as f:
                    st.download_button(label="📥 Unduh Excel (+ Foto)", data=f, file_name=f"Laporan_Lengkap_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            except Exception as e:
                st.error(f"Gagal memproses foto ke Excel: {e}")


# ============================================================
# 3. HALAMAN UTAMA (TRACKER)
# ============================================================
st.title("🏗️ Aplikasi Tracker Progres Proyek")

if not st.session_state.database_tasks:
    st.info("👈 Data kosong. Tambah data di panel kiri.")
else:
    daftar_area = list(st.session_state.database_tasks.keys())
    if "ingat_area" not in st.session_state or st.session_state.ingat_area not in daftar_area:
        st.session_state.ingat_area = daftar_area[0]
        
    area_terpilih = st.selectbox(
        "📍 Pilih Area Pekerjaan:", 
        daftar_area,
        key="ingat_area"
    )
    st.divider()

    pekerjaan_utama_dict = st.session_state.database_tasks[area_terpilih]

    if not pekerjaan_utama_dict:
        st.warning(f"Belum ada pekerjaan di area {area_terpilih}.")
    else:
        for main_task, sub_tasks in pekerjaan_utama_dict.items():
            with st.expander(f"🛠️ {main_task}", expanded=True):
                
                kunci_tukang = f"{area_terpilih}_{main_task}_tukang"
                kunci_peladen = f"{area_terpilih}_{main_task}_peladen"
                
                tukang_saat_ini = st.session_state.database_workers.get(kunci_tukang, "")
                peladen_saat_ini = st.session_state.database_workers.get(kunci_peladen, "")
                
                def simpan_tukang(k=kunci_tukang):
                    st.session_state.database_workers[k] = st.session_state[f"input_tukang_{k}"]
                    simpan_database()
                    
                def simpan_peladen(k=kunci_peladen):
                    st.session_state.database_workers[k] = st.session_state[f"input_peladen_{k}"]
                    simpan_database()
                
                col_t, col_p = st.columns(2)
                with col_t:
                    st.text_input("👷 Nama Tukang:", value=tukang_saat_ini, key=f"input_tukang_{kunci_tukang}", on_change=simpan_tukang, placeholder="Ketik nama & Enter")
                with col_p:
                    st.text_input("👷‍♂️ Nama Peladen:", value=peladen_saat_ini, key=f"input_peladen_{kunci_peladen}", on_change=simpan_peladen, placeholder="Ketik nama & Enter")
                
                if not sub_tasks:
                    st.write("*Belum ada checklist.*")
                else:
                    # Menghitung Rata-Rata Progres dari Slider Printilan
                    total_printilan = len(sub_tasks)
                    total_skor = 0
                    
                    for p_val in sub_tasks.values():
                        if isinstance(p_val, bool): p_val = 100 if p_val else 0
                        total_skor += p_val
                        
                    persentase = int(total_skor / total_printilan) if total_printilan > 0 else 0
                    
                    st.progress(persentase / 100)
                    st.markdown(f"**Progres Pekerjaan: {persentase}%**")
                    st.write("---")
                    
                    # === UPDATE: LOOP UNTUK SLIDER PRINTILAN ===
                    for printilan, val in sub_tasks.items():
                        
                        # Pengaman jika database lama terbaca bool (True/False)
                        if isinstance(val, bool):
                            val = 100 if val else 0
                            st.session_state.database_tasks[area_terpilih][main_task][printilan] = val
                            
                        unik_key = f"sld_{area_terpilih}_{main_task}_{printilan}"
                        
                        def update_status(a=area_terpilih, p=main_task, pr=printilan, key=unik_key):
                            st.session_state.database_tasks[a][p][pr] = st.session_state[key]
                            simpan_database()
                        
                        c1, c2 = st.columns([1, 10])
                        with c1:
                            if val == 100:
                                st.markdown("<h4 style='color:green; margin-top:20px;'>✅</h4>", unsafe_allow_html=True)
                            else:
                                st.markdown("<h4 style='color:gray; margin-top:20px;'>⚙️</h4>", unsafe_allow_html=True)
                        with c2:
                            st.slider(
                                printilan, 
                                min_value=0, max_value=100, value=val, step=5,
                                key=unik_key, 
                                on_change=update_status
                            )
                
                st.write("---")
                with st.form(f"form_foto_{main_task}", clear_on_submit=True):
                    foto_file = st.file_uploader("📸 Upload Foto Dokumentasi", type=["jpg", "png", "jpeg"])
                    submit_foto = st.form_submit_button("Simpan Foto")
                    if submit_foto and foto_file:
                        nama_file = f"{area_terpilih}_{main_task}_{foto_file.name}".replace(" ", "_")
                        path_simpan = os.path.join(FOLDER_FOTO, nama_file)
                        with open(path_simpan, "wb") as f:
                            f.write(foto_file.getbuffer())
                        st.success("✅ Foto tersimpan!")

# ==========================================
# 4. GENERATOR LAPORAN WHATSAPP 
# ==========================================
st.divider()
st.header("📱 Buat Laporan WhatsApp")
st.write("Teks di bawah ini dibuat otomatis berdasarkan hasil geser slider Anda.")

def generate_wa_text():
    lines = []
    lines.append("*ITEM PEKERJAAN DAN PROGRES*")
    lines.append("")
    
    area_idx = 1
    for area, dict_pekerjaan in st.session_state.database_tasks.items():
        lines.append(f"{area_idx}. *{area}*")
        
        for pekerjaan, dict_printilan in dict_pekerjaan.items():
            jumlah_printilan = len(dict_printilan)
            
            # Hitung persentase utama dari slider anak
            if jumlah_printilan > 0:
                total_skor = 0
                for v in dict_printilan.values():
                    if isinstance(v, bool): v = 100 if v else 0
                    total_skor += v
                persen = int(total_skor / jumlah_printilan)
            else:
                persen = 0
            
            nama_tukang = st.session_state.database_workers.get(f"{area}_{pekerjaan}_tukang", "")
            pekerja_info = f" ({nama_tukang})" if nama_tukang else ""
            
            if jumlah_printilan == 0:
                lines.append(f"👷🏻‍♂️ Pekerjaan {pekerjaan}{pekerja_info} ({persen}%)")
            else:
                lines.append(f"• *{pekerjaan}*{pekerja_info} ({persen}%)")
                
                prin_idx = 1
                for printilan, val in dict_printilan.items():
                    if isinstance(val, bool): val = 100 if val else 0
                    status_simbol = "✅" if val == 100 else ""
                    
                    # Menampilkan angka % di sebelah nama printilan
                    lines.append(f"   {prin_idx}. {printilan} ({val}%){status_simbol}")
                    prin_idx += 1
                    
        lines.append("")
        area_idx += 1
        
    return "\n".join(lines)

if st.session_state.database_tasks:
    teks_laporan_wa = generate_wa_text()
    st.text_area("Kolom Copy-Paste Laporan WA", value=teks_laporan_wa, height=400)
