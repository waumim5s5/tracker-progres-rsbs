import streamlit as st
import os
import json
import pandas as pd
from datetime import datetime

# Konfigurasi Halaman (Harus di baris paling atas)
st.set_page_config(page_title="Tracker Proyek RS", page_icon="🏗️", layout="wide")

# ==========================================
# 1. PERSIAPAN FOLDER & FILE DATABASE (ABSOLUTE PATH)
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
# 2. SIDEBAR (PANEL INPUT, HAPUS, & EXPORT)
# ==========================================
with st.sidebar:
    st.header("➕ Manajemen Data")
    
    # --- FORM TAMBAH AREA ---
    with st.expander("1. Tambah Area Baru"):
        with st.form("form_area"):
            area_baru = st.text_input("Nama Area (Contoh: RSBS LT 1)")
            submit_area = st.form_submit_button("Tambah Area")
            if submit_area and area_baru:
                if area_baru not in st.session_state.database_tasks:
                    st.session_state.database_tasks[area_baru] = {}
                    simpan_database()
                    st.success("Area Berhasil Ditambahkan!")
                else:
                    st.warning("Area tersebut sudah ada.")

    # --- FORM TAMBAH PEKERJAAN ---
    with st.expander("2. Tambah Pekerjaan Utama"):
        with st.form("form_pekerjaan"):
            if st.session_state.database_tasks:
                pilih_area_1 = st.selectbox("Pilih Area", list(st.session_state.database_tasks.keys()))
                pekerjaan_baru = st.text_input("Nama Pekerjaan (Contoh: Pasang Bata)")
                submit_pek = st.form_submit_button("Tambah Pekerjaan")
                if submit_pek and pekerjaan_baru:
                    if pekerjaan_baru not in st.session_state.database_tasks[pilih_area_1]:
                        st.session_state.database_tasks[pilih_area_1][pekerjaan_baru] = {}
                        simpan_database()
                        st.success("Pekerjaan Berhasil Ditambahkan!")
            else:
                st.write("Buat area dulu.")

    # --- FORM TAMBAH PRINTILAN (CHECKLIST) ---
    with st.expander("3. Tambah Printilan"):
        with st.form("form_printilan"):
            if st.session_state.database_tasks:
                pilih_area_2 = st.selectbox("Pilih Area", list(st.session_state.database_tasks.keys()))
                if st.session_state.database_tasks[pilih_area_2]:
                    pilih_pekerjaan = st.selectbox("Pilih Pekerjaan", list(st.session_state.database_tasks[pilih_area_2].keys()))
                    printilan_baru = st.text_input("Nama Printilan")
                    submit_prin = st.form_submit_button("Tambah Printilan")
                    if submit_prin and printilan_baru:
                        if printilan_baru not in st.session_state.database_tasks[pilih_area_2][pilih_pekerjaan]:
                            st.session_state.database_tasks[pilih_area_2][pilih_pekerjaan][printilan_baru] = False
                            simpan_database()
                            st.success("Printilan Berhasil Ditambahkan!")
                else:
                    st.write("Buat pekerjaan dulu.")
            else:
                st.write("Buat area dulu.")
    
    st.divider()

    # --- FITUR HAPUS DATA ---
    st.header("🗑️ Hapus Data")
    with st.expander("Panel Hapus Data (Hati-hati!)"):
        if st.session_state.database_tasks:
            # Pilihan Tingkat Hapus (Area / Pekerjaan / Printilan)
            jenis_hapus = st.radio("Apa yang ingin dihapus?", ["Area Seluruhnya", "Pekerjaan Utama", "Printilan (Checklist)"])
            
            with st.form("form_hapus"):
                # 1. Jika Hapus Area
                if jenis_hapus == "Area Seluruhnya":
                    area_hapus = st.selectbox("Pilih Area yang akan DIHAPUS", list(st.session_state.database_tasks.keys()))
                    submit_hapus = st.form_submit_button("🚨 HAPUS AREA INI")
                    if submit_hapus:
                        del st.session_state.database_tasks[area_hapus]
                        simpan_database()
                        st.success(f"Area {area_hapus} berhasil dihapus!")

                # 2. Jika Hapus Pekerjaan Utama
                elif jenis_hapus == "Pekerjaan Utama":
                    area_pilih = st.selectbox("Dari Area Mana?", list(st.session_state.database_tasks.keys()))
                    if st.session_state.database_tasks[area_pilih]:
                        pek_hapus = st.selectbox("Pilih Pekerjaan yang akan DIHAPUS", list(st.session_state.database_tasks[area_pilih].keys()))
                        submit_hapus = st.form_submit_button("🚨 HAPUS PEKERJAAN INI")
                        if submit_hapus:
                            del st.session_state.database_tasks[area_pilih][pek_hapus]
                            simpan_database()
                            st.success(f"Pekerjaan {pek_hapus} berhasil dihapus!")
                    else:
                        st.write("Tidak ada pekerjaan di area ini.")
                        submit_hapus = st.form_submit_button("Tutup", disabled=True)

                # 3. Jika Hapus Printilan
                elif jenis_hapus == "Printilan (Checklist)":
                    area_pilih2 = st.selectbox("Pilih Area", list(st.session_state.database_tasks.keys()))
                    if st.session_state.database_tasks[area_pilih2]:
                        pek_pilih = st.selectbox("Pilih Pekerjaan", list(st.session_state.database_tasks[area_pilih2].keys()))
                        if st.session_state.database_tasks[area_pilih2][pek_pilih]:
                            prin_hapus = st.selectbox("Pilih Printilan yang akan DIHAPUS", list(st.session_state.database_tasks[area_pilih2][pek_pilih].keys()))
                            submit_hapus = st.form_submit_button("🚨 HAPUS PRINTILAN INI")
                            if submit_hapus:
                                del st.session_state.database_tasks[area_pilih2][pek_pilih][prin_hapus]
                                simpan_database()
                                st.success(f"Printilan {prin_hapus} berhasil dihapus!")
                        else:
                            st.write("Tidak ada printilan di pekerjaan ini.")
                            submit_hapus = st.form_submit_button("Tutup", disabled=True)
                    else:
                        st.write("Tidak ada pekerjaan di area ini.")
                        submit_hapus = st.form_submit_button("Tutup", disabled=True)
        else:
            st.info("Database masih kosong.")

    st.divider()
    
    # Tombol Refresh Manual
    if st.button("🔄 Muat Ulang Halaman (Refresh)"):
        st.rerun()

    st.divider()
    
    # --- FITUR EXPORT KE EXCEL ---
    st.header("📊 Export Laporan")
    if st.button("Siapkan File Excel"):
        data_untuk_excel = []
        for area, dict_pekerjaan in st.session_state.database_tasks.items():
            for pekerjaan, dict_printilan in dict_pekerjaan.items():
                total = len(dict_printilan)
                selesai = sum(dict_printilan.values()) if total > 0 else 0
                progres_persen = (selesai / total) if total > 0 else 0
                
                data_untuk_excel.append({"Area": area, "Pekerjaan Utama": pekerjaan, "Item Printilan": "-", "Status": "PROGRES KESELURUHAN", "Progres (%)": progres_persen})
                
                for printilan, status in dict_printilan.items():
                    data_untuk_excel.append({"Area": area, "Pekerjaan Utama": pekerjaan, "Item Printilan": printilan, "Status": "Selesai" if status else "Belum", "Progres (%)": 1.0 if status else 0.0})
        
        if data_untuk_excel:
            df = pd.DataFrame(data_untuk_excel)
            file_excel = os.path.join(DIR_SAAT_INI, "Laporan_Progres.xlsx")
            df.to_excel(file_excel, index=False)
            with open(file_excel, "rb") as f:
                st.download_button(label="📥 Unduh Laporan Excel", data=f, file_name=f"Laporan_RSBS_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# 3. HALAMAN UTAMA (TRACKER)
# ==========================================
st.title("🏗️ Aplikasi Tracker Progres Proyek")

if not st.session_state.database_tasks:
    st.info("👈 Data kosong. Tambah Area di panel kiri.")
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
                        nama_file = f"{area_terpilih}_{main_task}_{foto_file.name}".replace(" ", "_")
                        path_simpan = os.path.join(FOLDER_FOTO, nama_file)
                        with open(path_simpan, "wb") as f:
                            f.write(foto_file.getbuffer())
                        st.success("✅ Foto dokumentasi tersimpan!")