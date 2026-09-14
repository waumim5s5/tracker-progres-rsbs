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


# ============================================================
# 3. PENUGASAN PEKERJA (AGENDA TIM)
# ============================================================
st.title("👷‍♂️ Manajemen Agenda Tim")
st.write("Atur penugasan Tukang dan Peladen untuk masing-masing pekerjaan hari ini.")

if st.session_state.database_tasks:
    with st.container(border=True):
        col1, col2 = st.columns(2)
        with col1:
            area_pekerja = st.selectbox("Pilih Area:", list(st.session_state.database_tasks.keys()), key="worker_area")
        
        if st.session_state.database_tasks[area_pekerja]:
            with col2:
                pekerjaan_pekerja = st.selectbox("Pilih Pekerjaan:", list(st.session_state.database_tasks[area_pekerja].keys()), key="worker_pek")
            
            kunci_tukang = f"{area_pekerja}_{pekerjaan_pekerja}_tukang"
            kunci_peladen = f"{area_pekerja}_{pekerjaan_pekerja}_peladen"
            
            val_tukang = st.session_state.database_workers.get(kunci_tukang, "")
            val_peladen = st.session_state.database_workers.get(kunci_peladen, "")
            
            with st.form("form_assign_workers"):
                col3, col4 = st.columns(2)
                with col3:
                    input_tukang = st.text_input("Nama Tukang (Opsional):", value=val_tukang, placeholder="contoh: pak pri")
                with col4:
                    input_peladen = st.text_input("Nama Laden (Opsional):", value=val_peladen, placeholder="contoh: pak ompong, & pak dar")
                
                if st.form_submit_button("Simpan Data Pekerja", use_container_width=True):
                    st.session_state.database_workers[kunci_tukang] = input_tukang.strip()
                    st.session_state.database_workers[kunci_peladen] = input_peladen.strip()
                    simpan_database()
                    st.success(f"Pekerja untuk '{pekerjaan_pekerja}' berhasil diperbarui!")
        else:
            st.warning("Belum ada daftar pekerjaan di area ini. Buat terlebih dahulu di panel kiri.")
else:
    st.info("Belum ada data area.")


# ============================================================
# 4. HALAMAN UTAMA (TRACKER PROGRES & FOTO)
# ============================================================
st.divider()
st.title("🏗️ Tracker Progres Pekerjaan")

if not st.session_state.database_tasks:
    st.info("👈 Data kosong. Tambah data di panel kiri.")
else:
    daftar_area = list(st.session_state.database_tasks.keys())
    if "ingat_area" not in st.session_state or st.session_state.ingat_area not in daftar_area:
        st.session_state.ingat_area = daftar_area[0]
        
    area_terpilih = st.selectbox(
        "📍 Tampilkan Progres Area:", 
        daftar_area,
        key="ingat_area"
    )

    pekerjaan_utama_dict = st.session_state.database_tasks[area_terpilih]

    if not pekerjaan_utama_dict:
        st.warning(f"Belum ada pekerjaan di area {area_terpilih}.")
    else:
        for main_task, sub_tasks in pekerjaan_utama_dict.items():
            with st.expander(f"🛠️ {main_task}", expanded=True):
                if not sub_tasks:
                    st.write("*Belum ada checklist printilan untuk pekerjaan ini.*")
                else:
                    total_printilan = len(sub_tasks)
                    total_skor = 0
                    
                    for p_val in sub_tasks.values():
                        if isinstance(p_val, bool): p_val = 100 if p_val else 0
                        total_skor += p_val
                        
                    persentase = int(total_skor / total_printilan) if total_printilan > 0 else 0
                    
                    st.progress(persentase / 100)
                    st.markdown(f"**Progres Pekerjaan: {persentase}%**")
                    st.write("---")
                    
                    for printilan, val in sub_tasks.items():
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
# 5. GENERATOR LAPORAN WHATSAPP (AGENDA)
# ==========================================
st.divider()
st.title("📱 Buat Laporan WhatsApp (Agenda Tim)")

col_wa1, col_wa2, col_wa3 = st.columns(3)
with col_wa1:
    input_nama = st.text_input("Nama Pelapor", value="samsul")
with col_wa2:
    input_tgl = st.text_input("Tanggal", value=datetime.now().strftime("%d %m %Y"))
with col_wa3:
    input_jam = st.text_input("Jam Laporan", value="14.00")

input_notes = st.text_area("NOTE (Catatan / Kendala / Pending):", placeholder="Ketik setiap catatan di baris baru\nContoh:\nkramik = pending\npembersihan kaca = pending")

def generate_wa_agenda(nama, tgl, jam, notes):
    lines = []
    lines.append(f"Assalamualikum pak {nama} izin melaporkan perkembngan proyek di RS Bina Sehat Tanggal {tgl}")
    lines.append("")
    
    for area, dict_pekerjaan in st.session_state.database_tasks.items():
        # Judul Area Harian
        lines.append(f"Agenda tim {area.lower()} ({jam})")
        
        pekerjaan_ditemukan = False
        
        for pekerjaan in dict_pekerjaan.keys():
            tukang = st.session_state.database_workers.get(f"{area}_{pekerjaan}_tukang", "").strip()
            peladen = st.session_state.database_workers.get(f"{area}_{pekerjaan}_peladen", "").strip()
            
            # Hanya memunculkan pekerjaan yang ada pekerjanya
            if tukang or peladen:
                pekerjaan_ditemukan = True
                lines.append(f"👷🏻‍♂️{pekerjaan.lower()}")
                if tukang:
                    lines.append(f"==> tukang = {tukang}")
                if peladen:
                    lines.append(f"==> laden = {peladen}")
        
        if not pekerjaan_ditemukan:
            lines.append("👷🏻‍♂️sementara tidak ada pekerjaan")
            
        lines.append("") 
    
    if notes.strip():
        lines.append("NOTE :")
        counter = 1
        for baris in notes.split("\n"):
            if baris.strip():
                lines.append(f"{counter}. {baris.strip()}")
                counter += 1
                
    return "\n".join(lines)

if st.session_state.database_tasks:
    teks_laporan_wa = generate_wa_agenda(input_nama, input_tgl, input_jam, input_notes)
    st.text_area("Kolom Copy-Paste Laporan WA", value=teks_laporan_wa, height=400)
