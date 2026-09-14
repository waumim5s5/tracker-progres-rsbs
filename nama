import streamlit as st
import json
import os
from datetime import datetime

st.set_page_config(page_title="Agenda Pekerja Proyek", page_icon="👷", layout="wide")

# ==========================================
# 1. PERSIAPAN DATABASE AGENDA
# ==========================================
DIR_SAAT_INI = os.path.dirname(os.path.abspath(__file__))
FILE_AGENDA = os.path.join(DIR_SAAT_INI, "database_agenda.json")

def simpan_agenda(data):
    with open(FILE_AGENDA, 'w') as f:
        json.dump(data, f, indent=4)

def muat_agenda():
    if os.path.exists(FILE_AGENDA):
        try:
            with open(FILE_AGENDA, 'r') as f:
                return json.load(f)
        except:
            return {}
    return {}

if 'db_agenda' not in st.session_state:
    st.session_state.db_agenda = muat_agenda()

# Struktur Data: { "Area A": { "Pekerjaan 1": {"tukang": "A", "laden": "B"} } }

# ==========================================
# 2. SIDEBAR - MANAJEMEN AREA & PEKERJAAN
# ==========================================
with st.sidebar:
    st.header("⚙️ Pengaturan Area & Pekerjaan")
    
    with st.expander("➕ Tambah Area Baru"):
        area_baru = st.text_input("Nama Area (contoh: Lantai 1)")
        if st.button("Tambah Area"):
            if area_baru and area_baru not in st.session_state.db_agenda:
                st.session_state.db_agenda[area_baru] = {}
                simpan_agenda(st.session_state.db_agenda)
                st.success("Area ditambahkan!")
                st.rerun()

    if st.session_state.db_agenda:
        with st.expander("➕ Tambah Pekerjaan"):
            area_pilih = st.selectbox("Pilih Area", list(st.session_state.db_agenda.keys()))
            pek_baru = st.text_input("Nama Pekerjaan")
            if st.button("Tambah Pekerjaan"):
                if pek_baru and pek_baru not in st.session_state.db_agenda[area_pilih]:
                    st.session_state.db_agenda[area_pilih][pek_baru] = {"tukang": "", "laden": ""}
                    simpan_agenda(st.session_state.db_agenda)
                    st.success("Pekerjaan ditambahkan!")
                    st.rerun()
                    
        with st.expander("🗑️ Hapus Area/Pekerjaan"):
            area_hapus = st.selectbox("Hapus dari Area", list(st.session_state.db_agenda.keys()), key="h_area")
            
            col1, col2 = st.columns(2)
            if col1.button("Hapus Area Ini"):
                del st.session_state.db_agenda[area_hapus]
                simpan_agenda(st.session_state.db_agenda)
                st.rerun()
                
            if st.session_state.db_agenda.get(area_hapus):
                pek_hapus = st.selectbox("Atau Hapus Pekerjaan", list(st.session_state.db_agenda[area_hapus].keys()), key="h_pek")
                if col2.button("Hapus Pekerjaan"):
                    del st.session_state.db_agenda[area_hapus][pek_hapus]
                    simpan_agenda(st.session_state.db_agenda)
                    st.rerun()
            
# ==========================================
# 3. HALAMAN UTAMA - INPUT PEKERJA
# ==========================================
st.title("👷‍♂️ Input Pekerja Harian")

if not st.session_state.db_agenda:
    st.info("👈 Silakan buat Area pekerjaan di menu sebelah kiri terlebih dahulu.")
else:
    for area, dict_pekerjaan in st.session_state.db_agenda.items():
        with st.expander(f"📍 {area}", expanded=True):
            if not dict_pekerjaan:
                st.write("*Belum ada pekerjaan di area ini. Tambahkan di menu kiri.*")
            else:
                for pekerjaan, pekerja in dict_pekerjaan.items():
                    st.markdown(f"**🏗️ {pekerjaan}**")
                    col_t, col_l = st.columns(2)
                    
                    # Key unik untuk setiap input
                    key_t = f"t_{area}_{pekerjaan}"
                    key_l = f"l_{area}_{pekerjaan}"
                    
                    # Callback saat input berubah
                    def update_pekerja(a=area, p=pekerjaan, kt=key_t, kl=key_l):
                        st.session_state.db_agenda[a][p]["tukang"] = st.session_state[kt]
                        st.session_state.db_agenda[a][p]["laden"] = st.session_state[kl]
                        simpan_agenda(st.session_state.db_agenda)

                    with col_t:
                        st.text_input("Nama Tukang:", value=pekerja["tukang"], key=key_t, on_change=update_pekerja)
                    with col_l:
                        st.text_input("Nama Laden:", value=pekerja["laden"], key=key_l, on_change=update_pekerja)
                    st.write("---")

# ==========================================
# 4. GENERATOR LAPORAN WA
# ==========================================
st.divider()
st.header("📱 Generate Laporan WA")

col_w1, col_w2, col_w3 = st.columns(3)
with col_w1:
    input_nama = st.text_input("Nama Pelapor", value="samsul")
with col_w2:
    # Set default date format DD MM YYYY as requested in prompt format
    input_tgl = st.text_input("Tanggal", value=datetime.now().strftime("%d %m %Y"))
with col_w3:
    input_jam = st.text_input("Jam Laporan", value="14.00")

input_notes = st.text_area("NOTE (opsional):", placeholder="1. kramik = pending
2. pembersihan kaca = pending")

def generate_wa_agenda(nama, tgl, jam, notes):
    lines = []
    lines.append(f"Assalamualikum pak {nama} izin melaporkan perkembngan proyek di RS Bina Sehat Tanggal {tgl}")
    lines.append("")
    
    for area, dict_pekerjaan in st.session_state.db_agenda.items():
        lines.append(f"Agenda tim {area.lower()} ({jam})")
        
        ada_pekerjaan_aktif = False
        
        for pekerjaan, pekerja in dict_pekerjaan.items():
            tukang = pekerja.get("tukang", "").strip()
            laden = pekerja.get("laden", "").strip()
            
            # Hanya memunculkan pekerjaan jika ada tukang ATAU laden yang diisi
            if tukang or laden:
                ada_pekerjaan_aktif = True
                lines.append(f"👷🏻‍♂️{pekerjaan.lower()}")
                if tukang:
                    lines.append(f"==> tukang = {tukang}")
                if laden:
                    lines.append(f"==> laden = {laden}")
        
        if not ada_pekerjaan_aktif:
            lines.append("👷🏻‍♂️sementara tidak ada pekerjaan")
            
        lines.append("") 
    
    if notes.strip():
        lines.append("NOTE :")
        for baris in notes.split("
"):
            if baris.strip():
                # Jika user belum ngasih nomor, kita bisa handle atau biarin sesuai input user
                # Untuk aman, biarkan sesuai input raw user agar fleksibel
                lines.append(baris.strip())
                
    return "
".join(lines)

if st.button("Generate Teks Laporan", type="primary"):
    teks_laporan_wa = generate_wa_agenda(input_nama, input_tgl, input_jam, input_notes)
    st.text_area("Silakan Copy Teks Berikut:", value=teks_laporan_wa, height=400)

