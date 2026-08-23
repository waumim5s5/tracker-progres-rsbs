from pyngrok import ngrok
import os

# Membuka port 8501 (port default Streamlit) ke internet publik
public_url = ngrok.connect(8501)
print(f"👉 APLIKASI BISA DIAKSES DARI HP MELALUI LINK INI: {public_url}")

# Menjalankan Streamlit secara otomatis
os.system("streamlit run app_progres_detail.py")