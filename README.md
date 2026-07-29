# 🎓 SIPG - Sistem Informasi Perizinan Guru (YPI Al Ghozali)

Sistem Informasi Perizinan Guru YPI Al Ghozali dibangun menggunakan 100% murni Python (Streamlit), siap di-deploy secara instan ke **Streamlit Community Cloud**, Railway, atau Render.

## 📁 Struktur Berkas Repo
```text
.
├── .streamlit/
│   └── config.toml          # Tema tampilan (Dark Mode Emerald Green)
├── DATA GURU DAN...xlsx      # File data guru awal untuk auto-seeding
├── Procfile                 # Konfigurasi deployment server
├── README.md                # Dokumentasi & panduan
├── requirements.txt         # Dependency modul Python
├── streamlit_app.py         # Kode utama aplikasi Streamlit
└── .gitignore               # Berkas yang diabaikan Git
```

## 🚀 Panduan Hosting Gratis di Streamlit Cloud (3 Menit)

1. **Upload ke GitHub**:
   - Buat repository baru di [GitHub](https://github.com/new) (misal: `sipg-alghozali`).
   - Push seluruh isi folder ini ke repository tersebut.

2. **Deploy di Streamlit Community Cloud**:
   - Buka [share.streamlit.io](https://share.streamlit.io/) dan login dengan GitHub.
   - Klik **"New app"**.
   - Masukkan repositori Anda (`username/sipg-alghozali`).
   - Masukkan Main file path: `streamlit_app.py`.
   - Klik **"Deploy!"** 🎉

---

## 💻 Cara Menjalankan di Komputer Lokal

1. Buka terminal di folder ini.
2. Install dependency:
   ```bash
   pip install -r requirements.txt
   ```
3. Jalankan aplikasi:
   ```bash
   streamlit run streamlit_app.py
   ```
