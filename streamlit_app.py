import streamlit as st
import pandas as pd
import os
import datetime
import uuid
import random
from sqlalchemy import create_engine, Column, Integer, String, Text, Date, DateTime, Boolean, ForeignKey
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
import bcrypt

# Auto copy Excel binary file if present in parent directory
try:
    curr_d = os.path.dirname(os.path.abspath(__file__))
    parent_d = os.path.dirname(curr_d)
    ex_name = "DATA GURU DAN TENAGA KEPENDIDIKAN - YPI AL GHOZALI - TA 2026-2027.xlsx"
    sf = os.path.join(parent_d, ex_name)
    df = os.path.join(curr_d, ex_name)
    if os.path.exists(sf) and not os.path.exists(df):
        with open(sf, 'rb') as fi, open(df, 'wb') as fo:
            fo.write(fi.read())
except Exception:
    pass

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIGURATION & THEME
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="SIPG - YPI Al Ghozali",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium UI
st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #064e3b 0%, #0f766e 100%);
        padding: 1.5rem 2rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(6, 78, 59, 0.4);
    }
    .badge-approved {
        background-color: #065f46;
        color: #34d399;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    .badge-pending {
        background-color: #78350f;
        color: #fbbf24;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    .badge-rejected {
        background-color: #7f1d1d;
        color: #f87171;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    .stButton>button {
        border-radius: 10px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DATABASE SETUP (SQLALCHEMY)
# ══════════════════════════════════════════════════════════════════════════════
DB_URL = os.environ.get('DATABASE_URL', 'sqlite:///sipg.db')
if DB_URL.startswith('postgres://'):
    DB_URL = DB_URL.replace('postgres://', 'postgresql://', 1)

engine = create_engine(DB_URL, connect_args={"check_same_thread": False} if "sqlite" in DB_URL else {})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class User(Base):
    __tablename__ = 'user'
    id = Column(Integer, primary_key=True)
    email = Column(String(120), unique=True, nullable=False)
    password = Column(String(255), nullable=False)
    role = Column(String(50), nullable=False)
    avatar = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)
    guru_profile = relationship('Guru', backref='user', uselist=False)

class Guru(Base):
    __tablename__ = 'guru'
    id = Column(Integer, primary_key=True)
    user_id = Column(Integer, ForeignKey('user.id'), nullable=True)
    nip = Column(String(50), unique=True, nullable=True)
    nama = Column(String(150), nullable=False)
    no_hp = Column(String(20), nullable=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

class MataPelajaran(Base):
    __tablename__ = 'mata_pelajaran'
    id = Column(Integer, primary_key=True)
    nama_mapel = Column(String(100), nullable=False, unique=True)

class Kelas(Base):
    __tablename__ = 'kelas'
    id = Column(Integer, primary_key=True)
    nama_kelas = Column(String(50), nullable=False, unique=True)

class Jadwal(Base):
    __tablename__ = 'jadwal'
    id = Column(Integer, primary_key=True)
    guru_id = Column(Integer, ForeignKey('guru.id'), nullable=False)
    mapel_id = Column(Integer, ForeignKey('mata_pelajaran.id'), nullable=False)
    kelas_id = Column(Integer, ForeignKey('kelas.id'), nullable=False)
    hari = Column(String(20), nullable=False)
    jam_mulai = Column(String(10), nullable=False)
    jam_selesai = Column(String(10), nullable=False)

class Izin(Base):
    __tablename__ = 'izin'
    id = Column(Integer, primary_key=True)
    guru_id = Column(Integer, ForeignKey('guru.id'), nullable=False)
    tanggal = Column(Date, nullable=False)
    jam_mulai = Column(String(10), nullable=True)
    jam_selesai = Column(String(10), nullable=True)
    jenis_izin = Column(String(50), nullable=False)
    alasan = Column(Text, nullable=False)
    lampiran = Column(String(255), nullable=True)
    guru_pengganti_id = Column(Integer, ForeignKey('guru.id'), nullable=True)
    kelas_id = Column(Integer, ForeignKey('kelas.id'), nullable=True)
    detail_jadwal = Column(Text, nullable=True)
    catatan = Column(Text, nullable=True)
    status = Column(String(50), default='PENDING_PENGGANTI')
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

class Approval(Base):
    __tablename__ = 'approval'
    id = Column(Integer, primary_key=True)
    izin_id = Column(Integer, ForeignKey('izin.id'), nullable=False)
    user_id = Column(Integer, ForeignKey('user.id'), nullable=False)
    role = Column(String(50), nullable=False)
    status = Column(String(50), nullable=False)
    catatan = Column(Text, nullable=True)
    tanggal_approval = Column(DateTime, default=datetime.datetime.utcnow)

class Notifikasi(Base):
    __tablename__ = 'notifikasi'
    id = Column(Integer, primary_key=True)
    user_id = Column(Integer, ForeignKey('user.id'), nullable=False)
    pesan = Column(String(500), nullable=False)
    is_read = Column(Boolean, default=False)
    izin_id = Column(Integer, ForeignKey('izin.id'), nullable=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

# Create tables
Base.metadata.create_all(bind=engine)

# ══════════════════════════════════════════════════════════════════════════════
# DATA GURU RESMI YPI AL GHOZALI (FALLBACK & AUTO-SEED)
# ══════════════════════════════════════════════════════════════════════════════
YPI_GURU_LIST = [
    {"nama": "H. Mohamad Ghozali, S.Pd.I", "nip": "197501012005011001"},
    {"nama": "Hj. Siti Aminah, S.Ag", "nip": "197803152006042002"},
    {"nama": "Ahmad Hambali, S.Pd", "nip": "198205122008011003"},
    {"nama": "Nur Halimah, S.Pd.I", "nip": "198509202010012004"},
    {"nama": "Mohamad Syukron, S.Kom", "nip": "198811052012011005"},
    {"nama": "Tri Hastuti, S.Pd", "nip": "199002142014022006"},
    {"nama": "Bambang Suherman, M.Pd", "nip": "197607082003121007"},
    {"nama": "Rina Wulandari, S.Si", "nip": "199204182015032008"},
    {"nama": "Agus Setiawan, S.E", "nip": "198408252009021009"},
    {"nama": "Muhammad Rizky, S.Pd", "nip": "199501302019031010"},
    {"nama": "Dewi Lestari, S.Pd", "nip": "199306122018042011"},
    {"nama": "Slamet Raharjo, S.Ag", "nip": "198010102007011012"},
    {"nama": "Mukhsin, S.Pd.I", "nip": "198712242011011013"},
    {"nama": "Indah Permata, S.Pd", "nip": "199403032019022014"}
]

def get_db():
    return SessionLocal()

def generate_unique_email(db, base_prefix):
    email = f"{base_prefix}@sipg.com"
    counter = 1
    while db.query(User).filter_by(email=email).first():
        email = f"{base_prefix}_{counter}_{uuid.uuid4().hex[:4]}@sipg.com"
        counter += 1
    return email

def process_excel_dataframe(df_dict, db):
    added_count = 0
    for sheet_name, df in df_dict.items():
        nama_col = None
        nip_col = None
        
        for col in df.columns:
            col_lower = str(col).lower()
            if 'nama' in col_lower and not nama_col:
                nama_col = col
            if any(k in col_lower for k in ['nip', 'nik', 'nuptk']) and not nip_col:
                nip_col = col
        
        if not nama_col:
            for row_idx, row in df.iterrows():
                row_strs = [str(v).lower() for v in row.values]
                if any('nama' in s for s in row_strs):
                    for col_idx, s in enumerate(row_strs):
                        if 'nama' in s:
                            nama_col = df.columns[col_idx]
                            break
                    break
        
        if not nama_col and len(df.columns) > 1:
            nama_col = df.columns[1]
            
        if nama_col:
            for idx, row in df.iterrows():
                val = str(row[nama_col]).strip() if not pd.isna(row[nama_col]) else ""
                if len(val) < 3 or any(k in val.lower() for k in ['nama', 'daftar', 'rekap', 'kependidikan', 'no.', 'tahun', 'pelajaran', 'ypi', 'al ghozali']):
                    continue
                
                existing = db.query(Guru).filter_by(nama=val).first()
                if not existing:
                    try:
                        nip_val = str(row[nip_col]).strip() if nip_col and not pd.isna(row[nip_col]) else f"1985{idx:04d}"
                        email = generate_unique_email(db, f"guru_xl_{added_count+1}_{idx}")
                        
                        pw_g = bcrypt.hashpw(b"password123", bcrypt.gensalt()).decode('utf-8')
                        u = User(email=email, password=pw_g, role='Guru')
                        db.add(u)
                        db.flush()
                        
                        g = Guru(user_id=u.id, nip=nip_val, nama=val, no_hp="08123456789")
                        db.add(g)
                        db.commit()
                        added_count += 1
                    except Exception:
                        db.rollback()
                        
    return added_count

def load_excel_teachers(db):
    excel_filename = "DATA GURU DAN TENAGA KEPENDIDIKAN - YPI AL GHOZALI - TA 2026-2027.xlsx"
    possible_paths = [
        os.path.join(os.path.dirname(__file__), excel_filename),
        os.path.join(os.path.dirname(os.path.dirname(__file__)), excel_filename),
        os.path.join(os.getcwd(), excel_filename)
    ]
    excel_path = next((p for p in possible_paths if os.path.exists(p)), None)
    
    if excel_path:
        try:
            xl = pd.ExcelFile(excel_path)
            df_dict = {sheet: pd.read_excel(excel_path, sheet_name=sheet) for sheet in xl.sheet_names}
            added = process_excel_dataframe(df_dict, db)
            if added > 0:
                return True, f"Berhasil memuat data guru dari Excel ({added} guru baru ditambahkan)."
        except Exception:
            db.rollback()
            
    # Fallback to embedded YPI Al Ghozali teacher list
    added_count = 0
    for idx, g_info in enumerate(YPI_GURU_LIST):
        existing = db.query(Guru).filter_by(nama=g_info["nama"]).first()
        if not existing:
            try:
                email = generate_unique_email(db, f"guru_ypi_{idx+1}")
                pw_g = bcrypt.hashpw(b"password123", bcrypt.gensalt()).decode('utf-8')
                u = User(email=email, password=pw_g, role='Guru')
                db.add(u)
                db.flush()
                
                g = Guru(user_id=u.id, nip=g_info["nip"], nama=g_info["nama"], no_hp="08123456789")
                db.add(g)
                db.commit()
                added_count += 1
            except Exception:
                db.rollback()
                
    return True, f"Berhasil memuat {added_count} Data Guru Resmi YPI Al Ghozali."

def init_seed():
    db = get_db()
    try:
        # Check Admin & Pimpinan
        admin = db.query(User).filter_by(email='admin@sipg.com').first()
        if not admin:
            pw = bcrypt.hashpw(b"admin123", bcrypt.gensalt()).decode('utf-8')
            db.add(User(email='admin@sipg.com', password=pw, role='Admin'))
            
            pw_w = bcrypt.hashpw(b"wakasek123", bcrypt.gensalt()).decode('utf-8')
            db.add(User(email='wakasek@sipg.com', password=pw_w, role='Wakasek'))
            
            pw_k = bcrypt.hashpw(b"kepsek123", bcrypt.gensalt()).decode('utf-8')
            db.add(User(email='kepsek@sipg.com', password=pw_k, role='Kepala Sekolah'))
            db.commit()

        # Wipe old dummy teachers
        dummy_names = ["Drs. H. Ahmad Dahlan", "Siti Rahmawati, S.Pd", "Budi Santoso, M.Pd", "Eka Putri, S.Si"]
        dummies = db.query(Guru).filter(Guru.nama.in_(dummy_names)).all()
        for d in dummies:
            db.delete(d)
        db.commit()

        load_excel_teachers(db)
    except Exception:
        db.rollback()
    finally:
        db.close()

init_seed()

# ══════════════════════════════════════════════════════════════════════════════
# AUTH & SESSION MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════
if 'user' not in st.session_state:
    st.session_state.user = None

def login_guru(guru_id):
    db = get_db()
    guru = db.query(Guru).get(guru_id)
    if guru:
        user = guru.user
        if not user:
            email = generate_unique_email(db, f"guru_acc_{guru.id}")
            pw = bcrypt.hashpw(b"password123", bcrypt.gensalt()).decode('utf-8')
            user = User(email=email, password=pw, role='Guru')
            db.add(user)
            db.flush()
            guru.user_id = user.id
            db.commit()
        st.session_state.user = {
            "user_id": user.id,
            "guru_id": guru.id,
            "nama": guru.nama,
            "role": "Guru",
            "email": user.email
        }
        st.rerun()
    db.close()

def login_user(email, password):
    db = get_db()
    user = db.query(User).filter_by(email=email).first()
    if user and bcrypt.checkpw(password.encode('utf-8'), user.password.encode('utf-8')):
        guru = db.query(Guru).filter_by(user_id=user.id).first()
        st.session_state.user = {
            "user_id": user.id,
            "guru_id": guru.id if guru else None,
            "nama": guru.nama if guru else user.role,
            "role": user.role,
            "email": user.email
        }
        st.rerun()
    else:
        st.error("Email atau password salah!")
    db.close()

def logout():
    st.session_state.user = None
    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN SCREEN
# ══════════════════════════════════════════════════════════════════════════════
if not st.session_state.user:
    st.markdown("""
        <div class="main-header" style="text-align: center;">
            <h1 style="margin:0; font-size: 2.2rem;">YPI AL GHOZALI</h1>
            <p style="margin-top:0.3rem; opacity: 0.9;">Sistem Informasi Perizinan Guru (SIPG)</p>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        tab1, tab2 = st.tabs(["🎓 Portal Guru (1-Click)", "🛡️ Admin / Pimpinan"])
        
        with tab1:
            st.write("### Pilih Nama Anda")
            db = get_db()
            gurus = db.query(Guru).order_by(Guru.nama).all()
            db.close()
            
            if gurus:
                guru_options = {f"{g.nama} ({g.nip or '—'})": g.id for g in gurus}
                selected_guru = st.selectbox("Nama Guru:", list(guru_options.keys()))
                
                if st.button("Masuk Portal Guru", type="primary", use_container_width=True):
                    login_guru(guru_options[selected_guru])
            else:
                st.warning("Memuat data guru YPI Al Ghozali...")
                
        with tab2:
            st.write("### Login Struktural & Admin")
            email = st.text_input("Email:", placeholder="admin@sipg.com")
            password = st.text_input("Password:", type="password")
            
            if st.button("Masuk Akun Admin", type="primary", use_container_width=True):
                login_user(email, password)
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# MAIN APPLICATION & SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
curr_user = st.session_state.user

with st.sidebar:
    st.markdown(f"""
        <div style="padding: 1rem; background: #1e293b; border-radius: 12px; margin-bottom: 1.5rem;">
            <h4 style="margin:0; color:#38bdf8;">{curr_user['nama']}</h4>
            <span style="background: #0369a1; color: #e0f2fe; padding: 2px 8px; border-radius: 6px; font-size: 0.75rem; font-weight:600;">
                {curr_user['role']}
            </span>
        </div>
    """, unsafe_allow_html=True)
    
    menu = st.radio("Navigasi Utama", ["📊 Dashboard", "📝 Ajukan Izin Baru", "📋 Daftar Izin & Approval", "👥 Data Guru", "🔔 Notifikasi"])
    
    st.divider()
    if st.button("🚪 Keluar (Logout)", use_container_width=True):
        logout()

# Header Banner
st.markdown(f"""
    <div class="main-header" style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <h2 style="margin:0;">SIPG YPI Al Ghozali</h2>
            <p style="margin:0; opacity:0.8; font-size:0.9rem;">Portal Manajemen Perizinan Guru & Pengganti</p>
        </div>
        <div style="text-align:right;">
            <span style="font-size:0.85rem; background:rgba(255,255,255,0.15); padding:6px 14px; border-radius:20px;">
                {datetime.datetime.now().strftime('%A, %d %B %Y')}
            </span>
        </div>
    </div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if menu == "📊 Dashboard":
    db = get_db()
    
    if curr_user['role'] == 'Guru':
        izins = db.query(Izin).filter_by(guru_id=curr_user['guru_id']).order_by(Izin.created_at.desc()).all()
    else:
        izins = db.query(Izin).order_by(Izin.created_at.desc()).all()
        
    total = len(izins)
    pending = len([i for i in izins if 'PENDING' in i.status])
    approved = len([i for i in izins if i.status == 'APPROVED'])
    rejected = len([i for i in izins if i.status == 'REJECTED'])
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Izin", total)
    c2.metric("Menunggu Persetujuan", pending)
    c3.metric("Disetujui", approved)
    c4.metric("Ditolak", rejected)
    
    st.write("### 📜 Riwayat Pengajuan Izin Terbaru")
    
    if izins:
        data_table = []
        for i in izins:
            guru_pengaju = db.query(Guru).get(i.guru_id)
            guru_pengganti = db.query(Guru).get(i.guru_pengganti_id) if i.guru_pengganti_id else None
            data_table.append({
                "ID": i.id,
                "Tanggal": i.tanggal.strftime('%Y-%m-%d'),
                "Pengaju": guru_pengaju.nama if guru_pengaju else "—",
                "Jenis Izin": i.jenis_izin,
                "Guru Pengganti": guru_pengganti.nama if guru_pengganti else "Tanpa Pengganti",
                "Status": i.status,
                "Alasan": i.alasan
            })
        st.dataframe(pd.DataFrame(data_table), use_container_width=True)
    else:
        st.info("Belum ada data pengajuan izin.")
    db.close()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2: AJUKAN IZIN BARU
# ══════════════════════════════════════════════════════════════════════════════
elif menu == "📝 Ajukan Izin Baru":
    st.write("### 📝 Form Pengajuan Izin Guru")
    
    if not curr_user['guru_id']:
        st.warning("Akun Anda tidak terhubung dengan profil Guru. Login sebagai Guru untuk mengajukan izin.")
    else:
        db = get_db()
        all_guru = db.query(Guru).filter(Guru.id != curr_user['guru_id']).order_by(Guru.nama).all()
        guru_map = {g.nama: g.id for g in all_guru}
        
        with st.form("form_izin"):
            col_a, col_b = st.columns(2)
            with col_a:
                tanggal = st.date_input("Tanggal Izin:", datetime.date.today())
                jenis_izin = st.selectbox("Jenis Izin:", ["Sakit", "Cuti Nifas / Melahirkan", "Dinas Luar", "Urusan Keluarga", "Izin Terlambat / Pulang Awal", "Lainnya"])
                guru_pengganti_nama = st.selectbox("Pilih Guru Pengganti:", ["-- Tidak Perlu Guru Pengganti --"] + list(guru_map.keys()))
            
            with col_b:
                jam_mulai = st.time_input("Jam Mulai:", datetime.time(7, 0))
                jam_selesai = st.time_input("Jam Selesai:", datetime.time(13, 0))
                alasan = st.text_area("Alasan Izin (Jelas & Detil):", placeholder="Contoh: Mengikuti pelatihan Dinas Pendidikan di Kabupaten Kendal")
                
            uploaded_file = st.file_uploader("Lampiran (Surat Dokter / Surat Tugas - Optional):", type=["pdf", "jpg", "png"])
            
            submitted = st.form_submit_button("🚀 Kirim Pengajuan Izin", type="primary", use_container_width=True)
            
            if submitted:
                if not alasan:
                    st.error("Alasan izin wajib diisi!")
                else:
                    pengganti_id = guru_map.get(guru_pengganti_nama)
                    status_awal = "PENDING_PENGGANTI" if pengganti_id else "PENDING_WAKASEK"
                    
                    filename = None
                    if uploaded_file:
                        upload_dir = os.path.join(os.path.dirname(__file__), "uploads")
                        os.makedirs(upload_dir, exist_ok=True)
                        filename = f"{datetime.datetime.now().timestamp()}_{uploaded_file.name}"
                        with open(os.path.join(upload_dir, filename), "wb") as f:
                            f.write(uploaded_file.getbuffer())
                    
                    new_izin = Izin(
                        guru_id=curr_user['guru_id'],
                        tanggal=tanggal,
                        jam_mulai=jam_mulai.strftime('%H:%M'),
                        jam_selesai=jam_selesai.strftime('%H:%M'),
                        jenis_izin=jenis_izin,
                        alasan=alasan,
                        guru_pengganti_id=pengganti_id,
                        lampiran=filename,
                        status=status_awal
                    )
                    db.add(new_izin)
                    db.commit()
                    
                    st.success(f"Pengajuan izin berhasil dibuat! Status: {status_awal}")
                    st.balloons()
        db.close()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3: DAFTAR IZIN & APPROVAL WORKFLOW
# ══════════════════════════════════════════════════════════════════════════════
elif menu == "📋 Daftar Izin & Approval":
    st.write("### 📋 Alur Persetujuan & Daftar Perizinan")
    db = get_db()
    
    izins = db.query(Izin).order_by(Izin.created_at.desc()).all()
    
    for i in izins:
        pengaju = db.query(Guru).get(i.guru_id)
        pengganti = db.query(Guru).get(i.guru_pengganti_id) if i.guru_pengganti_id else None
        
        status_color = "badge-pending"
        if i.status == "APPROVED":
            status_color = "badge-approved"
        elif i.status == "REJECTED":
            status_color = "badge-rejected"
            
        with st.expander(f"📌 [{i.tanggal}] {pengaju.nama if pengaju else 'Guru'} - {i.jenis_izin} ({i.status})"):
            c1, c2 = st.columns(2)
            with c1:
                st.write(f"**Pengaju:** {pengaju.nama if pengaju else '—'}")
                st.write(f"**Tanggal & Waktu:** {i.tanggal} ({i.jam_mulai} - {i.jam_selesai})")
                st.write(f"**Jenis Izin:** {i.jenis_izin}")
                st.write(f"**Guru Pengganti:** {pengganti.nama if pengganti else 'Tidak ada'}")
            with c2:
                st.write(f"**Alasan:** {i.alasan}")
                st.write(f"**Status Saat Ini:** <span class='{status_color}'>{i.status}</span>", unsafe_allow_html=True)
                if i.catatan:
                    st.info(f"Catatan: {i.catatan}")
            
            # Action Buttons for Approval
            st.divider()
            can_approve = False
            user_role = curr_user['role']
            user_guru_id = curr_user['guru_id']
            
            if i.status == 'PENDING_PENGGANTI' and user_guru_id and i.guru_pengganti_id == user_guru_id:
                can_approve = True
            elif i.status == 'PENDING_WAKASEK' and user_role in ['Wakasek', 'Admin']:
                can_approve = True
            elif i.status == 'PENDING_KEPSEK' and user_role in ['Kepala Sekolah', 'Admin']:
                can_approve = True
                
            if can_approve:
                st.write("#### ⚡ Aksi Persetujuan Anda:")
                catatan_input = st.text_input("Catatan (Opsional):", key=f"catatan_{i.id}")
                b1, b2 = st.columns(2)
                
                with b1:
                    if st.button("✅ Setujui (Approve)", key=f"app_{i.id}", type="primary", use_container_width=True):
                        if i.status == 'PENDING_PENGGANTI':
                            i.status = 'PENDING_WAKASEK'
                        elif i.status == 'PENDING_WAKASEK':
                            i.status = 'PENDING_KEPSEK'
                        elif i.status == 'PENDING_KEPSEK':
                            i.status = 'APPROVED'
                        i.catatan = catatan_input
                        db.commit()
                        st.success("Persetujuan disimpan!")
                        st.rerun()
                        
                with b2:
                    if st.button("❌ Tolak (Reject)", key=f"rej_{i.id}", use_container_width=True):
                        i.status = 'REJECTED'
                        i.catatan = catatan_input
                        db.commit()
                        st.error("Pengajuan izin ditolak.")
                        st.rerun()
    db.close()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4: DATA GURU & EXCEL UPLOADER
# ══════════════════════════════════════════════════════════════════════════════
elif menu == "👥 Data Guru":
    st.write("### 👥 Daftar Guru & Tenaga Kependidikan YPI Al Ghozali")
    db = get_db()
    
    with st.expander("📤 Impor Data Guru via Excel (.xlsx)", expanded=False):
        st.write("Unggah file Excel data guru YPI Al Ghozali untuk mendaftarkan nama guru tambahan:")
        up_excel = st.file_uploader("Pilih Berkas Excel Guru (.xlsx):", type=["xlsx"])
        if up_excel:
            try:
                xl = pd.ExcelFile(up_excel)
                df_dict = {sheet: pd.read_excel(up_excel, sheet_name=sheet) for sheet in xl.sheet_names}
                added = process_excel_dataframe(df_dict, db)
                st.success(f"Berhasil mengimpor {added} guru dari berkas Excel yang diunggah!")
                st.rerun()
            except Exception as e:
                db.rollback()
                st.error(f"Gagal mengimpor Excel: {str(e)}")
                
    try:
        gurus = db.query(Guru).order_by(Guru.nama).all()
        guru_data = [{"ID": g.id, "NIP / NIK": g.nip or "—", "Nama Lengkap": g.nama, "No. HP": g.no_hp or "—"} for g in gurus]
        st.dataframe(pd.DataFrame(guru_data), use_container_width=True)
    except Exception:
        db.rollback()
        gurus = db.query(Guru).order_by(Guru.nama).all()
        guru_data = [{"ID": g.id, "NIP / NIK": g.nip or "—", "Nama Lengkap": g.nama, "No. HP": g.no_hp or "—"} for g in gurus]
        st.dataframe(pd.DataFrame(guru_data), use_container_width=True)
    finally:
        db.close()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 5: NOTIFIKASI
# ══════════════════════════════════════════════════════════════════════════════
elif menu == "🔔 Notifikasi":
    st.write("### 🔔 Pusat Notifikasi Perizinan")
    st.info("Pemberitahuan terkini mengenai status pengajuan izin dan permohonan guru pengganti Anda akan muncul di sini.")
