from io import BytesIO
from docx import Document
from google import genai
from google.genai import types
from PIL import Image
pypdf_available = True
try:
  from pypdf import PdfReader
except ImportError:
  pypdf_available = False

import pandas as pd
import streamlit as st

# Konfigurasi Halaman & Responsif Mobile/Android
st.set_page_config(
    page_title="Smart Exam Grader Pro - Cyber Edition",
    page_icon="🤖",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Custom Styling CSS: Bumblebee Yellow + Hacker Matrix Glow + Transformer Mechanical Motion
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Share+Tech+Mono&family=Orbitron:wght@500;700;900&display=swap');

    /* Global Background Kuning Bumblebee dengan Tekstur Cyber */
    .stApp {
        background-color: #ffd700;
        background-image: radial-gradient(#cca300 1px, transparent 1px);
        background-size: 20px 20px;
        color: #111111;
        font-family: 'Share Tech Mono', monospace, sans-serif;
    }
    
    /* Header Instansi Gaya Transformer Cybertron */
    .instansi-header {
        background: linear-gradient(135deg, #000000 0%, #1a1a1a 100%);
        color: #ffd700;
        padding: 22px;
        border-radius: 6px;
        margin-bottom: 20px;
        border-left: 8px solid #ffcc00;
        border-right: 2px solid #ffcc00;
        box-shadow: 0 8px 25px rgba(0,0,0,0.3);
        transform: skewX(-1deg);
        animation: cyberGlitch 4s infinite alternate;
    }

    @keyframes cyberGlitch {
        0% { box-shadow: 0 8px 25px rgba(0,0,0,0.3); border-left-color: #ffd700; }
        50% { box-shadow: 0 8px 30px rgba(0, 255, 65, 0.4); border-left-color: #00ff41; }
        100% { box-shadow: 0 8px 25px rgba(0,0,0,0.3); border-left-color: #ffd700; }
    }

    .instansi-title {
        font-family: 'Orbitron', sans-serif;
        font-size: 1.35rem;
        font-weight: 900;
        color: #ffd700 !important;
        margin: 0;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        text-shadow: 2px 2px #000000;
    }

    /* Efek Teks Hacker Matrix Hijau Menyala */
    .hacker-text {
        font-family: 'Share Tech Mono', monospace;
        color: #00ff41 !important;
        background-color: #050505;
        padding: 4px 8px;
        border-radius: 4px;
        border: 1px dashed #00ff41;
        text-shadow: 0 0 8px rgba(0, 255, 65, 0.8);
        display: inline-block;
        font-size: 0.9rem;
    }

    .instansi-subtitle {
        font-size: 0.85rem;
        color: #f8fafc !important;
        margin-top: 8px;
        font-weight: 400;
    }

    /* Kotak Kartu Konten Transformer Armor */
    .card-box {
        background: #ffffff;
        padding: 20px;
        border-radius: 8px;
        border: 3px solid #111111;
        box-shadow: 5px 5px 0px #111111;
        margin-bottom: 20px;
        transition: transform 0.2s ease;
    }
    .card-box:hover {
        transform: translate(-2px, -2px);
        box-shadow: 7px 7px 0px #00ff41;
    }

    /* Label Input */
    label, .stTextInput label, .stTextArea label, .stSelectbox label, .stRadio label {
        color: #111111 !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        font-family: 'Orbitron', sans-serif;
    }

    /* Tombol Transformer Bumblebee (Hitam Pekat dengan border kuning/hijau hacker) */
    .stButton>button {
        width: 100%;
        background-color: #111111;
        color: #ffd700;
        font-family: 'Orbitron', sans-serif;
        font-weight: 700;
        border-radius: 4px;
        padding: 0.75rem 1rem;
        border: 2px solid #ffd700;
        font-size: 1rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        box-shadow: 3px 3px 0px #00ff41;
        transition: all 0.15s ease;
    }
    .stButton>button:hover {
        background-color: #00ff41;
        color: #000000;
        border-color: #111111;
        box-shadow: 3px 3px 0px #111111;
        transform: translate(-2px, -2px);
    }

    /* Footer Cyber */
    .footer {
        text-align: center;
        padding: 15px;
        font-size: 0.8rem;
        color: #ffd700;
        background-color: #111111;
        border: 2px solid #ffd700;
        border-radius: 6px;
        margin-top: 30px;
        margin-bottom: 20px;
        font-weight: 600;
        box-shadow: 4px 4px 0px #00ff41;
    }
    .footer a {
        color: #00ff41;
        text-decoration: none;
        font-weight: 700;
    }
    </style>
""",
    unsafe_allow_html=True,
)


def buat_file_docx(teks_konten):
  doc = Document()
  for line in teks_konten.split("\n"):
    if line.startswith("# "):
      doc.add_heading(line.replace("# ", ""), level=1)
    elif line.startswith("## "):
      doc.add_heading(line.replace("## ", ""), level=2)
    elif line.startswith("### "):
      doc.add_heading(line.replace("### ", ""), level=3)
    else:
      doc.add_paragraph(line)
  buffer = BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# AMBIL API KEY DARI SECRETS ATAU INPUT MANUAL
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")

# --- HEADER UTAMA CYBERNETIC ---
st.markdown(
    """
    <div class="instansi-header">
        <h1 class="instansi-title">🐝 BUMBLEBEE CYBER-GRADER PRO</h1>
        <p class="instansi-subtitle">Sistem Penilaian AI Autonomous & Administrasi Guru Berkecepatan Tinggi.</p>
        <br>
        <span class="hacker-text">SYSTEM STATUS: ONLINE // AI CORE: GEMINI-2.5-FLASH // SECURE PROTOCOL ACTIVE</span>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander(
    "🔑 Cyber Security & API Key Configuration", expanded=not GEMINI_API_KEY
):
  if not GEMINI_API_KEY:
    st.warning("⚠️ API Key belum terdeteksi di Secrets.")
    GEMINI_API_KEY = st.text_input(
        "Masukkan Gemini API Key", type="password", key="sidebar_key"
    )
  else:
    st.success("🔒 Enkripsi Saraf Aktif & Terhubung")

if not GEMINI_API_KEY:
  st.warning(
      "Mohon masukkan Google Gemini API Key di atas untuk menginisialisasi"
      " sistem."
  )
  st.stop()

if "kunci_pg" not in st.session_state:
  st.session_state.kunci_pg = ""
if "kunci_isian" not in st.session_state:
  st.session_state.kunci_isian = ""
if "kunci_essai" not in st.session_state:
  st.session_state.kunci_essai = ""
if "rekap_nilai" not in st.session_state:
  st.session_state.rekap_nilai = []
if "modul_hasil" not in st.session_state:
  st.session_state.modul_hasil = ""
if "soal_hasil" not in st.session_state:
  st.session_state.soal_hasil = ""

# Menu Navigasi Dropdown Cyber
st.markdown("### ⚡ SELECT OPERATION MODULE:")
menu_pilihan = st.selectbox(
    "Navigasi Utama",
    [
        "📖 1. Generator Modul Ajar",
        "📝 2. Generator Soal & Kunci",
        "⚙️ 3. Set Kunci Acuan",
        "🔍 4. Koreksi Siswa",
        "📊 5. Rekap Nilai",
    ],
    label_visibility="collapsed",
)

st.markdown("---")

# --- KONTROL TAMPILAN BERDASARKAN MENU PILIHAN ---

# 1. MODUL AJAR
if menu_pilihan == "📖 1. Generator Modul Ajar":
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📖 Generator Modul Ajar / RPP")
  st.markdown("Sintesis perangkat pembelajaran komprehensif berbasis AI.")

  nama_guru = st.text_input("Nama Guru & Gelar", "Ahmad Fauzi, S.Pd.")
  nama_sekolah = st.text_input("Nama Instansi / Sekolah", "SMP Negeri 1 Nusantara")
  mapel = st.text_input("Mata Pelajaran", "Ilmu Pengetahuan Alam (IPA)")
  kurikulum_aktif = st.text_input("Kurikulum", "Kurikulum Merdeka")
  fase_kelas = st.text_input("Fase / Kelas", "Fase D / Kelas VII")
  nama_ks = st.text_input("Nama Kepala Sekolah", "Dra. Hj. Siti Aminah, M.Pd.")
  topik = st.text_input("Topik / Materi Pokok", "Sistem Pencernaan Manusia")
  alokasi_waktu = st.text_input("Alokasi Waktu", "2 Pertemuan (4 x 40 Menit)")

  if st.button("⚡ EXECUTE: BUILD MODUL AJAR", type="primary"):
    with st.spinner("AI Transformer sedang merakit Modul Ajar..."):
      try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt_modul = f"""
                Buatkan Modul Ajar / RPP formal, sangat terstruktur, dan profesional untuk instansi pendidikan:
                - Guru: {nama_guru}, Sekolah: {nama_sekolah}
                - Mapel: {mapel}, Kurikulum: {kurikulum_aktif}
                - Kelas: {fase_kelas}, Kepala Sekolah: {nama_ks}
                - Topik: {topik}, Waktu: {alokasi_waktu}
                Sertakan komponen Identitas Instansi, Profil Pelajar Pancasila, Tujuan Pembelajaran, Kegiatan Pembelajaran, dan Tabel Rubrik Penilaian.
                """
        response = client.models.generate_content(
            model="gemini-2.5-flash", contents=prompt_modul
        )
        st.session_state.modul_hasil = response.text
        st.success("Modul Ajar berhasil disintesis!")
      except Exception as e:
        st.error(f"Error sistem: {e}")

  if st.session_state.modul_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Dokumen")
    st.markdown(st.session_state.modul_hasil)
    file_docx = buat_file_docx(st.session_state.modul_hasil)
    st.download_button(
        "📥 Unduh Modul Ajar (.docx)",
        data=file_docx,
        file_name=f"Modul_Ajar_{mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# 2. SOAL & KUNCI
elif menu_pilihan == "📝 2. Generator Soal & Kunci":
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📝 Generator Paket Soal")
  st.markdown("Pembuatan naskah ujian otomatis berkecepatan tinggi.")

  s_guru = st.text_input("Nama Pembuat Soal", "Ahmad Fauzi, S.Pd.", key="sg")
  s_sekolah = st.text_input("Instansi / Sekolah", "SMP Negeri 1 Nusantara", key="ss")
  s_mapel = st.text_input("Mata Pelajaran", "Matematika", key="sm")
  s_kur = st.text_input("Kurikulum", "Kurikulum Merdeka", key="sk")
  s_kelas = st.text_input("Kelas / Semester", "Kelas VII / Ganjil", key="skel")
  s_materi = st.text_input(
      "Materi / Bab Ujian", "Persamaan Linear Satu Variabel", key="smat"
  )
  s_komposisi = st.text_input(
      "Komposisi Soal",
      "5 Pilihan Ganda, 2 Isian Singkat, 1 Essai",
      key="skom",
  )

  if st.button("⚡ EXECUTE: GENERATE SOAL & KUNCI", type="primary"):
    with st.spinner("AI Transformer sedang merakit naskah ujian..."):
      try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt_soal = f"""
                Buatkan naskah soal ujian resmi instansi pendidikan lengkap dengan Kop Ujian, Petunjuk, Naskah Soal, Kunci Jawaban, & Rubrik Penilaian (Tabel) untuk:
                - Guru: {s_guru}, Sekolah: {s_sekolah}, Mapel: {s_mapel}
                - Kurikulum: {s_kur}, Kelas: {s_kelas}, Materi: {s_materi}
                - Komposisi: {s_komposisi}
                """
        response = client.models.generate_content(
            model="gemini-2.5-flash", contents=prompt_soal
        )
        st.session_state.soal_hasil = response.text
        st.success("Paket soal berhasil disintesis!")
      except Exception as e:
        st.error(f"Error sistem: {e}")

  if st.session_state.soal_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Naskah Soal")
    st.markdown(st.session_state.soal_hasil)
    file_docx_soal = buat_file_docx(st.session_state.soal_hasil)
    st.download_button(
        "📥 Unduh Paket Soal (.docx)",
        data=file_docx_soal,
        file_name=f"Soal_{s_mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# 3. SET KUNCI ACUAN
elif menu_pilihan == "⚙️ 3. Set Kunci Acuan":
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("⚙️ Konfigurasi Kunci Acuan")
  st.markdown("Injeksi parameter kunci untuk validasi pemindaian AI.")

  kunci_pg_input = st.text_area(
      "Kunci Pilihan Ganda (Contoh: 1.A, 2.B, 3.C)",
      value=st.session_state.kunci_pg,
      height=100,
  )
  kunci_isian_input = st.text_area(
      "Kunci Isian Singkat & Keyword",
      value=st.session_state.kunci_isian,
      height=100,
  )
  kunci_essai_input = st.text_area(
      "Rubrik Penilaian & Kunci Essai",
      value=st.session_state.kunci_essai,
      height=150,
  )

  if st.button("⚡ EXECUTE: SAVE MASTER KEYS", type="primary"):
    st.session_state.kunci_pg = kunci_pg_input
    st.session_state.kunci_isian = kunci_isian_input
    st.session_state.kunci_essai = kunci_essai_input
    st.success("Master keys berhasil di-commit ke neural core!")
  st.markdown("</div>", unsafe_allow_html=True)

# 4. KOREKSI SISWA
elif menu_pilihan == "🔍 4. Koreksi Siswa":
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("🔍 Sistem Koreksi Otomatis AI")
  if (
      not st.session_state.kunci_pg
      and not st.session_state.kunci_isian
      and not st.session_state.kunci_essai
  ):
    st.warning(
        "⚠️ Master keys belum dikonfigurasi! Harap isi di menu nomor 3 terlebih"
        " dahulu."
    )
  else:
    nama_siswa = st.text_input("Nama Lengkap / Nomor Induk Siswa")
    metode_siswa = st.radio(
        "Format Berkas Jawaban:",
        ["Unggah Foto / Scan", "Unggah Dokumen (PDF/Word)", "Ketik Teks"],
    )

    file_img, file_doc, teks_manual = None, None, ""
    if "Foto" in metode_siswa:
      file_img = st.file_uploader(
          "Unggah berkas foto/scan", type=["jpg", "png", "jpeg"]
      )
    elif "Dokumen" in metode_siswa:
      file_doc = st.file_uploader("Unggah berkas dokumen", type=["pdf", "docx"])
    else:
      teks_manual = st.text_area("Masukkan teks jawaban siswa")

    if st.button("⚡ EXECUTE: RUN AI VISION SCANNER", type="primary"):
      if not nama_siswa:
        st.error("Masukkan nama atau identitas siswa!")
      else:
        with st.spinner("AI Scanner sedang mendekripsi dan memindai jawaban..."):
          try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            payload = [f"""
                        Koreksi lembar jawaban siswa berdasarkan acuan berikut:
                        - Kunci PG: {st.session_state.kunci_pg}
                        - Kunci Isian: {st.session_state.kunci_isian}
                        - Kunci Essai: {st.session_state.kunci_essai}
                        
                        Format baris pertama WAJIB persis seperti ini:
                        NILAI_AKHIR: [Angka total nilai 0-100]
                        """]
            if file_img:
              payload.append(Image.open(file_img))
            elif file_doc:
              ext = file_doc.name.split(".")[-1].lower()
              ext_text = ""
              if ext == "pdf" and pypdf_available:
                reader = PdfReader(file_doc)
                for page in reader.pages:
                  ext_text += page.extract_text() or ""
              payload.append(f"Berkas Jawaban: {ext_text}")
            elif teks_manual:
              payload.append(f"Jawaban Siswa: {teks_manual}")

            resp = client.models.generate_content(
                model="gemini-2.5-flash", contents=payload
            )
            hasil = resp.text

            import re

            m = re.search(r"NILAI_AKHIR:\s*([0-9.]+)", hasil)
            skor = float(m.group(1)) if m else 0.0

            st.session_state.rekap_nilai.append(
                {"Nama Siswa": nama_siswa, "Nilai Akhir": skor, "Detail": hasil}
            )
            st.success(
                f"Scan Selesai! Siswa **{nama_siswa}** mendapat Nilai:"
                f" **{skor}**"
            )
            with st.expander("Lihat Rincian Analisis Cyber"):
              st.markdown(hasil)
          except Exception as e:
            st.error(f"Error: {e}")
  st.markdown("</div>", unsafe_allow_html=True)

# 5. REKAP NILAI
elif menu_pilihan == "📊 5. Rekap Nilai":
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📊 Rekapitulasi Nilai")
  if len(st.session_state.rekap_nilai) == 0:
    st.info("Belum ada data nilai siswa yang terekam.")
  else:
    df = pd.DataFrame(st.session_state.rekap_nilai)
    st.dataframe(df[["Nama Siswa", "Nilai Akhir"]], use_container_width=True)
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Unduh Rekap (.csv)",
        data=csv,
        file_name="rekap_nilai_akademik.csv",
        mime="text/csv",
    )
    if st.button("⚡ EXECUTE: PURGE DATABASE"):
      st.session_state.rekap_nilai = []
      st.rerun()
  st.markdown("</div>", unsafe_allow_html=True)

# --- FOOTER ---
st.markdown(
    """
    <div class="footer">
        <p>Bumblebee Cyber-System | Developed by <b>Zeeo</b><br>WhatsApp: <a href="https://wa.me/6282371729760" target="_blank">082371729760</a><br>© 2026 All Rights Reserved</p>
    </div>
    """,
    unsafe_allow_html=True,
)
