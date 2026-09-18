from io import BytesIO
import os
import re
from docx import Document
from google import genai
from PIL import Image

pypdf_available = True
try:
  from pypdf import PdfReader
except ImportError:
  pypdf_available = False

import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# KONFIGURASI HALAMAN (RESPONSIF ANDROID & DESKTOP)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Portal Akademik & Asesmen Elite",
    page_icon="✨",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------
# CUSTOM CSS: SENI VISUAL LUXURY CYBER-VELVET & PROFESIONAL
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    /* Global Dark Art Background */
    .stApp {
        background: radial-gradient(circle at 50% 0%, #1e1b4b 0%, #090d16 60%, #030712 100%);
        background-attachment: fixed;
        font-family: 'Plus Jakarta Sans', sans-serif;
        color: #f8fafc;
    }

    /* Sidebar Artsy & Elegan */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #090d16 0%, #111827 100%) !important;
        border-right: 1px solid rgba(139, 92, 246, 0.2);
    }
    section[data-testid="stSidebar"] .stMarkdown, 
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] span {
        color: #f3f4f6 !important;
    }

    /* Hero Banner Seni Visual Tinggi */
    .hero-banner {
        background: linear-gradient(135deg, #312e81 0%, #4c1d95 50%, #0f172a 100%);
        padding: 30px 24px;
        border-radius: 20px;
        color: #ffffff;
        margin-bottom: 24px;
        box-shadow: 0 20px 40px -15px rgba(109, 40, 217, 0.4);
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(167, 139, 250, 0.3);
    }
    .hero-banner::before {
        content: '';
        position: absolute;
        top: -50px;
        right: -50px;
        width: 150px;
        height: 150px;
        background: rgba(6, 182, 212, 0.25);
        border-radius: 50%;
        filter: blur(35px);
        pointer-events: none;
    }
    .hero-title {
        font-size: 1.6rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        color: #c084fc !important;
        margin: 0;
    }
    .hero-subtitle {
        font-size: 0.88rem;
        color: #cbd5e1 !important;
        margin-top: 8px;
        font-weight: 400;
        line-height: 1.5;
    }

    /* Kartu Dashboard Glassmorphism Berpendar */
    .dashboard-card {
        background: rgba(15, 23, 42, 0.85);
        backdrop-filter: blur(20px);
        padding: 24px;
        border-radius: 20px;
        border: 1px solid rgba(139, 92, 246, 0.25);
        box-shadow: 0 15px 35px -10px rgba(0, 0, 0, 0.7);
        margin-bottom: 20px;
        transition: all 0.3s ease;
    }
    .dashboard-card:hover {
        border-color: rgba(6, 182, 212, 0.5);
        box-shadow: 0 20px 40px -10px rgba(109, 40, 217, 0.3);
    }

    /* Judul Bagian dalam Kartu */
    .section-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: #f3f4f6;
        margin-bottom: 4px;
    }
    .section-desc {
        font-size: 0.86rem;
        color: #94a3b8;
        margin-bottom: 20px;
    }

    /* Label Formulir Responsif */
    label, .stTextInput label, .stTextArea label, .stSelectbox label, .stRadio label {
        color: #e2e8f0 !important;
        font-weight: 700 !important;
        font-size: 0.88rem !important;
        margin-bottom: 6px !important;
    }

    /* Input & Textarea Nyaman & Kontras Terang */
    .stTextInput input, .stTextArea textarea {
        background-color: #030712 !important;
        color: #38bdf8 !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        border-radius: 12px !important;
        border: 1.5px solid #334155 !important;
        padding: 12px 14px !important;
        box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.6) !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #a855f7 !important;
        box-shadow: 0 0 0 3px rgba(168, 85, 247, 0.25), inset 0 2px 4px rgba(0, 0, 0, 0.6) !important;
        background-color: #030712 !important;
    }
    .stTextInput input::placeholder, .stTextArea textarea::placeholder {
        color: #475569 !important;
        opacity: 1 !important;
    }

    /* Tombol Utama Gradien Cyber */
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #7c3aed 0%, #06b6d4 100%);
        color: #ffffff !important;
        font-weight: 800 !important;
        border-radius: 12px !important;
        padding: 0.85rem 1rem !important;
        border: none !important;
        font-size: 0.95rem !important;
        box-shadow: 0 6px 20px rgba(124, 58, 237, 0.4) !important;
        transition: all 0.25s ease !important;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #8b5cf6 0%, #22d3ee 100%) !important;
        box-shadow: 0 8px 25px rgba(6, 182, 212, 0.6) !important;
        transform: translateY(-2px);
    }

    /* Badge Estetik */
    .status-badge {
        display: inline-flex;
        align-items: center;
        padding: 5px 14px;
        border-radius: 9999px;
        font-size: 0.72rem;
        font-weight: 800;
        letter-spacing: 0.8px;
        text-transform: uppercase;
        background: rgba(168, 85, 247, 0.15);
        color: #d8b4fe;
        border: 1px solid rgba(168, 85, 247, 0.35);
        margin-bottom: 12px;
    }

    /* Footer */
    .footer-container {
        text-align: center;
        padding: 20px;
        font-size: 0.8rem;
        color: #94a3b8;
        background: rgba(15, 23, 42, 0.6);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        margin-top: 30px;
        margin-bottom: 20px;
        font-weight: 500;
    }
    .footer-container a {
        color: #38bdf8;
        text-decoration: none;
        font-weight: 700;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# FUNGSI PEMBENTUK DOKUMEN WORD (.DOCX) DENGAN PARSER TABEL PRESISI
# ---------------------------------------------------------
def buat_file_docx(teks_konten):
  doc = Document()

  # Bersihkan tag HTML mentah
  teks_bersih = re.sub(
      r"<\s*br\s*/?>", "\n", teks_konten, flags=re.IGNORECASE
  )
  teks_bersih = re.sub(r"</?[a-zA-Z]+[^>]*>", "", teks_bersih)

  lines = teks_bersih.split("\n")
  table_rows = []

  def flush_table():
    nonlocal table_rows
    if table_rows:
      # Saring baris pemisah markdown (seperti |---|---|)
      filtered_rows = [
          row
          for row in table_rows
          if not all(
              c in "-:| "
              for c in "".join(
                  [cell for cell in row if isinstance(cell, str)]
              )
          )
      ]
      if filtered_rows:
        num_cols = max(len(row) for row in filtered_rows)
        table = doc.add_table(rows=len(filtered_rows), cols=num_cols)
        table.style = "Table Grid"
        for r_idx, row_data in enumerate(filtered_rows):
          for c_idx, cell_value in enumerate(row_data):
            if c_idx < num_cols:
              table.cell(r_idx, c_idx).text = str(cell_value).strip()
      table_rows = []

  for line in lines:
    stripped = line.strip()

    # Deteksi baris tabel markdown yang diawali dan diakhiri karakter pipa (|)
    if stripped.startswith("|") and stripped.endswith("|"):
      cells = [c.strip() for c in stripped.split("|")[1:-1]]
      table_rows.append(cells)
      continue
    else:
      flush_table()

    if not stripped:
      continue

    # Format Dokumen (Heading & Penomoran)
    if stripped.startswith("# "):
      doc.add_heading(stripped.replace("# ", "").strip(), level=1)
    elif stripped.startswith("## "):
      doc.add_heading(stripped.replace("## ", "").strip(), level=2)
    elif stripped.startswith("### "):
      doc.add_heading(stripped.replace("### ", "").strip(), level=3)
    elif re.match(r"^(\d+[\.\)]|[a-zA-Z][\.\)]|\•|\-|\*)\s+", stripped):
      doc.add_paragraph(stripped, style="List Bullet")
    else:
      doc.add_paragraph(stripped)

  flush_table()  # Pastikan tabel terakhir ikut diproses
  buffer = BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# ---------------------------------------------------------
# FUNGSI AUTO-RETRY UNTUK ERROR 503
# ---------------------------------------------------------
def generate_content_with_retry(
    client, model_name, contents, max_retries=3, delay=3
):
  import time

  for attempt in range(max_retries):
    try:
      return client.models.generate_content(model=model_name, contents=contents)
    except Exception as e:
      if "503" in str(e) or "UNAVAILABLE" in str(e):
        if attempt < max_retries - 1:
          time.sleep(delay * (attempt + 1))
          continue
      raise e


# ---------------------------------------------------------
# INISIALISASI ENVIRONMENT & API KEY
# ---------------------------------------------------------
for var in [
    "GOOGLE_GENAI_USE_VERTEXAI",
    "VERTEXAI_PROJECT",
    "VERTEXAI_LOCATION",
    "GOOGLE_CLOUD_PROJECT",
]:
  if var in os.environ:
    del os.environ[var]

raw_key = st.secrets.get("GEMINI_API_KEY", "") or st.secrets.get(
    "GOOGLE_API_KEY", ""
)
GEMINI_API_KEY = str(raw_key).strip().strip('"').strip("'")

# ---------------------------------------------------------
# NAVIGASI SIDEBAR
# ---------------------------------------------------------
with st.sidebar:
  st.markdown("### ✨ Elite Academic AI")
  st.markdown(
      "<p style='color: #c084fc; font-size: 0.78rem; margin-top:"
      " -10px;'>Mobile & Desktop Optimized</p>",
      unsafe_allow_html=True,
  )
  st.markdown("---")

  menu_pilihan = st.radio(
      "NAVIGASI UTAMA:",
      [
          "📖 Generator Modul Ajar",
          "📝 Generator Soal Asesmen",
          "⚙️ Set Kunci Acuan",
          "🔍 Koreksi Siswa",
          "📊 Rekap Nilai",
      ],
      index=0,
  )

  st.markdown("---")
  st.markdown("### 🔑 Status Koneksi")

  if not GEMINI_API_KEY:
    st.warning("⚠️ Belum terhubung")
    input_manual = st.text_input("Gemini API Key", type="password")
    if input_manual:
      GEMINI_API_KEY = input_manual.strip().strip('"').strip("'")
  else:
    st.success("🔒 Sistem Aktif & Aman")

if not GEMINI_API_KEY:
  st.warning(
      "Mohon masukkan Gemini API Key di panel menu samping untuk mengakses"
      " aplikasi."
  )
  st.stop()

try:
  client = genai.Client(api_key=GEMINI_API_KEY)
except Exception as e:
  st.error(f"Gagal inisialisasi Client AI: {e}")
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

# ---------------------------------------------------------
# HERO BANNER UTAMA
# ---------------------------------------------------------
st.markdown(
    """
    <div class="hero-banner">
        <h1 class="hero-title">Portal Akademik & Asesmen Elite</h1>
        <p class="hero-subtitle">Sistem Kecerdasan Buatan Terintegrasi untuk Penyusunan Perangkat Pembelajaran, Naskah Soal Asesmen, dan Penilaian Objektif.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# MENU 1: MODUL AJAR
# ---------------------------------------------------------
if menu_pilihan == "📖 Generator Modul Ajar":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">PERANGKAT PEMBELAJARAN</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-title">📖 Generator Modul Ajar / RPP</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-desc">Lengkapi formulir di bawah ini untuk merancang'
      ' Modul Ajar formal berstandar Kurikulum Merdeka.</div>',
      unsafe_allow_html=True,
  )

  nama_guru = st.text_input(
      "Nama Guru & Gelar", placeholder="Contoh: Ahmad Fauzi, S.Pd."
  )
  nama_sekolah = st.text_input(
      "Nama Instansi / Sekolah", placeholder="Contoh: SMP Negeri 1 Nusantara"
  )
  mapel = st.text_input(
      "Mata Pelajaran", placeholder="Contoh: Ilmu Pengetahuan Alam (IPA)"
  )
  kurikulum_aktif = st.text_input(
      "Kurikulum", placeholder="Contoh: Kurikulum Merdeka"
  )
  fase_kelas = st.text_input(
      "Fase / Kelas", placeholder="Contoh: Fase D / Kelas VII"
  )
  nama_ks = st.text_input(
      "Nama Kepala Sekolah", placeholder="Contoh: Dra. Hj. Siti Aminah, M.Pd."
  )
  topik = st.text_input(
      "Topik / Materi Pokok", placeholder="Contoh: Sistem Pencernaan Manusia"
  )
  alokasi_waktu = st.text_input(
      "Alokasi Waktu", placeholder="Contoh: 2 Pertemuan (4 x 40 Menit)"
  )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("✨ Buat Modul Ajar Sekarang", type="primary"):
    with st.spinner("Sistem Gemini sedang merancang Modul Ajar..."):
      try:
        prompt_modul = f"""
                Buatkan Modul Ajar / RPP formal, sangat terstruktur, dan profesional untuk instansi pendidikan:
                - Guru: {nama_guru}, Sekolah: {nama_sekolah}
                - Mapel: {mapel}, Kurikulum: {kurikulum_aktif}
                - Kelas: {fase_kelas}, Kepala Sekolah: {nama_ks}
                - Topik: {topik}, Waktu: {alokasi_waktu}
                Sertakan komponen Identitas Instansi, Profil Pelajar Pancasila, Tujuan Pembelajaran, Kegiatan Pembelajaran, serta Tabel Rubrik Penilaian dalam bentuk tabel markdown standar lengkap menggunakan garis vertikal (|) untuk kolom dan barisnya.
                PENTING: Gunakan teks bersih murni tanpa tag HTML sama sekali (seperti <br> atau <p>).
                """
        response = generate_content_with_retry(
            client, "gemini-3.5-flash", prompt_modul
        )
        st.session_state.modul_hasil = response.text
        st.success("Modul Ajar berhasil disusun!")
      except Exception as e:
        st.error(f"Error sistem: {e}")

  if st.session_state.modul_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Dokumen Modul Ajar")
    st.markdown(st.session_state.modul_hasil)
    file_docx = buat_file_docx(st.session_state.modul_hasil)
    st.download_button(
        "📥 Unduh Modul Ajar (.docx)",
        data=file_docx,
        file_name=f"Modul_Ajar_{mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# MENU 2: GENERATOR SOAL ASESMEN
# ---------------------------------------------------------
elif menu_pilihan == "📝 Generator Soal Asesmen":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">ASESMEN & EVALUASI</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-title">📝 Generator Soal Asesmen & Kunci</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-desc">Penyusunan naskah soal asesmen resmi lengkap'
      ' dengan penomoran, kunci jawaban, dan rubrik bobot nilai berbentuk'
      ' tabel.</div>',
      unsafe_allow_html=True,
  )

  s_guru = st.text_input("Nama Pembuat Asesmen", placeholder="Ahmad Fauzi, S.Pd.")
  s_sekolah = st.text_input(
      "Instansi / Sekolah", placeholder="SMP Negeri 1 Nusantara"
  )
  s_mapel = st.text_input("Mata Pelajaran", placeholder="Matematika")
  s_kur = st.text_input("Kurikulum", placeholder="Kurikulum Merdeka")
  s_kelas = st.text_input("Kelas / Semester", placeholder="Kelas VII / Ganjil")
  s_materi = st.text_input(
      "Materi Asesmen", placeholder="Persamaan Linear Satu Variabel"
  )
  s_komposisi = st.text_input(
      "Komposisi Soal Asesmen",
      placeholder="5 Pilihan Ganda, 2 Isian Singkat, 1 Essai",
  )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("✨ Susun Naskah Soal Asesmen", type="primary"):
    with st.spinner("Sistem Gemini sedang menyusun naskah asesmen..."):
      try:
        prompt_soal = f"""
                Buatkan naskah soal asesmen resmi instansi pendidikan lengkap dengan Kop Soal, Petunjuk, Naskah Soal Asesmen (setiap nomor soal menggunakan penomoran tegas seperti 1., 2., 3. dan pilihan ganda menggunakan A., B., C., D.), Kunci Jawaban, & Tabel Rubrik Penilaian.
                PENTING UNTUK TABEL: Buat tabel rubrik penilaian menggunakan format tabel markdown standar dengan garis vertikal (|), contoh:
                | Jenis Soal | Jumlah Soal | Bobot per Soal | Skor Maksimal |
                | :--- | :--- | :--- | :--- |
                | Pilihan Ganda | 5 | 10 | 50 |
                
                Data Asesmen:
                - Guru: {s_guru}, Sekolah: {s_sekolah}, Mapel: {s_mapel}
                - Kurikulum: {s_kur}, Kelas: {s_kelas}, Materi: {s_materi}
                - Komposisi: {s_komposisi}
                PENTING: Gunakan teks bersih murni tanpa tag HTML sama sekali (seperti <br> atau <p>).
                """
        response = generate_content_with_retry(
            client, "gemini-3.5-flash", prompt_soal
        )
        st.session_state.soal_hasil = response.text
        st.success("Paket soal asesmen berhasil disusun!")
      except Exception as e:
        st.error(f"Error sistem: {e}")

  if st.session_state.soal_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Naskah Soal Asesmen")
    st.markdown(st.session_state.soal_hasil)
    file_docx_soal = buat_file_docx(st.session_state.soal_hasil)
    st.download_button(
        "📥 Unduh Soal Asesmen (.docx)",
        data=file_docx_soal,
        file_name=f"Soal_Asesmen_{s_mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# MENU 3: SET KUNCI ACUAN
# ---------------------------------------------------------
elif menu_pilihan == "⚙️ Set Kunci Acuan":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">KONFIGURASI PENILAIAN</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-title">⚙️ Konfigurasi Kunci Acuan Asesmen</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-desc">Atur acuan kunci jawaban asesmen sebelum'
      ' melakukan koreksi jawaban siswa secara otomatis.</div>',
      unsafe_allow_html=True,
  )

  kunci_pg_input = st.text_area(
      "Kunci Pilihan Ganda",
      value=st.session_state.kunci_pg,
      placeholder="Contoh: 1.A, 2.B, 3.C, 4.D, 5.B",
      height=100,
  )
  kunci_isian_input = st.text_area(
      "Kunci Isian Singkat & Keyword Jawaban",
      value=st.session_state.kunci_isian,
      placeholder="Contoh: 1. y = 6; 2. x = 5",
      height=100,
  )
  kunci_essai_input = st.text_area(
      "Rubrik Penilaian & Kunci Essai",
      value=st.session_state.kunci_essai,
      placeholder=(
          "Contoh: Soal 1: Model matematika 8x + 12 = 52 (Skor 10), x = 5 (Skor"
          " 10)"
      ),
      height=150,
  )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("💾 Simpan Kunci Acuan", type="primary"):
    st.session_state.kunci_pg = kunci_pg_input
    st.session_state.kunci_isian = kunci_isian_input
    st.session_state.kunci_essai = kunci_essai_input
    st.success("Kunci acuan berhasil disimpan dalam sistem!")
  st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# MENU 4: KOREKSI SISWA
# ---------------------------------------------------------
elif menu_pilihan == "🔍 Koreksi Siswa":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">KOREKSI OTOMATIS</div>', unsafe_allow_html=True
  )
  st.markdown(
      '<div class="section-title">🔍 Koreksi Lembar Jawaban Asesmen Siswa</div>',
      unsafe_allow_html=True,
  )

  if (
      not st.session_state.kunci_pg
      and not st.session_state.kunci_isian
      and not st.session_state.kunci_essai
  ):
    st.warning("⚠️ Harap simpan Kunci Jawaban terlebih dahulu pada Menu 3!")
  else:
    nama_siswa = st.text_input(
        "Identitas Siswa", placeholder="Masukkan Nama Lengkap / NISN Siswa"
    )
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
      teks_manual = st.text_area(
          "Masukkan teks jawaban siswa",
          placeholder="Ketik atau tempel lembar jawaban siswa di sini...",
      )

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Jalankan Analisis & Koreksi AI", type="primary"):
      if not nama_siswa:
        st.error("Masukkan identitas atau nama siswa!")
      else:
        with st.spinner("Sistem Gemini sedang menganalisis lembar asesmen..."):
          try:
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

            resp = generate_content_with_retry(
                client, "gemini-3.5-flash", contents=payload
            )
            hasil = resp.text

            m = re.search(r"NILAI_AKHIR:\s*([0-9.]+)", hasil)
            skor = float(m.group(1)) if m else 0.0

            st.session_state.rekap_nilai.append({
                "Nama Siswa": nama_siswa,
                "Nilai Akhir": skor,
                "Detail": hasil,
            })
            st.success(
                f"Koreksi Selesai! Siswa **{nama_siswa}** mendapat Nilai:"
                f" **{skor}**"
            )
            with st.expander("Lihat Rincian Analisis Penilaian"):
              st.markdown(hasil)
          except Exception as e:
            st.error(f"Error: {e}")
  st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# MENU 5: REKAP NILAI
# ---------------------------------------------------------
elif menu_pilihan == "📊 Rekap Nilai":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">REKAPITULASI DOKUMEN</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-title">📊 Rekapitulasi Nilai Asesmen Siswa</div>',
      unsafe_allow_html=True,
  )

  if len(st.session_state.rekap_nilai) == 0:
    st.info("Belum ada data rekapitulasi nilai asesmen yang terekam.")
  else:
    df = pd.DataFrame(st.session_state.rekap_nilai)
    st.dataframe(df[["Nama Siswa", "Nilai Akhir"]], use_container_width=True)
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Unduh Rekap (.csv)",
        data=csv,
        file_name="rekap_nilai_asesmen.csv",
        mime="text/csv",
    )
    if st.button("🗑️ Kosongkan Rekap"):
      st.session_state.rekap_nilai = []
      st.rerun()
  st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# FOOTER
# ---------------------------------------------------------
st.markdown(
    """
    <div class="footer-container">
        <p>Portal Akademik Elite | Developed by <b>Zeeo</b><br>
        WhatsApp Support: <a href="https://wa.me/6282371729760" target="_blank">082371729760</a><br>
        © 2026 All Rights Reserved</p>
    </div>
    """,
    unsafe_allow_html=True,
)
