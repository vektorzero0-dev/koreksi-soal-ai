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
# KONFIGURASI HALAMAN STREAMLIT
# ---------------------------------------------------------
st.set_page_config(
    page_title="Portal Akademik Guru AI",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------
# CUSTOM CSS: DESAIN MODERN SAAS DASHBOARD (EDTECH PREMIUM)
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    /* Global Body & Background */
    .stApp {
        background: #f1f5f9;
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #0f172a;
    }

    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background: #0f172a !important;
        border-right: 1px solid #1e293b;
    }
    section[data-testid="stSidebar"] .stMarkdown, 
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] span {
        color: #f8fafc !important;
    }

    /* Header Banner Executive Dashboard */
    .hero-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #3b82f6 100%);
        padding: 36px 40px;
        border-radius: 24px;
        color: #ffffff;
        margin-bottom: 28px;
        box-shadow: 0 20px 25px -5px rgba(15, 23, 42, 0.15), 0 8px 10px -6px rgba(15, 23, 42, 0.1);
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .hero-banner::after {
        content: '';
        position: absolute;
        top: -40px;
        right: -40px;
        width: 200px;
        height: 200px;
        background: rgba(255, 255, 255, 0.08);
        border-radius: 50%;
        pointer-events: none;
    }
    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        letter-spacing: -0.8px;
        color: #ffffff !important;
        margin: 0;
    }
    .hero-subtitle {
        font-size: 0.98rem;
        color: #e2e8f0 !important;
        margin-top: 8px;
        font-weight: 400;
        max-width: 650px;
        line-height: 1.6;
    }

    /* Main Content Card Container */
    .dashboard-card {
        background: #ffffff;
        padding: 32px;
        border-radius: 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.05);
        margin-bottom: 24px;
    }

    /* Section Headers inside Cards */
    .section-title {
        font-size: 1.35rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .section-desc {
        font-size: 0.88rem;
        color: #64748b;
        margin-bottom: 24px;
    }

    /* Form Labels & High-Contrast Input Fields */
    label, .stTextInput label, .stTextArea label, .stSelectbox label, .stRadio label {
        color: #334155 !important;
        font-weight: 700 !important;
        font-size: 0.88rem !important;
        letter-spacing: -0.2px;
        margin-bottom: 6px !important;
    }

    .stTextInput input, .stTextArea textarea {
        background-color: #ffffff !important;
        color: #0f172a !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        border-radius: 12px !important;
        border: 1.5px solid #cbd5e1 !important;
        padding: 12px 16px !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05) !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #2563eb !important;
        box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.15) !important;
    }
    .stTextInput input::placeholder, .stTextArea textarea::placeholder {
        color: #64748b !important;
        font-weight: 400 !important;
        opacity: 0.85 !important;
    }

    /* Primary Interactive Buttons */
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
        color: #ffffff !important;
        font-weight: 700 !important;
        border-radius: 12px !important;
        padding: 0.85rem 1.5rem !important;
        border: none !important;
        font-size: 0.95rem !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.3) !important;
        transition: all 0.25s ease !important;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.4) !important;
        transform: translateY(-2px);
    }

    /* Badge Indicators */
    .status-badge {
        display: inline-flex;
        align-items: center;
        padding: 6px 14px;
        border-radius: 9999px;
        font-size: 0.8rem;
        font-weight: 700;
        background: #eff6ff;
        color: #1d4ed8;
        border: 1px solid #bfdbfe;
        margin-bottom: 16px;
    }

    /* Footer Box */
    .footer-container {
        text-align: center;
        padding: 24px;
        font-size: 0.85rem;
        color: #64748b;
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        margin-top: 40px;
        margin-bottom: 20px;
        font-weight: 500;
    }
    .footer-container a {
        color: #2563eb;
        text-decoration: none;
        font-weight: 700;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# FUNGSI PEMBENTUK DOKUMEN WORD (.DOCX) DENGAN PARSER TUNGGAL
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
    if stripped.startswith("|") and stripped.endswith("|"):
      cells = [c.strip() for c in stripped.split("|")[1:-1]]
      table_rows.append(cells)
      continue
    else:
      flush_table()

    if not stripped:
      continue

    if stripped.startswith("# "):
      doc.add_heading(stripped.replace("# ", "").strip(), level=1)
    elif stripped.startswith("## "):
      doc.add_heading(stripped.replace("## ", "").strip(), level=2)
    elif stripped.startswith("### "):
      doc.add_heading(stripped.replace("### ", "").strip(), level=3)
    elif re.match(r"^(\d+[\.\)]|[a-zA-Z][\.\)]|\•|\-)\s+", stripped):
      doc.add_paragraph(stripped, style="List Bullet")
    else:
      doc.add_paragraph(stripped)

  flush_table()
  buffer = BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# ---------------------------------------------------------
# FUNGSI AUTO-RETRY UNTUK ERROR 503 (LONJAKAN SERVER)
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
# NAVIGASI SIDEBAR DASHBOARD
# ---------------------------------------------------------
with st.sidebar:
  st.markdown("### 🎓 Academic AI Studio")
  st.markdown("---")

  menu_pilihan = st.radio(
      "PILIH LAYANAN UTAMA:",
      [
          "📖 Generator Modul Ajar",
          "📝 Generator Soal & Kunci",
          "⚙️ Set Kunci Acuan",
          "🔍 Koreksi Siswa",
          "📊 Rekap Nilai",
      ],
      index=0,
  )

  st.markdown("---")
  st.markdown("### 🔑 Status API Key")

  if not GEMINI_API_KEY:
    st.warning("⚠️ Belum terhubung")
    input_manual = st.text_input("Gemini API Key", type="password")
    if input_manual:
      GEMINI_API_KEY = input_manual.strip().strip('"').strip("'")
  else:
    st.success("🔒 Terhubung & Siap Digunakan")

  st.markdown("---")
  st.caption("Engine: **Gemini 3.5 Flash**")
  st.caption("Versi Aplikasi: **3.0 Executive Edition**")

if not GEMINI_API_KEY:
  st.warning(
      "Mohon masukkan Gemini API Key di panel sebelah kiri untuk mulai"
      " menggunakan portal."
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
        <h1 class="hero-title">Portal Asisten Akademik Guru</h1>
        <p class="hero-subtitle">Platform Kecerdasan Buatan Terintegrasi untuk Menyusun Perangkat Pembelajaran, Bank Soal Standar Ujian, dan Koreksi Jawaban Otomatis secara Objektif.</p>
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
      ' Modul Ajar formal berbasis Kurikulum Merdeka.</div>',
      unsafe_allow_html=True,
  )

  col1, col2 = st.columns(2)
  with col1:
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
        "Kurikulum Validasi", placeholder="Contoh: Kurikulum Merdeka"
    )
  with col2:
    fase_kelas = st.text_input(
        "Fase / Kelas", placeholder="Contoh: Fase D / Kelas VII"
    )
    nama_ks = st.text_input(
        "Nama Kepala Sekolah",
        placeholder="Contoh: Dra. Hj. Siti Aminah, M.Pd.",
    )
    topik = st.text_input(
        "Topik / Materi Pokok", placeholder="Contoh: Sistem Pencernaan Manusia"
    )
    alokasi_waktu = st.text_input(
        "Alokasi Waktu", placeholder="Contoh: 2 Pertemuan (4 x 40 Menit)"
    )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("🚀 Buat Modul Ajar Sekarang", type="primary"):
    with st.spinner("Sistem Gemini sedang merancang Modul Ajar..."):
      try:
        prompt_modul = f"""
                Buatkan Modul Ajar / RPP formal, sangat terstruktur, dan profesional untuk instansi pendidikan:
                - Guru: {nama_guru}, Sekolah: {nama_sekolah}
                - Mapel: {mapel}, Kurikulum: {kurikulum_aktif}
                - Kelas: {fase_kelas}, Kepala Sekolah: {nama_ks}
                - Topik: {topik}, Waktu: {alokasi_waktu}
                Sertakan komponen Identitas Instansi, Profil Pelajar Pancasila, Tujuan Pembelajaran, Kegiatan Pembelajaran, serta Tabel Rubrik Penilaian dalam bentuk tabel markdown standar (menggunakan garis vertikal |).
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
# MENU 2: GENERATOR SOAL & KUNCI
# ---------------------------------------------------------
elif menu_pilihan == "📝 Generator Soal & Kunci":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">EVALUASI AKADEMIK</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-title">📝 Generator Paket Ujian & Kunci</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-desc">Penyusunan naskah soal ujian resmi lengkap'
      ' dengan KOP, kisi-kisi, kunci jawaban, dan rubrik bobot nilai.</div>',
      unsafe_allow_html=True,
  )

  col1, col2 = st.columns(2)
  with col1:
    s_guru = st.text_input("Nama Pembuat Soal", placeholder="Ahmad Fauzi, S.Pd.")
    s_sekolah = st.text_input(
        "Instansi / Sekolah", placeholder="SMP Negeri 1 Nusantara"
    )
    s_mapel = st.text_input("Mata Pelajaran", placeholder="Matematika")
    s_kur = st.text_input("Kurikulum", placeholder="Kurikulum Merdeka")
  with col2:
    s_kelas = st.text_input("Kelas / Semester", placeholder="Kelas VII / Ganjil")
    s_materi = st.text_input(
        "Materi Pokok", placeholder="Persamaan Linear Satu Variabel"
    )
    s_komposisi = st.text_input(
        "Komposisi Soal",
        placeholder="5 Pilihan Ganda, 2 Isian Singkat, 1 Essai",
    )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("🚀 Susun Naskah Soal Ujian", type="primary"):
    with st.spinner("Sistem Gemini sedang menyusun naskah ujian..."):
      try:
        prompt_soal = f"""
                Buatkan naskah soal ujian resmi instansi pendidikan lengkap dengan Kop Ujian, Petunjuk, Naskah Soal, Kunci Jawaban, & Rubrik Penilaian (gunakan tabel markdown standar dengan garis vertikal |) untuk:
                - Guru: {s_guru}, Sekolah: {s_sekolah}, Mapel: {s_mapel}
                - Kurikulum: {s_kur}, Kelas: {s_kelas}, Materi: {s_materi}
                - Komposisi: {s_komposisi}
                PENTING: Gunakan teks bersih murni tanpa tag HTML sama sekali (seperti <br> atau <p>).
                """
        response = generate_content_with_retry(
            client, "gemini-3.5-flash", prompt_soal
        )
        st.session_state.soal_hasil = response.text
        st.success("Paket soal berhasil disusun!")
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
      '<div class="section-title">⚙️ Konfigurasi Kunci Acuan</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-desc">Atur acuan kunci jawaban sebelum melakukan'
      ' koreksi jawaban siswa secara otomatis.</div>',
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
          " 10), P = 17m, L = 9m (Skor 10)"
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
      '<div class="section-title">🔍 Koreksi Lembar Jawaban Siswa</div>',
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
        with st.spinner("Sistem Gemini sedang menganalisis jawaban..."):
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
      '<div class="section-title">📊 Rekapitulasi Nilai Ujian Siswa</div>',
      unsafe_allow_html=True,
  )

  if len(st.session_state.rekap_nilai) == 0:
    st.info("Belum ada data rekapitulasi nilai siswa yang terekam.")
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
        <p>Sistem Portal Akademik Terintegrasi | Developed by <b>Zeeo</b><br>
        WhatsApp Support: <a href="https://wa.me/6282371729760" target="_blank">082371729760</a><br>
        © 2026 All Rights Reserved</p>
    </div>
    """,
    unsafe_allow_html=True,
)
