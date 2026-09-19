from io import BytesIO
import os
import re
import docx
from docx import Document
from google import genai
from PIL import Image

pypdf_available = True
try:
  from pypdf import PdfReader
except ImportError:
  pypdf_available = False

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------
# KONFIGURASI HALAMAN (STABIL & RESPONSIF)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Portal Akademik & Asesmen Instansi Pendidikan",
    page_icon="🏛️",
    layout="centered",
    initial_sidebar_state="collapsed",
)


# ---------------------------------------------------------
# FUNGSI KONEKSI GOOGLE SHEETS (HARDCODED SERVICE ACCOUNT)
# ---------------------------------------------------------
def get_google_sheet_data():
  try:
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds_dict = {
        "type": "service_account",
        "project_id": "cedar-router-509114-q4",
        "private_key_id": "f5b6dfe116cf51f4d5259b723fafe06b3a121daa",
        "private_key": r"""-----BEGIN PRIVATE KEY-----
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQD10EbEvwaHzJFX
K+zs8Ee15BmMT8m6RnZL4fExow0KBVtBd3nrQtO6WYp/wjYFaMWwAGAPGVQ1jPJM
jOasWRssjOU2U6JyoEljr5+DJMAX5dPaf15PVNKjwAuKGqn3UInrjMjglY41zbhW
nn4NWDLYrJTO1G0Zxfuwnb6e+Krj9oRBJBYkleYAcyuKs2CP6W++yQCZwWLxmKCJO
rommmrzuvkywyu40FktUTNdDgY9/Zk5uNS6XkD0eaNmrhzKr9F/PHQe/xgUID0Ov
XFa71J6//Nmk/L0kFPdFH/UkWryoo0035BUMUsjJyGKH2NCeVKkWmpV49bZcn+7T
DBIBiIL9AgMBAAECggEAGCtBFg2xOu90d+V3/2j1qA70Kx1aOJA+h+IAiNCfEtXp
fdOdP3I75qXwrfSewqPlUhOLXlivdK910GyHLrl5KEV3MQMCpTIY3S4SqT6XC5J7
pM9iqrqsllCm4c6S2SAIW4plYt2ZlLU0M4IuE+seizn2bk9u+vBPkGYUUGibdUZs
QbNZbJ7KlZqIMpOpD3nV4kjUwxk6b7sW0AtBHUDYQg/Yo241LJFFBla5tpt1v/Dj
LJiYq4Fbfz/PHobcR6eC/oBpYwK2jbIIUZGbfhRtELmW5ONuySPsjZZCmfv0b7Qa
4fvf82pUX05/Opoy0VZ4qhCFTjqmys/g0lsBjLzPcQKBgQD+GkDt6tm+ZxJDPhAj
rb1Iw65mcZ+vZAvx2M/peG2joDYLwhWa9q6Z2udlHh6AfHb2u2+iGs/77jHnEwTC
LsWyORo/e4ooPm3GDbqOmyEwpmWUiw3a5qm/tMV6CcEhT5URlwMoAXrh3PHoZ5KQ
bz4GfK3yRn9ypXPiwV6RXE3bMQKBgQD3pi1vTO4nD6vv/NsQu0R7689B9ZsGUM0Z
JP1mJZObj2J+ZrREqJWqUZu2fE5L2XuZY9IQBikbl1FNgDnxVmz01p95wZ9ozdvF
k1UpvjKQr/8dg0eo8E676KMtKC9WInofOBWtjgP1mXFvtA6dPw7E02yEvEBzlvck
uJ+XZm8ZjQKBgQDpt04XRxbF6VnD3XbMykW6grmLYmEE2lmeNdRuIqV9haOQRxDG
OrS3sL96oyxc854cLKRuDolUaG8f4b9Tt9+AoMMCtueJQnqHWyNHfWoWrEXsTcYN
nnHFvcZ7dM9GeiOtMhYCSsGHNEwKxx2noTVlYcB8yIyOgWIvxefg4bRTzUQKBgQDl
nrGx+yK4l89bV33+baNH+y5eP6KQ5mz5bj36i+T6ICtahu8Z71o3XQ5BSEb7bgXur
qnPrAIunVxLD+aPDOxAZkeKdHQEmRaUI+7cD260xmsfTKymOeC/M/dg3zQj5rUft
JCqWpxrs773QhwD2vMCJsjr2b1Cm4t+aYs8/rnRjEQKBgEWnBVEJ8sBgXO5BQg2i
6zWbownr4WgH1GWWfPkMF6h1Nuu3GMdFOINC/6BgAXMoSfXpjoRKmwQ7SVMDE4S1
YeTHVd7RfdQ/xPqkUmU5ZSeoVBvHLqL+li1RLZ0MgJqhR2Kdhf6fHYrziA9Glw5t
fg38cGt8VuHOP9xdfzI+zY/I=
-----END PRIVATE KEY-----""",
        "client_email": (
            "bot-kuota@cedar-router-509114-q4.iam.gserviceaccount.com"
        ),
        "client_id": "108263579617939653999",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": (
            "https://www.googleapis.com/oauth2/v1/certs"
        ),
        "client_x509_cert_url": (
            "https://www.googleapis.com/robot/v1/metadata/x509/bot-kuota%40cedar-router-509114-q4.iam.gserviceaccount.com"
        ),
        "universe_domain": "googleapis.com",
    }

    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    sheet_name = st.secrets["sheet_config"]["sheet_name"]
    sheet = client.open(sheet_name).sheet1
    return sheet
  except Exception as e:
    st.error(f"Gagal terhubung ke Google Sheets: {e}")
    return None


# ---------------------------------------------------------
# CUSTOM CSS: ANIMASI BACKGROUND BERGERAK DINAMIS & WARNA KONTRAST TINGGI
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    .stApp {
        background: linear-gradient(125deg, #030712 0%, #0f172a 35%, #1e1b4b 70%, #064e3b 100%);
        background-size: 400% 400%;
        animation: gradientAnimation 18s ease infinite;
        font-family: 'Plus Jakarta Sans', sans-serif;
        color: #f8fafc;
    }

    @keyframes gradientAnimation {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    .stApp::before {
        content: '';
        position: fixed;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background-image: 
            radial-gradient(circle, rgba(251, 191, 36, 0.15) 1.5px, transparent 1.5px),
            radial-gradient(circle, rgba(56, 189, 248, 0.12) 2px, transparent 2px),
            linear-gradient(rgba(139, 92, 246, 0.05) 1px, transparent 1px),
            linear-gradient(90deg, rgba(139, 92, 246, 0.05) 1px, transparent 1px);
        background-size: 60px 60px, 90px 90px, 45px 45px, 45px 45px;
        animation: gridMove 20s linear infinite;
        z-index: 0;
        pointer-events: none;
    }

    @keyframes gridMove {
        0% { transform: translate(0, 0) rotate(0deg); }
        100% { transform: translate(60px, 60px) rotate(3deg); }
    }

    section[data-testid="stSidebar"] {
        background: rgba(3, 7, 18, 0.95) !important;
        border-right: 1.5px solid rgba(251, 191, 36, 0.3);
        backdrop-filter: blur(16px);
    }
    section[data-testid="stSidebar"] .stMarkdown, 
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] span {
        color: #f3f4f6 !important;
    }

    .hero-banner {
        background: linear-gradient(135deg, rgba(30, 58, 138, 0.9) 0%, rgba(15, 23, 42, 0.95) 100%);
        backdrop-filter: blur(12px);
        padding: 32px 24px;
        border-radius: 20px;
        color: #ffffff;
        margin-bottom: 24px;
        box-shadow: 0 20px 40px -15px rgba(30, 58, 138, 0.7);
        position: relative;
        overflow: hidden;
        border-left: 6px solid #fbbf24;
        border-top: 1px solid rgba(255, 255, 255, 0.2);
        z-index: 1;
    }
    .hero-title {
        font-size: 1.55rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        color: #fbbf24 !important;
        margin: 0;
    }
    .hero-subtitle {
        font-size: 0.86rem;
        color: #cbd5e1 !important;
        margin-top: 8px;
        font-weight: 400;
        line-height: 1.5;
    }

    .dashboard-card {
        background: rgba(15, 23, 42, 0.88);
        backdrop-filter: blur(24px);
        padding: 24px;
        border-radius: 20px;
        border: 1.5px solid rgba(251, 191, 36, 0.35);
        box-shadow: 0 20px 40px -10px rgba(0, 0, 0, 0.8);
        margin-bottom: 20px;
        position: relative;
        z-index: 1;
    }

    .section-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: #f8fafc;
        margin-bottom: 4px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .section-desc {
        font-size: 0.86rem;
        color: #94a3b8;
        margin-bottom: 20px;
    }

    label, .stTextInput label, .stTextArea label, .stSelectbox label, .stRadio label {
        color: #f1f5f9 !important;
        font-weight: 700 !important;
        font-size: 0.88rem !important;
        margin-bottom: 6px !important;
    }

    .stTextInput input, .stTextArea textarea {
        background-color: #030712 !important;
        color: #fbbf24 !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        border-radius: 12px !important;
        border: 1.5px solid #475569 !important;
        padding: 12px 14px !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #fbbf24 !important;
        box-shadow: 0 0 0 3px rgba(251, 191, 36, 0.3) !important;
        background-color: #030712 !important;
    }

    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #d97706 0%, #fbbf24 100%);
        color: #030712 !important;
        font-weight: 800 !important;
        border-radius: 12px !important;
        padding: 0.85rem 1rem !important;
        border: none !important;
        font-size: 0.95rem !important;
        box-shadow: 0 6px 20px rgba(245, 158, 11, 0.4) !important;
        transition: all 0.25s ease !important;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #fbbf24 0%, #fef08a 100%) !important;
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(251, 191, 36, 0.6) !important;
    }

    .status-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 5px 14px;
        border-radius: 9999px;
        font-size: 0.72rem;
        font-weight: 800;
        letter-spacing: 0.8px;
        text-transform: uppercase;
        background: rgba(251, 191, 36, 0.2);
        color: #fef08a;
        border: 1px solid rgba(251, 191, 36, 0.4);
        margin-bottom: 12px;
    }

    .footer-container {
        text-align: center;
        padding: 20px;
        font-size: 0.8rem;
        color: #cbd5e1;
        background: rgba(15, 23, 42, 0.75);
        border: 1px solid rgba(251, 191, 36, 0.2);
        border-radius: 16px;
        margin-top: 30px;
        margin-bottom: 20px;
        position: relative;
        z-index: 1;
    }
    .footer-container a {
        color: #fbbf24;
        text-decoration: none;
        font-weight: 700;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# FUNGSI PEMBENTUK DOKUMEN WORD (.DOCX) DENGAN PARSER TABEL
# ---------------------------------------------------------
def buat_file_docx(teks_konten):
  doc = Document()
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
    elif re.match(r"^([a-zA-Z][\.\)]|\•|\-|\*)", stripped):
      p = doc.add_paragraph(stripped, style="List Bullet")
      p.paragraph_format.left_indent = docx.shared.Inches(0.4)
    elif re.match(r"^\d+[\.\)]", stripped):
      doc.add_paragraph(stripped)
    else:
      doc.add_paragraph(stripped)

  flush_table()
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
# INISIALISASI ENVIRONMENT & API KEY (DARI SECRETS)
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

if not GEMINI_API_KEY:
  st.warning("Mohon masukkan Gemini API Key di panel Secrets.")
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
if "cpatp_hasil" not in st.session_state:
  st.session_state.cpatp_hasil = ""
if "kktp_hasil" not in st.session_state:
  st.session_state.kktp_hasil = ""
if "prosemprota_hasil" not in st.session_state:
  st.session_state.prosemprota_hasil = ""
if "soal_hasil" not in st.session_state:
  st.session_state.soal_hasil = ""

# ---------------------------------------------------------
# NAVIGASI SIDEBAR & VALIDASI TOKEN GOOGLE SHEETS
# ---------------------------------------------------------
with st.sidebar:
  st.markdown("### 🏛️ Portal Akademik Resmi")
  st.markdown(
      "<p style='color: #fbbf24; font-size: 0.78rem; margin-top:"
      " -10px;'>Instansi Pendidikan Formal</p>",
      unsafe_allow_html=True,
  )
  st.markdown("---")

  st.markdown("### 🎟️ Aktivasi Token Kuota")
  input_token = st.text_input(
      "Masukkan Token Anda", type="password", placeholder="Contoh: VIP-SITI-10"
  )

  token_aktif = False
  sisa_kuota_sekarang = 0
  row_index = None
  sheet = None

  cleaned_token = input_token.strip()

  if cleaned_token != "":
    sheet = get_google_sheet_data()
    if sheet:
      try:
        records = sheet.get_all_records()
        for idx, row in enumerate(records, start=2):
          if str(row.get("token")).strip() == cleaned_token:
            sisa_kuota_sekarang = int(row.get("kuota", 0))
            row_index = idx
            break

        if row_index is not None:
          if sisa_kuota_sekarang > 0:
            token_aktif = True
            st.success(f"✅ Token Aktif! Sisa Kuota: **{sisa_kuota_sekarang}x**")
          else:
            st.error("❌ Kuota Anda sudah habis di Google Sheets!")
        else:
          st.error("❌ Token tidak ditemukan.")
      except Exception as e:
        st.error(f"Gagal membaca database: {e}")
    else:
      st.error("Gagal terhubung ke Google Sheets.")
  else:
    st.warning("⚠️ Masukkan token untuk mulai menggunakan generator.")

  st.markdown("---")
  menu_pilihan = st.radio(
      "NAVIGASI UTAMA:",
      [
          "📖 Generator Modul Ajar",
          "🎯 Generator CP & ATP",
          "📊 Generator KKTP",
          "📅 Generator Prosem & Prota",
          "📝 Generator Soal Asesmen",
          "⚙️ Set Kunci Acuan",
          "🔍 Koreksi Siswa",
          "📊 Rekap Nilai",
      ],
      index=0,
  )

# ---------------------------------------------------------
# HERO BANNER UTAMA
# ---------------------------------------------------------
st.markdown(
    """
    <div class="hero-banner">
        <h1 class="hero-title">🏛️ Portal Asisten Akademik & Asesmen</h1>
        <p class="hero-subtitle">Sistem Terintegrasi Resmi Instansi Pendidikan untuk Penyusunan Perangkat Pembelajaran, CP/ATP/KKTP/Prosem/Prota, Naskah Asesmen, dan Evaluasi.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if not token_aktif:
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown("### 🔒 Akses Terkunci")
  st.markdown(
      "Masukkan **Token Kuota** yang valid di panel *sidebar* untuk mulai"
      " menggunakan aplikasi."
  )
  st.markdown("</div>", unsafe_allow_html=True)
  st.stop()

# ---------------------------------------------------------
# MENU 1: MODUL AJAR (DENGAN PENGURANGAN KUOTA OTOMATIS)
# ---------------------------------------------------------
if menu_pilihan == "📖 Generator Modul Ajar":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">📜 Perangkat Pembelajaran</div>',
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
      "👤 Nama Guru & Gelar", placeholder="Contoh: Ahmad Fauzi, S.Pd."
  )
  nama_sekolah = st.text_input(
      "🏫 Nama Instansi / Sekolah", placeholder="Contoh: SMP Negeri 1 Nusantara"
  )
  mapel = st.text_input(
      "📚 Mata Pelajaran", placeholder="Contoh: Ilmu Pengetahuan Alam (IPA)"
  )
  kurikulum_aktif = st.text_input(
      "📑 Kurikulum", placeholder="Contoh: Kurikulum Merdeka"
  )
  fase_kelas = st.text_input(
      "🎓 Fase / Kelas", placeholder="Contoh: Fase D / Kelas VII"
  )
  nama_ks = st.text_input(
      "✍️ Nama Kepala Sekolah",
      placeholder="Contoh: Dra. Hj. Siti Aminah, M.Pd.",
  )
  topik = st.text_input(
      "💡 Topik / Materi Pokok", placeholder="Contoh: Sistem Pencernaan Manusia"
  )
  alokasi_waktu = st.text_input(
      "⏱️ Alokasi Waktu", placeholder="Contoh: 2 Pertemuan (4 x 40 Menit)"
  )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("🚀 Buat Modul Ajar Sekarang", type="primary"):
    with st.spinner(
        "Sistem AI Akademik sedang merancang Modul Ajar formal..."
    ):
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

        new_quota = sisa_kuota_sekarang - 1
        sheet.update_cell(row_index, 3, new_quota)

        st.success(
            f"Modul Ajar berhasil disusun! Sisa kuota Anda diperbarui di Sheets"
            f" jadi: {new_quota}x"
        )
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
# MENU 2: GENERATOR CP & ATP
# ---------------------------------------------------------
elif menu_pilihan == "🎯 Generator CP & ATP":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">🎯 Capaian & Alur Pembelajaran</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-title">🎯 Generator CP & ATP (Capaian Pembelajaran &'
      ' Alur Tujuan Pembelajaran)</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-desc">Susun Capaian Pembelajaran (CP) dan Alur'
      ' Tujuan Pembelajaran (ATP) resmi instansi pendidikan.</div>',
      unsafe_allow_html=True,
  )

  cp_guru = st.text_input(
      "👤 Nama Guru & Gelar",
      placeholder="Contoh: Ahmad Fauzi, S.Pd.",
      key="cp_guru",
  )
  cp_sekolah = st.text_input(
      "🏫 Nama Instansi / Sekolah",
      placeholder="Contoh: SMP Negeri 1 Nusantara",
      key="cp_sekolah",
  )
  cp_mapel = st.text_input(
      "📚 Mata Pelajaran",
      placeholder="Contoh: Ilmu Pengetahuan Alam (IPA)",
      key="cp_mapel",
  )
  cp_kur = st.text_input(
      "📑 Kurikulum", placeholder="Contoh: Kurikulum Merdeka", key="cp_kur"
  )
  cp_fase = st.text_input(
      "🎓 Fase / Kelas", placeholder="Contoh: Fase D / Kelas VII", key="cp_fase"
  )
  cp_ks = st.text_input(
      "✍️ Nama Kepala Sekolah",
      placeholder="Contoh: Dra. Hj. Siti Aminah, M.Pd.",
      key="cp_ks",
  )
  cp_materi = st.text_input(
      "💡 Lingkup Materi / Elemen",
      placeholder="Contoh: Pemahaman Sains & Keterampilan Proses",
      key="cp_materi",
  )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("🚀 Susun Dokumen CP & ATP", type="primary"):
    with st.spinner("Sistem Asesor Pintar sedang menyusun dokumen CP & ATP..."):
      try:
        prompt_cpatp = f"""
                Buatkan dokumen Capaian Pembelajaran (CP) dan Alur Tujuan Pembelajaran (ATP) formal, sangat terstruktur, dan profesional untuk instansi pendidikan:
                - Guru: {cp_guru}, Sekolah: {cp_sekolah}
                - Mapel: {cp_mapel}, Kurikulum: {cp_kur}
                - Fase/Kelas: {cp_fase}, Kepala Sekolah: {cp_ks}
                - Lingkup Materi/Elemen: {cp_materi}
                Sertakan komponen Identitas Instansi, Rasional, Capaian Pembelajaran (CP) per Elemen, serta Tabel Alur Tujuan Pembelajaran (ATP) yang merinci Tujuan Pembelajaran, Kelas/Semester, dan Estimasi Jam Pelajaran dalam bentuk tabel markdown standar lengkap menggunakan garis vertikal (|) untuk kolom dan barisnya.
                PENTING: Gunakan teks bersih murni tanpa tag HTML sama sekali (seperti <br> atau <p>).
                """
        response = generate_content_with_retry(
            client, "gemini-3.5-flash", prompt_cpatp
        )
        st.session_state.cpatp_hasil = response.text

        new_quota = sisa_kuota_sekarang - 1
        sheet.update_cell(row_index, 3, new_quota)

        st.success(
            "Dokumen CP & ATP berhasil disusun! Sisa kuota Anda diperbarui di"
            f" Sheets jadi: {new_quota}x"
        )
      except Exception as e:
        st.error(f"Error sistem: {e}")

  if st.session_state.cpatp_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Dokumen CP & ATP")
    st.markdown(st.session_state.cpatp_hasil)
    file_docx_cpatp = buat_file_docx(st.session_state.cpatp_hasil)
    st.download_button(
        "📥 Unduh Dokumen CP & ATP (.docx)",
        data=file_docx_cpatp,
        file_name=f"CP_ATP_{cp_mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# MENU 3: GENERATOR KKTP
# ---------------------------------------------------------
elif menu_pilihan == "📊 Generator KKTP":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">📊 Kriteria Ketercapaian</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-title">📊 Generator KKTP (Kriteria Ketercapaian'
      ' Tujuan Pembelajaran)</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-desc">Rancang instrumen dan rubrik Kriteria'
      ' Ketercapaian Tujuan Pembelajaran (KKTP) asesmen formatif & sumatif.</div>',
      unsafe_allow_html=True,
  )

  kktp_guru = st.text_input(
      "👤 Nama Guru & Gelar",
      placeholder="Contoh: Ahmad Fauzi, S.Pd.",
      key="kktp_guru",
  )
  kktp_sekolah = st.text_input(
      "🏫 Nama Instansi / Sekolah",
      placeholder="Contoh: SMP Negeri 1 Nusantara",
      key="kktp_sekolah",
  )
  kktp_mapel = st.text_input(
      "📚 Mata Pelajaran", placeholder="Contoh: Matematika", key="kktp_mapel"
  )
  kktp_kur = st.text_input(
      "📑 Kurikulum", placeholder="Contoh: Kurikulum Merdeka", key="kktp_kur"
  )
  kktp_fase = st.text_input(
      "🎓 Fase / Kelas", placeholder="Contoh: Fase D / Kelas VII", key="kktp_fase"
  )
  kktp_ks = st.text_input(
      "✍️ Nama Kepala Sekolah",
      placeholder="Contoh: Dra. Hj. Siti Aminah, M.Pd.",
      key="kktp_ks",
  )
  kktp_tujuan = st.text_input(
      "💡 Tujuan Pembelajaran",
      placeholder=(
          "Contoh: Memahami dan menyelesaikan persamaan linear satu variabel"
      ),
      key="kktp_tujuan",
  )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("🚀 Susun Dokumen KKTP", type="primary"):
    with st.spinner("Sistem Asesor Pintar sedang menyusun dokumen KKTP..."):
      try:
        prompt_kktp = f"""
                Buatkan dokumen Kriteria Ketercapaian Tujuan Pembelajaran (KKTP) formal, sangat terstruktur, dan profesional untuk instansi pendidikan:
                - Guru: {kktp_guru}, Sekolah: {kktp_sekolah}
                - Mapel: {kktp_mapel}, Kurikulum: {kktp_kur}
                - Fase/Kelas: {kktp_fase}, Kepala Sekolah: {kktp_ks}
                - Tujuan Pembelajaran: {kktp_tujuan}
                Sertakan komponen Identitas Instansi, Pendekatan KKTP (Deskripsi Kriteria, Rubrik Interval Nilai, atau Kriteria Ceklis), serta Tabel Interval Ketercapaian (Baru Berkembang, Layak, Cakap, Mahir) dalam bentuk tabel markdown standar lengkap menggunakan garis vertikal (|) untuk kolom dan barisnya.
                PENTING: Gunakan teks bersih murni tanpa tag HTML sama sekali (seperti <br> atau <p>).
                """
        response = generate_content_with_retry(
            client, "gemini-3.5-flash", prompt_kktp
        )
        st.session_state.kktp_hasil = response.text

        new_quota = sisa_kuota_sekarang - 1
        sheet.update_cell(row_index, 3, new_quota)

        st.success(
            "Dokumen KKTP berhasil disusun! Sisa kuota Anda diperbarui di Sheets"
            f" jadi: {new_quota}x"
        )
      except Exception as e:
        st.error(f"Error sistem: {e}")

  if st.session_state.kktp_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Dokumen KKTP")
    st.markdown(st.session_state.kktp_hasil)
    file_docx_kktp = buat_file_docx(st.session_state.kktp_hasil)
    st.download_button(
        "📥 Unduh Dokumen KKTP (.docx)",
        data=file_docx_kktp,
        file_name=f"KKTP_{kktp_mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# MENU 4: GENERATOR PROSEM & PROTA
# ---------------------------------------------------------
elif menu_pilihan == "📅 Generator Prosem & Prota":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">📅 Program Semester & Tahunan</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-title">📅 Generator Prosem & Prota (Program Semester'
      ' & Program Tahunan)</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="section-desc">Penyusunan Program Semester (Prosem) dan'
      ' Program Tahunan (Prota) resmi instansi pendidikan secara terintegrasi.</div>',
      unsafe_allow_html=True,
  )

  pr_guru = st.text_input(
      "👤 Nama Guru & Gelar",
      placeholder="Contoh: Ahmad Fauzi, S.Pd.",
      key="pr_guru",
  )
  pr_sekolah = st.text_input(
      "🏫 Nama Instansi / Sekolah",
      placeholder="Contoh: SMP Negeri 1 Nusantara",
      key="pr_sekolah",
  )
  pr_mapel = st.text_input(
      "📚 Mata Pelajaran",
      placeholder="Contoh: Ilmu Pengetahuan Alam (IPA)",
      key="pr_mapel",
  )
  pr_kur = st.text_input(
      "📑 Kurikulum", placeholder="Contoh: Kurikulum Merdeka", key="pr_kur"
  )
  pr_fase = st.text_input(
      "🎓 Fase / Kelas", placeholder="Contoh: Fase D / Kelas VII", key="pr_fase"
  )
  pr_ks = st.text_input(
      "✍️ Nama Kepala Sekolah",
      placeholder="Contoh: Dra. Hj. Siti Aminah, M.Pd.",
      key="pr_ks",
  )
  pr_Tahun = st.text_input(
      "📅 Tahun Pelajaran", placeholder="Contoh: 2026/2027", key="pr_Tahun"
  )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("🚀 Susun Dokumen Prosem & Prota", type="primary"):
    with st.spinner(
        "Sistem Asesor Pintar sedang menyusun dokumen Prosem & Prota..."
    ):
      try:
        prompt_prosemprota = f"""
                Buatkan dokumen Program Tahunan (Prota) dan Program Semester (Prosem) formal, sangat terstruktur, dan profesional untuk instansi pendidikan:
                - Guru: {pr_guru}, Sekolah: {pr_sekolah}
                - Mapel: {pr_mapel}, Kurikulum: {pr_kur}
                - Fase/Kelas: {pr_fase}, Kepala Sekolah: {pr_ks}, Tahun Pelajaran: {pr_Tahun}
                Sertakan komponen Identitas Instansi, Tabel Program Tahunan (Alokasi waktu per unit/bab), serta Tabel Program Semester (distribusi alokasi waktu per bulan dalam semester ganjil dan genap) dalam bentuk tabel markdown standar lengkap menggunakan garis vertikal (|) untuk kolom dan barisnya.
                PENTING: Gunakan teks bersih murni tanpa tag HTML sama sekali (seperti <br> atau <p>).
                """
        response = generate_content_with_retry(
            client, "gemini-3.5-flash", prompt_prosemprota
        )
        st.session_state.prosemprota_hasil = response.text

        new_quota = sisa_kuota_sekarang - 1
        sheet.update_cell(row_index, 3, new_quota)

        st.success(
            "Dokumen Prosem & Prota berhasil disusun! Sisa kuota Anda diperbarui"
            f" di Sheets jadi: {new_quota}x"
        )
      except Exception as e:
        st.error(f"Error sistem: {e}")

  if st.session_state.prosemprota_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Dokumen Prosem & Prota")
    st.markdown(st.session_state.prosemprota_hasil)
    file_docx_pr = buat_file_docx(st.session_state.prosemprota_hasil)
    st.download_button(
        "📥 Unduh Dokumen Prosem & Prota (.docx)",
        data=file_docx_pr,
        file_name=f"Prosem_Prota_{pr_mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# MENU 5: GENERATOR SOAL ASESMEN
# ---------------------------------------------------------
elif menu_pilihan == "📝 Generator Soal Asesmen":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">📋 Asesmen & Evaluasi</div>',
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

  s_guru = st.text_input(
      "👤 Nama Pembuat Asesmen",
      placeholder="Ahmad Fauzi, S.Pd.",
      key="s_guru",
  )
  s_sekolah = st.text_input(
      "🏫 Instansi / Sekolah",
      placeholder="SMP Negeri 1 Nusantara",
      key="s_sekolah",
  )
  s_mapel = st.text_input(
      "📚 Mata Pelajaran", placeholder="Matematika", key="s_mapel"
  )
  s_kur = st.text_input(
      "📑 Kurikulum", placeholder="Kurikulum Merdeka", key="s_kur"
  )
  s_kelas = st.text_input(
      "🎓 Kelas / Semester", placeholder="Kelas VII / Ganjil", key="s_kelas"
  )
  s_materi = st.text_input(
      "💡 Materi Asesmen",
      placeholder="Persamaan Linear Satu Variabel",
      key="s_materi",
  )
  s_komposisi = st.text_input(
      "📊 Komposisi Soal Asesmen",
      placeholder="5 Pilihan Ganda, 2 Isian Singkat, 1 Essai",
      key="s_komposisi",
  )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("🚀 Susun Naskah Soal Asesmen", type="primary"):
    with st.spinner("Sistem Asesor Pintar sedang menyusun naskah asesmen..."):
      try:
        prompt_soal = f"""
                Buatkan naskah soal asesmen resmi instansi pendidikan lengkap dengan Kop Soal, Petunjuk, Naskah Soal Asesmen (setiap nomor soal menggunakan penomoran tegas seperti 1., 2., 3. dan pilihan ganda ditulis tepat di bawah pertanyaan dengan format terindentasi A., B., C., D.), Kunci Jawaban, & Tabel Rubrik Penilaian.
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

        new_quota = sisa_kuota_sekarang - 1
        sheet.update_cell(row_index, 3, new_quota)

        st.success(
            "Paket soal asesmen berhasil disusun! Sisa kuota Anda diperbarui di"
            f" Sheets jadi: {new_quota}x"
        )
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
# MENU 6: SET KUNCI ACUAN
# ---------------------------------------------------------
elif menu_pilihan == "⚙️ Set Kunci Acuan":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">🎛️ Konfigurasi Penilaian</div>',
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
      "🎯 Kunci Pilihan Ganda (Cukup ketik huruf abjadnya saja secara berurutan,"
      " contoh: A,B,C,D,A atau ABCD)",
      value=st.session_state.kunci_pg,
      placeholder="Contoh: A B C D A B C D A B",
      height=100,
  )
  kunci_isian_input = st.text_area(
      "📝 Kunci Isian Singkat & Keyword Jawaban",
      value=st.session_state.kunci_isian,
      placeholder="Contoh: 1. y = 6; 2. x = 5",
      height=100,
  )
  kunci_essai_input = st.text_area(
      "📋 Rubrik Penilaian & Kunci Essai",
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
# MENU 7: KOREKSI SISWA
# ---------------------------------------------------------
elif menu_pilihan == "🔍 Koreksi Siswa":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">🔍 Verifikasi & Grading</div>',
      unsafe_allow_html=True,
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
    st.warning("⚠️ Harap simpan Kunci Jawaban terlebih dahulu pada Menu Lainnya!")
  else:
    nama_siswa = st.text_input(
        "👤 Identitas Siswa", placeholder="Masukkan Nama Lengkap / NISN Siswa"
    )
    metode_siswa = st.radio(
        "📂 Format Berkas Jawaban:",
        ["Unggah Foto / Scan", "Unggah Dokumen (PDF/Word)", "Ketik Teks"],
    )

    file_img, file_doc, teks_manual = None, None, ""
    if "Foto" in metode_siswa:
      file_img = st.file_uploader(
          "📸 Unggah berkas foto/scan", type=["jpg", "png", "jpeg"]
      )
    elif "Dokumen" in metode_siswa:
      file_doc = st.file_uploader("📎 Unggah berkas dokumen", type=["pdf", "docx"])
    else:
      teks_manual = st.text_area(
          "✏️ Masukkan teks jawaban siswa",
          placeholder="Ketik atau tempel lembar jawaban siswa di sini...",
      )

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Jalankan Analisis & Koreksi AI", type="primary"):
      if not nama_siswa:
        st.error("Masukkan identitas atau nama siswa!")
      else:
        with st.spinner("Sistem Asesor Pintar sedang menganalisis asesmen..."):
          try:
            payload = [f"""
                        Koreksi lembar jawaban siswa berdasarkan acuan berikut:
                        - Kunci Pilihan Ganda (Hanya huruf abjad jawaban benar): {st.session_state.kunci_pg}
                        - Kunci Isian Singkat: {st.session_state.kunci_isian}
                        - Rubrik Kunci Essai: {st.session_state.kunci_essai}
                        
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

            new_quota = sisa_kuota_sekarang - 1
            sheet.update_cell(row_index, 3, new_quota)

            st.success(
                f"Koreksi Selesai! Siswa **{nama_siswa}** mendapat Nilai:"
                f" **{skor}**. Sisa kuota Sheets: **{new_quota}x**"
            )
            with st.expander("📊 Lihat Rincian Analisis Penilaian"):
              st.markdown(hasil)
          except Exception as e:
            st.error(f"Error: {e}")
  st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# MENU 8: REKAP NILAI
# ---------------------------------------------------------
elif menu_pilihan == "📊 Rekap Nilai":
  st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="status-badge">📈 Rekapitulasi Data</div>',
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
        <p>🏛️ Portal Akademik Resmi Instansi Pendidikan | Developed by <b>Zeeo</b><br>
        WhatsApp Support: <a href="https://wa.me/6282371729760" target="_blank">082371729760</a><br>
        © 2026 All Rights Reserved</p>
    </div>
    """,
    unsafe_allow_html=True,
)
