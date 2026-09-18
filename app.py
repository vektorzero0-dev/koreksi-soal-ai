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

# Konfigurasi Halaman
st.set_page_config(
    page_title="Platform Asisten Guru AI Pro",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling CSS: Tema Dark Mode Profesional (Adem di Mata, Elegan, Kontras Tinggi)
st.markdown(
    """
    <style>
    /* Global Background Gelap Profesional (Arang / Charcoal Dark) */
    .stApp {
        background-color: #0f172a;
        color: #f8fafc;
    }
    
    /* Header Utama */
    .main-title {
        font-size: 2.2rem;
        color: #f1f5f9;
        font-weight: 700;
        margin-bottom: 0px;
    }
    .sub-title {
        font-size: 1rem;
        color: #94a3b8;
        margin-bottom: 25px;
    }

    /* Kotak Kartu Konten (Dark Card dengan Border Tipis Elegan) */
    .content-card {
        background: #1e293b;
        padding: 24px;
        border-radius: 10px;
        border: 1px solid #334155;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
    }

    /* Warna Teks & Label di dalam Streamlit agar kontras di Dark Mode */
    .stTextInput label, .stTextArea label, .stSelectbox label, .stRadio label, .stFileUploader label {
        color: #e2e8f0 !important;
        font-weight: 500;
    }

    /* Tombol Utama */
    .stButton>button {
        background-color: #3b82f6;
        color: #ffffff;
        font-weight: 600;
        border-radius: 6px;
        border: none;
        padding: 0.5rem 1.2rem;
        transition: background-color 0.2s ease;
    }
    .stButton>button:hover {
        background-color: #2563eb;
        color: #ffffff;
    }

    /* Sidebar Gelap */
    section[data-testid="stSidebar"] {
        background-color: #111827;
        border-right: 1px solid #1f2937;
    }
    section[data-testid="stSidebar"] * {
        color: #f3f4f6 !important;
    }

    /* Footer Gelap Melayang */
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #0b0f19;
        color: #9ca3af;
        text-align: center;
        padding: 10px;
        font-size: 0.85rem;
        border-top: 1px solid #1f2937;
        z-index: 1000;
    }
    .footer a {
        color: #60a5fa;
        text-decoration: none;
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

with st.sidebar:
  st.title("Panel Guru AI")
  st.markdown("---")
  if not GEMINI_API_KEY:
    st.warning("API Key belum diatur di Secrets.")
    GEMINI_API_KEY = st.text_input(
        "Masukkan Gemini API Key", type="password", key="sidebar_key"
    )
  else:
    st.success("API Key Terhubung Aman")
  st.markdown("---")
  st.markdown(
      "**Menu Utama:**\n1. Generator Modul Ajar\n2. Generator Soal & Kunci\n3."
      " Set Kunci Terpisah\n4. Koreksi Siswa\n5. Rekap Nilai"
  )

if not GEMINI_API_KEY:
  st.warning("Mohon masukkan Google Gemini API Key terlebih dahulu di sidebar.")
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

st.markdown(
    '<p class="main-title">🎓 Platform Asisten Cerdas Guru (AI)</p>',
    unsafe_allow_html=True,
)
st.markdown(
    '<p class="sub-title">Solusi administrasi pembelajaran dan koreksi ujian'
    " otomatis berstandar profesional.</p>",
    unsafe_allow_html=True,
)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📖 1. Modul Ajar",
    "📝 2. Soal & Kunci",
    "⚙️ 3. Set Kunci Acuan",
    "🔍 4. Koreksi Siswa",
    "📊 5. Rekap Nilai",
])

# --- TAB 1: MODUL AJAR ---
with tab1:
  st.markdown('<div class="content-card">', unsafe_allow_html=True)
  st.header("Generator Modul Ajar / RPP")
  col_m1, col_m2 = st.columns(2)
  with col_m1:
    nama_guru = st.text_input("Nama Guru & Gelar", "Ahmad Fauzi, S.Pd.")
    nama_sekolah = st.text_input("Nama Sekolah", "SMP Negeri 1 Nusantara")
    mapel = st.text_input("Mata Pelajaran", "IPA")
    kurikulum_aktif = st.text_input("Kurikulum", "Kurikulum Merdeka")
  with col_m2:
    fase_kelas = st.text_input("Fase / Kelas", "Fase D / Kelas VII")
    nama_ks = st.text_input("Nama Kepala Sekolah", "Dra. Siti Aminah, M.Pd.")
    topik = st.text_input("Topik / Materi", "Sistem Pencernaan")
    alokasi_waktu = st.text_input("Alokasi Waktu", "2 Pertemuan (4 x 40 Menit)")

  if st.button("Buat Modul Ajar", type="primary"):
    with st.spinner("Memproses modul ajar..."):
      try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt_modul = f"""
                Buatkan Modul Ajar / RPP lengkap, terstruktur, dan rapi untuk:
                - Guru: {nama_guru}, Sekolah: {nama_sekolah}
                - Mapel: {mapel}, Kurikulum: {kurikulum_aktif}
                - Kelas: {fase_kelas}, Kepala Sekolah: {nama_ks}
                - Topik: {topik}, Waktu: {alokasi_waktu}
                Sertakan komponen Identitas, Tujuan Pembelajaran, Kegiatan, dan Tabel Asesmen.
                """
        response = client.models.generate_content(
            model="gemini-2.5-flash", contents=prompt_modul
        )
        st.session_state.modul_hasil = response.text
        st.success("Berhasil dibuat!")
      except Exception as e:
        st.error(f"Error: {e}")

  if st.session_state.modul_hasil:
    st.markdown("---")
    st.markdown(st.session_state.modul_hasil)
    file_docx = buat_file_docx(st.session_state.modul_hasil)
    st.download_button(
        "Unduh Modul Ajar (.docx)",
        data=file_docx,
        file_name=f"Modul_Ajar_{mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# --- TAB 2: SOAL & KUNCI ---
with tab2:
  st.markdown('<div class="content-card">', unsafe_allow_html=True)
  st.header("Generator Soal & Kunci Jawaban")
  col_s1, col_s2 = st.columns(2)
  with col_s1:
    s_guru = st.text_input("Nama Guru", "Ahmad Fauzi, S.Pd.", key="sg")
    s_sekolah = st.text_input("Sekolah", "SMP Negeri 1 Nusantara", key="ss")
    s_mapel = st.text_input("Mata Pelajaran", "Matematika", key="sm")
    s_kur = st.text_input("Kurikulum", "Kurikulum Merdeka", key="sk")
  with col_s2:
    s_kelas = st.text_input("Kelas", "Kelas VII", key="skel")
    s_materi = st.text_input("Materi", "Persamaan Linear", key="smat")
    s_komposisi = st.text_input(
        "Komposisi", "5 Pilihan Ganda, 2 Isian, 1 Essai", key="skom"
    )

  if st.button("Buat Paket Soal", type="primary"):
    with st.spinner("Memproses soal..."):
      try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt_soal = f"""
                Buatkan paket soal ujian lengkap dengan Kunci Jawaban & Rubrik Penilaian (Tabel) untuk:
                - Guru: {s_guru}, Sekolah: {s_sekolah}, Mapel: {s_mapel}
                - Kurikulum: {s_kur}, Kelas: {s_kelas}, Materi: {s_materi}
                - Komposisi: {s_komposisi}
                """
        response = client.models.generate_content(
            model="gemini-2.5-flash", contents=prompt_soal
        )
        st.session_state.soal_hasil = response.text
        st.success("Soal berhasil dibuat!")
      except Exception as e:
        st.error(f"Error: {e}")

  if st.session_state.soal_hasil:
    st.markdown("---")
    st.markdown(st.session_state.soal_hasil)
    file_docx_soal = buat_file_docx(st.session_state.soal_hasil)
    st.download_button(
        "Unduh Paket Soal (.docx)",
        data=file_docx_soal,
        file_name=f"Soal_{s_mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# --- TAB 3: SET KUNCI TERPISAH ---
with tab3:
  st.markdown('<div class="content-card">', unsafe_allow_html=True)
  st.header("Set Kunci Jawaban Acuan Koreksi")
  col_k1, col_k2 = st.columns(2)
  with col_k1:
    kunci_pg_input = st.text_area(
        "Kunci Pilihan Ganda (Contoh: 1.A, 2.B)",
        value=st.session_state.kunci_pg,
        height=100,
    )
    kunci_isian_input = st.text_area(
        "Kunci Isian Singkat",
        value=st.session_state.kunci_isian,
        height=100,
    )
  with col_k2:
    kunci_essai_input = st.text_area(
        "Rubrik & Kunci Essai",
        value=st.session_state.kunci_essai,
        height=240,
    )

  if st.button("Simpan Kunci Jawaban", type="primary"):
    st.session_state.kunci_pg = kunci_pg_input
    st.session_state.kunci_isian = kunci_isian_input
    st.session_state.kunci_essai = kunci_essai_input
    st.success("Kunci jawaban berhasil disimpan!")
  st.markdown("</div>", unsafe_allow_html=True)

# --- TAB 4: KOREKSI SISWA ---
with tab4:
  st.markdown('<div class="content-card">', unsafe_allow_html=True)
  st.header("Koreksi Lembar Jawaban Siswa")
  if (
      not st.session_state.kunci_pg
      and not st.session_state.kunci_isian
      and not st.session_state.kunci_essai
  ):
    st.warning("Harap isi Kunci Jawaban terlebih dahulu di Tab 3!")
  else:
    nama_siswa = st.text_input("Nama Siswa")
    metode_siswa = st.radio(
        "Metode Jawaban:", ["Unggah Foto", "Unggah Dokumen (PDF/Word)", "Ketik"]
    )

    file_img, file_doc, teks_manual = None, None, ""
    if "Foto" in metode_siswa:
      file_img = st.file_uploader("Foto Lembar Jawab", type=["jpg", "png"])
    elif "Dokumen" in metode_siswa:
      file_doc = st.file_uploader("Dokumen Jawab", type=["pdf", "docx"])
    else:
      teks_manual = st.text_area("Tulis Jawaban Siswa")

    if st.button("Proses Koreksi", type="primary"):
      if not nama_siswa:
        st.error("Masukkan nama siswa!")
      else:
        with st.spinner("Mengoreksi..."):
          try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            payload = [f"""
                        Koreksi jawaban siswa berdasarkan kunci berikut:
                        - PG: {st.session_state.kunci_pg}
                        - Isian: {st.session_state.kunci_isian}
                        - Essai: {st.session_state.kunci_essai}
                        
                        Format baris pertama WAJIB:
                        NILAI_AKHIR: [Angka 0-100]
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
              payload.append(f"Jawaban: {ext_text}")
            elif teks_manual:
              payload.append(f"Jawaban: {teks_manual}")

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
                f"Selesai! Nilai {nama_siswa}: **{skor}**"
            )
            with st.expander("Detail Koreksi"):
              st.markdown(hasil)
          except Exception as e:
            st.error(f"Error: {e}")
  st.markdown("</div>", unsafe_allow_html=True)

# --- TAB 5: REKAP NILAI ---
with tab5:
  st.markdown('<div class="content-card">', unsafe_allow_html=True)
  st.header("Rekapitulasi Nilai")
  if len(st.session_state.rekap_nilai) == 0:
    st.info("Belum ada data nilai.")
  else:
    df = pd.DataFrame(st.session_state.rekap_nilai)
    st.dataframe(df[["Nama Siswa", "Nilai Akhir"]], use_container_width=True)
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Unduh Rekap (CSV)",
        data=csv,
        file_name="rekap_nilai.csv",
        mime="text/csv",
    )
    if st.button("Hapus Rekap"):
      st.session_state.rekap_nilai = []
      st.rerun()
  st.markdown("</div>", unsafe_allow_html=True)

# --- FOOTER ---
st.markdown(
    """
    <div class="footer">
        <p>Developed by <b>Zeeo</b> | WhatsApp: <a href="https://wa.me/6282371729760" target="_blank">082371729760</a> | © 2026 All Rights Reserved</p>
    </div>
    """,
    unsafe_allow_html=True,
)
