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

# Konfigurasi Halaman Instansi Pendidikan
st.set_page_config(
    page_title="Sistem Akademik & Asisten Guru AI",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling CSS: Modern, Profesional, Khas Instansi Pendidikan (Navy & Soft Clean)
st.markdown(
    """
    <style>
    /* Global Background Profesional */
    .stApp {
        background-color: #f1f5f9;
        color: #1e293b;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Header Utama Khas Instansi */
    .instansi-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        color: white;
        padding: 30px;
        border-radius: 12px;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .instansi-title {
        font-size: 2.2rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .instansi-subtitle {
        font-size: 1.05rem;
        color: #93c5fd;
        margin-top: 8px;
        font-weight: 400;
    }

    /* Kotak Kartu Konten Modern (Card Layout) */
    .card-box {
        background: #ffffff;
        padding: 28px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
        margin-bottom: 20px;
        transition: all 0.3s ease;
    }
    .card-box:hover {
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
        border-color: #cbd5e1;
    }

    /* Tombol Utama Khas Korporat/Instansi */
    .stButton>button {
        background-color: #1e3a8a;
        color: #ffffff;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.55rem 1.2rem;
        border: none;
        transition: all 0.2s ease;
        box-shadow: 0 2px 4px rgba(30, 58, 138, 0.2);
    }
    .stButton>button:hover {
        background-color: #1d4ed8;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(29, 78, 216, 0.3);
    }

    /* Sidebar Profesional */
    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }

    /* Footer Instansi */
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #ffffff;
        color: #475569;
        text-align: center;
        padding: 12px;
        font-size: 0.85rem;
        border-top: 1px solid #e2e8f0;
        box-shadow: 0 -4px 6px -1px rgba(0, 0, 0, 0.02);
        z-index: 1000;
    }
    .footer a {
        color: #1e3a8a;
        text-decoration: none;
        font-weight: 600;
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
  st.image("https://img.icons8.com/color/96/school.png", width=65)
  st.title("Panel Akademik AI")
  st.markdown("---")
  if not GEMINI_API_KEY:
    st.warning("⚠️ API Key belum diatur di Secrets.")
    GEMINI_API_KEY = st.text_input(
        "Masukkan Gemini API Key", type="password", key="sidebar_key"
    )
  else:
    st.success("🔒 Sistem Keamanan Aktif")
  st.markdown("---")
  st.markdown(
      "**Modul Layanan:**\n1. Generator Modul Ajar\n2. Generator Soal & Kunci\n3."
      " Set Kunci Acuan\n4. Koreksi Lembar Siswa\n5. Rekapitulasi Nilai"
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

# Header Instansi Pendidikan Modern
st.markdown(
    """
    <div class="instansi-header">
        <p class="instansi-title">🏛️ PORTAL ASISTEN AKADEMIK & GURU PROFESIONAL</p>
        <p class="instansi-subtitle">Sistem Terintegrasi Kecerdasan Buatan untuk Penyusunan Perangkat Pembelajaran, Bank Soal, dan Penilaian Objektif.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📖 1. Generator Modul Ajar",
    "📝 2. Generator Soal & Kunci",
    "⚙️ 3. Set Kunci Acuan",
    "🔍 4. Koreksi Siswa",
    "📊 5. Rekap Nilai",
])

# --- TAB 1: MODUL AJAR ---
with tab1:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📖 Generator Modul Ajar / RPP Formal")
  st.markdown(
      "Penyusunan perangkat pembelajaran komprehensif berbasis standar kurikulum"
      " nasional."
  )
  col_m1, col_m2 = st.columns(2)
  with col_m1:
    nama_guru = st.text_input("Nama Guru & Gelar", "Ahmad Fauzi, S.Pd.")
    nama_sekolah = st.text_input("Nama Instansi / Sekolah", "SMP Negeri 1 Nusantara")
    mapel = st.text_input("Mata Pelajaran", "Ilmu Pengetahuan Alam (IPA)")
    kurikulum_aktif = st.text_input("Kurikulum", "Kurikulum Merdeka")
  with col_m2:
    fase_kelas = st.text_input("Fase / Kelas", "Fase D / Kelas VII")
    nama_ks = st.text_input("Nama Kepala Sekolah", "Dra. Hj. Siti Aminah, M.Pd.")
    topik = st.text_input("Topik / Materi Pokok", "Sistem Pencernaan Manusia")
    alokasi_waktu = st.text_input("Alokasi Waktu", "2 Pertemuan (4 x 40 Menit)")

  if st.button("🚀 Proses Pembuatan Modul Ajar", type="primary"):
    with st.spinner("Sistem sedang merancang Modul Ajar formal..."):
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

# --- TAB 2: SOAL & KUNCI ---
with tab2:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📝 Generator Paket Soal & Kunci Jawaban Resmi")
  st.markdown(
      "Pembuatan naskah ujian berstandar asesmen nasional lengkap dengan rubrik"
      " skor."
  )
  col_s1, col_s2 = st.columns(2)
  with col_s1:
    s_guru = st.text_input(
        "Nama Pembuat Soal", "Ahmad Fauzi, S.Pd.", key="sg"
    )
    s_sekolah = st.text_input(
        "Instansi / Sekolah", "SMP Negeri 1 Nusantara", key="ss"
    )
    s_mapel = st.text_input("Mata Pelajaran", "Matematika", key="sm")
    s_kur = st.text_input("Kurikulum", "Kurikulum Merdeka", key="sk")
  with col_s2:
    s_kelas = st.text_input("Kelas / Semester", "Kelas VII / Ganjil", key="skel")
    s_materi = st.text_input(
        "Materi / Bab Ujian", "Persamaan Linear Satu Variabel", key="smat"
    )
    s_komposisi = st.text_input(
        "Komposisi Soal",
        "5 Pilihan Ganda, 2 Isian Singkat, 1 Essai",
        key="skom",
    )

  if st.button("🚀 Proses Pembuatan Paket Soal", type="primary"):
    with st.spinner("Sistem sedang menyusun naskah ujian & kunci jawaban..."):
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
        st.success("Paket soal berhasil disusun!")
      except Exception as e:
        st.error(f"Error sistem: {e}")

  if st.session_state.soal_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Naskah Soal & Kunci")
    st.markdown(st.session_state.soal_hasil)
    file_docx_soal = buat_file_docx(st.session_state.soal_hasil)
    st.download_button(
        "📥 Unduh Paket Soal (.docx)",
        data=file_docx_soal,
        file_name=f"Soal_{s_mapel}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)

# --- TAB 3: SET KUNCI TERPISAH ---
with tab3:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("⚙️ Konfigurasi Kunci Jawaban Acuan Penilaian")
  st.markdown(
      "Pemisahan parameter kunci jawaban untuk mendukung akurasi penilaian"
      " otomatis."
  )
  col_k1, col_k2 = st.columns(2)
  with col_k1:
    kunci_pg_input = st.text_area(
        "Kunci Pilihan Ganda (Contoh: 1.A, 2.B, 3.C)",
        value=st.session_state.kunci_pg,
        height=110,
    )
    kunci_isian_input = st.text_area(
        "Kunci Isian Singkat & Keyword",
        value=st.session_state.kunci_isian,
        height=110,
    )
  with col_k2:
    kunci_essai_input = st.text_area(
        "Rubrik Penilaian & Kunci Essai",
        value=st.session_state.kunci_essai,
        height=250,
    )

  if st.button("💾 Simpan Kunci Acuan ke Sistem", type="primary"):
    st.session_state.kunci_pg = kunci_pg_input
    st.session_state.kunci_isian = kunci_isian_input
    st.session_state.kunci_essai = kunci_essai_input
    st.success("Parameter kunci jawaban berhasil disimpan dalam memori sistem!")
  st.markdown("</div>", unsafe_allow_html=True)

# --- TAB 4: KOREKSI SISWA ---
with tab4:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("🔍 Sistem Koreksi Lembar Jawaban Siswa (AI Vision & Dokumen)")
  if (
      not st.session_state.kunci_pg
      and not st.session_state.kunci_isian
      and not st.session_state.kunci_essai
  ):
    st.warning(
        "⚠️ Harap tentukan dan simpan Kunci Jawaban terlebih dahulu di Tab 3!"
    )
  else:
    nama_siswa = st.text_input("Nama Lengkap / Nomor Induk Siswa")
    metode_siswa = st.radio(
        "Pilih Format Berkas Jawaban:",
        [
            "Unggah Foto / Scan Lembar Jawab",
            "Unggah Dokumen (PDF / Word)",
            "Ketik Teks Jawaban Langsung",
        ],
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

    if st.button("🚀 Jalankan Proses Koreksi AI", type="primary"):
      if not nama_siswa:
        st.error("Mohon masukkan nama atau identitas siswa!")
      else:
        with st.spinner(
            "Sistem AI sedang menganalisis jawaban berdasarkan rubrik acuan..."
        ):
          try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            payload = [f"""
                        Bertindaklah sebagai tim penguji akademik instansi pendidikan yang objektif. Koreksi lembar jawaban siswa berdasarkan acuan berikut:
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
              payload.append(f"Berkas Jawaban Siswa: {ext_text}")
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
                f"Koreksi Selesai! Siswa **{nama_siswa}** memperoleh Nilai Akhir:"
                f" **{skor}**"
            )
            with st.expander("Lihat Rincian Analisis Koreksi"):
              st.markdown(hasil)
          except Exception as e:
            st.error(f"Terjadi kendala sistem: {e}")
  st.markdown("</div>", unsafe_allow_html=True)

# --- TAB 5: REKAP NILAI ---
with tab5:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📊 Rekapitulasi Nilai Akademik Kelas")
  if len(st.session_state.rekap_nilai) == 0:
    st.info("Belum ada data nilai siswa yang terekam dalam sesi ini.")
  else:
    df = pd.DataFrame(st.session_state.rekap_nilai)
    st.dataframe(df[["Nama Siswa", "Nilai Akhir"]], use_container_width=True)
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Unduh Rekap Nilai Akademik (CSV)",
        data=csv,
        file_name="rekap_nilai_akademik.csv",
        mime="text/csv",
    )
    if st.button("🗑️ Kosongkan Data Rekap"):
      st.session_state.rekap_nilai = []
      st.rerun()
  st.markdown("</div>", unsafe_allow_html=True)

# --- FOOTER ---
st.markdown(
    """
    <div class="footer">
        <p>Sistem Akademik Terintegrasi | Developed by <b>Zeeo</b> | WhatsApp: <a href="https://wa.me/6282371729760" target="_blank">082371729760</a> | © 2026 All Rights Reserved</p>
    </div>
    """,
    unsafe_allow_html=True,
)
