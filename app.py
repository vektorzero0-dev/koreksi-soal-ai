from io import BytesIO
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

# Konfigurasi Halaman Instansi Pendidikan Formal & Responsif Mobile
st.set_page_config(
    page_title="Portal Akademik & Asisten Guru AI",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Custom Styling CSS: Desain Instansi Pendidikan Formal, Modern, Bersih, dan Profesional
st.markdown(
    """
    <style>
    /* Global Background Standar Instansi Korporat/Pendidikan */
    .stApp {
        background-color: #f1f5f9;
        color: #0f172a;
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* Header Resmi Instansi Pendidikan */
    .instansi-header {
        background-color: #1e3a8a;
        color: #ffffff;
        padding: 24px 28px;
        border-radius: 12px;
        margin-bottom: 24px;
        border-left: 6px solid #3b82f6;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .instansi-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #ffffff !important;
        margin: 0;
        letter-spacing: -0.3px;
    }
    .instansi-subtitle {
        font-size: 0.9rem;
        color: #cbd5e1 !important;
        margin-top: 6px;
        font-weight: 400;
        line-height: 1.4;
    }

    /* Kotak Kartu Konten Bersih & Elegan */
    .card-box {
        background: #ffffff;
        padding: 24px;
        border-radius: 12px;
        border: 1px solid #cbd5e1;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.03);
        margin-bottom: 20px;
    }

    /* Label Input Tegas & Profesional */
    label, .stTextInput label, .stTextArea label, .stSelectbox label, .stRadio label {
        color: #1e293b !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
    }

    /* Tombol Navigasi & Aksi Utama */
    .stButton>button {
        width: 100%;
        background-color: #1e3a8a;
        color: #ffffff;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.7rem 1rem;
        border: none;
        font-size: 0.95rem;
        transition: background-color 0.2s ease;
    }
    .stButton>button:hover {
        background-color: #1d4ed8;
        color: #ffffff;
    }

    /* Footer Instansi Formal */
    .footer {
        text-align: center;
        padding: 16px;
        font-size: 0.82rem;
        color: #64748b;
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        margin-top: 30px;
        margin-bottom: 20px;
        font-weight: 500;
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

# --- HEADER UTAMA INSTANSI PENDIDIKAN ---
st.markdown(
    """
    <div class="instansi-header">
        <h1 class="instansi-title">🏛️ PORTAL ASISTEN AKADEMIK & GURU</h1>
        <p class="instansi-subtitle">Sistem Terintegrasi Kecerdasan Buatan untuk Penyusunan Perangkat Pembelajaran, Bank Soal, dan Penilaian Objektif.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("🔑 Pengaturan API Key & Keamanan", expanded=not GEMINI_API_KEY):
  if not GEMINI_API_KEY:
    st.warning("⚠️ API Key belum terdeteksi di Secrets.")
    GEMINI_API_KEY = st.text_input(
        "Masukkan Gemini API Key", type="password", key="sidebar_key"
    )
  else:
    st.success("🔒 Sistem Keamanan Aktif & Terhubung")

if not GEMINI_API_KEY:
  st.warning(
      "Mohon masukkan Google Gemini API Key di atas untuk mulai menggunakan"
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

# Menu Navigasi Dropdown Formal & Ramah Android
st.markdown("### 📌 Pilih Menu Layanan Akademik:")
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
  st.markdown(
      "Penyusunan perangkat pembelajaran komprehensif berbasis kurikulum"
      " nasional."
  )

  nama_guru = st.text_input("Nama Guru & Gelar", "Ahmad Fauzi, S.Pd.")
  nama_sekolah = st.text_input("Nama Instansi / Sekolah", "SMP Negeri 1 Nusantara")
  mapel = st.text_input("Mata Pelajaran", "Ilmu Pengetahuan Alam (IPA)")
  kurikulum_aktif = st.text_input("Kurikulum", "Kurikulum Merdeka")
  fase_kelas = st.text_input("Fase / Kelas", "Fase D / Kelas VII")
  nama_ks = st.text_input("Nama Kepala Sekolah", "Dra. Hj. Siti Aminah, M.Pd.")
  topik = st.text_input("Topik / Materi Pokok", "Sistem Pencernaan Manusia")
  alokasi_waktu = st.text_input("Alokasi Waktu", "2 Pertemuan (4 x 40 Menit)")

  if st.button("🚀 Proses Pembuatan Modul Ajar", type="primary"):
    with st.spinner("Sistem Gemini 3.5 Flash sedang merancang Modul Ajar..."):
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
            model="gemini-3.5-flash", contents=prompt_modul
        )
        st.session_state.modul_hasil = response.text
        st.success("Modul Ajar berhasil disusun!")
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
  st.markdown("Pembuatan naskah ujian resmi lengkap dengan rubrik skor.")

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

  if st.button("🚀 Proses Pembuatan Paket Soal", type="primary"):
    with st.spinner("Sistem Gemini 3.5 Flash sedang menyusun naskah ujian..."):
      try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt_soal = f"""
                Buatkan naskah soal ujian resmi instansi pendidikan lengkap dengan Kop Ujian, Petunjuk, Naskah Soal, Kunci Jawaban, & Rubrik Penilaian (Tabel) untuk:
                - Guru: {s_guru}, Sekolah: {s_sekolah}, Mapel: {s_mapel}
                - Kurikulum: {s_kur}, Kelas: {s_kelas}, Materi: {s_materi}
                - Komposisi: {s_komposisi}
                """
        response = client.models.generate_content(
            model="gemini-3.5-flash", contents=prompt_soal
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

# 3. SET KUNCI ACUAN
elif menu_pilihan == "⚙️ 3. Set Kunci Acuan":
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("⚙️ Konfigurasi Kunci Acuan")
  st.markdown("Pemisahan parameter kunci untuk akurasi penilaian.")

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

  if st.button("💾 Simpan Kunci Acuan", type="primary"):
    st.session_state.kunci_pg = kunci_pg_input
    st.session_state.kunci_isian = kunci_isian_input
    st.session_state.kunci_essai = kunci_essai_input
    st.success("Kunci acuan berhasil disimpan!")
  st.markdown("</div>", unsafe_allow_html=True)

# 4. KOREKSI SISWA
elif menu_pilihan == "🔍 4. Koreksi Siswa":
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("🔍 Koreksi Lembar Jawaban Siswa")
  if (
      not st.session_state.kunci_pg
      and not st.session_state.kunci_isian
      and not st.session_state.kunci_essai
  ):
    st.warning("⚠️ Harap simpan Kunci Jawaban terlebih dahulu di menu nomor 3!")
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

    if st.button("🚀 Jalankan Koreksi AI", type="primary"):
      if not nama_siswa:
        st.error("Masukkan nama atau identitas siswa!")
      else:
        with st.spinner("Sistem Gemini 3.5 Flash sedang menganalisis jawaban..."):
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
                model="gemini-3.5-flash", contents=payload
            )
            hasil = resp.text

            import re

            m = re.search(r"NILAI_AKHIR:\s*([0-9.]+)", hasil)
            skor = float(m.group(1)) if m else 0.0

            st.session_state.rekap_nilai.append(
                {"Nama Siswa": nama_siswa, "Nilai Akhir": skor, "Detail": hasil}
            )
            st.success(
                f"Koreksi Selesai! Siswa **{nama_siswa}** mendapat Nilai:"
                f" **{skor}**"
            )
            with st.expander("Lihat Rincian Analisis"):
              st.markdown(hasil)
          except Exception as e:
            st.error(f"Error: {e}")
  st.markdown("</div>", unsafe_allow_html=True)

# 5. REKAP NILAI
elif menu_pilihan == "📊 5. Rekap Nilai":
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📊 Rekapitulasi Nilai Akademik")
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
    if st.button("🗑️ Kosongkan Rekap"):
      st.session_state.rekap_nilai = []
      st.rerun()
  st.markdown("</div>", unsafe_allow_html=True)

# --- FOOTER ---
st.markdown(
    """
    <div class="footer">
        <p>Sistem Akademik Terintegrasi | Developed by <b>Zeeo</b><br>WhatsApp: <a href="https://wa.me/6282371729760" target="_blank">082371729760</a><br>© 2026 All Rights Reserved</p>
    </div>
    """,
    unsafe_allow_html=True,
)
