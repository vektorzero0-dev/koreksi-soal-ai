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

# Konfigurasi Halaman & Tema Profesional
st.set_page_config(
    page_title="Platform Pintar Guru AI Pro",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling CSS: Modern, Clean, Solid Color, Tanpa Gradiasi, dengan Efek Animasi Halus
st.markdown(
    """
    <style>
    /* Global Background & Font */
    .stApp {
        background-color: #f4f6f9;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Header Utama Clean & Modern dengan Animasi Masuk */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(-10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .main-header {
        font-size: 2.4rem;
        color: #0f172a;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin-bottom: 0px;
        animation: fadeIn 0.6s ease-out;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #64748b;
        margin-bottom: 25px;
        font-weight: 400;
        animation: fadeIn 0.8s ease-out;
    }

    /* Kotak-kotak Kartu (Card Layout) dengan Efek Animasi Hover & Transisi */
    .card-box {
        background: #ffffff;
        padding: 28px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05), 0 1px 2px 0 rgba(0, 0, 0, 0.03);
        margin-bottom: 20px;
        transition: transform 0.25s ease, box-shadow 0.25s ease;
        animation: fadeIn 0.5s ease-out;
    }
    .card-box:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.08), 0 8px 10px -6px rgba(0, 0, 0, 0.05);
        border-color: #cbd5e1;
    }

    /* Tombol Interaktif dengan Animasi Halus (Tanpa Gradiasi) */
    .stButton>button {
        background-color: #0f172a;
        color: #ffffff;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.55rem 1.2rem;
        border: none;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
    }
    .stButton>button:hover {
        background-color: #1e293b;
        color: #ffffff;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.15);
    }

    /* Styling Sidebar Profesional */
    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }

    /* Footer Profesional Fixed dengan Animasi */
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #ffffff;
        color: #475569;
        text-align: center;
        padding: 10px;
        font-size: 0.85rem;
        border-top: 1px solid #e2e8f0;
        box-shadow: 0 -4px 6px -1px rgba(0, 0, 0, 0.02);
        z-index: 1000;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# Fungsi untuk membuat file Word (.docx) dari teks AI
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
  st.image(
      "https://img.icons8.com/color/96/artificial-intelligence.png", width=65
  )
  st.title("Panel Guru AI Pro")
  st.markdown("---")
  if not GEMINI_API_KEY:
    st.warning("⚠️ API Key belum diatur di Secrets.")
    GEMINI_API_KEY = st.text_input(
        "Masukkan Gemini API Key", type="password", key="sidebar_key"
    )
  else:
    st.success("🔒 API Key Terhubung Aman")
  st.markdown("---")
  st.markdown(
      "**Fitur Sistem:**\n1. Generator Modul Ajar\n2. Generator Soal &\n"
      "   Kunci\n3. Set Kunci Terpisah (PG, Isian, Essai)\n4. Koreksi Cerdas AI\n5."
      " Rekap Nilai Kelas"
  )

if not GEMINI_API_KEY:
  st.warning("Mohon masukkan Google Gemini API Key terlebih dahulu di sidebar.")
  st.stop()

# Inisialisasi State Session
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

# Judul Utama Aplikasi
st.markdown(
    '<p class="main-header">🎓 Platform Asisten Cerdas Guru (AI)</p>',
    unsafe_allow_html=True,
)
st.markdown(
    '<p class="sub-header">Solusi administrasi pembelajaran modern, generator'
    " soal terstruktur, dan koreksi otomatis berstandar AI.</p>",
    unsafe_allow_html=True,
)

# Navigasi Tab Utama
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📖 1. Generator Modul Ajar",
    "📝 2. Generator Soal & Kunci",
    "⚙️ 3. Set Kunci Jawaban Terpisah",
    "🔍 4. Koreksi Lembar Siswa",
    "📊 5. Rekap Nilai Kelas",
])


# --- TAB 1: GENERATOR MODUL AJAR ---
with tab1:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📖 Generator Modul Ajar / RPP Lengkap")
  st.markdown("Isi identitas lengkap untuk menghasilkan perangkat ajar formal.")

  col_m1, col_m2 = st.columns(2)
  with col_m1:
    nama_guru = st.text_input("Nama Guru & Gelar", "Ahmad Fauzi, S.Pd.")
    nama_sekolah = st.text_input("Nama Sekolah", "SMP Negeri 1 Nusantara")
    mapel = st.text_input("Mata Pelajaran", "Ilmu Pengetahuan Alam (IPA)")
    pilihan_kurikulum = st.selectbox(
        "Pilih Kurikulum",
        [
            "Kurikulum Merdeka",
            "Kurikulum 2013 (K-13)",
            "Kurikulum Darurat",
            "Lainnya (Ketik Manual)",
        ],
    )
    if pilihan_kurikulum == "Lainnya (Ketik Manual)":
      kurikulum_aktif = st.text_input(
          "Masukkan Nama Kurikulum Lainnya", "Kurikulum Pesantren Terpadu"
      )
    else:
      kurikulum_aktif = pilihan_kurikulum

  with col_m2:
    fase_kelas = st.text_input("Fase / Kelas", "Fase D / Kelas VII")
    nama_ks = st.text_input("Nama Kepala Sekolah", "Dra. Hj. Siti Aminah, M.Pd.")
    topik = st.text_input("Topik / Materi Pokok", "Sistem Pencernaan Manusia")
    alokasi_waktu = st.text_input("Alokasi Waktu", "2 Pertemuan (4 x 40 Menit)")

  if st.button("🚀 Buat Modul Ajar Sekarang", type="primary"):
    with st.spinner("AI sedang merancang Modul Ajar profesional..."):
      try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt_modul = f"""
                Buatkan Modul Ajar / RPP lengkap, terstruktur, profesional, dan rapi tanpa teks kode mentah untuk:
                - Nama Guru: {nama_guru}
                - Nama Sekolah: {nama_sekolah}
                - Mata Pelajaran: {mapel}
                - Kurikulum: {kurikulum_aktif}
                - Fase/Kelas: {fase_kelas}
                - Kepala Sekolah: {nama_ks}
                - Topik: {topik}
                - Alokasi Waktu: {alokasi_waktu}

                Sertakan komponen formal: Identitas, Profil Pelajar Pancasila, Tujuan Pembelajaran, Kegiatan Pembelajaran, Asesmen & Rubrik Penilaian (Tabel).
                """
        response = client.models.generate_content(
            model="gemini-2.5-flash", contents=prompt_modul
        )
        st.session_state.modul_hasil = response.text
        st.success("Modul Ajar berhasil dibuat!")
      except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

  if st.session_state.modul_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Modul Ajar")
    st.markdown(st.session_state.modul_hasil)
    file_docx_modul = buat_file_docx(st.session_state.modul_hasil)
    st.download_button(
        label="📥 Unduh Modul Ajar (Format .DOCX / Word)",
        data=file_docx_modul,
        file_name=f"Modul_Ajar_{mapel}_{topik}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)


# --- TAB 2: GENERATOR SOAL & KUNCI JAWABAN ---
with tab2:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📝 Generator Paket Soal & Kunci Jawaban Lengkap")
  st.markdown(
      "Hasilkan paket soal ujian formal beserta rubrik dan pedoman penskoran."
  )

  col_s1, col_s2 = st.columns(2)
  with col_s1:
    s_nama_guru = st.text_input(
        "Nama Pembuat Soal / Guru", "Ahmad Fauzi, S.Pd.", key="s_guru"
    )
    s_sekolah = st.text_input(
        "Nama Instansi / Sekolah",
        "SMP Negeri 1 Nusantara",
        key="s_sekolah",
    )
    s_mapel = st.text_input(
        "Mata Pelajaran Ujian", "Matematika", key="s_mapel"
    )
    s_pilihan_kurikulum = st.selectbox(
        "Kurikulum Acuan Soal",
        [
            "Kurikulum Merdeka",
            "Kurikulum 2013 (K-13)",
            "Lainnya (Ketik Manual)",
        ],
        key="s_kur",
    )
    if s_pilihan_kurikulum == "Lainnya (Ketik Manual)":
      s_kur_aktif = st.text_input(
          "Ketik Nama Kurikulum Soal",
          "Kurikulum Pesantren",
          key="s_kur_manual",
      )
    else:
      s_kur_aktif = s_pilihan_kurikulum

  with col_s2:
    s_kelas = st.text_input("Kelas / Semester", "Kelas VII / Semester Ganjil")
    s_materi = st.text_input(
        "Materi / Bab Ujian", "Persamaan Linear Satu Variabel"
    )
    s_komposisi = st.text_input(
        "Komposisi Soal",
        "5 Pilihan Ganda (A-E), 2 Uraian Singkat, 1 Essai",
    )
    s_kesulitan = st.selectbox(
        "Tingkat Kesulitan",
        ["Mudah", "Sedang (HOTS)", "Tinggi / Kompleks (HOTS)"],
        key="s_kesulitan",
    )

  if st.button("🚀 Buat Paket Soal & Kunci Jawaban", type="primary"):
    with st.spinner("AI sedang menyusun paket soal dan kunci jawaban terbaik..."):
      try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt_soal = f"""
                Buatkan paket soal ujian resmi lengkap dengan identitas, naskah soal, kunci jawaban, dan rubrik penilaian yang sangat rapi (gunakan format Markdown, tabel untuk skor, dan penomoran tegas, tanpa teks kode mentah) untuk:
                - Guru Pengampu: {s_nama_guru}
                - Sekolah: {s_sekolah}
                - Mata Pelajaran: {s_mapel}
                - Kurikulum: {s_kur_aktif}
                - Kelas/Semester: {s_kelas}
                - Materi: {s_materi}
                - Komposisi: {s_komposisi}
                - Tingkat Kesulitan: {s_kesulitan}

                Struktur Output: KOP Ujian, Petunjuk Pengerjaan, Naskah Soal (PG, Uraian, Essai), dan Kunci Jawaban serta Pedoman Penskoran (Tabel Markdown).
                """
        response = client.models.generate_content(
            model="gemini-2.5-flash", contents=prompt_soal
        )
        st.session_state.soal_hasil = response.text
        st.success("Paket Soal & Kunci Jawaban Berhasil Dibuat!")
      except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

  if st.session_state.soal_hasil:
    st.markdown("---")
    st.subheader("📄 Pratinjau Paket Soal & Kunci Jawaban")
    st.markdown(st.session_state.soal_hasil)
    file_docx_soal = buat_file_docx(st.session_state.soal_hasil)
    st.download_button(
        label="📥 Unduh Paket Soal & Kunci (Format .DOCX / Word)",
        data=file_docx_soal,
        file_name=f"Paket_Soal_Kunci_{s_mapel}_{s_materi}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
  st.markdown("</div>", unsafe_allow_html=True)


# --- TAB 3: SET KUNCI JAWABAN TERPISAH ---
with tab3:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("⚙️ Set Kunci Jawaban Terpisah (Acuan Koreksi)")
  st.markdown(
      "Pisahkan kunci jawaban berdasarkan jenis soal agar AI dapat membaca dan"
      " mengoreksi dengan presisi tinggi."
  )

  col_k1, col_k2 = st.columns(2)

  with col_k1:
    st.subheader("📌 1. Kunci Pilihan Ganda")
    kunci_pg_input = st.text_area(
        "Masukkan Kunci PG (Contoh: 1.A, 2.C, 3.B, dst)",
        value=st.session_state.kunci_pg,
        height=120,
    )

    st.subheader("📌 2. Kunci Isian Singkat")
    kunci_isian_input = st.text_area(
        "Masukkan Kunci Isian Singkat & Kata Kunci Utama",
        value=st.session_state.kunci_isian,
        height=120,
    )

  with col_k2:
    st.subheader("📌 3. Rubrik & Kunci Essai")
    kunci_essai_input = st.text_area(
        "Masukkan Rubrik/Kunci Essai & Bobot Penilaian",
        value=st.session_state.kunci_essai,
        height=275,
    )

  if st.button("Simpan Seluruh Kunci Jawaban", type="primary"):
    st.session_state.kunci_pg = kunci_pg_input
    st.session_state.kunci_isian = kunci_isian_input
    st.session_state.kunci_essai = kunci_essai_input
    st.success(
        "Semua kategori kunci jawaban (PG, Isian, Essai) berhasil disimpan dan"
        " disinkronkan ke sistem koreksi!"
    )
  st.markdown("</div>", unsafe_allow_html=True)


# --- TAB 4: KOREKSI LEMBAR SISWA ---
with tab4:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("🔍 Koreksi Lembar Jawaban Siswa (AI Vision & Dokumen)")

  if (
      not st.session_state.kunci_pg
      and not st.session_state.kunci_isian
      and not st.session_state.kunci_essai
  ):
    st.warning(
        "⚠️ Harap atur dan simpan minimal salah satu Kunci Jawaban di Tab 3"
        " terlebih dahulu!"
    )
  else:
    nama_siswa = st.text_input("Nama Lengkap / Nomor Peserta Siswa")
    metode_siswa = st.radio(
        "Format Lembar Jawaban Siswa:",
        [
            "Unggah Foto / Scan (Tulisan Tangan)",
            "Unggah Dokumen (PDF / Word)",
            "Ketik Teks Jawaban Langsung",
        ],
    )

    file_siswa_img = None
    file_siswa_doc = None
    teks_siswa_manual = ""

    if "Foto" in metode_siswa:
      file_siswa_img = st.file_uploader(
          "Unggah foto lembar jawaban", type=["jpg", "jpeg", "png"]
      )
      if file_siswa_img:
        st.image(
            Image.open(file_siswa_img),
            caption=f"Lembar Jawab: {nama_siswa}",
            width=350,
        )
    elif "Dokumen" in metode_siswa:
      file_siswa_doc = st.file_uploader(
          "Unggah file PDF / Word", type=["pdf", "docx", "txt"]
      )
    else:
      teks_siswa_manual = st.text_area("Ketik teks jawaban siswa di sini:")

    if st.button("Proses Koreksi Komprehensif Sekarang", type="primary"):
      if not nama_siswa:
        st.error("Mohon masukkan nama siswa!")
      else:
        with st.spinner(
            "AI sedang mencocokkan lembar jawaban dengan seluruh kategori"
            " kunci..."
        ):
          try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            contents_payload = []

            master_kunci_gabungan = f"""
            - KUNCI PILIHAN GANDA:
            {st.session_state.kunci_pg if st.session_state.kunci_pg else "Tidak ada"}

            - KUNCI ISIAN SINGKAT:
            {st.session_state.kunci_isian if st.session_state.kunci_isian else "Tidak ada"}

            - RUBRIK & KUNCI ESSAI:
            {st.session_state.kunci_essai if st.session_state.kunci_essai else "Tidak ada"}
            """

            prompt_instruksi = f"""
            Anda adalah guru profesional yang sangat teliti, objektif, dan adil. 
            Tugas Anda adalah mengoreksi lembar jawaban siswa secara menyeluruh (mencakup Pilihan Ganda, Isian Singkat, dan Essai) berdasarkan KUNCI JAWABAN TERPISAH BERIKUT:
            
            {master_kunci_gabungan}
            
            INSTRUKSI FORMAT OUTPUT:
            Baris pertama WAJIB menuliskan persis format ini agar terbaca sistem:
            NILAI_AKHIR: [Angka total nilai akhir skala 0 sampai 100]
            
            Setelah baris itu, berikan rincian evaluasi terstruktur dalam format Markdown (Tabel atau list per bagian: Pilihan Ganda, Isian Singkat, dan Essai), beserta ulasan atau poin perbaikan untuk siswa.
            """
            contents_payload.append(prompt_instruksi)

            if file_siswa_img:
              contents_payload.append(Image.open(file_siswa_img))
            elif file_siswa_doc:
              ext = file_siswa_doc.name.split(".")[-1].lower()
              extracted = ""
              if ext == "pdf" and pypdf_available:
                reader = PdfReader(file_siswa_doc)
                for page in reader.pages:
                  extracted += page.extract_text() or ""
              elif ext in ["docx", "doc"]:
                doc = Document(file_siswa_doc)
                for para in doc.paragraphs:
                  extracted += para.text + "\n"
              contents_payload.append(f"JAWABAN SISWA:\n{extracted}")
            elif teks_siswa_manual:
              contents_payload.append(f"JAWABAN SISWA:\n{teks_siswa_manual}")

            response = client.models.generate_content(
                model="gemini-2.5-flash", contents=contents_payload
            )
            teks_hasil = response.text

            import re
            match_nilai = re.search(r"NILAI_AKHIR:\s*([0-9.]+)", teks_hasil)
            skor_angka = float(match_nilai.group(1)) if match_nilai else 0.0

            st.session_state.rekap_nilai.append({
                "Nama Siswa": nama_siswa,
                "Nilai Akhir": skor_angka,
                "Detail Koreksi": teks_hasil,
            })

            st.success(
                f"Koreksi Selesai! Siswa **{nama_siswa}** memperoleh Nilai Akhir:"
                f" **{skor_angka}**"
            )
            with st.expander("Lihat Rincian Koreksi Lengkap"):
              st.markdown(teks_hasil)

          except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses: {e}")
  st.markdown("</div>", unsafe_allow_html=True)


# --- TAB 5: REKAP NILAI KELAS ---
with tab5:
  st.markdown('<div class="card-box">', unsafe_allow_html=True)
  st.header("📊 Rekapitulasi Nilai Kelas")

  if len(st.session_state.rekap_nilai) == 0:
    st.info(
        "Belum ada data siswa yang dikoreksi. Lakukan koreksi pada Tab 4"
        " terlebih dahulu."
    )
  else:
    df_rekap = pd.DataFrame(st.session_state.rekap_nilai)
    st.dataframe(
        df_rekap[["Nama Siswa", "Nilai Akhir"]], use_container_width=True
    )

    csv_data = df_rekap.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Unduh Rekap Nilai (CSV / Excel)",
        data=csv_data,
        file_name="rekap_nilai_kelas.csv",
        mime="text/csv",
    )

    if st.button("🗑️ Hapus Semua Data Rekap"):
      st.session_state.rekap_nilai = []
      st.rerun()
  st.markdown("</div>", unsafe_allow_html=True)

# --- FOOTER APLIKASI ---
st.markdown(
    """
    <div class="footer">
        <p>Developed with ❤️ by <b>Zeeo</b> | WhatsApp: <a href="https://wa.me/6282371729760" target="_blank">082371729760</a> | © 2026 All Rights Reserved</p>
    </div>
    """,
    unsafe_allow_html=True,
)
