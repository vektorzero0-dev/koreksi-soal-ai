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
    page_title="AI Smart Exam Grader Pro", page_icon="📝", layout="wide"
)

# AMBIL API KEY TERTANAM DARI SECRETS STREAMLIT
try:
  GEMINI_API_KEY = st.secrets["AQ.Ab8RN6I2VxjP9z33FXLKlwKccV8EWA6ULZkGoO1MEzmxgpzg4g"]
except Exception:
  st.error(
      "⚠️ API Key belum diatur di Streamlit Secrets! Harap atur GEMINI_API_KEY"
      " di menu pengaturan Cloud."
  )
  st.stop()

# Inisialisasi Session State
if "kunci_master" not in st.session_state:
  st.session_state.kunci_master = ""
if "rekap_nilai" not in st.session_state:
  st.session_state.rekap_nilai = []

st.title("📝 Aplikasi Koreksi Soal Otomatis & Fleksibel (Gemini 3.5 Flash)")
st.markdown(
    "Mendukung input kunci jawaban & jawaban siswa via **Ketik Manual, PDF,"
    " Word, atau Foto**."
)

tab1, tab2, tab3 = st.tabs(
    ["1. Upload / Set Kunci Jawaban", "2. Koreksi Siswa", "3. Rekap Nilai"]
)


# Fungsi pembantu untuk membaca teks dari file PDF atau Word
def ekstrak_teks_dari_file(uploaded_file):
  file_extension = uploaded_file.name.split(".")[-1].lower()
  teks_terekstrak = ""

  if file_extension == "pdf":
    if not pypdf_available:
      return "[Error: Pustaka pypdf belum terinstal]"
    try:
      reader = PdfReader(uploaded_file)
      for page in reader.pages:
        teks_terekstrak += page.extract_text() or ""
    except Exception as e:
      teks_terekstrak = f"[Gagal membaca PDF: {e}]"

  elif file_extension in ["docx", "doc"]:
    try:
      doc = Document(uploaded_file)
      for para in doc.paragraphs:
        teks_terekstrak += para.text + "\n"
    except Exception as e:
      teks_terekstrak = f"[Gagal membaca Word: {e}]"

  elif file_extension in ["txt"]:
    teks_terekstrak = uploaded_file.read().decode("utf-8")

  return teks_terekstrak


# --- TAB 1: KUNCI JAWABAN ---
with tab1:
  st.header("Langkah 1: Tentukan Kunci Jawaban & Rubrik Acuan")
  st.markdown(
      "Anda bisa menuliskan kunci jawaban secara manual ATAU mengunggah"
      " dokumen/foto."
  )

  metode_kunci = st.radio(
      "Pilih Metode Masukan Kunci Jawaban:",
      ["Ketik Manual", "Unggah Dokumen (PDF/Word/TXT)", "Unggah Foto/Scan"],
  )

  teks_kunci_input = ""
  file_kunci_obj = None

  if metode_kunci == "Ketik Manual":
    teks_kunci_input = st.text_area(
        "Ketik Kunci Jawaban & Rubrik Lengkap (PG, Uraian, Essai):",
        value=st.session_state.kunci_master,
        height=150,
    )
  elif metode_kunci == "Unggah Dokumen (PDF/Word/TXT)":
    uploaded_doc_kunci = st.file_uploader(
        "Pilih file PDF atau Word berisi kunci jawaban", type=["pdf", "docx", "txt"]
    )
    if uploaded_doc_kunci:
      teks_kunci_input = ekstrak_teks_dari_file(uploaded_doc_kunci)
      st.text_area(
          "Pratinjau Teks Kunci dari Dokumen:",
          value=teks_kunci_input,
          height=150,
      )
  else:
    file_kunci_obj = st.file_uploader(
        "Unggah foto/scan kunci jawaban", type=["jpg", "jpeg", "png"]
    )
    if file_kunci_obj:
      st.image(
          Image.open(file_kunci_obj), caption="Pratinjau Foto Kunci", width=300
      )

  if st.button("Simpan Kunci Jawaban ke Sistem", type="primary"):
    if (
        not teks_kunci_input.strip()
        and metode_kunci != "Unggah Foto/Scan"
        and not file_kunci_obj
    ):
      st.warning("Kunci jawaban masih kosong!")
    else:
      if metode_kunci == "Unggah Foto/Scan" and file_kunci_obj:
        st.session_state.kunci_master = (
            "[Kunci Jawaban dilampirkan dalam bentuk Gambar/Foto]"
        )
        st.session_state.file_kunci_obj = file_kunci_obj
      else:
        st.session_state.kunci_master = teks_kunci_input
        st.session_state.file_kunci_obj = None
      st.success("Kunci jawaban berhasil disimpan sebagai acuan utama!")

# --- TAB 2: KOREKSI SISWA ---
with tab2:
  st.header("Langkah 2: Koreksi Lembar Jawaban Siswa")

  if not st.session_state.kunci_master and "file_kunci_obj" not in st.session_state:
    st.warning("⚠️ Harap atur dan simpan Kunci Jawaban terlebih dahulu di Tab 1!")
  else:
    nama_siswa = st.text_input("Nama Lengkap / Nomor Peserta Siswa")
    metode_siswa = st.radio(
        "Pilih Format Jawaban Siswa:",
        ["Unggah Foto/Scan (Tulisan Tangan/Lembar Jawab)", "Unggah Dokumen (PDF/Word)", "Ketik Teks Jawaban"],
    )

    file_siswa_img = None
    file_siswa_doc = None
    teks_siswa_manual = ""

    if "Foto" in metode_siswa:
      file_siswa_img = st.file_uploader(
          "Unggah foto lembar jawaban siswa", type=["jpg", "jpeg", "png"]
      )
      if file_siswa_img:
        st.image(
            Image.open(file_siswa_img), caption=f"Lembar Jawab: {nama_siswa}", width=400
        )
    elif "Dokumen" in metode_siswa:
      file_siswa_doc = st.file_uploader(
          "Unggah file PDF / Word jawaban siswa", type=["pdf", "docx", "txt"]
      )
    else:
      teks_siswa_manual = st.text_area("Ketik atau salin jawaban siswa di sini:")

    if st.button("Proses Koreksi dengan Gemini 3.5 Flash", type="primary"):
      if not nama_siswa:
        st.error("Mohon masukkan nama siswa!")
      else:
        with st.spinner("Gemini 3.5 Flash sedang menganalisis dan mengoreksi jawaban..."):
          try:
            client = genai.Client(api_key=GEMINI_API_KEY)

            # Siapkan muatan isi (contents) untuk Gemini 3.5 Flash
            contents_payload = []

            # Tambahkan instruksi sistem/prompt utama
            prompt_instruksi = f"""
            Anda adalah guru profesional yang sangat teliti dan objektif. 
            Tugas Anda adalah mengoreksi jawaban siswa berdasarkan KUNCI JAWABAN ACUAN.
            
            KUNCI JAWABAN ACUAN:
            {st.session_state.kunci_master}
            
            INSTRUKSI FORMAT OUTPUT:
            Baris pertama WAJIB menuliskan persis format ini untuk dibaca sistem:
            NILAI_AKHIR: [Angka total nilai dari rentang 0 sampai 100]
            
            Setelah baris itu, berikan rincian ulasan, poin perbaikan, dan evaluasi mendalam per nomor soal (Pilihan Ganda, Uraian Singkat, dan Essai) dalam format Markdown yang rapi.
            """
            contents_payload.append(prompt_instruksi)

            # Jika kunci berbentuk file gambar
            if "file_kunci_obj" in st.session_state and st.session_state.file_kunci_obj:
              contents_payload.append(
                  Image.open(st.session_state.file_kunci_obj)
              )

            # Tambahkan data jawaban siswa berdasarkan metodenya
            if file_siswa_img:
              contents_payload.append(Image.open(file_siswa_img))
            elif file_siswa_doc:
              teks_extracted = ekstrak_teks_dari_file(file_siswa_doc)
              contents_payload.append(f"JAWABAN SISWA DARI DOKUMEN:\n{teks_extracted}")
            elif teks_siswa_manual:
              contents_payload.append(f"JAWABAN SISWA:\n{teks_siswa_manual}")

            # Panggil model Gemini 3.5 Flash (mendukung multimodal & reasoning efisien)
            response = client.models.generate_content(
                model="gemini-3.5-flash", contents=contents_payload
            )

            teks_hasil = response.text

            # Ekstraksi angka nilai akhir secara otomatis
            import re
            match_nilai = re.search(r"NILAI_AKHIR:\s*([0-9.]+)", teks_hasil)
            skor_angka = float(match_nilai.group(1)) if match_nilai else 0.0

            # Simpan ke memori rekap kelas
            st.session_state.rekap_nilai.append({
                "Nama Siswa": nama_siswa,
                "Nilai Akhir": skor_angka,
                "Detail Koreksi": teks_hasil,
            })

            st.success(
                f"Koreksi Selesai! Siswa **{nama_siswa}** memperoleh Nilai Akhir:"
                f" **{skor_angka}**"
            )
            with st.expander("Lihat Detail Rincian Koreksi AI"):
              st.markdown(teks_hasil)

          except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses dengan Gemini: {e}")

# --- TAB 3: REKAP NILAI ---
with tab3:
  st.header("Langkah 3: Rekapitulasi Nilai Kelas")

  if len(st.session_state.rekap_nilai) == 0:
    st.info("Belum ada data siswa yang dikoreksi.")
  else:
    df_rekap = pd.DataFrame(st.session_state.rekap_nilai)
    st.dataframe(df_rekap[["Nama Siswa", "Nilai Akhir"]], use_container_width=True)

    csv_data = df_rekap.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Unduh Rekap Nilai (CSV/Excel)",
        data=csv_data,
        file_name="rekap_nilai_kelas.csv",
        mime="text/csv",
    )

    if st.button("🗑️ Hapus Semua Rekap"):
      st.session_state.rekap_nilai = []
      st.rerun()
