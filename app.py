from google import genai
from PIL import Image
import pandas as pd
import streamlit as st

# Konfigurasi Halaman
st.set_page_config(
    page_title="Aplikasi Koreksi & Rekap Otomatis", page_icon="📊", layout="wide"
)

# Inisialisasi Session State
if "kunci_jawaban" not in st.session_state:
  st.session_state.kunci_jawaban = None
if "rekap_nilai" not in st.session_state:
  st.session_state.rekap_nilai = []

st.title("📊 Aplikasi Koreksi Soal & Rekap Nilai Otomatis (AI)")

tab1, tab2, tab3 = st.tabs(
    ["1. Set Kunci Jawaban", "2. Koreksi Lembar Siswa", "3. Rekap Nilai Kelas"]
)

# --- TAB 1: SET KUNCI JAWABAN ---
with tab1:
  st.header("Langkah 1: Tentukan Kunci Jawaban Acuan")
  st.markdown(
      "Masukkan kunci jawaban dan rubrik penilaian acuan untuk seluruh siswa."
  )

  with st.form("form_kunci"):
    api_key_input = st.text_input(
        "Google Gemini API Key",
        type="password",
        value=st.session_state.get("api_key", ""),
    )
    kunci_pg = st.text_area(
        "Kunci Pilihan Ganda",
        value=st.session_state.get(
            "kunci_pg", "1. A, 2. C, 3. B, 4. D, 5. E"
        ),
    )
    kunci_uraian = st.text_area(
        "Kunci Uraian Singkat",
        value=st.session_state.get(
            "kunci_uraian",
            "1. Fotosintesis (Kata kunci: klorofil, cahaya matahari)",
        ),
    )
    kunci_essai = st.text_area(
        "Rubrik/Kunci Essai",
        value=st.session_state.get(
            "kunci_essai",
            "1. Siklus air (Kunci: Evaporasi, Kondensasi, Presipitasi). Bobot"
            " maks: 20",
        ),
    )

    submitted_kunci = st.form_submit_button("Simpan Kunci Jawaban")
    if submitted_kunci:
      st.session_state.api_key = api_key_input
      st.session_state.kunci_pg = kunci_pg
      st.session_state.kunci_uraian = kunci_uraian
      st.session_state.kunci_essai = kunci_essai
      st.session_state.kunci_jawaban = True
      st.success("Kunci jawaban berhasil disimpan!")

# --- TAB 2: KOREKSI LEMBAR SISWA ---
with tab2:
  st.header("Langkah 2: Unggah Lembar Jawaban Siswa")

  if not st.session_state.kunci_jawaban:
    st.warning("⚠️ Harap simpan 'Kunci Jawaban' terlebih dahulu pada Tab 1!")
  else:
    nama_siswa = st.text_input("Nama Lengkap / Nomor Peserta Siswa")
    uploaded_image = st.file_uploader(
        "Unggah Foto Lembar Jawaban Siswa", type=["jpg", "jpeg", "png"]
    )

    if uploaded_image is not None:
      image = Image.open(uploaded_image)
      st.image(image, caption=f"Lembar Jawaban: {nama_siswa}", width=400)

    if st.button("Koreksi & Masukkan ke Rekap", type="primary"):
      if not nama_siswa or not uploaded_image:
        st.error("Mohon masukkan nama siswa dan unggah foto lembar jawabannya!")
      else:
        with st.spinner("AI sedang mencocokkan jawaban dengan kunci..."):
          try:
            client = genai.Client(api_key=st.session_state.api_key)

            prompt = f"""
                        Anda adalah asisten guru yang objektif. Koreksi lembar jawaban siswa di gambar berdasarkan KUNCI JAWABAN ACUAN berikut:
                        - Kunci PG: {st.session_state.kunci_pg}
                        - Kunci Uraian: {st.session_state.kunci_uraian}
                        - Kunci Essai: {st.session_state.kunci_essai}

                        Berikan format output persis seperti ini di baris pertama agar bisa dibaca sistem:
                        NILAI_AKHIR: [Angka total skala 0-100]
                        
                        Setelah itu, berikan rincian ulasan atau catatan perbaikan yang mendetail untuk siswa.
                        """

            response = client.models.generate_content(
                model="gemini-2.5-flash", contents=[image, prompt]
            )

            teks_hasil = response.text

            # Ekstraksi nilai angka dari teks AI
            import re

            match_nilai = re.search(r"NILAI_AKHIR:\s*([0-9.]+)", teks_hasil)
            skor_angka = (
                float(match_nilai.group(1)) if match_nilai else 0.0
            )

            # Simpan ke rekap nilai
            st.session_state.rekap_nilai.append({
                "Nama Siswa": nama_siswa,
                "Nilai Akhir": skor_angka,
                "Detail AI": teks_hasil,
            })

            st.success(
                f"Berhasil! Siswa **{nama_siswa}** mendapat Nilai Akhir:"
                f" **{skor_angka}**"
            )
            with st.expander("Lihat Rincian Koreksi AI"):
              st.markdown(teks_hasil)

          except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses: {e}")

# --- TAB 3: REKAP NILAI KELAS ---
with tab3:
  st.header("Langkah 3: Rekapitulasi Nilai Seluruh Siswa")

  if len(st.session_state.rekap_nilai) == 0:
    st.info("Belum ada data siswa yang dikoreksi.")
  else:
    df_rekap = pd.DataFrame(st.session_state.rekap_nilai)
    st.dataframe(df_rekap[["Nama Siswa", "Nilai Akhir"]], use_container_width=True)

    csv_data = df_rekap.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Unduh Rekap Nilai (CSV/Excel)",
        data=csv_data,
        file_name="rekap_nilai_ujian.csv",
        mime="text/csv",
    )

    if st.button("🗑️ Bersihkan Semua Data Rekap"):
      st.session_state.rekap_nilai = []
      st.rerun()
