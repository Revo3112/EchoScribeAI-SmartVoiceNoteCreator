# EchoScribe AI - Smart Voice Note Creator

![EchoScribe AI Logo](https://via.placeholder.com/300x150.png?text=EchoScribe+AI+Logo)
*(Ganti dengan link ke logo aplikasi Anda)*

**EchoScribe AI** adalah aplikasi desktop canggih yang dirancang untuk mengubah rekaman suara Anda menjadi catatan teks terstruktur dan profesional secara otomatis. Dengan dukungan teknologi AI terkini, aplikasi ini tidak hanya mentranskripsikan audio, tetapi juga meningkatkan, memformat, dan menyajikannya dalam dokumen Word (.docx) yang siap pakai.

[![Python Version](https://img.shields.io/badge/python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## Daftar Isi

1.  [Demo Singkat](#demo-singkat)
2.  [Fitur Utama](#fitur-utama)
3.  [Prasyarat](#prasyarat)
4.  [Instalasi](#instalasi)
5.  [Konfigurasi](#konfigurasi)
    *   [FFmpeg](#ffmpeg)
    *   [Groq API Key](#groq-api-key)
6.  [Cara Penggunaan](#cara-penggunaan)
    *   [Tab Rekaman](#tab-rekaman-recording)
    *   [Tab Pengaturan](#tab-pengaturan-settings)
    *   [Tab Output](#tab-output)
    *   [Status Bar](#status-bar)
7.  [Bagaimana Sistem Bekerja](#bagaimana-sistem-bekerja)
    *   [Alur Kerja Utama](#alur-kerja-utama)
    *   [Komponen Kunci](#komponen-kunci)
8.  [Troubleshooting](#troubleshooting)
9.  [Dependensi Utama](#dependensi-utama)
10. [Kontribusi](#kontribusi)
11. [Lisensi](#lisensi)

---

## Demo Singkat

Berikut adalah gambaran tampilan antarmuka EchoScribe AI:

![EchoScribe AI Screenshot](https://via.placeholder.com/600x400.png?text=Screenshot+Aplikasi+EchoScribe+AI)
*(Ganti dengan link ke screenshot aplikasi Anda)*

Dengarkan contoh hasil peningkatan audio dan transkripsi (jika ada):
[Link ke Contoh Audio Demo](https://example.com/audio_demo.mp3)
*(Ganti dengan link ke demo audio Anda)*

---

## Fitur Utama

*   🎙️ **Perekaman Audio Fleksibel**: Rekam audio dari mikrofon, audio sistem (khusus Windows dengan PyAudioWPatch), atau keduanya secara bersamaan (dual recording).
*   📊 **Visualisasi Audio Real-time**: Pantau input audio Anda secara visual dengan berbagai mode (waveform, bars, spectrum, fill) dan sensitivitas yang dapat disesuaikan.
*   🔊 **Transkripsi Akurat**: Pilih antara mesin pengenalan suara Google Speech Recognition atau Groq API dengan model Whisper canggih untuk transkripsi berkualitas tinggi.
*   🤖 **Peningkatan AI Cerdas**: Manfaatkan kekuatan AI Groq untuk secara otomatis:
    *   Menganalisis konteks audio (rapat, kuliah, dikte, dll.).
    *   Memilih model AI yang optimal untuk transkripsi dan peningkatan.
    *   Menyusun ulang transkrip mentah menjadi catatan yang terstruktur, koheren, dan profesional.
    *   Menambahkan judul, subjudul, poin-poin, dan format lain yang relevan.
*   📄 **Ekspor ke Dokumen Word (.docx)**: Hasil akhir disimpan dalam format DOCX dengan styling profesional, termasuk:
    *   Tema dokumen yang adaptif berdasarkan jenis konten.
    *   Header dan footer otomatis.
    *   Styling heading dengan ikon kontekstual.
    *   Format daftar (bullet & numbered), tabel, kutipan, dan blok kode.
    *   Callout/admonition block (Note, Warning, Important, dll.) yang ditingkatkan.
*   ⚙️ **Pengaturan Komprehensif**:
    *   Manajemen API Key Groq (gunakan API key sendiri atau default).
    *   Pilihan bahasa transkripsi (Indonesia, Inggris, Jepang, Mandarin).
    *   Mode ekonomis untuk penggunaan AI yang lebih hemat.
    *   Konfigurasi output (folder, awalan nama file).
    *   Pengaturan rekaman lanjutan (dukungan rekaman panjang, ukuran chunk, jeda API).
*   💾 **Manajemen Konfigurasi**: Pengaturan aplikasi disimpan secara otomatis dan dimuat saat startup.
*   🛠️ **Penanganan Error & Logging**: Sistem penanganan error yang tangguh dan logging detail untuk troubleshooting.
*   🎨 **Antarmuka Modern & Intuitif**: Dibangun dengan CustomTkinter untuk tampilan yang menarik dan mudah digunakan.
*   🛡️ **Kompatibilitas Sistem**: Pengecekan kompatibilitas awal untuk memastikan dependensi terpenuhi.

---

## Prasyarat

Sebelum menginstal dan menjalankan EchoScribe AI, pastikan sistem Anda memenuhi persyaratan berikut:

1.  **Python**: Versi 3.7 atau lebih baru.
2.  **FFmpeg**:
    *   Dibutuhkan oleh `pydub` untuk pemrosesan format audio tertentu.
    *   Pastikan FFmpeg terinstal dan `ffmpeg.exe` (atau `ffmpeg` di Linux/macOS) dapat diakses melalui PATH sistem Anda, atau ditempatkan di salah satu lokasi pencarian standar aplikasi (lihat bagian [Konfigurasi FFmpeg](#ffmpeg)).
    *   Unduh dari [ffmpeg.org](https://ffmpeg.org/download.html).
3.  **Groq API Key (Opsional, Sangat Direkomendasikan)**:
    *   Untuk fungsionalitas transkripsi Whisper dan peningkatan AI penuh, Anda disarankan menggunakan API key Groq Anda sendiri.
    *   Aplikasi menyediakan API key default dengan batasan penggunaan.
    *   Dapatkan API key dari [Groq Console](https://console.groq.com/).
4.  **Sistem Operasi**:
    *   **Windows**: Direkomendasikan untuk fungsionalitas penuh, terutama perekaman audio sistem (menggunakan `PyAudioWPatch`).
    *   **Linux/macOS**: Fitur inti seperti perekaman mikrofon dan pemrosesan AI akan berfungsi. Perekaman audio sistem mungkin memerlukan konfigurasi tambahan atau tidak berfungsi secara optimal.
5.  **Koneksi Internet**: Diperlukan untuk layanan transkripsi cloud (Google, Groq) dan peningkatan AI.
6.  **Mikrofon**: Diperlukan untuk merekam audio dari mikrofon.

---

## Instalasi

1.  **Clone Repository (Jika dari source code)**:
    ```bash
    git clone https://link_ke_repository_anda.git
    cd nama_folder_repository
    ```

2.  **Buat dan Aktifkan Virtual Environment (Direkomendasikan)**:
    ```bash
    python -m venv venv
    # Windows
    venv\Scripts\activate
    # macOS/Linux
    source venv/bin/activate
    ```

3.  **Install Dependensi**:
    Pastikan Anda memiliki file `requirements.txt` atau install manual:
    ```bash
    pip install -r requirements.txt
    ```
    Jika `requirements.txt` tidak tersedia, Anda perlu menginstal pustaka utama:
    ```bash
    pip install speechrecognition pydub customtkinter python-docx groq matplotlib numpy pyaudio sounddevice
    # Untuk Windows, coba instal pyaudiowpatch untuk perekaman audio sistem yang lebih baik:
    pip install pyaudiowpatch
    ```
    Lihat bagian [Dependensi Utama](#dependensi-utama) untuk daftar lengkap.

4.  **Pastikan FFmpeg Terinstal**:
    Ikuti petunjuk di [Prasyarat](#prasyarat) dan [Konfigurasi FFmpeg](#ffmpeg).

5.  **Jalankan Aplikasi**:
    ```bash
    python nama_file_utama_anda.py
    ```
    (Ganti `nama_file_utama_anda.py` dengan nama file Python utama, misal `main.py` atau `app.py` sesuai kode Anda).

---

## Konfigurasi

### FFmpeg

Aplikasi ini menggunakan `pydub` yang bergantung pada FFmpeg untuk memproses beberapa format audio.
Aplikasi akan mencoba menemukan FFmpeg secara otomatis di lokasi berikut:
*   Direktori yang sama dengan file executable aplikasi.
*   `C:\FFmpeg\bin\ffmpeg.exe` (Windows)
*   `%USERPROFILE%\FFmpeg\bin\ffmpeg.exe` (Windows)
*   `%LOCALAPPDATA%\FFmpeg\bin\ffmpeg.exe` (Windows)
*   `%PROGRAMFILES%\FFmpeg\bin\ffmpeg.exe` (Windows)
*   `%PROGRAMFILES(X86)%\FFmpeg\bin\ffmpeg.exe` (Windows)
*   Melalui PATH sistem.

**Jika FFmpeg tidak ditemukan:**
1.  Unduh FFmpeg dari [https://ffmpeg.org/download.html](https://ffmpeg.org/download.html).
2.  Ekstrak arsipnya.
3.  Tambahkan folder `bin` (yang berisi `ffmpeg.exe`) ke PATH environment variable sistem Anda, ATAU letakkan `ffmpeg.exe` di salah satu lokasi di atas.

### Groq API Key

Untuk menggunakan fitur transkripsi Whisper dan peningkatan AI tanpa batasan, Anda dapat mengkonfigurasi API key Groq Anda sendiri.

1.  **Dapatkan API Key**:
    *   Kunjungi [https://console.groq.com/](https://console.groq.com/).
    *   Daftar atau login ke akun Anda.
    *   Buat API Key baru di dashboard Anda.
    *   Salin API key tersebut.

2.  **Konfigurasi di Aplikasi**:
    *   Buka tab **Pengaturan** di aplikasi EchoScribe AI.
    *   Klik tombol **"Kelola API Key"**.
    *   Sebuah dialog akan muncul. Tempelkan API key Anda di kolom input.
    *   Klik **"Simpan"**. Aplikasi akan mencoba memverifikasi API key Anda.
    *   Jika Anda tidak ingin menggunakan API key sendiri, Anda dapat memilih **"Gunakan Default"**.

API key kustom Anda akan disimpan secara lokal di komputer Anda dalam file `config.json` di folder `.echoscribe` (dalam direktori home pengguna Anda).

---

## Cara Penggunaan

Antarmuka EchoScribe AI dibagi menjadi tiga tab utama: **Rekaman**, **Pengaturan**, dan **Output**.

### Tab Rekaman (Recording)

Ini adalah area utama untuk mengontrol proses perekaman.

![Tab Rekaman](https://via.placeholder.com/500x350.png?text=Tab+Rekaman)
*(Ganti dengan screenshot Tab Rekaman)*

*   **Timer**: Menampilkan durasi rekaman saat ini (`00:00:00`).
*   **Visualisasi Audio**: Menampilkan representasi visual dari input audio secara real-time. Anda dapat memilih mode visualisasi (Waveform, Bars, Spectrum, Fill) dan mengatur sensitivitasnya melalui kontrol di bawah area visualisasi. Visualisasi dapat diaktifkan/dinonaktifkan.
*   **Tombol Mulai/Berhenti Rekaman**:
    *   **"Mulai Rekaman"**: Klik untuk memulai sesi perekaman baru.
    *   **"Berhenti Rekaman"**: Klik untuk mengakhiri sesi perekaman dan memulai proses transkripsi serta peningkatan AI.
*   **Progress Bar**: Menampilkan kemajuan proses transkripsi dan peningkatan setelah rekaman dihentikan.
*   **Pengaturan Cepat**:
    *   **Mikrofon**: Pilih perangkat mikrofon yang ingin digunakan dari daftar dropdown. Tombol `⟳` di sebelahnya berfungsi untuk menyegarkan daftar mikrofon.
    *   **Sumber Audio**:
        *   **Mikrofon saja**: Merekam hanya dari mikrofon yang dipilih.
        *   **Audio sistem saja**: Merekam semua suara yang diputar oleh sistem komputer Anda (misalnya, dari video YouTube, musik, panggilan online). Fitur ini berfungsi optimal di Windows dengan PyAudioWPatch.
        *   **Mikrofon + Audio sistem**: Merekam dari mikrofon dan audio sistem secara bersamaan, menggabungkannya menjadi satu rekaman.

### Tab Pengaturan (Settings)

Di sini Anda dapat menyesuaikan berbagai aspek perilaku aplikasi.

![Tab Pengaturan](https://via.placeholder.com/500x350.png?text=Tab+Pengaturan)
*(Ganti dengan screenshot Tab Pengaturan)*

*   **Konfigurasi API**:
    *   **Status API Key**: Menampilkan status API key yang sedang digunakan (Default atau Custom).
    *   **Kelola API Key**: Membuka dialog untuk memasukkan, menyimpan, atau menghapus API key Groq kustom Anda.
*   **Pengaturan Pengenalan Suara**:
    *   **Bahasa**: Pilih bahasa utama dari audio yang akan direkam (misalnya, `id-ID` untuk Bahasa Indonesia, `en-US` untuk Bahasa Inggris).
    *   **Mesin Pengenalan**: Pilih mesin transkripsi (`Google` atau `Whisper` via Groq).
    *   **Mode Ekonomis**: Jika dicentang dan menggunakan Groq, aplikasi akan mencoba menggunakan model AI yang lebih hemat (misalnya, `distil-whisper` untuk Bahasa Inggris).
*   **Pengaturan Peningkatan AI**:
    *   **Gunakan AI**: Aktifkan atau nonaktifkan fitur peningkatan catatan menggunakan AI setelah transkripsi. Jika dinonaktifkan, Anda hanya akan mendapatkan transkrip mentah.
*   **Pengaturan Output**:
    *   **Folder Output**: Tentukan folder tempat dokumen Word (.docx) hasil akan disimpan. Klik **"Browse"** untuk memilih folder.
    *   **Awalan Nama File**: Tentukan awalan yang akan digunakan untuk nama file output (misalnya, `catatan_rapat_`).
*   **Pengaturan Lanjutan**:
    *   **Rekaman Panjang**: Aktifkan dukungan untuk rekaman yang sangat panjang. Aplikasi akan memecah audio menjadi chunk dan memprosesnya secara bertahap untuk menghindari batasan memori dan API.
    *   **Ukuran Penggalan (Chunk Size)**: Jika "Rekaman Panjang" aktif, atur durasi setiap chunk audio (dalam detik atau menit) yang akan diproses. Nilai lebih kecil berarti lebih sering diproses, lebih besar berarti lebih jarang.
    *   **Jeda API (detik)**: Mengatur waktu tunggu antar permintaan ke API Groq untuk menghindari pembatasan laju (rate limiting).

### Tab Output

Area ini menampilkan hasil transkripsi dan teks yang telah ditingkatkan oleh AI.

![Tab Output](https://via.placeholder.com/500x350.png?text=Tab+Output)
*(Ganti dengan screenshot Tab Output)*

*   **Area Tampilan Hasil**: Menampilkan pratinjau teks yang telah diproses. Untuk rekaman panjang, pratinjau mungkin hanya menampilkan sebagian awal konten. Konten lengkap akan tersedia di file `.docx`.
*   **Tombol Ekspor**:
    *   **Salin ke Clipboard**: Menyalin seluruh teks dari area tampilan ke clipboard.
    *   **Ekspor ke Word**: Menyimpan teks saat ini dari area tampilan langsung ke file Word baru (tanpa melalui proses peningkatan AI lebih lanjut dari tab Rekaman).
    *   **Buka Folder Output**: Membuka folder yang telah Anda tentukan di Pengaturan Output menggunakan file explorer sistem Anda.

### Status Bar

Terletak di bagian bawah jendela aplikasi, status bar memberikan informasi real-time mengenai:
*   Status aplikasi saat ini (Siap, Merekam, Memproses, Error, dll.).
*   Waktu sistem saat ini.

---

## Bagaimana Sistem Bekerja

EchoScribe AI mengintegrasikan beberapa teknologi dan alur kerja untuk mengubah audio menjadi catatan terstruktur.

### Alur Kerja Utama

1.  **Input Audio**:
    *   Pengguna memilih sumber audio (mikrofon, sistem, atau dual).
    *   Saat perekaman dimulai, audio ditangkap dalam format PCM.
    *   Untuk "Rekaman Panjang", audio disimpan dalam chunk-chunk sementara berformat WAV. Jika tidak, disimpan sebagai satu file WAV sementara.
    *   Selama perekaman, data audio dikirim ke modul visualisasi untuk umpan balik real-time.

2.  **Penghentian Rekaman & Pra-pemrosesan**:
    *   Ketika rekaman dihentikan, chunk-chunk audio (jika ada) atau file audio tunggal disiapkan untuk transkripsi.
    *   Aplikasi mendeteksi konteks audio (`detect_audio_context`) seperti durasi, tingkat volume, rasio keheningan, untuk membantu memilih model transkripsi yang optimal.

3.  **Transkripsi (Speech-to-Text)**:
    *   Setiap chunk audio atau file audio tunggal dikirim ke mesin pengenalan suara yang dipilih:
        *   **Google Speech Recognition**: Menggunakan `recognizer.recognize_google()`.
        *   **Groq API (Whisper)**: Menggunakan `groq_client.audio.transcriptions.create()` dengan model Whisper yang dipilih (`select_optimal_transcription_model` berdasarkan konteks audio dan pilihan pengguna).
    *   Hasilnya adalah teks mentah dari audio.

4.  **Peningkatan AI (AI Enhancement)**:
    *   Jika diaktifkan, teks mentah dari setiap chunk (atau keseluruhan teks jika bukan rekaman panjang) dikirim ke API Groq (`groq_client.chat.completions.create()`) untuk peningkatan.
    *   `_analyze_content_characteristics`: Aplikasi menganalisis karakteristik konten (tipe, bahasa, kompleksitas, elemen seperti tabel/daftar) menggunakan kombinasi rule-based dan AI (DeepSeek via Groq) untuk klasifikasi yang lebih akurat.
    *   `_select_optimal_model`: Berdasarkan analisis konten, model AI Groq yang paling sesuai dipilih (misalnya, model yang lebih baik untuk konten teknis, rapat, atau naratif).
    *   `_create_content_adaptive_prompts`: Prompt yang sangat spesifik dan adaptif dibuat untuk LLM Groq, menginstruksikan AI bagaimana cara menyusun, memformat, dan meningkatkan teks berdasarkan jenis konten yang terdeteksi.
    *   AI memproses teks untuk:
        *   Memperbaiki tata bahasa dan ejaan.
        *   Menambahkan struktur (judul, subjudul, poin).
        *   Menghilangkan redundansi sambil mempertahankan detail penting.
        *   Memformat istilah teknis, data, dll.
    *   `enhance_document_cohesion`: Untuk rekaman panjang yang terdiri dari banyak chunk, setelah setiap chunk ditingkatkan, keseluruhan teks gabungan dapat diproses lagi untuk meningkatkan koherensi dan alur antar bagian.

5.  **Pembuatan Dokumen Word (.docx)**:
    *   Teks yang telah ditingkatkan kemudian diformat menjadi dokumen Word menggunakan pustaka `python-docx`.
    *   `_process_markdown_content`: Fungsi ini berperan seperti parser Markdown yang canggih. Ia mengurai teks yang telah ditingkatkan (yang diharapkan memiliki sintaks mirip Markdown dari AI) dan menerjemahkannya menjadi elemen-elemen Word:
        *   Heading (`#`, `##`, dll.) dengan ikon kontekstual dan styling adaptif.
        *   Daftar bullet dan bernomor (dengan nested level).
        *   Task list (`[ ]`, `[x]`).
        *   Kutipan (`> `).
        *   Blok kode (``` ```) dengan styling.
        *   Tabel (format Markdown).
        *   Callout/admonition (`:::note`, `:::warning`, dll.) dengan styling visual.
        *   Berbagai format inline (bold, italic, underline, strikethrough, code, highlight, dll.).
    *   `_setup_document_styles`, `_configure_page_layout`, `_add_document_header`, `_add_document_footer`: Fungsi-fungsi ini menyiapkan tema, style, layout halaman, header, dan footer yang profesional dan konsisten untuk dokumen.
    *   `finalize_document_formatting_enhanced`: Melakukan sentuhan akhir pada pemformatan dokumen.

6.  **Output & Penyimpanan**:
    *   Pratinjau teks ditampilkan di tab "Output".
    *   Dokumen `.docx` disimpan ke folder yang ditentukan pengguna.

### Komponen Kunci

*   **`VoiceToMarkdownApp`**: Kelas utama aplikasi, mengelola UI, state, dan alur kerja.
*   **`APIKeyDialog`**: Dialog untuk manajemen API key Groq.
*   **`ErrorHandler`**: Kelas terpusat untuk menangani kesalahan dan menampilkan pesan ke pengguna.
*   **Manajemen Konfigurasi (`setup_config_management`, `load_config`, `save_config`)**: Menyimpan dan memuat preferensi pengguna.
*   **Perekaman Audio**:
    *   `record_microphone_audio()`: Merekam dari mikrofon.
    *   `record_system_audio()`: Merekam audio sistem (mengandalkan `pyaudiowpatch`).
    *   `record_dual_audio()`: Merekam keduanya dan menggabungkannya.
    *   `save_audio_chunk()`: Menyimpan bagian audio untuk rekaman panjang.
*   **Visualisasi Audio**:
    *   Menggunakan `matplotlib` dan `numpy` untuk menampilkan waveform, bars, dll. secara real-time.
    *   `update_visualization_loop()` berjalan di thread terpisah.
*   **Pemrosesan Audio**:
    *   `process_audio_thread()`: Thread utama untuk transkripsi dan peningkatan.
    *   `process_standard_recording_enhanced()` dan `process_extended_recording_optimized()`: Logika untuk memproses rekaman pendek dan panjang.
    *   `transcribe_with_groq_whisper()`: Antarmuka ke API transkripsi Groq.
*   **Peningkatan AI**:
    *   `enhance_with_ai()`: Fungsi utama untuk mengirim teks ke Groq LLM untuk peningkatan.
    *   `_analyze_content_characteristics()`: Menganalisis teks untuk menentukan jenis dan struktur konten.
    *   `_select_optimal_model()`: Memilih model Groq yang paling sesuai.
    *   `_create_content_adaptive_prompts()`: Membuat prompt dinamis untuk LLM.
    *   `enhance_document_cohesion()`: Meningkatkan alur keseluruhan dokumen.
*   **Pembuatan Dokumen Word**:
    *   `save_as_word_document()`: Fungsi utama untuk membuat file `.docx`.
    *   `_process_markdown_content()`: Mengurai teks yang ditingkatkan dan memetakannya ke elemen Word.
    *   Berbagai fungsi `_add_...` dan `_style_...` untuk memformat elemen spesifik (heading, list, tabel, callout, dll.).

---

## Troubleshooting

*   **FFmpeg tidak ditemukan**:
    *   Pastikan FFmpeg terinstal dengan benar dan PATH-nya sudah diatur. Lihat bagian [Konfigurasi FFmpeg](#ffmpeg).
    *   Aplikasi mungkin tetap berjalan tetapi beberapa fitur pemrosesan audio mungkin gagal.
*   **Perekaman Audio Sistem Tidak Berfungsi**:
    *   Fitur ini sangat bergantung pada `PyAudioWPatch` dan umumnya hanya berfungsi di Windows.
    *   Pastikan `PyAudioWPatch` terinstal.
    *   Periksa pengaturan suara Windows Anda. Pastikan perangkat output default sudah benar dan "Stereo Mix" (atau sejenisnya) diaktifkan jika tersedia.
    *   Jika gagal, coba gunakan mode "Dual Recording" sebagai alternatif, atau rekam audio sistem menggunakan software lain dan impor filenya (fitur impor belum ada, ini saran umum).
    *   Aplikasi memiliki dialog troubleshooting (`_show_enhanced_system_audio_troubleshooting`) yang akan muncul jika ada masalah.
*   **Error API Key Groq**:
    *   Pastikan API key Anda valid dan memiliki kuota yang cukup.
    *   Periksa koneksi internet Anda.
    *   Dialog API key di Pengaturan memungkinkan Anda memasukkan ulang atau menggunakan API key default.
*   **Mikrofon Tidak Terdeteksi**:
    *   Pastikan mikrofon terhubung dengan benar dan diizinkan oleh sistem operasi Anda.
    *   Gunakan tombol `⟳` (Refresh) di tab Rekaman untuk memperbarui daftar mikrofon.
*   **Kualitas Transkripsi Rendah**:
    *   Pastikan lingkungan rekaman minim noise.
    *   Gunakan mikrofon berkualitas baik dan dekat dengan sumber suara.
    *   Pilih bahasa yang sesuai di Pengaturan.
*   **Aplikasi Lambat atau Macet Saat Memproses Rekaman Panjang**:
    *   Rekaman yang sangat panjang memerlukan sumber daya signifikan. Pastikan "Rekaman Panjang" diaktifkan di Pengaturan.
    *   Coba kurangi "Ukuran Penggalan" agar pemrosesan dilakukan lebih sering dalam porsi lebih kecil, atau perbesar untuk mengurangi frekuensi panggilan API (tapi meningkatkan risiko timeout jika chunk terlalu besar).
*   **File Log**:
    *   Aplikasi menyimpan log detail di folder `.echoscribe` di direktori home pengguna Anda (misalnya, `C:\Users\NamaAnda\.echoscribe\echoscribe.log`). File log ini sangat berguna untuk mendiagnosis masalah.

---

## Dependensi Utama

Aplikasi ini menggunakan beberapa pustaka Python utama:

*   `customtkinter`: Untuk antarmuka pengguna grafis modern.
*   `speech_recognition`: Untuk interaksi dengan API pengenalan suara.
*   `pydub` & `wave` & `audioop`: Untuk manipulasi dan pemrosesan file audio.
*   `groq`: Untuk berinteraksi dengan API Groq (Whisper dan LLM).
*   `python-docx`: Untuk membuat dan memanipulasi file Word (.docx).
*   `matplotlib` & `numpy`: Untuk visualisasi audio real-time.
*   `pyaudio` & `sounddevice`: Untuk perekaman dan pemutaran audio.
*   `pyaudiowpatch` (Windows): Untuk perekaman audio sistem yang lebih baik di Windows.
*   `threading`: Untuk operasi latar belakang agar UI tetap responsif.

Disarankan untuk menginstal semua dependensi menggunakan `pip install -r requirements.txt` (jika file tersebut disediakan) atau secara manual.

---

## Kontribusi

Kontribusi untuk EchoScribe AI sangat diterima! Jika Anda ingin berkontribusi, silakan:
1.  Fork repository ini.
2.  Buat branch fitur baru (`git checkout -b fitur/NamaFitur`).
3.  Commit perubahan Anda (`git commit -am 'Menambahkan fitur X'`).
4.  Push ke branch (`git push origin fitur/NamaFitur`).
5.  Buat Pull Request baru.

Harap pastikan kode Anda mengikuti standar kualitas dan menyertakan dokumentasi yang relevan.

---

**Catatan Penting:**

1.  **Ganti Placeholder**:
    *   `link_ke_logo.png`: Ganti dengan URL atau path ke file logo Anda.
    *   `link_ke_screenshot.png`: Ganti dengan URL atau path ke screenshot aplikasi Anda.
    *   `link_ke_audio_demo.mp3`: Ganti dengan URL ke demo audio jika ada.
    *   `https://link_ke_repository_anda.git`: Ganti dengan URL repository Git Anda.
    *   `nama_folder_repository`: Ganti dengan nama folder setelah clone.
    *   `nama_file_utama_anda.py`: Ganti dengan nama file Python skrip utama Anda.
    *   File `LICENSE`: Buat file ini jika belum ada.

