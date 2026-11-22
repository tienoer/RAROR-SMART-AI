<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>RAPOR SMART AI — Dari Data Menjadi Aksi, Dari Rapor Menjadi Perubahan</title>

  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>

  <!-- Google Fonts: Plus Jakarta Sans, Inter & JetBrains Mono -->
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600;700&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap" rel="stylesheet" />

  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>

  <!-- SheetJS (XLSX) for Direct Official Excel Rapor Pendidikan Import & Export -->
  <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>

  <!-- PDF.js for Official PDF Rapor Pendidikan Extraction -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>

  <script>
    tailwind.config = {
      theme: {
        extend: {
          fontFamily: {
            sans: ['Inter', 'system-ui', 'sans-serif'],
            display: ['Plus Jakarta Sans', 'Inter', 'sans-serif'],
            mono: ['JetBrains Mono', 'monospace'],
          },
          colors: {
            emerald: {
              850: '#064e3b',
              925: '#033629',
              950: '#022c22',
            },
            gold: {
              50: '#fffbeb',
              100: '#fef3c7',
              200: '#fde68a',
              400: '#fbbf24',
              500: '#f59e0b',
              600: '#d97706',
              700: '#b45309',
            }
          }
        }
      }
    };
  </script>

  <style>
    body {
      font-family: 'Inter', system-ui, sans-serif;
      background-color: #f4f7f6;
      color: #0f172a;
    }
    h1, h2, h3, .font-heading {
      font-family: 'Plus Jakarta Sans', 'Inter', sans-serif;
    }
    .emerald-gradient {
      background: linear-gradient(135deg, #064e3b 0%, #047857 100%);
    }
    .emerald-dark-gradient {
      background: linear-gradient(160deg, #022c22 0%, #064e3b 60%, #065f46 100%);
    }
    .gold-gradient {
      background: linear-gradient(135deg, #d97706 0%, #f59e0b 100%);
    }
    .card-elevated {
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 1rem;
      box-shadow: 0 2px 10px -2px rgba(6, 78, 59, 0.05);
    }
    .custom-scrollbar::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    .custom-scrollbar::-webkit-scrollbar-track {
      background: transparent;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb {
      background: rgba(148, 163, 184, 0.45);
      border-radius: 9999px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover {
      background: rgba(100, 116, 139, 0.7);
    }
    @media print {
      .no-print { display: none !important; }
      body { background: #ffffff !important; }
    }
  </style>
</head>
<body class="antialiased selection:bg-emerald-800 selection:text-white">

  <!-- TOP DEMO BANNER (Visible only when Demo Mode is active) -->
  <div id="demoBanner" class="hidden no-print bg-amber-500 text-slate-950 px-4 py-2 text-xs sm:text-sm font-semibold flex flex-wrap items-center justify-between gap-2 shadow-sm">
    <div class="flex items-center gap-2">
      <i data-lucide="flask-conical" class="w-4 h-4 shrink-0"></i>
      <span><strong>MODE DEMO AKTIF:</strong> Anda sedang melihat simulasi data contoh. Jangan gunakan data demo sebagai laporan resmi sekolah.</span>
    </div>
    <button onclick="exitDemoMode()" class="px-3 py-1 rounded-lg bg-slate-900 text-white hover:bg-slate-800 text-xs font-bold transition">
      Kembali ke Data Sekolah Asli
    </button>
  </div>

  <!-- MAIN APP SHELL -->
  <div class="min-h-screen flex flex-col lg:flex-row">

    <!-- MOBILE SIDEBAR OVERLAY -->
    <div id="mobileSidebarOverlay" onclick="toggleMobileSidebar(false)" class="fixed inset-0 bg-slate-950/60 backdrop-blur-xs z-40 hidden lg:hidden"></div>

    <!-- SIDEBAR NAVIGATION -->
    <aside id="appSidebar" class="no-print fixed lg:sticky top-0 left-0 z-50 h-screen w-72 emerald-dark-gradient text-white flex flex-col transition-transform duration-300 -translate-x-full lg:translate-x-0 shrink-0 shadow-xl">
      <!-- Brand Header -->
      <div class="p-5 border-b border-emerald-800/70">
        <div class="flex items-center justify-between">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 rounded-xl gold-gradient flex items-center justify-center text-slate-950 font-extrabold shadow-md shrink-0">
              <i data-lucide="sparkles" class="w-5 h-5"></i>
            </div>
            <div>
              <div class="font-heading font-extrabold text-base tracking-tight text-white flex items-center gap-1.5">
                RAPOR SMART <span class="text-amber-400">AI</span>
              </div>
              <p class="text-[10px] text-emerald-200 leading-tight mt-0.5">Dari Data Menjadi Aksi, Dari Rapor Menjadi Perubahan.</p>
            </div>
          </div>
          <button onclick="toggleMobileSidebar(false)" class="lg:hidden p-1.5 rounded-lg text-emerald-200 hover:bg-emerald-800/60">
            <i data-lucide="x" class="w-5 h-5"></i>
          </button>
        </div>
      </div>

      <!-- Active Planning Year Selector -->
      <div class="px-4 py-3 border-b border-emerald-800/60 bg-emerald-950/40">
        <label class="block text-[10px] uppercase tracking-wider font-bold text-emerald-300 mb-1">Tahun Perencanaan Aktif</label>
        <div class="flex items-center gap-1.5">
          <select id="activeYearSelect" onchange="switchPlanningYear(this.value)" class="w-full bg-emerald-900/90 border border-emerald-700 text-white text-xs font-semibold rounded-lg px-2.5 py-2 focus:outline-none focus:border-amber-400">
          </select>
          <button onclick="openNewYearModal()" title="Tambah / Duplikasi Tahun Perencanaan" class="p-2 rounded-lg bg-emerald-800 hover:bg-emerald-700 text-amber-300 border border-emerald-700 transition shrink-0">
            <i data-lucide="folder-plus" class="w-4 h-4"></i>
          </button>
        </div>
      </div>

      <!-- Navigation Menu -->
      <nav id="sidebarNav" class="flex-1 overflow-y-auto custom-scrollbar px-3 py-3 space-y-1 text-xs">
        <!-- Rendered dynamically -->
      </nav>

      <!-- Sidebar Footer Actions -->
      <div class="p-3 border-t border-emerald-800/70 bg-emerald-950/50 space-y-2">
        <button onclick="toggleAssistantDrawer(true)" class="w-full py-2.5 px-3 rounded-xl gold-gradient text-slate-950 font-bold text-xs flex items-center justify-center gap-2 shadow-md hover:brightness-105 transition">
          <i data-lucide="bot" class="w-4 h-4"></i>
          <span>ASISTEN RAPOR SMART AI</span>
        </button>
        <div class="grid grid-cols-2 gap-1.5">
          <button onclick="openGuideModal()" class="py-1.5 px-2 rounded-lg bg-emerald-900/80 hover:bg-emerald-800 text-emerald-100 text-[11px] font-medium flex items-center justify-center gap-1 border border-emerald-800">
            <i data-lucide="help-circle" class="w-3.5 h-3.5 text-amber-400"></i>
            <span>Panduan</span>
          </button>
          <button id="demoToggleBtn" onclick="toggleDemoMode()" class="py-1.5 px-2 rounded-lg bg-emerald-900/80 hover:bg-emerald-800 text-emerald-100 text-[11px] font-medium flex items-center justify-center gap-1 border border-emerald-800">
            <i data-lucide="play-circle" class="w-3.5 h-3.5 text-amber-400"></i>
            <span>Mode Demo</span>
          </button>
        </div>
      </div>
    </aside>

    <!-- MAIN CONTENT WORKSPACE -->
    <div class="flex-1 flex flex-col min-w-0">

      <!-- TOP HEADER BAR -->
      <header class="no-print sticky top-0 z-30 bg-white/95 backdrop-blur-md border-b border-slate-200 px-4 sm:px-6 py-3">
        <div class="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3">
          <div class="flex items-center gap-3 min-w-0">
            <button onclick="toggleMobileSidebar(true)" class="lg:hidden p-2 rounded-xl border border-slate-200 text-slate-700 hover:bg-slate-100">
              <i data-lucide="menu" class="w-5 h-5"></i>
            </button>
            <div class="min-w-0">
              <div class="flex items-center gap-2 flex-wrap">
                <h1 id="headerSchoolName" class="font-heading font-bold text-sm sm:text-base text-slate-900 truncate">Satuan Pendidikan Belum Diisi</h1>
                <span id="headerNpsnBadge" class="px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-800 border border-emerald-200 font-mono text-[11px] font-semibold">NPSN: -</span>
                <span id="headerJenjangBadge" class="hidden sm:inline-block px-2 py-0.5 rounded-md bg-amber-50 text-amber-800 border border-amber-200 text-[11px] font-semibold">Jenjang: -</span>
              </div>
              <p id="headerSubtitle" class="text-[11px] text-slate-500 truncate mt-0.5">Tahun Rapor: - • Tahun Anggaran/Perencanaan: -</p>
            </div>
          </div>

          <div class="flex items-center gap-2 flex-wrap">
            <span id="autoSaveBadge" class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-800 text-[11px] font-semibold border border-emerald-200">
              <span class="w-2 h-2 rounded-full bg-emerald-600 animate-pulse"></span>
              <span>Tersimpan Otomatis</span>
            </span>
            <button onclick="setActiveTab('ekspor')" title="Unduh Dokumen PDF, Word, dan Excel" class="px-3 py-2 rounded-xl gold-gradient text-slate-950 hover:brightness-105 text-xs font-bold flex items-center gap-1.5 shadow-xs">
              <i data-lucide="download" class="w-4 h-4"></i>
              <span>Unduh Laporan</span>
            </button>
            <button onclick="exportProjectJson()" title="Ekspor Cadangan Data Proyek (.JSON)" class="p-2 rounded-xl border border-slate-200 text-slate-700 hover:bg-slate-50 text-xs font-medium flex items-center gap-1.5">
              <i data-lucide="download-cloud" class="w-4 h-4 text-emerald-700"></i>
              <span class="hidden md:inline">Cadangkan</span>
            </button>
            <label title="Impor Cadangan Data Proyek (.JSON)" class="cursor-pointer p-2 rounded-xl border border-slate-200 text-slate-700 hover:bg-slate-50 text-xs font-medium flex items-center gap-1.5">
              <i data-lucide="upload-cloud" class="w-4 h-4 text-amber-600"></i>
              <span class="hidden md:inline">Pulihkan</span>
              <input type="file" accept=".json" onchange="importProjectJson(event)" class="hidden" />
            </label>
          </div>
        </div>
      </header>

      <!-- DYNAMIC MODULE CONTAINER -->
      <main id="mainModuleContainer" class="flex-1 p-4 sm:p-6 max-w-7xl w-full mx-auto space-y-6">
        <!-- Rendered dynamically by activeTab -->
      </main>

      <!-- FOOTER -->
      <footer class="no-print border-t border-slate-200 bg-white px-4 sm:px-6 py-4 text-xs text-slate-500">
        <div class="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-2">
          <div>
            <strong class="text-emerald-900">RAPOR SMART AI</strong> — Platform Perencanaan Berbasis Data (PBD) Terintegrasi untuk Kepala Sekolah, Guru, Bendahara BOS, &amp; Tim Pengembang Sekolah.
          </div>
          <div class="text-[11px] text-slate-400">
            Catatan: Dokumen RKJM, RKT, dan Rancangan ARKAS merupakan kertas kerja perencanaan internal sekolah sebelum diinput ke sistem resmi.
          </div>
        </div>
      </footer>
    </div>
  </div>

  <!-- FLOATING AI ASSISTANT DRAWER -->
  <div id="assistantOverlay" onclick="toggleAssistantDrawer(false)" class="fixed inset-0 bg-slate-950/50 backdrop-blur-xs z-50 hidden"></div>
  <aside id="assistantDrawer" class="no-print fixed top-0 right-0 z-50 h-screen w-full sm:w-[430px] bg-white shadow-2xl border-l border-slate-200 flex flex-col transition-transform duration-300 translate-x-full">
    <div class="emerald-gradient text-white p-4 flex items-center justify-between">
      <div class="flex items-center gap-2.5">
        <div class="w-9 h-9 rounded-xl gold-gradient text-slate-950 flex items-center justify-center font-bold shadow">
          <i data-lucide="bot" class="w-5 h-5"></i>
        </div>
        <div>
          <h3 class="font-heading font-bold text-sm">ASISTEN RAPOR SMART AI</h3>
          <p class="text-[11px] text-emerald-100">Berbasis data aktual sekolah Anda • Tanpa mengarang data</p>
        </div>
      </div>
      <button onclick="toggleAssistantDrawer(false)" class="p-1.5 rounded-lg hover:bg-white/15">
        <i data-lucide="x" class="w-5 h-5"></i>
      </button>
    </div>

    <!-- Quick Prompt Chips -->
    <div class="p-3 bg-slate-50 border-b border-slate-200 flex gap-1.5 overflow-x-auto custom-scrollbar shrink-0">
      <button onclick="askAssistantQuick('Ringkas hasil Rapor Pendidikan & nilai capaian sekolah saya saat ini')" class="whitespace-nowrap px-2.5 py-1 rounded-full bg-white border border-emerald-200 text-emerald-900 text-[11px] font-semibold hover:bg-emerald-50">Ringkas Nilai Capaian</button>
      <button onclick="askAssistantQuick('Apa indikator prioritas utama dan dugaan akar masalahnya?')" class="whitespace-nowrap px-2.5 py-1 rounded-full bg-white border border-emerald-200 text-emerald-900 text-[11px] font-semibold hover:bg-emerald-50">Cek Prioritas &amp; Akar Masalah</button>
      <button onclick="askAssistantQuick('Periksa keterhubungan RKJM, RKT, dan anggaran ARKAS saya')" class="whitespace-nowrap px-2.5 py-1 rounded-full bg-white border border-emerald-200 text-emerald-900 text-[11px] font-semibold hover:bg-emerald-50">Audit Benang Merah</button>
      <button onclick="askAssistantQuick('Hitung rekap rancangan anggaran ARKAS per komponen')" class="whitespace-nowrap px-2.5 py-1 rounded-full bg-white border border-emerald-200 text-emerald-900 text-[11px] font-semibold hover:bg-emerald-50">Rekap Anggaran ARKAS</button>
    </div>

    <!-- Chat Messages -->
    <div id="assistantMessages" class="flex-1 overflow-y-auto custom-scrollbar p-4 space-y-3 bg-slate-50/50 text-xs">
    </div>

    <!-- Chat Input -->
    <form onsubmit="handleAssistantSubmit(event)" class="p-3 bg-white border-t border-slate-200 flex items-center gap-2">
      <input id="assistantInput" type="text" placeholder="Tanyakan analisis rapor, RKJM, RKT, atau anggaran..." class="flex-1 rounded-xl border border-slate-300 px-3.5 py-2.5 text-xs focus:border-emerald-700 focus:outline-none" />
      <button type="submit" class="p-2.5 rounded-xl emerald-gradient text-white hover:opacity-95 transition shrink-0">
        <i data-lucide="send" class="w-4 h-4"></i>
      </button>
    </form>
  </aside>

  <!-- CUSTOM MODAL CONTAINER -->
  <div id="customModalBackdrop" class="fixed inset-0 bg-slate-950/60 backdrop-blur-xs z-50 hidden items-center justify-center p-3 sm:p-4 overflow-y-auto">
    <div id="customModalBox" class="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-2xl overflow-hidden my-auto">
      <!-- Modal content injected here -->
    </div>
  </div>

  <!-- TOAST NOTIFICATION CONTAINER -->
  <div id="toastContainer" class="no-print fixed bottom-4 right-4 z-50 flex flex-col gap-2 max-w-sm w-full pointer-events-none px-4 sm:px-0"></div>

  <!-- APPLICATION SCRIPT -->
  <script>
    // Configure PDF.js worker if available
    if (window.pdfjsLib) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }

    // =========================================================================
    // 1. MASTER REFERENCE CATALOGS (Standard PBD Indicators & ARKAS Components)
    // =========================================================================
    const KATALOG_INDIKATOR_STANDAR = [
      { kode: "A.1", nama: "Kemampuan Literasi", dimensi: "D.1 Mutu dan Relevansi Hasil Belajar Peserta Didik" },
      { kode: "A.2", nama: "Kemampuan Numerasi", dimensi: "D.1 Mutu dan Relevansi Hasil Belajar Peserta Didik" },
      { kode: "A.3", nama: "Karakter", dimensi: "D.1 Mutu dan Relevansi Hasil Belajar Peserta Didik" },
      { kode: "D.1", nama: "Kualitas Pembelajaran", dimensi: "D.2 Mutu dan Relevansi Pembelajaran" },
      { kode: "D.2", nama: "Refleksi dan Perbaikan Pembelajaran oleh Guru", dimensi: "D.2 Mutu dan Relevansi Pembelajaran" },
      { kode: "D.3", nama: "Kepemimpinan Instruksional", dimensi: "D.2 Mutu dan Relevansi Pembelajaran" },
      { kode: "D.4", nama: "Iklim Keamanan Satuan Pendidikan", dimensi: "D.3 Lingkungan Belajar yang Aman, Nyaman, dan Inklusif" },
      { kode: "D.6", nama: "Iklim Kesetaraan Gender", dimensi: "D.3 Lingkungan Belajar yang Aman, Nyaman, dan Inklusif" },
      { kode: "D.8", nama: "Iklim Kebinekaan", dimensi: "D.3 Lingkungan Belajar yang Aman, Nyaman, dan Inklusif" },
      { kode: "D.10", nama: "Iklim Inklusivitas", dimensi: "D.3 Lingkungan Belajar yang Aman, Nyaman, dan Inklusif" },
      { kode: "E.1", nama: "Partisipasi Warga Satuan Pendidikan", dimensi: "E. Pengelolaan Satuan Pendidikan yang Partisipatif, Transparan, dan Akuntabel" },
      { kode: "E.2", nama: "Proporsi Pemanfaatan Sumber Daya Sekolah untuk Peningkatan Mutu", dimensi: "E. Pengelolaan Satuan Pendidikan yang Partisipatif, Transparan, dan Akuntabel" },
      { kode: "E.3", nama: "Pemanfaatan TIK untuk Pengelolaan Anggaran", dimensi: "E. Pengelolaan Satuan Pendidikan yang Partisipatif, Transparan, dan Akuntabel" },
      { kode: "C.1", nama: "Proporsi GTK Bersertifikat", dimensi: "C. Kompetensi dan Kinerja PTK" },
      { kode: "C.2", nama: "Proporsi GTK Penggerak", dimensi: "C. Kompetensi dan Kinerja PTK" }
    ];

    const KOMPONEN_BOSP_STANDAR = [
      "01. Pengembangan Perpustakaan dan/atau Layanan Pojok Baca",
      "02. Penerimaan Peserta Didik Baru (PPDB)",
      "03. Pelaksanaan Kegiatan Pembelajaran dan Bermain",
      "04. Pelaksanaan Kegiatan Evaluasi/Asesmen Pembelajaran",
      "05. Pelaksanaan Administrasi Kegiatan Satuan Pendidikan",
      "06. Pengembangan Profesi Pendidik dan Tenaga Kependidikan",
      "07. Pembiayaan Langganan Daya dan Jasa",
      "08. Pemeliharaan Sarana dan Prasarana Satuan Pendidikan",
      "09. Penyediaan Alat Multimedia Pembelajaran",
      "10. Pembayaran Honor (Sesuai Ketentuan BOSP Berlaku — Perlu Verifikasi Juknis)"
    ];

    const SUMBER_DANA_STANDAR = [
      "BOSP Reguler",
      "BOSP Kinerja",
      "BOSDA / APBD",
      "Swadaya / Komite Sekolah",
      "Sumber Sah Lainnya (Perlu Verifikasi)"
    ];

    const NAV_ITEMS = [
      { id: "dashboard", label: "Dashboard Utama", icon: "layout-dashboard", step: "0" },
      { id: "input", label: "1. Input Rapor & Nilai Capaian", icon: "file-spreadsheet", step: "1" },
      { id: "analisis", label: "2. Analisis Capaian", icon: "bar-chart-3", step: "2" },
      { id: "akar", label: "3. Akar Masalah (5-Why)", icon: "git-branch", step: "3" },
      { id: "prioritas", label: "4. Penentuan Prioritas", icon: "list-ordered", step: "4" },
      { id: "rekomendasi", label: "5. Rekomendasi Program", icon: "lightbulb", step: "5" },
      { id: "rkjm", label: "6. Penyusunan RKJM", icon: "compass", step: "6" },
      { id: "rkt", label: "7. Penyusunan RKT", icon: "calendar-days", step: "7" },
      { id: "arkas", label: "8. Rancangan ARKAS", icon: "wallet", step: "8" },
      { id: "koneksi", label: "Benang Merah Dokumen", icon: "network", step: "9" },
      { id: "monev", label: "Monitoring & Evaluasi", icon: "activity", step: "10" },
      { id: "ekspor", label: "Laporan & Unduh", icon: "printer", step: "11" }
    ];

    // =========================================================================
    // 2. IN-MEMORY STATE ARCHITECTURE (Multi-Year + Separate Demo State)
    // =========================================================================
    function createEmptyYearData(yearLabel = "2026") {
      const yNum = parseInt(yearLabel, 10) || 2026;
      return {
        yearLabel: String(yNum),
        school: {
          nama: "",
          npsn: "",
          alamat: "",
          kecamatan: "",
          kabupaten: "",
          provinsi: "",
          jenjang: "SD",
          tahunRapor: String(yNum),
          tahunPerencanaan: String(yNum),
          periodeRkjmAwal: String(yNum),
          periodeRkjmAkhir: String(yNum + 3),
          kepalaSekolah: "",
          nipKepalaSekolah: "",
          bendaharaBos: ""
        },
        indicators: [],
        rootCauses: [],
        priorities: [],
        recommendations: [],
        rkjm: [],
        rkt: [],
        arkas: [],
        monev: [],
        history: [
          {
            waktu: new Date().toLocaleString("id-ID"),
            modul: "Sistem",
            aksi: "Inisialisasi Tahun Perencanaan",
            detail: `Membuat ruang kerja perencanaan tahun ${yNum}.`
          }
        ]
      };
    }

    function createDemoYearData() {
      const demo = createEmptyYearData("2026");
      demo.school = {
        nama: "SD NEGERI HARAPAN BANGSA 01 (DATA DEMO)",
        npsn: "20299901",
        alamat: "Jl. Pendidikan Merdeka No. 45",
        kecamatan: "Sukamaju",
        kabupaten: "Kab. Bandung",
        provinsi: "Jawa Barat",
        jenjang: "SD",
        tahunRapor: "2026",
        tahunPerencanaan: "2026",
        periodeRkjmAwal: "2026",
        periodeRkjmAkhir: "2029",
        kepalaSekolah: "Dra. Siti Aminah, M.Pd.",
        nipKepalaSekolah: "19740512 199803 2 004",
        bendaharaBos: "Budi Santoso, S.Pd."
      };

      const ind1 = {
        id: "demo-ind-1",
        kode: "A.2",
        nama: "Kemampuan Numerasi",
        dimensi: "D.1 Mutu dan Relevansi Hasil Belajar Peserta Didik",
        nilai: 52.4,
        kategori: "Kurang",
        perubahan: -4.2,
        peringkat: "Peringkat Bawah (Kuantil 2 Kab/Kota)",
        rekomendasiRapor: "Peningkatan kompetensi GTK dalam pembelajaran numerasi kontekstual dan pemanfaatan platform belajar.",
        terverifikasi: true
      };
      const ind2 = {
        id: "demo-ind-2",
        kode: "D.1",
        nama: "Kualitas Pembelajaran",
        dimensi: "D.2 Mutu dan Relevansi Pembelajaran",
        nilai: 58.1,
        kategori: "Sedang",
        perubahan: -1.8,
        peringkat: "Peringkat Menengah Bawah (Kuantil 3)",
        rekomendasiRapor: "Aktivasi komunitas belajar sekolah untuk refleksi metode pembelajaran interaktif dan manajemen kelas.",
        terverifikasi: true
      };
      const ind3 = {
        id: "demo-ind-3",
        kode: "A.1",
        nama: "Kemampuan Literasi",
        dimensi: "D.1 Mutu dan Relevansi Hasil Belajar Peserta Didik",
        nilai: 66.8,
        kategori: "Sedang",
        perubahan: 2.5,
        peringkat: "Peringkat Menengah Atas (Kuantil 4)",
        rekomendasiRapor: "Penguatan pojok baca kelas dan penyediaan buku bacaan nonteks berjenjang.",
        terverifikasi: true
      };
      const ind4 = {
        id: "demo-ind-4",
        kode: "D.4",
        nama: "Iklim Keamanan Satuan Pendidikan",
        dimensi: "D.3 Lingkungan Belajar yang Aman, Nyaman, dan Inklusif",
        nilai: 78.5,
        kategori: "Baik",
        perubahan: 4.1,
        peringkat: "Peringkat Atas (Kuantil 5)",
        rekomendasiRapor: "Pertahankan program Tim Pencegahan dan Penanganan Kekerasan (TPPK) sekolah.",
        terverifikasi: true
      };

      demo.indicators = [ind1, ind2, ind3, ind4];

      demo.rootCauses = [
        {
          id: "demo-rc-1",
          indicatorId: ind1.id,
          masalahUtama: "Capaian Kemampuan Numerasi berada pada kategori Kurang (52,40) dan mengalami penurunan 4,20 poin dibanding tahun sebelumnya.",
          dugaanPenyebab: [
            { id: "c1", teks: "Pembelajaran matematika masih berfokus pada hafalan rumus prosedural tanpa konteks pemecahan masalah sehari-hari.", terkonfirmasi: true },
            { id: "c2", teks: "Keterbatasan alat peraga manipulatif numerasi di kelas rendah dan kelas tinggi.", terkonfirmasi: true },
            { id: "c3", teks: "Belum optimalnya analisis diagnostik awal kemampuan numerasi siswa di awal semester.", terkonfirmasi: false }
          ],
          whyChain: [
            "Why 1: Skor kompetensi domain bilangan dan geometri pada ANBK menurun.",
            "Why 2: Siswa kesulitan memahami soal cerita numerasi berbasis konteks kehidupan nyata.",
            "Why 3: Guru jarang menyusun soal asesmen formatif berbasis literasi-numerasi.",
            "Why 4: KKG internal/Komunitas Belajar belum membedah modul ajar numerasi secara rutin.",
            "Why 5: Belum ada jadwal supervisi klinis dan pendampingan perangkat ajar numerasi yang terprogram."
          ],
          buktiPendukung: "Nilai Rapor Pendidikan A.2 = 52,40 (turun -4,20); hasil asesmen diagnostik awal semester ganjil menunjukkan 48% siswa kelas IV-V kesulitan menyelesaikan soal penalaran dua langkah.",
          pertanyaanRefleksi: [
            "Bagaimana strategi pembelajaran numerasi yang selama ini kita terapkan di kelas?",
            "Media konkret apa yang paling dibutuhkan siswa agar konsep bilangan dan geometri mudah dipahami?",
            "Seberapa rutin kita mendiskusikan hasil kesalahan umum siswa dalam komunitas belajar?"
          ],
          dataTambahan: "Rekap nilai formatif matematika per fase, daftar inventaris alat peraga matematika, dan catatan supervisi kelas.",
          dampakPesertaDidik: "Peserta didik mengalami kesulitan berpikir logis, menganalisis informasi kuantitatif, dan berisiko tertinggal pada jenjang pendidikan berikutnya.",
          alternatifSolusi: "1) Workshop & pendampingan Kombel penyusunan modul ajar numerasi kontekstual; 2) Pengadaan alat peraga numerasi; 3) Klinik numerasi 15 menit sebelum pelajaran."
        },
        {
          id: "demo-rc-2",
          indicatorId: ind2.id,
          masalahUtama: "Kualitas Pembelajaran berada pada kategori Sedang (58,10) dan turun 1,80 poin terutama pada sub-indikator metode pembelajaran interaktif.",
          dugaanPenyebab: [
            { id: "c4", teks: "Metode ceramah masih dominan sehingga partisipasi aktif siswa dalam diskusi kelompok rendah.", terkonfirmasi: true },
            { id: "c5", teks: "Umpan balik konstruktif guru terhadap hasil kerja siswa belum terdokumentasi rutin.", terkonfirmasi: false }
          ],
          whyChain: [
            "Why 1: Skor dukungan psikologis dan aktivasi kognitif pada survei lingkungan belajar menurun.",
            "Why 2: Guru kesulitan mengelola pembelajaran berdiferensiasi pada kelas dengan kemampuan beragam.",
            "Why 3: Berbagi praktik baik antar guru melalui observasi rekan sejawat belum berjalan konsisten."
          ],
          buktiPendukung: "Nilai Rapor Pendidikan D.1 = 58,10 (-1,80); catatan observasi kelas semester lalu.",
          pertanyaanRefleksi: [
            "Apakah siswa sudah diberi ruang bertanya dan berdiskusi secara aktif dalam setiap pertemuan?",
            "Bagaimana mekanisme umpan balik yang membuat siswa tahu cara memperbaiki pekerjaannya?"
          ],
          dataTambahan: "Instrumen observasi pembelajaran teman sejawat dan survei umpan balik murid.",
          dampakPesertaDidik: "Motivasi belajar siswa menurun dan pemahaman konsep menjadi dangkal.",
          alternatifSolusi: "Pelaksanaan Lesson Study / observasi rekan sejawat bulanan melalui Komunitas Belajar sekolah."
        }
      ];

      demo.priorities = [
        {
          id: "demo-prio-1",
          indicatorId: ind1.id,
          skorCapaian: 5,
          skorTren: 5,
          skorDampak: 5,
          skorUrgensi: 5,
          skorKemampuan: 4,
          totalSkor: 24,
          peringkatManual: 1,
          alasanPrioritas: "Indikator standar pelayanan minimal (SPM) pendidikan dengan nilai terendah (52,40 — Kurang) dan mengalami penurunan tajam (-4,20), berdampak langsung pada daya nalar seluruh peserta didik."
        },
        {
          id: "demo-prio-2",
          indicatorId: ind2.id,
          skorCapaian: 4,
          skorTren: 4,
          skorDampak: 5,
          skorUrgensi: 4,
          skorKemampuan: 5,
          totalSkor: 22,
          peringkatManual: 2,
          alasanPrioritas: "Menjadi pengungkit utama perbaikan hasil belajar (Literasi & Numerasi) di kelas dan dapat segera dibenahi melalui penguatan Komunitas Belajar guru."
        }
      ];

      demo.recommendations = [
        {
          id: "demo-rec-1",
          indicatorId: ind1.id,
          permasalahan: "Rendahnya capaian Kemampuan Numerasi (52,40) akibat pembelajaran belum kontekstual dan minimnya media konkret.",
          tujuanPerbaikan: "Meningkatkan skor Kemampuan Numerasi minimal 8 poin menjadi kategori Sedang/Baik pada Rapor Pendidikan berikutnya.",
          sasaranProgram: "Seluruh guru kelas I–VI (12 orang) dan 180 peserta didik.",
          bentukKegiatan: "IHT Pembelajaran Numerasi Kontekstual, Penguatan Kombel Bulanan, dan Pengadaan Kit Alat Peraga Numerasi.",
          langkahPelaksanaan: "1. Asesmen diagnostik awal numerasi; 2. IHT penyusunan modul ajar & soal numerasi; 3. Pengadaan alat peraga; 4. Implementasi & supervisi kelas; 5. Evaluasi sumatif.",
          penanggungJawab: "Wakil Kepala Sekolah Bidang Kurikulum & Koordinator Kombel",
          indikatorKeberhasilan: "100% guru menerapkan modul ajar numerasi kontekstual; skor rata-rata numerasi naik >= 8 poin.",
          targetCapaian: "Nilai Rapor A.2 meningkat dari 52,40 menjadi >= 60,50.",
          waktuPelaksanaan: "Juli – November 2026",
          kebutuhanSumberDaya: "Narasumber praktisi numerasi, Kit Alat Peraga Matematika, ATK pelatihan.",
          perkiraanAnggaran: 8500000,
          disetujui: true
        },
        {
          id: "demo-rec-2",
          indicatorId: ind2.id,
          permasalahan: "Kualitas pembelajaran (58,10) mengalami penurunan pada aspek pembelajaran interaktif dan berdiferensiasi.",
          tujuanPerbaikan: "Meningkatkan skor Kualitas Pembelajaran menjadi >= 65,00 melalui budaya refleksi dan observasi rekan sejawat.",
          sasaranProgram: "Seluruh pendidik satuan pendidikan (12 guru).",
          bentukKegiatan: "Workshop Pembelajaran Interaktif Berdiferensiasi & Siklus Observasi Rekan Sejawat.",
          langkahPelaksanaan: "1. Penyusunan jadwal observasi kelas; 2. Lokakarya praktik baik pembelajaran aktif; 3. Pelaksanaan observasi & umpan balik; 4. Refleksi akhir semester.",
          penanggungJawab: "Kepala Sekolah & Tim Pengembang Kurikulum",
          indikatorKeberhasilan: "Minimal 2 kali observasi rekan sejawat per guru dalam satu semester.",
          targetCapaian: "Nilai Rapor D.1 meningkat dari 58,10 menjadi >= 65,00.",
          waktuPelaksanaan: "Agustus – Desember 2026",
          kebutuhanSumberDaya: "Panduan observasi kelas, ATK, konsumsi kegiatan Kombel.",
          perkiraanAnggaran: 3200000,
          disetujui: true
        }
      ];

      demo.rkjm = [
        {
          id: "demo-rkjm-1",
          indicatorId: ind1.id,
          dimensi: "D.1 Mutu dan Relevansi Hasil Belajar Peserta Didik",
          indikatorNama: "[A.2] Kemampuan Numerasi",
          akarMasalah: "Pembelajaran numerasi belum kontekstual dan keterbatasan alat peraga manipulatif.",
          tujuanStrategis: "Mewujudkan pembelajaran numerasi bermakna yang meningkatkan daya nalar peserta didik secara berkelanjutan.",
          programPengembangan: "Program Penguatan Numerasi Kontekstual & Digital Sekolah",
          targetTahun1: "Skor 60,50",
          targetTahun2: "Skor 66,00",
          targetTahun3: "Skor 72,00",
          targetTahun4: "Skor 78,00",
          indikatorKeberhasilan: "Kenaikan skor Rapor Pendidikan A.2 dan ketuntasan asesmen diagnostik numerasi.",
          penanggungJawab: "Koordinator Kurikulum",
          sumberPendanaan: "BOSP Reguler"
        },
        {
          id: "demo-rkjm-2",
          indicatorId: ind2.id,
          dimensi: "D.2 Mutu dan Relevansi Pembelajaran",
          indikatorNama: "[D.1] Kualitas Pembelajaran",
          akarMasalah: "Dominasi metode ceramah dan belum optimalnya refleksi rekan sejawat di Kombel.",
          tujuanStrategis: "Membangun ekosistem pembelajaran interaktif, berdiferensiasi, dan berpusat pada peserta didik.",
          programPengembangan: "Program Transformasi Pembelajaran Aktif & Komunitas Belajar",
          targetTahun1: "Skor 65,00",
          targetTahun2: "Skor 70,00",
          targetTahun3: "Skor 75,00",
          targetTahun4: "Skor 80,00",
          indikatorKeberhasilan: "Terlaksananya observasi rekan sejawat berkala dan kenaikan skor D.1.",
          penanggungJawab: "Kepala Sekolah",
          sumberPendanaan: "BOSP Reguler"
        }
      ];

      demo.rkt = [
        {
          id: "demo-rkt-1",
          rkjmId: "demo-rkjm-1",
          indicatorId: ind1.id,
          namaProgram: "Program Penguatan Numerasi Kontekstual & Digital Sekolah",
          namaKegiatan: "IHT Penyusunan Modul Ajar & Asesmen Numerasi Kontekstual",
          latarBelakang: "Capaian numerasi berada di angka 52,40 (Kurang) dan turun 4,20 poin.",
          tujuan: "Meningkatkan kompetensi 12 guru dalam merancang pembelajaran numerasi berbasis pemecahan masalah.",
          sasaran: "12 Guru Kelas dan Guru Mata Pelajaran",
          indikatorKeberhasilan: "Tersusunnya 12 modul ajar dan bank soal numerasi tervalidasi.",
          target: "100% guru menghasilkan minimal 2 modul ajar numerasi kontekstual.",
          bulanPelaksanaan: [7, 8],
          jadwalTeks: "Juli – Agustus 2026",
          penanggungJawab: "Koordinator Kurikulum",
          kebutuhanAnggaran: 3500000,
          sumberDana: "BOSP Reguler",
          buktiPelaksanaan: "Daftar hadir, notulen, foto kegiatan, dan dokumen modul ajar."
        },
        {
          id: "demo-rkt-2",
          rkjmId: "demo-rkjm-1",
          indicatorId: ind1.id,
          namaProgram: "Program Penguatan Numerasi Kontekstual & Digital Sekolah",
          namaKegiatan: "Pengadaan Alat Peraga Manipulatif Numerasi Kelas",
          latarBelakang: "Minimnya media pembelajaran konkret untuk konsep bilangan, pecahan, dan geometri.",
          tujuan: "Menyediakan kit alat peraga numerasi yang layak di setiap rombongan belajar.",
          sasaran: "6 Rombongan Belajar (Kelas I–VI)",
          indikatorKeberhasilan: "Tersedianya 6 paket alat peraga numerasi siap pakai.",
          target: "6 Paket Kit Numerasi terdistribusi ke kelas.",
          bulanPelaksanaan: [8, 9],
          jadwalTeks: "Agustus – September 2026",
          penanggungJawab: "Wakasek Sarpras & Bendahara BOS",
          kebutuhanAnggaran: 5000000,
          sumberDana: "BOSP Reguler",
          buktiPelaksanaan: "Faktur pembelian SIPLah, BAST barang, dan kartu inventaris."
        },
        {
          id: "demo-rkt-3",
          rkjmId: "demo-rkjm-2",
          indicatorId: ind2.id,
          namaProgram: "Program Transformasi Pembelajaran Aktif & Komunitas Belajar",
          namaKegiatan: "Lokakarya Pembelajaran Berdiferensiasi & Observasi Rekan Sejawat",
          latarBelakang: "Skor Kualitas Pembelajaran 58,10 (-1,80) perlu penguatan interaksi kelas.",
          tujuan: "Menerapkan siklus observasi rekan sejawat dan refleksi pembelajaran di Kombel.",
          sasaran: "12 Pendidik",
          indikatorKeberhasilan: "24 laporan observasi kelas dan rencana tindak lanjut guru.",
          target: "2 siklus observasi per guru selama semester berjalan.",
          bulanPelaksanaan: [9, 10, 11],
          jadwalTeks: "September – November 2026",
          penanggungJawab: "Ketua Komunitas Belajar",
          kebutuhanAnggaran: 3200000,
          sumberDana: "BOSP Reguler",
          buktiPelaksanaan: "Lembar observasi kelas, daftar hadir Kombel, dan laporan refleksi."
        }
      ];

      demo.arkas = [
        {
          id: "demo-arkas-1",
          rktId: "demo-rkt-1",
          rkjmId: "demo-rkjm-1",
          indicatorId: ind1.id,
          program: "Program Penguatan Numerasi Kontekstual & Digital Sekolah",
          kegiatan: "IHT Penyusunan Modul Ajar & Asesmen Numerasi Kontekstual",
          komponen: "06. Pengembangan Profesi Pendidik dan Tenaga Kependidikan",
          uraianBelanja: "Honorarium Narasumber Praktisi Pendidikan Numerasi (2 Hari x 4 JP)",
          volume: 8,
          satuan: "OJ",
          hargaSatuan: 250000,
          sumberDana: "BOSP Reguler",
          waktuPelaksanaan: "Juli 2026",
          statusVerifikasi: "Sesuai Standar Biaya (Perlu Cek SBU Daerah)"
        },
        {
          id: "demo-arkas-2",
          rktId: "demo-rkt-1",
          rkjmId: "demo-rkjm-1",
          indicatorId: ind1.id,
          program: "Program Penguatan Numerasi Kontekstual & Digital Sekolah",
          kegiatan: "IHT Penyusunan Modul Ajar & Asesmen Numerasi Kontekstual",
          komponen: "06. Pengembangan Profesi Pendidik dan Tenaga Kependidikan",
          uraianBelanja: "Konsumsi Rapat/Pelatihan Peserta IHT (15 Orang x 2 Hari)",
          volume: 30,
          satuan: "Kotak",
          hargaSatuan: 35000,
          sumberDana: "BOSP Reguler",
          waktuPelaksanaan: "Juli 2026",
          statusVerifikasi: "Terverifikasi Juknis BOSP"
        },
        {
          id: "demo-arkas-3",
          rktId: "demo-rkt-1",
          rkjmId: "demo-rkjm-1",
          indicatorId: ind1.id,
          program: "Program Penguatan Numerasi Kontekstual & Digital Sekolah",
          kegiatan: "IHT Penyusunan Modul Ajar & Asesmen Numerasi Kontekstual",
          komponen: "05. Pelaksanaan Administrasi Kegiatan Satuan Pendidikan",
          uraianBelanja: "Penggandaan Modul & Panduan Asesmen Numerasi Peserta",
          volume: 15,
          satuan: "Paket",
          hargaSatuan: 30000,
          sumberDana: "BOSP Reguler",
          waktuPelaksanaan: "Juli 2026",
          statusVerifikasi: "Terverifikasi Juknis BOSP"
        },
        {
          id: "demo-arkas-4",
          rktId: "demo-rkt-2",
          rkjmId: "demo-rkjm-1",
          indicatorId: ind1.id,
          program: "Program Penguatan Numerasi Kontekstual & Digital Sekolah",
          kegiatan: "Pengadaan Alat Peraga Manipulatif Numerasi Kelas",
          komponen: "03. Pelaksanaan Kegiatan Pembelajaran dan Bermain",
          uraianBelanja: "Kit Alat Peraga Matematika Dasar & Geometri (Pengadaan SIPLah)",
          volume: 6,
          satuan: "Paket",
          hargaSatuan: 800000,
          sumberDana: "BOSP Reguler",
          waktuPelaksanaan: "Agustus 2026",
          statusVerifikasi: "Terverifikasi Juknis BOSP"
        },
        {
          id: "demo-arkas-5",
          rktId: "demo-rkt-3",
          rkjmId: "demo-rkjm-2",
          indicatorId: ind2.id,
          program: "Program Transformasi Pembelajaran Aktif & Komunitas Belajar",
          kegiatan: "Lokakarya Pembelajaran Berdiferensiasi & Observasi Rekan Sejawat",
          komponen: "06. Pengembangan Profesi Pendidik dan Tenaga Kependidikan",
          uraianBelanja: "Konsumsi & ATK Pertemuan Komunitas Belajar (4 Sesi x 15 Peserta)",
          volume: 60,
          satuan: "Paket",
          hargaSatuan: 40000,
          sumberDana: "BOSP Reguler",
          waktuPelaksanaan: "September – November 2026",
          statusVerifikasi: "Terverifikasi Juknis BOSP"
        }
      ];

      demo.monev = [
        {
          id: "demo-monev-1",
          rktId: "demo-rkt-1",
          status: "Selesai",
          realisasiFisik: 100,
          realisasiAnggaran: 3500000,
          hambatan: "Penyesuaian jadwal narasumber agar tidak mengganggu jam efektif kelas.",
          buktiKegiatan: "Daftar hadir 100%, dokumentasi foto, 12 modul ajar numerasi terkumpul.",
          evaluasiHasil: "Seluruh guru mampu menyusun skenario pembelajaran numerasi berbasis masalah nyata.",
          tindakLanjut: "Pendampingan implementasi modul ajar di kelas melalui supervisi akademik."
        },
        {
          id: "demo-monev-2",
          rktId: "demo-rkt-2",
          status: "Sedang Berjalan",
          realisasiFisik: 75,
          realisasiAnggaran: 4800000,
          hambatan: "Pengiriman 1 paket alat peraga geometri menunggu konfirmasi penyedia SIPLah.",
          buktiKegiatan: "Bukti pemesanan SIPLah dan BAST 5 paket awal.",
          evaluasiHasil: "5 kelas sudah mulai menggunakan alat peraga dalam pembelajaran pecahan.",
          tindakLanjut: "Koordinasi dengan penyedia untuk penyelesaian pengiriman 1 paket tersisa."
        },
        {
          id: "demo-monev-3",
          rktId: "demo-rkt-3",
          status: "Belum Dilaksanakan",
          realisasiFisik: 0,
          realisasiAnggaran: 0,
          hambatan: "Menunggu jadwal siklus observasi semester berjalan.",
          buktiKegiatan: "-",
          evaluasiHasil: "Belum dimulai sesuai jadwal perencanaan bulan September.",
          tindakLanjut: "Finalisasi instrumen observasi rekan sejawat bersama Koordinator Kombel."
        }
      ];

      return demo;
    }

    const APP_STATE = {
      activeTab: "dashboard",
      activeYear: "2026",
      isDemoMode: false,
      analisisFilter: "all",
      rkjmEditMode: false,
      years: {
        "2026": createEmptyYearData("2026")
      },
      demoData: createDemoYearData(),
      chatHistory: [
        {
          role: "assistant",
          text: "Halo Bapak/Ibu Tim Pengembang Sekolah! Saya **ASISTEN RAPOR SMART AI**. Silakan unggah file Excel/PDF Rapor Pendidikan atau isi **Nilai Capaian** pada menu **1. Input Rapor & Nilai Capaian**. Saya siap membantu menganalisis data aktual sekolah Anda secara runtut hingga ke RKJM, RKT, dan rancangan ARKAS."
        }
      ]
    };

    // =========================================================================
    // 3. CORE HELPER FUNCTIONS
    // =========================================================================
    function uid() {
      return "id-" + Math.random().toString(36).substring(2, 9) + "-" + Date.now().toString(36);
    }

    function escapeHtml(val) {
      if (val === null || val === undefined) return "";
      return String(val)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }

    function formatRupiah(num) {
      const n = Number(num) || 0;
      return "Rp " + n.toLocaleString("id-ID");
    }

    function formatScore(val) {
      if (val === "" || val === null || val === undefined || isNaN(Number(val))) return "-";
      return Number(val).toLocaleString("id-ID", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
    }

    function getCurrentData() {
      if (APP_STATE.isDemoMode) return APP_STATE.demoData;
      if (!APP_STATE.years[APP_STATE.activeYear]) {
        APP_STATE.years[APP_STATE.activeYear] = createEmptyYearData(APP_STATE.activeYear);
      }
      return APP_STATE.years[APP_STATE.activeYear];
    }

    function classifyCategory(nilai, existingCategory = "") {
      if (nilai !== "" && nilai !== null && nilai !== undefined && !isNaN(parseFloat(nilai))) {
        const n = parseFloat(nilai);
        if (n >= 70) return "Baik";
        if (n >= 55) return "Sedang";
        return "Kurang";
      }
      if (existingCategory) {
        const low = existingCategory.toLowerCase();
        if (low.includes("baik") || low.includes("di atas") || low.includes("mencapai")) return "Baik";
        if (low.includes("sedang")) return "Sedang";
        if (low.includes("kurang") || low.includes("rendah") || low.includes("di bawah") || low.includes("perlu")) return "Kurang";
      }
      return "Belum Terverifikasi";
    }

    function getCategoryBadgeHtml(kat) {
      if (kat === "Baik") {
        return `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-900 border border-emerald-200"><span class="w-1.5 h-1.5 rounded-full bg-emerald-600"></span>Baik</span>`;
      }
      if (kat === "Sedang") {
        return `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-900 border border-amber-200"><span class="w-1.5 h-1.5 rounded-full bg-amber-500"></span>Sedang</span>`;
      }
      if (kat === "Kurang") {
        return `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-rose-100 text-rose-900 border border-rose-200"><span class="w-1.5 h-1.5 rounded-full bg-rose-600"></span>Kurang</span>`;
      }
      return `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-slate-100 text-slate-600 border border-slate-200">Perlu Konfirmasi</span>`;
    }

    function getDeltaBadgeHtml(perubahan) {
      if (perubahan === "" || perubahan === null || perubahan === undefined || isNaN(parseFloat(perubahan))) {
        return `<span class="text-slate-400 font-mono text-xs">-</span>`;
      }
      const d = parseFloat(perubahan);
      if (d > 0) {
        return `<span class="inline-flex items-center gap-0.5 px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-800 font-mono font-bold text-xs border border-emerald-200">▲ +${d.toFixed(2)}</span>`;
      }
      if (d < 0) {
        return `<span class="inline-flex items-center gap-0.5 px-2 py-0.5 rounded-md bg-rose-50 text-rose-800 font-mono font-bold text-xs border border-rose-200">▼ ${d.toFixed(2)}</span>`;
      }
      return `<span class="inline-flex items-center px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 font-mono font-semibold text-xs">0.00</span>`;
    }

    function triggerAutoSave(modul, aksi, detail) {
      const data = getCurrentData();
      data.history.unshift({
        waktu: new Date().toLocaleString("id-ID"),
        modul,
        aksi,
        detail
      });
      if (data.history.length > 50) data.history.pop();

      const badge = document.getElementById("autoSaveBadge");
      if (badge) {
        badge.innerHTML = `<span class="w-2 h-2 rounded-full bg-emerald-600"></span><span>Tersimpan (${new Date().toLocaleTimeString("id-ID", { hour: "2-digit", minute: "2-digit" })})</span>`;
      }
    }

    function showToast(message, type = "info") {
      const container = document.getElementById("toastContainer");
      if (!container) return;
      const styles = {
        success: "bg-emerald-900 text-white border-emerald-700",
        warning: "bg-amber-500 text-slate-950 border-amber-600",
        error: "bg-rose-700 text-white border-rose-800",
        info: "bg-slate-900 text-white border-slate-700"
      };
      const el = document.createElement("div");
      el.className = `pointer-events-auto p-3.5 rounded-xl border shadow-xl text-xs font-semibold flex items-start gap-2.5 transition-all duration-300 ${styles[type] || styles.info}`;
      el.innerHTML = `<i data-lucide="info" class="w-4 h-4 shrink-0 mt-0.5"></i><div class="flex-1 leading-relaxed">${escapeHtml(message)}</div>`;
      container.appendChild(el);
      if (window.lucide) window.lucide.createIcons();
      setTimeout(() => {
        el.style.opacity = "0";
        setTimeout(() => el.remove(), 300);
      }, 4500);
    }

    function openCustomModal(innerHtml, maxWidthClass = "max-w-2xl") {
      const backdrop = document.getElementById("customModalBackdrop");
      const box = document.getElementById("customModalBox");
      if (!backdrop || !box) return;
      box.className = `bg-white rounded-2xl shadow-2xl border border-slate-200 w-full ${maxWidthClass} overflow-hidden my-auto`;
      box.innerHTML = innerHtml;
      backdrop.classList.remove("hidden");
      backdrop.classList.add("flex");
      if (window.lucide) window.lucide.createIcons();
    }

    function closeCustomModal() {
      const backdrop = document.getElementById("customModalBackdrop");
      if (!backdrop) return;
      backdrop.classList.add("hidden");
      backdrop.classList.remove("flex");
    }

    function toggleMobileSidebar(open) {
      const sb = document.getElementById("appSidebar");
      const ov = document.getElementById("mobileSidebarOverlay");
      if (!sb || !ov) return;
      if (open) {
        sb.classList.remove("-translate-x-full");
        ov.classList.remove("hidden");
      } else {
        sb.classList.add("-translate-x-full");
        ov.classList.add("hidden");
      }
    }

    function setActiveTab(tabId) {
      APP_STATE.activeTab = tabId;
      toggleMobileSidebar(false);
      renderAll();
      window.scrollTo({ top: 0, behavior: "smooth" });
    }

    function toggleAssistantDrawer(open) {
      const drawer = document.getElementById("assistantDrawer");
      const overlay = document.getElementById("assistantOverlay");
      if (!drawer || !overlay) return;
      if (open) {
        drawer.classList.remove("translate-x-full");
        overlay.classList.remove("hidden");
        renderAssistantMessages();
      } else {
        drawer.classList.add("translate-x-full");
        overlay.classList.add("hidden");
      }
    }

    function exportProjectJson() {
      const data = getCurrentData();
      const payload = {
        app: "RAPOR SMART AI",
        exportedAt: new Date().toISOString(),
        activeYear: APP_STATE.activeYear,
        isDemoMode: APP_STATE.isDemoMode,
        years: APP_STATE.years
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const safeName = (data.school.nama || "Sekolah").replace(/[^a-zA-Z0-9]+/g, "_");
      a.download = `Cadangan_RAPOR_SMART_AI_${safeName}_${APP_STATE.activeYear}.json`;
      a.click();
      URL.revokeObjectURL(url);
      showToast("File cadangan proyek (.JSON) berhasil diunduh.", "success");
    }

    function importProjectJson(event) {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const parsed = JSON.parse(e.target.result);
          if (!parsed || !parsed.years || typeof parsed.years !== "object") {
            showToast("Format file cadangan .JSON tidak valid.", "error");
            return;
          }
          APP_STATE.years = parsed.years;
          APP_STATE.activeYear = parsed.activeYear || Object.keys(parsed.years)[0] || "2026";
          APP_STATE.isDemoMode = false;
          renderAll();
          showToast(`Data proyek berhasil dipulihkan (Tahun Aktif: ${APP_STATE.activeYear}).`, "success");
        } catch (err) {
          console.error(err);
          showToast("Gagal membaca file cadangan .JSON.", "error");
        }
      };
      reader.readAsText(file);
    }

    // =========================================================================
    // 4. DEMO MODE & MULTI-YEAR MANAGEMENT
    // =========================================================================
    function toggleDemoMode() {
      APP_STATE.isDemoMode = !APP_STATE.isDemoMode;
      renderAll();
      showToast(
        APP_STATE.isDemoMode
          ? "Mode Demo diaktifkan. Menampilkan contoh data lengkap SD Negeri Harapan Bangsa 01."
          : "Kembali ke ruang kerja data asli sekolah Anda.",
        APP_STATE.isDemoMode ? "warning" : "success"
      );
    }

    function exitDemoMode() {
      APP_STATE.isDemoMode = false;
      renderAll();
      showToast("Anda telah kembali ke data sekolah asli.", "success");
    }

    function switchPlanningYear(yearVal) {
      if (APP_STATE.isDemoMode) {
        APP_STATE.isDemoMode = false;
      }
      APP_STATE.activeYear = yearVal;
      renderAll();
      showToast(`Beralih ke Tahun Perencanaan ${yearVal}.`, "info");
    }

    function openNewYearModal() {
      const years = Object.keys(APP_STATE.years).sort();
      openCustomModal(`
        <div class="p-5 sm:p-6">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Tambah / Duplikasi Tahun Perencanaan</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="createNewPlanningYear(event)" class="mt-4 space-y-4 text-xs">
            <div>
              <label class="block font-bold text-slate-700 mb-1">Tahun Perencanaan Baru</label>
              <input type="number" name="newYear" min="2023" max="2035" value="${parseInt(APP_STATE.activeYear, 10) + 1}" required class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono font-bold" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Sumber Data Awal</label>
              <select name="copyFrom" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                <option value="">Mulai Kosong (Hanya salin identitas sekolah)</option>
                ${years.map(y => `<option value="${y}">Duplikasi Indikator &amp; Program dari Tahun ${y}</option>`).join("")}
              </select>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-4 py-2 rounded-xl emerald-gradient text-white font-bold">Buat Tahun Perencanaan</button>
            </div>
          </form>
        </div>
      `, "max-w-md");
    }

    function createNewPlanningYear(e) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const newYear = String(fd.get("newYear")).trim();
      const copyFrom = fd.get("copyFrom");
      if (!newYear) return;

      if (APP_STATE.years[newYear]) {
        showToast(`Tahun ${newYear} sudah tersedia.`, "warning");
        return;
      }

      if (copyFrom && APP_STATE.years[copyFrom]) {
        const cloned = JSON.parse(JSON.stringify(APP_STATE.years[copyFrom]));
        cloned.yearLabel = newYear;
        cloned.school.tahunPerencanaan = newYear;
        cloned.school.tahunRapor = newYear;
        cloned.history = [{
          waktu: new Date().toLocaleString("id-ID"),
          modul: "Manajemen Data",
          aksi: "Duplikasi Tahun",
          detail: `Menduplikasi data perencanaan dari tahun ${copyFrom} ke ${newYear}.`
        }];
        APP_STATE.years[newYear] = cloned;
      } else {
        const fresh = createEmptyYearData(newYear);
        const cur = getCurrentData();
        fresh.school = { ...cur.school, tahunPerencanaan: newYear, tahunRapor: newYear };
        APP_STATE.years[newYear] = fresh;
      }

      APP_STATE.isDemoMode = false;
      APP_STATE.activeYear = newYear;
      closeCustomModal();
      renderAll();
      showToast(`Tahun perencanaan ${newYear} berhasil dibuat.`, "success");
    }

    // =========================================================================
    // 5. NAVIGATION & HEADER RENDERING
    // =========================================================================
    function renderNavigationAndHeader() {
      const data = getCurrentData();

      // Demo Banner
      const demoBanner = document.getElementById("demoBanner");
      const demoToggleBtn = document.getElementById("demoToggleBtn");
      if (APP_STATE.isDemoMode) {
        demoBanner?.classList.remove("hidden");
        if (demoToggleBtn) demoToggleBtn.classList.add("ring-2", "ring-amber-400", "bg-emerald-800");
      } else {
        demoBanner?.classList.add("hidden");
        if (demoToggleBtn) demoToggleBtn.classList.remove("ring-2", "ring-amber-400", "bg-emerald-800");
      }

      // Year Select
      const yearSelect = document.getElementById("activeYearSelect");
      if (yearSelect) {
        const years = Object.keys(APP_STATE.years).sort();
        yearSelect.innerHTML = years.map(y => `<option value="${y}" ${y === APP_STATE.activeYear ? "selected" : ""}>Tahun ${y}</option>`).join("");
      }

      // Top Header School Metadata
      const hName = document.getElementById("headerSchoolName");
      const hNpsn = document.getElementById("headerNpsnBadge");
      const hJenjang = document.getElementById("headerJenjangBadge");
      const hSub = document.getElementById("headerSubtitle");

      if (hName) hName.textContent = data.school.nama || "Satuan Pendidikan Belum Diisi (Klik Input Rapor)";
      if (hNpsn) hNpsn.textContent = `NPSN: ${data.school.npsn || "-"}`;
      if (hJenjang) hJenjang.textContent = `Jenjang: ${data.school.jenjang || "-"}`;
      if (hSub) {
        const loc = [data.school.kecamatan, data.school.kabupaten, data.school.provinsi].filter(Boolean).join(", ");
        hSub.textContent = `Tahun Rapor: ${data.school.tahunRapor || "-"} • Tahun Anggaran: ${data.school.tahunPerencanaan || "-"}${loc ? " • " + loc : ""}`;
      }

      // Sidebar Nav Items with Counts
      const counts = {
        input: data.indicators.length,
        analisis: data.indicators.filter(i => classifyCategory(i.nilai, i.kategori) !== "Baik").length,
        akar: data.rootCauses.length,
        prioritas: data.priorities.length,
        rekomendasi: data.recommendations.length,
        rkjm: data.rkjm.length,
        rkt: data.rkt.length,
        arkas: data.arkas.length,
        monev: data.monev.length
      };

      const navEl = document.getElementById("sidebarNav");
      if (navEl) {
        navEl.innerHTML = NAV_ITEMS.map(item => {
          const active = APP_STATE.activeTab === item.id;
          const badgeCount = counts[item.id];
          return `
            <button onclick="setActiveTab('${item.id}')"
              class="w-full flex items-center justify-between gap-2 px-3 py-2.5 rounded-xl font-medium transition text-left ${active ? 'bg-amber-400 text-slate-950 font-bold shadow-sm' : 'text-emerald-100 hover:bg-emerald-800/70'}">
              <span class="flex items-center gap-2.5 min-w-0">
                <i data-lucide="${item.icon}" class="w-4 h-4 shrink-0 ${active ? 'text-slate-950' : 'text-amber-300'}"></i>
                <span class="truncate">${item.label}</span>
              </span>
              ${typeof badgeCount === "number" && badgeCount > 0
                ? `<span class="px-1.5 py-0.5 rounded-md text-[10px] font-mono font-bold ${active ? 'bg-slate-950 text-amber-300' : 'bg-emerald-950/80 text-emerald-200'}">${badgeCount}</span>`
                : ""}
            </button>
          `;
        }).join("");
      }
    }

    // =========================================================================
    // 6. MODULE 0: DASHBOARD UTAMA (With Nilai Capaian Summary & Chart)
    // =========================================================================
    function renderDashboard() {
      const data = getCurrentData();
      const indicators = data.indicators;

      const scoredIndicators = indicators.filter(i => i.nilai !== "" && !isNaN(parseFloat(i.nilai)));
      const avgScore = scoredIndicators.length > 0
        ? (scoredIndicators.reduce((acc, i) => acc + parseFloat(i.nilai), 0) / scoredIndicators.length)
        : null;

      const perluBenahi = indicators.filter(i => {
        const k = classifyCategory(i.nilai, i.kategori);
        return k === "Kurang" || k === "Sedang" || (i.perubahan !== "" && parseFloat(i.perubahan) < 0);
      });

      const totalAnggaran = data.arkas.reduce((sum, item) => sum + ((Number(item.volume) || 0) * (Number(item.hargaSatuan) || 0)), 0);

      // 8-Step Readiness Checklist
      const steps = [
        { label: "Identitas & Data Rapor", done: Boolean(data.school.nama && indicators.length > 0), tab: "input" },
        { label: "Analisis Capaian", done: indicators.length > 0, tab: "analisis" },
        { label: "Akar Masalah (5-Why)", done: data.rootCauses.length > 0, tab: "akar" },
        { label: "Penentuan Prioritas", done: data.priorities.length > 0, tab: "prioritas" },
        { label: "Rekomendasi Program", done: data.recommendations.length > 0, tab: "rekomendasi" },
        { label: "Dokumen RKJM", done: data.rkjm.length > 0, tab: "rkjm" },
        { label: "Dokumen RKT", done: data.rkt.length > 0, tab: "rkt" },
        { label: "Rancangan ARKAS", done: data.arkas.length > 0, tab: "arkas" }
      ];
      const doneSteps = steps.filter(s => s.done).length;
      const pctComplete = Math.round((doneSteps / steps.length) * 100);

      return `
        <!-- HERO BANNER -->
        <section class="emerald-gradient text-white rounded-2xl p-5 sm:p-7 shadow-lg relative overflow-hidden">
          <div class="relative z-10 flex flex-col lg:flex-row lg:items-center justify-between gap-5">
            <div class="space-y-2 max-w-2xl">
              <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-white/15 text-amber-300 text-xs font-bold">
                <i data-lucide="sparkles" class="w-3.5 h-3.5"></i>
                <span>Perencanaan Berbasis Data (PBD) Terintegrasi</span>
              </div>
              <h2 class="font-heading text-xl sm:text-2xl lg:text-3xl font-extrabold tracking-tight">
                ${escapeHtml(data.school.nama || "Selamat Datang di RAPOR SMART AI")}
              </h2>
              <p class="text-xs sm:text-sm text-emerald-100 leading-relaxed">
                ${data.school.nama
                  ? `NPSN <strong>${escapeHtml(data.school.npsn || "-")}</strong> • Jenjang <strong>${escapeHtml(data.school.jenjang || "-")}</strong> • Tahun Ajaran/Rapor <strong>${escapeHtml(data.school.tahunRapor)}</strong> • Tahun Anggaran <strong>${escapeHtml(data.school.tahunPerencanaan)}</strong>`
                  : "Mulai dengan mengunggah file Excel/PDF Rapor Pendidikan atau mengisi Nilai Capaian indikator sekolah Anda untuk menyusun perencanaan otomatis."}
              </p>
              <div class="pt-2 flex flex-wrap gap-2.5">
                <button onclick="setActiveTab('input')" class="px-4 py-2.5 rounded-xl gold-gradient text-slate-950 font-bold text-xs sm:text-sm shadow hover:brightness-105 transition flex items-center gap-2">
                  <i data-lucide="upload" class="w-4 h-4"></i>
                  <span>Input / Unggah Rapor Pendidikan</span>
                </button>
                ${indicators.length === 0 ? `
                  <button onclick="toggleDemoMode()" class="px-4 py-2.5 rounded-xl bg-white/15 hover:bg-white/25 text-white font-semibold text-xs sm:text-sm border border-white/25 transition flex items-center gap-2">
                    <i data-lucide="play-circle" class="w-4 h-4 text-amber-300"></i>
                    <span>Lihat Simulasi Mode Demo</span>
                  </button>
                ` : `
                  <button onclick="setActiveTab('analisis')" class="px-4 py-2.5 rounded-xl bg-white/15 hover:bg-white/25 text-white font-semibold text-xs sm:text-sm border border-white/25 transition flex items-center gap-2">
                    <i data-lucide="bar-chart-3" class="w-4 h-4 text-amber-300"></i>
                    <span>Lihat Analisis Capaian</span>
                  </button>
                `}
                <button onclick="setActiveTab('ekspor')" class="px-4 py-2.5 rounded-xl bg-emerald-950/60 hover:bg-emerald-950/80 text-amber-300 font-bold text-xs sm:text-sm border border-amber-400/40 transition flex items-center gap-2">
                  <i data-lucide="download" class="w-4 h-4"></i>
                  <span>Unduh Dokumen (PDF / Word / Excel)</span>
                </button>
              </div>
            </div>

            <!-- Progress Card -->
            <div class="bg-emerald-950/60 border border-emerald-700/70 rounded-2xl p-4 sm:p-5 lg:w-80 shrink-0">
              <div class="flex items-center justify-between text-xs mb-1.5">
                <span class="font-bold text-emerald-200">Status Pengisian Data PBD</span>
                <span class="font-mono font-extrabold text-amber-300 text-sm">${pctComplete}%</span>
              </div>
              <div class="w-full h-2.5 bg-emerald-900 rounded-full overflow-hidden mb-3">
                <div class="h-full gold-gradient rounded-full transition-all duration-500" style="width: ${pctComplete}%"></div>
              </div>
              <div class="grid grid-cols-2 gap-1.5 text-[11px]">
                ${steps.map(st => `
                  <button onclick="setActiveTab('${st.tab}')" class="flex items-center gap-1.5 text-left truncate hover:text-amber-300 transition ${st.done ? 'text-emerald-100' : 'text-emerald-400/80'}">
                    <i data-lucide="${st.done ? 'check-circle-2' : 'circle'}" class="w-3.5 h-3.5 shrink-0 ${st.done ? 'text-amber-400' : 'text-emerald-600'}"></i>
                    <span class="truncate">${st.label}</span>
                  </button>
                `).join("")}
              </div>
            </div>
          </div>
        </section>

        <!-- KPI STAT CARDS (Including Rata-Rata Nilai Capaian) -->
        <section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
          <div class="card-elevated p-4 border-l-4 border-l-emerald-700">
            <div class="flex items-center justify-between">
              <span class="text-xs font-semibold text-slate-500">Rata-Rata Nilai Capaian</span>
              <div class="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-800 flex items-center justify-center">
                <i data-lucide="gauge" class="w-4 h-4"></i>
              </div>
            </div>
            <div class="mt-2 font-mono text-2xl font-extrabold text-emerald-900">
              ${avgScore !== null ? avgScore.toFixed(2) : "-"} <span class="text-xs font-sans font-normal text-slate-400">/ 100</span>
            </div>
            <p class="text-[11px] text-slate-500 mt-1">Dari ${scoredIndicators.length} indikator bernilai</p>
          </div>

          <div class="card-elevated p-4 border-l-4 border-l-rose-600">
            <div class="flex items-center justify-between">
              <span class="text-xs font-semibold text-slate-500">Indikator Perlu Dibenahi</span>
              <div class="w-8 h-8 rounded-lg bg-rose-50 text-rose-700 flex items-center justify-center">
                <i data-lucide="alert-triangle" class="w-4 h-4"></i>
              </div>
            </div>
            <div class="mt-2 font-mono text-2xl font-extrabold text-rose-700">${perluBenahi.length}</div>
            <p class="text-[11px] text-slate-500 mt-1">Dari total ${indicators.length} indikator Rapor</p>
          </div>

          <div class="card-elevated p-4 border-l-4 border-l-amber-500">
            <div class="flex items-center justify-between">
              <span class="text-xs font-semibold text-slate-500">Program Prioritas &amp; RKJM</span>
              <div class="w-8 h-8 rounded-lg bg-amber-50 text-amber-700 flex items-center justify-center">
                <i data-lucide="compass" class="w-4 h-4"></i>
              </div>
            </div>
            <div class="mt-2 font-mono text-2xl font-extrabold text-slate-900">${data.rkjm.length}</div>
            <p class="text-[11px] text-slate-500 mt-1">${data.priorities.length} indikator diprioritaskan</p>
          </div>

          <div class="card-elevated p-4 border-l-4 border-l-teal-600">
            <div class="flex items-center justify-between">
              <span class="text-xs font-semibold text-slate-500">Jumlah Kegiatan RKT</span>
              <div class="w-8 h-8 rounded-lg bg-teal-50 text-teal-700 flex items-center justify-center">
                <i data-lucide="calendar-check" class="w-4 h-4"></i>
              </div>
            </div>
            <div class="mt-2 font-mono text-2xl font-extrabold text-slate-900">${data.rkt.length}</div>
            <p class="text-[11px] text-slate-500 mt-1">Kegiatan tahunan terprogram</p>
          </div>

          <div class="card-elevated p-4 border-l-4 border-l-amber-600">
            <div class="flex items-center justify-between">
              <span class="text-xs font-semibold text-slate-500">Total Rancangan ARKAS</span>
              <div class="w-8 h-8 rounded-lg bg-amber-50 text-amber-800 flex items-center justify-center">
                <i data-lucide="wallet" class="w-4 h-4"></i>
              </div>
            </div>
            <div class="mt-2 font-mono text-lg sm:text-xl font-extrabold text-emerald-950 truncate" title="${formatRupiah(totalAnggaran)}">
              ${formatRupiah(totalAnggaran)}
            </div>
            <p class="text-[11px] text-slate-500 mt-1">${data.arkas.length} rincian belanja</p>
          </div>
        </section>

        <!-- CHART & NILAI CAPAIAN TABLE ON DASHBOARD -->
        <section class="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <!-- Bar Chart Nilai Capaian -->
          <div class="lg:col-span-6 card-elevated p-5">
            <div class="flex items-center justify-between gap-2 pb-3 border-b border-slate-100">
              <div>
                <h3 class="font-bold text-slate-900 text-sm sm:text-base">Grafik Nilai Capaian Indikator Rapor</h3>
                <p class="text-xs text-slate-500">Skala 0 – 100 (Hijau: Baik ≥70 | Emas: Sedang 55–69 | Merah: Kurang &lt;55)</p>
              </div>
              <button onclick="setActiveTab('input')" class="text-xs font-bold text-emerald-800 hover:underline shrink-0">+ Edit Nilai</button>
            </div>

            ${indicators.length === 0 ? `
              <div class="py-12 text-center space-y-3">
                <div class="w-12 h-12 rounded-2xl bg-emerald-50 text-emerald-800 flex items-center justify-center mx-auto">
                  <i data-lucide="bar-chart-2" class="w-6 h-6"></i>
                </div>
                <p class="text-xs text-slate-500 max-w-xs mx-auto">Belum ada data indikator Rapor Pendidikan. Unggah file Excel/PDF atau tambahkan indikator beserta <strong>Nilai Capaian</strong>.</p>
                <button onclick="setActiveTab('input')" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Mulai Input Data</button>
              </div>
            ` : `
              <div class="mt-4 space-y-3.5 max-h-[360px] overflow-y-auto custom-scrollbar pr-1">
                ${indicators.map(ind => {
                  const hasScore = ind.nilai !== "" && !isNaN(parseFloat(ind.nilai));
                  const val = hasScore ? Math.min(100, Math.max(0, parseFloat(ind.nilai))) : 0;
                  const kat = classifyCategory(ind.nilai, ind.kategori);
                  const barColor = kat === "Baik" ? "bg-emerald-600" : kat === "Sedang" ? "bg-amber-500" : kat === "Kurang" ? "bg-rose-600" : "bg-slate-300";
                  return `
                    <div>
                      <div class="flex items-center justify-between text-xs mb-1 gap-2">
                        <span class="font-semibold text-slate-800 truncate">${ind.kode ? `<span class="font-mono text-emerald-800">[${escapeHtml(ind.kode)}]</span> ` : ""}${escapeHtml(ind.nama)}</span>
                        <div class="flex items-center gap-2 shrink-0">
                          <span class="font-mono font-extrabold text-slate-900">${hasScore ? val.toFixed(2) : "Kosong"}</span>
                          ${getCategoryBadgeHtml(kat)}
                        </div>
                      </div>
                      <div class="w-full h-3 bg-slate-100 rounded-full overflow-hidden border border-slate-200/70">
                        <div class="h-full ${barColor} rounded-full transition-all duration-500" style="width: ${hasScore ? val : 0}%"></div>
                      </div>
                    </div>
                  `;
                }).join("")}
              </div>
            `}
          </div>

          <!-- Ringkasan Tabel Nilai Capaian -->
          <div class="lg:col-span-6 card-elevated p-5 flex flex-col">
            <div class="flex items-center justify-between gap-2 pb-3 border-b border-slate-100">
              <div>
                <h3 class="font-bold text-slate-900 text-sm sm:text-base">Ringkasan Nilai Capaian Indikator</h3>
                <p class="text-xs text-slate-500">Daftar indikator beserta skor capaian dan tren perubahan</p>
              </div>
              <button onclick="setActiveTab('analisis')" class="text-xs font-bold text-emerald-800 hover:underline shrink-0">Analisis Detail →</button>
            </div>

            ${indicators.length === 0 ? `
              <div class="py-12 text-center text-xs text-slate-400 my-auto">
                Data nilai capaian akan tampil otomatis setelah Anda memasukkan data Rapor Pendidikan.
              </div>
            ` : `
              <div class="mt-3 overflow-x-auto flex-1">
                <table class="w-full text-left text-xs">
                  <thead>
                    <tr class="border-b border-slate-200 text-slate-500 uppercase text-[10px]">
                      <th class="py-2 pr-2">Indikator</th>
                      <th class="py-2 px-2 text-center">Nilai Capaian</th>
                      <th class="py-2 px-2 text-center">Kategori</th>
                      <th class="py-2 pl-2 text-right">Perubahan</th>
                    </tr>
                  </thead>
                  <tbody class="divide-y divide-slate-100">
                    ${indicators.map(ind => `
                      <tr class="hover:bg-slate-50">
                        <td class="py-2.5 pr-2 font-semibold text-slate-800">
                          ${ind.kode ? `<span class="font-mono text-emerald-800">[${escapeHtml(ind.kode)}]</span> ` : ""}${escapeHtml(ind.nama)}
                        </td>
                        <td class="py-2.5 px-2 text-center font-mono font-extrabold text-slate-900 text-sm">
                          ${formatScore(ind.nilai)}
                        </td>
                        <td class="py-2.5 px-2 text-center">${getCategoryBadgeHtml(classifyCategory(ind.nilai, ind.kategori))}</td>
                        <td class="py-2.5 pl-2 text-right">${getDeltaBadgeHtml(ind.perubahan)}</td>
                      </tr>
                    `).join("")}
                  </tbody>
                </table>
              </div>
            `}
          </div>
        </section>
      `;
    }

    // =========================================================================
    // 7. MODULE 1: INPUT DATA RAPOR PENDIDIKAN (Excel, PDF, Manual + Nilai Capaian)
    // =========================================================================
    function renderInputModule() {
      const data = getCurrentData();
      const s = data.school;

      return `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 1 • Sumber Data Utama</span>
            <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Input Data Rapor Pendidikan &amp; Nilai Capaian</h2>
            <p class="text-xs text-slate-600 mt-0.5">Unggah langsung file <strong>Excel Rapor Pendidikan</strong> resmi, dokumen <strong>PDF</strong>, atau isi <strong>Nilai Capaian</strong> secara manual.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button onclick="downloadExcelTemplate()" class="px-3.5 py-2 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-slate-800 font-semibold text-xs flex items-center gap-1.5">
              <i data-lucide="file-down" class="w-4 h-4 text-emerald-700"></i>
              <span>Unduh Templat Excel</span>
            </button>
            <button onclick="setActiveTab('analisis')" class="px-4 py-2 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center gap-1.5 shadow">
              <span>Mulai Analisis</span>
              <i data-lucide="arrow-right" class="w-4 h-4"></i>
            </button>
          </div>
        </div>

        <!-- SCHOOL IDENTITY FORM -->
        <section class="card-elevated p-5">
          <div class="flex items-center justify-between pb-3 border-b border-slate-100">
            <div class="flex items-center gap-2">
              <div class="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-800 flex items-center justify-center">
                <i data-lucide="building-2" class="w-4 h-4"></i>
              </div>
              <div>
                <h3 class="font-bold text-slate-900 text-sm">Identitas Satuan Pendidikan</h3>
                <p class="text-[11px] text-slate-500">Otomatis terisi saat mengunggah file Rapor Pendidikan atau dapat disunting langsung.</p>
              </div>
            </div>
            <button type="button" onclick="saveSchoolIdentityForm()" class="px-3.5 py-1.5 rounded-xl bg-emerald-900 text-white text-xs font-bold hover:bg-emerald-800 flex items-center gap-1.5">
              <i data-lucide="save" class="w-3.5 h-3.5"></i>
              <span>Simpan Identitas</span>
            </button>
          </div>

          <form id="schoolIdentityForm" onsubmit="event.preventDefault(); saveSchoolIdentityForm();" class="mt-4 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
            <div class="sm:col-span-2">
              <label class="block font-bold text-slate-700 mb-1">Nama Sekolah *</label>
              <input type="text" name="nama" value="${escapeHtml(s.nama)}" placeholder="Contoh: SD Negeri 01 Merdeka" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">NPSN *</label>
              <input type="text" name="npsn" value="${escapeHtml(s.npsn)}" placeholder="8 digit NPSN" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Jenjang Pendidikan</label>
              <select name="jenjang" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none">
                ${["PAUD / TK", "SD", "SMP", "SMA", "SMK", "SLB", "PKBM"].map(j => `<option value="${j}" ${s.jenjang === j ? "selected" : ""}>${j}</option>`).join("")}
              </select>
            </div>
            <div class="sm:col-span-2">
              <label class="block font-bold text-slate-700 mb-1">Alamat Sekolah</label>
              <input type="text" name="alamat" value="${escapeHtml(s.alamat)}" placeholder="Jalan, Desa/Kelurahan" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Kecamatan</label>
              <input type="text" name="kecamatan" value="${escapeHtml(s.kecamatan)}" placeholder="Nama Kecamatan" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Kabupaten / Kota</label>
              <input type="text" name="kabupaten" value="${escapeHtml(s.kabupaten)}" placeholder="Kabupaten / Kota" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Provinsi</label>
              <input type="text" name="provinsi" value="${escapeHtml(s.provinsi)}" placeholder="Nama Provinsi" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Tahun Rapor Pendidikan</label>
              <input type="text" name="tahunRapor" value="${escapeHtml(s.tahunRapor)}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Tahun Perencanaan / Anggaran</label>
              <input type="text" name="tahunPerencanaan" value="${escapeHtml(s.tahunPerencanaan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono focus:border-emerald-700 focus:outline-none" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Kepala Sekolah</label>
              <input type="text" name="kepalaSekolah" value="${escapeHtml(s.kepalaSekolah)}" placeholder="Nama & Gelar Kepala Sekolah" class="w-full rounded-xl border border-slate-300 px-3 py-2 focus:border-emerald-700 focus:outline-none" />
            </div>
          </form>
        </section>

        <!-- 3 WAYS TO INPUT RAPOR PENDIDIKAN DATA -->
        <section class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <!-- Method 1: Direct Excel Rapor Pendidikan Upload -->
          <div class="card-elevated p-5 border-2 border-emerald-700/30 bg-emerald-50/30 flex flex-col justify-between">
            <div>
              <div class="flex items-center justify-between mb-3">
                <div class="w-10 h-10 rounded-xl emerald-gradient text-white flex items-center justify-center shadow-sm">
                  <i data-lucide="file-spreadsheet" class="w-5 h-5"></i>
                </div>
                <span class="px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-900 text-[10px] font-extrabold">DIREKOMENDASIKAN</span>
              </div>
              <h3 class="font-bold text-slate-900 text-sm">1. Unggah Langsung Excel Rapor</h3>
              <p class="text-xs text-slate-600 mt-1 leading-relaxed">
                Mendukung langsung file unduhan resmi <strong>Rapor Pendidikan (.xlsx / .xls / .csv)</strong>. Otomatis melewati sheet <em>Panduan</em>, membaca <strong>Nilai Capaian</strong>, kategori, dan perubahan skor.
              </p>
            </div>
            <div class="mt-4">
              <label class="w-full cursor-pointer py-2.5 px-4 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center justify-center gap-2 shadow hover:opacity-95 transition">
                <i data-lucide="upload" class="w-4 h-4"></i>
                <span>Pilih File Excel Rapor (.xlsx)</span>
                <input type="file" accept=".xlsx,.xls,.csv" onchange="handleExcelUpload(event)" class="hidden" />
              </label>
            </div>
          </div>

          <!-- Method 2: PDF Rapor Pendidikan Upload -->
          <div class="card-elevated p-5 flex flex-col justify-between">
            <div>
              <div class="flex items-center justify-between mb-3">
                <div class="w-10 h-10 rounded-xl bg-amber-100 text-amber-800 flex items-center justify-center">
                  <i data-lucide="file-text" class="w-5 h-5"></i>
                </div>
                <span class="px-2.5 py-0.5 rounded-full bg-amber-50 text-amber-800 text-[10px] font-bold border border-amber-200">PDF EKSTRAKSI</span>
              </div>
              <h3 class="font-bold text-slate-900 text-sm">2. Unggah Dokumen PDF Rapor</h3>
              <p class="text-xs text-slate-600 mt-1 leading-relaxed">
                Mengekstrak teks indikator dari file PDF Rapor Pendidikan dan menampilkan jendela konfirmasi untuk memverifikasi <strong>Nilai Capaian</strong> yang terbaca.
              </p>
            </div>
            <div class="mt-4">
              <label class="w-full cursor-pointer py-2.5 px-4 rounded-xl bg-amber-500 hover:bg-amber-600 text-slate-950 font-bold text-xs flex items-center justify-center gap-2 shadow-xs transition">
                <i data-lucide="file-up" class="w-4 h-4"></i>
                <span>Pilih File PDF Rapor (.pdf)</span>
                <input type="file" accept=".pdf" onchange="handlePdfUpload(event)" class="hidden" />
              </label>
            </div>
          </div>

          <!-- Method 3: Manual Input & Standard PBD Catalog -->
          <div class="card-elevated p-5 flex flex-col justify-between">
            <div>
              <div class="flex items-center justify-between mb-3">
                <div class="w-10 h-10 rounded-xl bg-slate-900 text-amber-400 flex items-center justify-center">
                  <i data-lucide="edit-3" class="w-5 h-5"></i>
                </div>
                <span class="px-2.5 py-0.5 rounded-full bg-slate-100 text-slate-700 text-[10px] font-bold">MANUAL &amp; KATALOG</span>
              </div>
              <h3 class="font-bold text-slate-900 text-sm">3. Input Manual &amp; Katalog PBD</h3>
              <p class="text-xs text-slate-600 mt-1 leading-relaxed">
                Pilih indikator dari daftar standar Kemendikdasmen dan langsung ketik <strong>Nilai Capaian (0–100)</strong> sesuai Rapor Pendidikan sekolah Anda.
              </p>
            </div>
            <div class="mt-4 grid grid-cols-2 gap-2">
              <button onclick="openBulkCatalogPickerModal()" class="py-2.5 px-3 rounded-xl bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs flex items-center justify-center gap-1.5">
                <i data-lucide="list-checks" class="w-4 h-4 text-amber-400"></i>
                <span>Katalog PBD</span>
              </button>
              <button onclick="openAddIndicatorModal()" class="py-2.5 px-3 rounded-xl border border-slate-300 hover:bg-slate-50 text-slate-800 font-bold text-xs flex items-center justify-center gap-1.5">
                <i data-lucide="plus" class="w-4 h-4 text-emerald-700"></i>
                <span>Tambah Baru</span>
              </button>
            </div>
          </div>
        </section>

        <!-- INDICATORS VERIFICATION & INLINE NILAI CAPAIAN TABLE -->
        <section class="card-elevated overflow-hidden">
          <div class="p-4 sm:p-5 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Indikator &amp; Nilai Capaian Rapor Pendidikan (${data.indicators.length})</h3>
              <p class="text-xs text-slate-500">Anda dapat langsung mengetik atau mengubah angka pada kolom <strong>Nilai Capaian (0–100)</strong> di bawah ini.</p>
            </div>
            <div class="flex flex-wrap gap-2">
              <button onclick="openBulkCatalogPickerModal()" class="px-3 py-1.5 rounded-xl bg-emerald-50 text-emerald-900 border border-emerald-200 hover:bg-emerald-100 text-xs font-bold flex items-center gap-1.5">
                <i data-lucide="list-plus" class="w-3.5 h-3.5"></i>
                <span>Pilih Indikator &amp; Isi Nilai</span>
              </button>
              ${data.indicators.length > 0 ? `
                <button onclick="confirmClearAllIndicators()" class="px-3 py-1.5 rounded-xl bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-100 text-xs font-semibold">
                  Kosongkan Tabel
                </button>
              ` : ""}
            </div>
          </div>

          ${data.indicators.length === 0 ? `
            <div class="p-10 text-center space-y-3">
              <div class="w-12 h-12 rounded-2xl bg-slate-100 text-slate-500 flex items-center justify-center mx-auto">
                <i data-lucide="clipboard-list" class="w-6 h-6"></i>
              </div>
              <h4 class="font-bold text-slate-800 text-sm">Belum Ada Data Indikator Rapor Pendidikan</h4>
              <p class="text-xs text-slate-500 max-w-md mx-auto">Silakan unggah file Excel Rapor Pendidikan Anda atau klik tombol <strong>Pilih Indikator &amp; Isi Nilai</strong> untuk memasukkan data capaian sekolah.</p>
            </div>
          ` : `
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs">
                <thead class="bg-slate-50 border-b border-slate-200 text-slate-600 uppercase text-[10px]">
                  <tr>
                    <th class="py-3 px-4">Indikator &amp; Dimensi</th>
                    <th class="py-3 px-3 text-center">Nilai Capaian (0–100)</th>
                    <th class="py-3 px-3 text-center">Kategori</th>
                    <th class="py-3 px-3 text-center">Perubahan</th>
                    <th class="py-3 px-3">Peringkat / Rekomendasi Rapor</th>
                    <th class="py-3 px-3 text-center">Status Data</th>
                    <th class="py-3 px-4 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                  ${data.indicators.map(ind => {
                    const kat = classifyCategory(ind.nilai, ind.kategori);
                    const hasScore = ind.nilai !== "" && !isNaN(parseFloat(ind.nilai));
                    return `
                      <tr class="hover:bg-slate-50/80 align-middle">
                        <td class="py-3 px-4 min-w-[220px]">
                          <div class="font-bold text-slate-900">
                            ${ind.kode ? `<span class="font-mono text-emerald-800 mr-1">[${escapeHtml(ind.kode)}]</span>` : ""}${escapeHtml(ind.nama)}
                          </div>
                          <div class="text-[11px] text-slate-500 mt-0.5">${escapeHtml(ind.dimensi || "-")}</div>
                        </td>
                        <td class="py-3 px-3 text-center">
                          <input type="number" step="0.01" min="0" max="100"
                            value="${hasScore ? escapeHtml(ind.nilai) : ""}"
                            placeholder="Isi nilai"
                            aria-label="Nilai Capaian ${escapeHtml(ind.nama)}"
                            onchange="updateIndicatorScoreInline('${ind.id}', this.value)"
                            class="w-24 text-center font-mono font-bold text-xs rounded-lg border ${hasScore ? "border-slate-300 bg-white text-slate-900" : "border-rose-300 bg-rose-50 text-rose-800 placeholder:text-rose-400"} px-2 py-1.5 focus:border-emerald-700 focus:outline-none" />
                        </td>
                        <td class="py-3 px-3 text-center">${getCategoryBadgeHtml(kat)}</td>
                        <td class="py-3 px-3 text-center">${getDeltaBadgeHtml(ind.perubahan)}</td>
                        <td class="py-3 px-3 min-w-[180px] text-[11px] text-slate-600">
                          ${ind.peringkat ? `<div class="font-semibold text-slate-800">${escapeHtml(ind.peringkat)}</div>` : ""}
                          <div class="line-clamp-2">${escapeHtml(ind.rekomendasiRapor || "-")}</div>
                        </td>
                        <td class="py-3 px-3 text-center">
                          ${ind.terverifikasi && hasScore
                            ? `<span class="inline-flex items-center gap-1 text-[11px] font-bold text-emerald-800"><i data-lucide="check-circle-2" class="w-3.5 h-3.5"></i> Terverifikasi</span>`
                            : `<button onclick="verifyIndicatorRow('${ind.id}')" class="px-2 py-1 rounded-lg bg-amber-100 text-amber-900 font-bold text-[10px] hover:bg-amber-200">Perlu Konfirmasi</button>`}
                        </td>
                        <td class="py-3 px-4 text-right whitespace-nowrap">
                          <button onclick="openEditIndicatorModal('${ind.id}')" class="p-1.5 rounded-lg hover:bg-emerald-50 text-emerald-800" title="Edit Lengkap Indikator">
                            <i data-lucide="pencil" class="w-3.5 h-3.5"></i>
                          </button>
                          <button onclick="deleteIndicator('${ind.id}')" class="p-1.5 rounded-lg hover:bg-rose-50 text-rose-700" title="Hapus Indikator">
                            <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                          </button>
                        </td>
                      </tr>
                    `;
                  }).join("")}
                </tbody>
              </table>
            </div>
          `}
        </section>
      `;
    }

    function saveSchoolIdentityForm() {
      const form = document.getElementById("schoolIdentityForm");
      if (!form) return;
      const fd = new FormData(form);
      const data = getCurrentData();
      data.school.nama = String(fd.get("nama") || "").trim();
      data.school.npsn = String(fd.get("npsn") || "").trim();
      data.school.jenjang = String(fd.get("jenjang") || "SD").trim();
      data.school.alamat = String(fd.get("alamat") || "").trim();
      data.school.kecamatan = String(fd.get("kecamatan") || "").trim();
      data.school.kabupaten = String(fd.get("kabupaten") || "").trim();
      data.school.provinsi = String(fd.get("provinsi") || "").trim();
      data.school.tahunRapor = String(fd.get("tahunRapor") || "").trim();
      data.school.tahunPerencanaan = String(fd.get("tahunPerencanaan") || "").trim();
      data.school.kepalaSekolah = String(fd.get("kepalaSekolah") || "").trim();

      triggerAutoSave("Input Data", "Simpan Identitas Sekolah", `Memperbarui profil ${data.school.nama || "Satuan Pendidikan"}.`);
      renderAll();
      showToast("Identitas satuan pendidikan berhasil disimpan.", "success");
    }

    function updateIndicatorScoreInline(indicatorId, rawVal) {
      const data = getCurrentData();
      const ind = data.indicators.find(i => i.id === indicatorId);
      if (!ind) return;

      const trimmed = String(rawVal || "").trim();
      if (trimmed === "") {
        ind.nilai = "";
        ind.kategori = "Belum Terverifikasi";
        ind.terverifikasi = false;
      } else {
        const parsed = parseFloat(trimmed);
        if (isNaN(parsed) || parsed < 0 || parsed > 100) {
          showToast("Nilai capaian harus berupa angka antara 0 sampai 100.", "warning");
          renderAll();
          return;
        }
        ind.nilai = parsed;
        ind.kategori = classifyCategory(parsed);
        ind.terverifikasi = true;
      }

      triggerAutoSave("Input Data", "Update Nilai Capaian", `Memperbarui nilai capaian "${ind.nama}" menjadi ${ind.nilai !== "" ? ind.nilai : "kosong"}.`);
      renderAll();
      if (ind.nilai !== "") {
        showToast(`Nilai capaian "${ind.nama}" (${ind.nilai} — ${ind.kategori}) disimpan.`, "success");
      }
    }

    function verifyIndicatorRow(indicatorId) {
      const data = getCurrentData();
      const ind = data.indicators.find(i => i.id === indicatorId);
      if (!ind) return;
      if (ind.nilai === "" || isNaN(parseFloat(ind.nilai))) {
        openEditIndicatorModal(indicatorId);
        showToast("Nilai capaian masih kosong. Silakan isi nilai capaian terlebih dahulu.", "warning");
        return;
      }
      ind.terverifikasi = true;
      ind.kategori = classifyCategory(ind.nilai, ind.kategori);
      triggerAutoSave("Input Data", "Konfirmasi Indikator", `Mengonfirmasi keakuratan indikator "${ind.nama}".`);
      renderAll();
      showToast(`Indikator "${ind.nama}" telah dikonfirmasi.`, "success");
    }

    function openBulkCatalogPickerModal() {
      const data = getCurrentData();
      const existingMap = new Map(data.indicators.map(i => [i.kode, i]));

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div>
              <h3 class="font-bold text-slate-900 text-base">Pilih Indikator &amp; Isi Nilai Capaian Rapor Pendidikan</h3>
              <p class="text-xs text-slate-500">Centang indikator dan langsung masukkan <strong>Nilai Capaian (0–100)</strong> sesuai Rapor Pendidikan sekolah Anda.</p>
            </div>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveBulkCatalogSelection(event)" class="mt-4 space-y-2 text-xs">
            <div class="max-h-80 overflow-y-auto divide-y divide-slate-100 border border-slate-200 rounded-xl">
              ${KATALOG_INDIKATOR_STANDAR.map(std => {
                const existing = existingMap.get(std.kode);
                const already = Boolean(existing);
                return `
                  <div class="flex items-center justify-between gap-3 p-3 hover:bg-slate-50 ${already ? "bg-slate-50/80" : ""}">
                    <label class="flex items-start gap-3 flex-1 cursor-pointer min-w-0">
                      <input type="checkbox" id="chk_${std.kode}" name="selectedCodes" value="${std.kode}" ${already ? "disabled checked" : ""} class="mt-0.5 rounded text-emerald-700 shrink-0" />
                      <div class="min-w-0">
                        <div class="font-bold text-slate-900 truncate">[${std.kode}] ${escapeHtml(std.nama)} ${already ? '<span class="text-[10px] text-emerald-700 font-semibold">(Sudah ada)</span>' : ""}</div>
                        <div class="text-[11px] text-slate-500 truncate">${escapeHtml(std.dimensi)}</div>
                      </div>
                    </label>
                    <div class="shrink-0 flex items-center gap-1.5">
                      <span class="text-[10px] font-semibold text-slate-500 hidden sm:inline">Nilai Capaian:</span>
                      <input type="number" step="0.01" min="0" max="100"
                        name="score_${std.kode}"
                        value="${already && existing.nilai !== "" ? escapeHtml(existing.nilai) : ""}"
                        ${already ? "disabled" : ""}
                        oninput="if(this.value !== ''){ const cb = document.getElementById('chk_${std.kode}'); if(cb && !cb.disabled) cb.checked = true; }"
                        placeholder="0 - 100"
                        class="w-24 text-center font-mono font-bold text-xs rounded-lg border border-slate-300 px-2 py-1.5 focus:border-emerald-700 focus:outline-none disabled:bg-slate-100 disabled:text-slate-500" />
                    </div>
                  </div>
                `;
              }).join("")}
            </div>
            <div class="pt-3 flex justify-end gap-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Indikator &amp; Nilai Capaian</button>
            </div>
          </form>
        </div>
      `, "max-w-2xl");
    }

    function saveBulkCatalogSelection(e) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const selected = fd.getAll("selectedCodes");
      const data = getCurrentData();
      let added = 0;

      selected.forEach(kode => {
        const std = KATALOG_INDIKATOR_STANDAR.find(s => s.kode === kode);
        if (std && !data.indicators.some(i => i.kode === kode)) {
          const rawScore = String(fd.get(`score_${kode}`) || "").trim();
          const parsedScore = rawScore !== "" ? parseFloat(rawScore) : NaN;
          const hasValidScore = !isNaN(parsedScore) && parsedScore >= 0 && parsedScore <= 100;

          data.indicators.push({
            id: uid(),
            kode: std.kode,
            nama: std.nama,
            dimensi: std.dimensi,
            nilai: hasValidScore ? parsedScore : "",
            kategori: hasValidScore ? classifyCategory(parsedScore) : "Belum Terverifikasi",
            perubahan: "",
            peringkat: "",
            rekomendasiRapor: "",
            terverifikasi: hasValidScore
          });
          added++;
        }
      });

      closeCustomModal();
      if (added > 0) {
        triggerAutoSave("Input Data", "Tambah dari Katalog", `Menambahkan ${added} indikator standar PBD.`);
        renderAll();
        showToast(`${added} indikator berhasil ditambahkan beserta nilai capaiannya.`, "success");
      }
    }

    function openAddIndicatorModal() {
      openCustomModal(`
        <div class="p-5 sm:p-6">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Tambah Indikator Rapor Pendidikan Manual</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveIndicatorModalForm(event, null)" class="mt-4 space-y-3 text-xs">
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Kode Indikator</label>
                <input type="text" name="kode" placeholder="Misal: A.1" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
              <div class="sm:col-span-2">
                <label class="block font-bold text-slate-700 mb-1">Nama Indikator *</label>
                <input type="text" name="nama" required placeholder="Contoh: Kemampuan Literasi" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Dimensi Rapor Pendidikan</label>
              <input type="text" name="dimensi" placeholder="Contoh: D.1 Mutu dan Relevansi Hasil Belajar Peserta Didik" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-emerald-900 mb-1">Nilai Capaian (0–100) *</label>
                <input type="number" step="0.01" min="0" max="100" name="nilai" required placeholder="Contoh: 64.50" class="w-full rounded-xl border-2 border-emerald-600 px-3 py-2 font-mono font-bold" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Kategori Capaian</label>
                <select name="kategori" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  <option value="">Otomatis dari Nilai</option>
                  <option value="Baik">Baik</option>
                  <option value="Sedang">Sedang</option>
                  <option value="Kurang">Kurang</option>
                </select>
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Perubahan Skor</label>
                <input type="number" step="0.01" name="perubahan" placeholder="Misal: -2.40 atau 3.15" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Peringkat / Kuantil (Jika tersedia)</label>
              <input type="text" name="peringkat" placeholder="Contoh: Peringkat Menengah Atas (Kuantil 4 Kab/Kota)" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Rekomendasi dari Rapor Pendidikan</label>
              <textarea name="rekomendasiRapor" rows="2" placeholder="Salin inspirasi benahi dari Rapor Pendidikan..." class="w-full rounded-xl border border-slate-300 px-3 py-2"></textarea>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Indikator</button>
            </div>
          </form>
        </div>
      `);
    }

    function openEditIndicatorModal(indicatorId) {
      const data = getCurrentData();
      const ind = data.indicators.find(i => i.id === indicatorId);
      if (!ind) return;

      openCustomModal(`
        <div class="p-5 sm:p-6">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Edit Data &amp; Nilai Capaian Indikator</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveIndicatorModalForm(event, '${ind.id}')" class="mt-4 space-y-3 text-xs">
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Kode Indikator</label>
                <input type="text" name="kode" value="${escapeHtml(ind.kode)}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
              <div class="sm:col-span-2">
                <label class="block font-bold text-slate-700 mb-1">Nama Indikator *</label>
                <input type="text" name="nama" required value="${escapeHtml(ind.nama)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Dimensi Rapor Pendidikan</label>
              <input type="text" name="dimensi" value="${escapeHtml(ind.dimensi)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold text-emerald-900 mb-1">Nilai Capaian (0–100) *</label>
                <input type="number" step="0.01" min="0" max="100" name="nilai" required value="${escapeHtml(ind.nilai)}" class="w-full rounded-xl border-2 border-emerald-600 px-3 py-2 font-mono font-bold" />
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Kategori Capaian</label>
                <select name="kategori" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  <option value="">Otomatis dari Nilai</option>
                  <option value="Baik" ${ind.kategori === "Baik" ? "selected" : ""}>Baik</option>
                  <option value="Sedang" ${ind.kategori === "Sedang" ? "selected" : ""}>Sedang</option>
                  <option value="Kurang" ${ind.kategori === "Kurang" ? "selected" : ""}>Kurang</option>
                </select>
              </div>
              <div>
                <label class="block font-bold text-slate-700 mb-1">Perubahan Skor</label>
                <input type="number" step="0.01" name="perubahan" value="${escapeHtml(ind.perubahan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" />
              </div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Peringkat / Kuantil</label>
              <input type="text" name="peringkat" value="${escapeHtml(ind.peringkat)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" />
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Rekomendasi Rapor Pendidikan</label>
              <textarea name="rekomendasiRapor" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(ind.rekomendasiRapor)}</textarea>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Perubahan</button>
            </div>
          </form>
        </div>
      `);
    }

    function saveIndicatorModalForm(e, indicatorId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();

      const rawNilai = String(fd.get("nilai") || "").trim();
      const parsedNilai = rawNilai !== "" ? parseFloat(rawNilai) : "";
      const selectedKat = String(fd.get("kategori") || "").trim();
      const finalKat = selectedKat || classifyCategory(parsedNilai);

      const rawDelta = String(fd.get("perubahan") || "").trim();
      const parsedDelta = rawDelta !== "" ? parseFloat(rawDelta) : "";

      const obj = {
        id: indicatorId || uid(),
        kode: String(fd.get("kode") || "").trim(),
        nama: String(fd.get("nama") || "").trim(),
        dimensi: String(fd.get("dimensi") || "").trim(),
        nilai: parsedNilai,
        kategori: finalKat,
        perubahan: parsedDelta,
        peringkat: String(fd.get("peringkat") || "").trim(),
        rekomendasiRapor: String(fd.get("rekomendasiRapor") || "").trim(),
        terverifikasi: parsedNilai !== "" && !isNaN(parsedNilai)
      };

      if (indicatorId) {
        const idx = data.indicators.findIndex(i => i.id === indicatorId);
        if (idx >= 0) data.indicators[idx] = obj;
      } else {
        data.indicators.push(obj);
      }

      closeCustomModal();
      triggerAutoSave("Input Data", indicatorId ? "Edit Indikator" : "Tambah Indikator", `Menyimpan indikator "${obj.nama}" dengan Nilai Capaian ${obj.nilai}.`);
      renderAll();
      showToast(`Indikator "${obj.nama}" berhasil disimpan.`, "success");
    }

    function deleteIndicator(indicatorId) {
      const data = getCurrentData();
      const ind = data.indicators.find(i => i.id === indicatorId);
      if (!ind) return;

      openCustomModal(`
        <div class="p-5 sm:p-6">
          <h3 class="font-bold text-slate-900 text-base">Konfirmasi Hapus Indikator</h3>
          <p class="text-xs text-slate-600 mt-2 leading-relaxed">
            Apakah Anda yakin ingin menghapus indikator <strong>"${escapeHtml(ind.nama)}"</strong>? Data analisis terkait indikator ini juga akan dilepas.
          </p>
          <div class="mt-5 flex justify-end gap-2 text-xs">
            <button onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
            <button onclick="executeDeleteIndicator('${indicatorId}')" class="px-4 py-2 rounded-xl bg-rose-600 text-white font-bold">Ya, Hapus Indikator</button>
          </div>
        </div>
      `, "max-w-md");
    }

    function executeDeleteIndicator(indicatorId) {
      const data = getCurrentData();
      data.indicators = data.indicators.filter(i => i.id !== indicatorId);
      data.rootCauses = data.rootCauses.filter(r => r.indicatorId !== indicatorId);
      data.priorities = data.priorities.filter(p => p.indicatorId !== indicatorId);
      closeCustomModal();
      triggerAutoSave("Input Data", "Hapus Indikator", "Menghapus 1 indikator dari tabel Rapor Pendidikan.");
      renderAll();
      showToast("Indikator telah dihapus.", "info");
    }

    function confirmClearAllIndicators() {
      openCustomModal(`
        <div class="p-5 sm:p-6">
          <h3 class="font-bold text-slate-900 text-base">Konfirmasi Kosongkan Tabel Indikator</h3>
          <p class="text-xs text-slate-600 mt-2">Seluruh indikator Rapor Pendidikan pada tahun aktif akan dihapus. Tindakan ini tidak dapat dibatalkan.</p>
          <div class="mt-5 flex justify-end gap-2 text-xs">
            <button onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
            <button onclick="executeClearAllIndicators()" class="px-4 py-2 rounded-xl bg-rose-600 text-white font-bold">Kosongkan Semua</button>
          </div>
        </div>
      `, "max-w-md");
    }

    function executeClearAllIndicators() {
      const data = getCurrentData();
      data.indicators = [];
      closeCustomModal();
      triggerAutoSave("Input Data", "Kosongkan Indikator", "Menghapus seluruh baris indikator.");
      renderAll();
      showToast("Tabel indikator telah dikosongkan.", "info");
    }

    // =========================================================================
    // 8. DIRECT EXCEL RAPOR PENDIDIKAN PARSER (Multi-Sheet Official Kemendikdasmen)
    // =========================================================================
    window._pendingExcelWorkbook = null;
    window._pendingExcelFileName = "";
    window._pendingExcelSchoolMeta = {};
    window._pendingExcelSheetName = "";
    window._pendingExcelItems = [];

    function downloadExcelTemplate() {
      if (!window.XLSX) {
        showToast("Pustaka Excel sedang dimuat, silakan coba beberapa detik lagi.", "warning");
        return;
      }
      const rows = [
        ["TEMPLATE INPUT RAPOR PENDIDIKAN — RAPOR SMART AI"],
        ["Nama Sekolah:", "", "NPSN:", ""],
        [],
        ["Kode Indikator", "Nama Indikator", "Dimensi", "Nilai Capaian", "Kategori Capaian", "Perubahan Skor", "Peringkat", "Rekomendasi Rapor Pendidikan"],
        ...KATALOG_INDIKATOR_STANDAR.slice(0, 8).map(k => [k.kode, k.nama, k.dimensi, "", "", "", "", ""])
      ];
      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Laporan Rapor Pendidikan");
      XLSX.writeFile(wb, "Template_Rapor_Pendidikan_RAPOR_SMART_AI.xlsx");
      showToast("Templat Excel berhasil diunduh.", "success");
    }

    function handleExcelUpload(event) {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (!file) return;

      if (!window.XLSX) {
        showToast("Pustaka pembaca Excel belum siap. Pastikan koneksi internet aktif lalu muat ulang.", "error");
        return;
      }

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const arrayBuffer = new Uint8Array(e.target.result);
          const wb = XLSX.read(arrayBuffer, { type: "array" });

          if (!wb.SheetNames || wb.SheetNames.length === 0) {
            showToast("File Excel kosong atau tidak memiliki lembar kerja (sheet).", "error");
            return;
          }

          window._pendingExcelWorkbook = wb;
          window._pendingExcelFileName = file.name;
          window._pendingExcelSchoolMeta = extractSchoolMetaFromWorkbook(wb);

          // Choose best sheet automatically (skip Panduan/Petunjuk/Cover)
          let bestSheet = wb.SheetNames[0];
          let bestCount = -1;
          wb.SheetNames.forEach(name => {
            const lower = name.toLowerCase();
            if (lower.includes("panduan") || lower.includes("petunjuk") || lower.includes("cover") || lower.includes("daftar isi")) {
              return;
            }
            const parsed = parseExcelSheetForRapor(wb.Sheets[name]);
            if (parsed.items.length > bestCount) {
              bestCount = parsed.items.length;
              bestSheet = name;
            }
          });

          // Fallback if all sheets were checked
          if (bestCount <= 0) {
            wb.SheetNames.forEach(name => {
              const parsed = parseExcelSheetForRapor(wb.Sheets[name]);
              if (parsed.items.length > bestCount) {
                bestCount = parsed.items.length;
                bestSheet = name;
              }
            });
          }

          window._pendingExcelSheetName = "";
          openExcelImportPreviewModal(bestSheet);
        } catch (err) {
          console.error(err);
          showToast("Gagal membaca file Excel. Pastikan file berformat .xlsx, .xls, atau .csv yang valid.", "error");
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function extractSchoolMetaFromWorkbook(wb) {
      const meta = { nama: "", npsn: "", kecamatan: "", kabupaten: "", provinsi: "" };
      wb.SheetNames.forEach(sName => {
        const sheet = wb.Sheets[sName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        for (let r = 0; r < Math.min(30, rows.length); r++) {
          const row = rows[r] || [];
          for (let c = 0; c < row.length; c++) {
            const cell = String(row[c] || "").trim();
            const low = cell.toLowerCase();
            const nextVal = String(row[c + 1] || row[c + 2] || "").replace(/^[:=\s]+/, "").trim();

            if (!meta.nama && (low === "nama sekolah" || low === "nama satuan pendidikan" || low.startsWith("nama satuan pendidikan:") || low.startsWith("nama sekolah:"))) {
              meta.nama = cell.includes(":") && cell.split(":")[1].trim() ? cell.split(":")[1].trim() : nextVal;
            }
            if (!meta.npsn && (low === "npsn" || low.startsWith("npsn:"))) {
              const rawNpsn = cell.includes(":") && cell.split(":")[1].trim() ? cell.split(":")[1].trim() : nextVal;
              const match8 = rawNpsn.match(/\d{8}/);
              meta.npsn = match8 ? match8[0] : rawNpsn;
            }
            if (!meta.kecamatan && (low === "kecamatan" || low.startsWith("kecamatan:"))) {
              meta.kecamatan = cell.includes(":") && cell.split(":")[1].trim() ? cell.split(":")[1].trim() : nextVal;
            }
            if (!meta.kabupaten && (low.includes("kabupaten") || low.includes("kab/kota") || low.includes("kota"))) {
              if (low.length < 30) {
                meta.kabupaten = cell.includes(":") && cell.split(":")[1].trim() ? cell.split(":")[1].trim() : nextVal;
              }
            }
            if (!meta.provinsi && (low === "provinsi" || low.startsWith("provinsi:"))) {
              meta.provinsi = cell.includes(":") && cell.split(":")[1].trim() ? cell.split(":")[1].trim() : nextVal;
            }
          }
        }
      });
      return meta;
    }

    function parseIndonesianScoreValue(rawScore) {
      if (rawScore === null || rawScore === undefined || rawScore === "") return NaN;
      if (typeof rawScore === "number") return rawScore;
      const cleaned = String(rawScore).replace(/%/g, "").replace(/,/g, ".").trim();
      const match = cleaned.match(/-?\d+(?:\.\d+)?/);
      if (!match) return NaN;
      return parseFloat(match[0]);
    }

    function parseDeltaValue(rawDelta) {
      if (rawDelta === null || rawDelta === undefined || rawDelta === "") return "";
      if (typeof rawDelta === "number") return rawDelta;
      const str = String(rawDelta).replace(/,/g, ".").trim();
      const match = str.match(/[+-]?\d+(?:\.\d+)?/);
      if (!match) {
        if (str.toLowerCase().includes("tetap")) return 0;
        return "";
      }
      let num = parseFloat(match[0]);
      if (str.toLowerCase().includes("turun") && num > 0) num = -num;
      return num;
    }

    function parseExcelSheetForRapor(sheet) {
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      if (!rows || rows.length === 0) return { items: [] };

      // 1. Detect header row dynamically within first 40 rows
      let headerIdx = -1;
      let bestScore = 0;

      for (let r = 0; r < Math.min(40, rows.length); r++) {
        const row = (rows[r] || []).map(v => String(v || "").toLowerCase().trim());
        let score = 0;
        row.forEach(cell => {
          if (!cell) return;
          if (cell.includes("indikator") || cell === "aspek" || cell === "komponen") score += 3;
          if (cell.includes("skor") || cell.includes("nilai") || cell.includes("capaian")) score += 3;
          if (cell.includes("label") || cell.includes("kategori") || cell.includes("predikat")) score += 2;
          if (cell.includes("perubahan") || cell.includes("delta") || cell.includes("selisih")) score += 2;
          if (cell.includes("dimensi") || cell.includes("peringkat") || cell.includes("benahi") || cell.includes("rekomendasi")) score += 1;
        });
        if (score > bestScore && score >= 4) {
          bestScore = score;
          headerIdx = r;
        }
      }

      if (headerIdx === -1) return { items: [] };

      const headers = (rows[headerIdx] || []).map(v => String(v || "").toLowerCase().trim());

      const findCol = (keywords, excludeKeywords = []) => {
        for (const kw of keywords) {
          const idx = headers.findIndex(h =>
            h.includes(kw) && !excludeKeywords.some(ex => h.includes(ex))
          );
          if (idx !== -1) return idx;
        }
        return -1;
      };

      const colKode = findCol(["kode indikator", "kode", "no. indikator", "no indikator", "id indikator"]);
      const colNama = findCol(["nama indikator", "indikator prioritas", "indikator", "aspek"], ["kode", "id", "definisi", "penjelasan"]);
      const colDimensi = findCol(["dimensi", "standar"]);
      const colNilai = findCol(["nilai capaian", "skor rapor", "skor tahun ini", "skor 2026", "skor 2025", "skor 2024", "nilai", "skor", "capaian"], ["label", "kategori", "perubahan", "sebelum", "lalu", "definisi"]);
      const colKategori = findCol(["kategori capaian", "label capaian", "kategori", "predikat", "status capaian", "tingkat capaian"]);
      const colPerubahan = findCol(["perubahan skor", "perubahan", "delta", "selisih", "tren"]);
      const colPeringkat = findCol(["peringkat", "kuantil", "posisi"]);
      const colRekom = findCol(["rekomendasi", "inspirasi benahi", "benahi", "tindak lanjut"]);

      if (colNama === -1) return { items: [] };

      const items = [];
      for (let r = headerIdx + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        let rawNama = String(row[colNama] || "").trim();
        if (!rawNama || rawNama.length < 3) continue;
        if (rawNama.toLowerCase() === "indikator" || rawNama.toLowerCase() === "nama indikator") continue;

        let kode = colKode !== -1 ? String(row[colKode] || "").trim() : "";
        // Extract leading code like "A.1 Kemampuan Literasi" if kode column is absent
        if (!kode) {
          const codeMatch = rawNama.match(/^([A-Z]\.\d+(?:\.\d+)?)\s*[-:.]?\s*(.+)$/i);
          if (codeMatch) {
            kode = codeMatch[1].toUpperCase();
            rawNama = codeMatch[2].trim();
          }
        }

        const dimensi = colDimensi !== -1 ? String(row[colDimensi] || "").trim() : "";
        let parsedScore = colNilai !== -1 ? parseIndonesianScoreValue(row[colNilai]) : NaN;

        // If score wasn't found in colNilai, check adjacent column in case of merged header cells
        if (isNaN(parsedScore) && colNilai !== -1 && row[colNilai + 1] !== undefined) {
          const adj = parseIndonesianScoreValue(row[colNilai + 1]);
          if (!isNaN(adj) && adj >= 0 && adj <= 100) parsedScore = adj;
        }

        const hasValidScore = !isNaN(parsedScore) && parsedScore >= 0 && parsedScore <= 100;
        const rawKat = colKategori !== -1 ? String(row[colKategori] || "").trim() : "";
        const finalKat = classifyCategory(hasValidScore ? parsedScore : "", rawKat);
        const perubahan = colPerubahan !== -1 ? parseDeltaValue(row[colPerubahan]) : "";
        const peringkat = colPeringkat !== -1 ? String(row[colPeringkat] || "").trim() : "";
        const rekomendasiRapor = colRekom !== -1 ? String(row[colRekom] || "").trim() : "";

        items.push({
          id: uid(),
          kode,
          nama: rawNama,
          dimensi,
          nilai: hasValidScore ? parsedScore : "",
          kategori: finalKat,
          perubahan,
          peringkat,
          rekomendasiRapor,
          terverifikasi: hasValidScore
        });
      }

      return { items };
    }

    function openExcelImportPreviewModal(selectedSheetName) {
      const wb = window._pendingExcelWorkbook;
      if (!wb) return;

      if (window._pendingExcelSheetName !== selectedSheetName) {
        const sheet = wb.Sheets[selectedSheetName];
        const parsed = parseExcelSheetForRapor(sheet);
        window._pendingExcelSheetName = selectedSheetName;
        window._pendingExcelItems = parsed.items;
      }

      const items = window._pendingExcelItems || [];
      const meta = window._pendingExcelSchoolMeta || {};
      const hasMeta = Boolean(meta.nama || meta.npsn || meta.kabupaten || meta.provinsi);

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[90vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div>
              <h3 class="font-bold text-slate-900 text-base">Konfirmasi Impor Excel Rapor Pendidikan</h3>
              <p class="text-xs text-slate-500">File: <strong>${escapeHtml(window._pendingExcelFileName)}</strong> • Ditemukan <strong>${items.length}</strong> indikator</p>
            </div>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>

          <div class="mt-4 space-y-4 text-xs">
            <!-- Sheet Selector -->
            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-slate-50 p-3 rounded-xl border border-slate-200">
              <span class="font-bold text-slate-700">Lembar Kerja (Sheet) Terpilih:</span>
              <select onchange="openExcelImportPreviewModal(this.value)" class="rounded-lg border border-slate-300 bg-white px-3 py-1.5 font-semibold">
                ${wb.SheetNames.map(s => `<option value="${escapeHtml(s)}" ${s === selectedSheetName ? "selected" : ""}>${escapeHtml(s)}</option>`).join("")}
              </select>
            </div>

            ${hasMeta ? `
              <div class="p-3 rounded-xl bg-emerald-50 border border-emerald-200 flex items-start gap-2.5">
                <input type="checkbox" id="applyDetectedSchoolMeta" checked class="mt-0.5 rounded text-emerald-700" />
                <label for="applyDetectedSchoolMeta" class="cursor-pointer">
                  <div class="font-bold text-emerald-950">Identitas Sekolah Terdeteksi dari Dokumen:</div>
                  <div class="text-emerald-800 mt-0.5">${escapeHtml(meta.nama || "-")} (NPSN: ${escapeHtml(meta.npsn || "-")}) ${meta.kabupaten ? "• " + escapeHtml(meta.kabupaten) : ""}</div>
                </label>
              </div>
            ` : ""}

            ${items.length === 0 ? `
              <div class="p-6 text-center rounded-xl border border-amber-200 bg-amber-50 text-amber-900">
                <p class="font-bold">Tidak ada tabel indikator yang terbaca pada sheet "${escapeHtml(selectedSheetName)}".</p>
                <p class="mt-1">Silakan pilih nama sheet lain di atas (misalnya <em>Laporan Rapor Pendidikan</em> atau <em>Capaian</em>).</p>
              </div>
            ` : `
              <div class="border border-slate-200 rounded-xl overflow-hidden">
                <div class="max-h-72 overflow-y-auto">
                  <table class="w-full text-left text-[11px]">
                    <thead class="bg-slate-100 border-b border-slate-200 sticky top-0">
                      <tr>
                        <th class="py-2 px-3">Kode</th>
                        <th class="py-2 px-3">Nama Indikator</th>
                        <th class="py-2 px-3 text-center">Nilai Capaian</th>
                        <th class="py-2 px-3 text-center">Kategori</th>
                        <th class="py-2 px-3 text-center">Perubahan</th>
                      </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                      ${items.map((it, idx) => `
                        <tr>
                          <td class="py-2 px-3 font-mono font-bold text-emerald-800">${escapeHtml(it.kode || "-")}</td>
                          <td class="py-2 px-3 font-semibold text-slate-900">${escapeHtml(it.nama)}</td>
                          <td class="py-2 px-3 text-center">
                            <input type="number" step="0.01" min="0" max="100"
                              value="${it.nilai !== "" ? it.nilai : ""}"
                              placeholder="Isi nilai"
                              oninput="updatePendingExcelScore(${idx}, this.value)"
                              class="w-20 text-center font-mono font-bold text-[11px] rounded-lg border ${it.nilai !== "" ? "border-slate-300 bg-white text-slate-900" : "border-rose-300 bg-rose-50 text-rose-800"} px-1.5 py-1 focus:border-emerald-700 focus:outline-none" />
                          </td>
                          <td id="previewExcelKat_${idx}" class="py-2 px-3 text-center font-semibold">${escapeHtml(it.kategori)}</td>
                          <td class="py-2 px-3 text-center font-mono">${it.perubahan !== "" ? it.perubahan : "-"}</td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              </div>

              <div class="flex flex-wrap items-center gap-4 pt-1">
                <label class="flex items-center gap-2 cursor-pointer font-semibold text-slate-700">
                  <input type="radio" name="excelImportMode" value="replace" checked class="text-emerald-700" />
                  <span>Ganti Data Tabel Saat Ini</span>
                </label>
                <label class="flex items-center gap-2 cursor-pointer font-semibold text-slate-700">
                  <input type="radio" name="excelImportMode" value="append" class="text-emerald-700" />
                  <span>Tambahkan ke Data yang Ada</span>
                </label>
              </div>
            `}

            <div class="pt-3 border-t border-slate-100 flex justify-end gap-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              ${items.length > 0 ? `
                <button type="button" onclick="confirmExcelImport('${escapeHtml(selectedSheetName)}')" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold shadow">
                  Konfirmasi &amp; Simpan (${items.length} Indikator)
                </button>
              ` : ""}
            </div>
          </div>
        </div>
      `, "max-w-3xl");
    }

    function updatePendingExcelScore(idx, rawVal) {
      const item = window._pendingExcelItems?.[idx];
      if (!item) return;
      const trimmed = String(rawVal || "").trim();
      if (trimmed === "") {
        item.nilai = "";
        item.kategori = "Belum Terverifikasi";
        item.terverifikasi = false;
      } else {
        const num = parseFloat(trimmed);
        if (!isNaN(num) && num >= 0 && num <= 100) {
          item.nilai = num;
          item.kategori = classifyCategory(num);
          item.terverifikasi = true;
        }
      }
      const katCell = document.getElementById(`previewExcelKat_${idx}`);
      if (katCell) katCell.textContent = item.kategori;
    }

    function confirmExcelImport(sheetName) {
      const wb = window._pendingExcelWorkbook;
      if (!wb || !wb.Sheets[sheetName]) return;

      const itemsToImport = window._pendingExcelItems && window._pendingExcelSheetName === sheetName
        ? window._pendingExcelItems
        : parseExcelSheetForRapor(wb.Sheets[sheetName]).items;

      const data = getCurrentData();
      const applyMeta = document.getElementById("applyDetectedSchoolMeta")?.checked;
      const meta = window._pendingExcelSchoolMeta || {};
      if (applyMeta) {
        if (meta.nama) data.school.nama = meta.nama;
        if (meta.npsn) data.school.npsn = meta.npsn;
        if (meta.kecamatan) data.school.kecamatan = meta.kecamatan;
        if (meta.kabupaten) data.school.kabupaten = meta.kabupaten;
        if (meta.provinsi) data.school.provinsi = meta.provinsi;
      }

      const mode = document.querySelector('input[name="excelImportMode"]:checked')?.value || "replace";
      if (mode === "replace") {
        data.indicators = itemsToImport;
      } else {
        itemsToImport.forEach(item => data.indicators.push(item));
      }

      const unverifiedCount = itemsToImport.filter(i => !i.terverifikasi).length;
      window._pendingExcelWorkbook = null;
      window._pendingExcelItems = [];
      window._pendingExcelSheetName = "";
      closeCustomModal();
      triggerAutoSave("Input Data", "Impor Excel Rapor", `Mengimpor ${itemsToImport.length} indikator dari sheet "${sheetName}".`);
      renderAll();

      if (unverifiedCount > 0) {
        showToast(`Berhasil mengimpor ${itemsToImport.length} indikator. ${unverifiedCount} indikator belum memiliki angka skor dan dapat langsung diisi pada kolom Nilai Capaian.`, "warning");
      } else {
        showToast(`Berhasil mengimpor ${itemsToImport.length} indikator Rapor Pendidikan beserta Nilai Capaiannya!`, "success");
      }
    }

    // =========================================================================
    // 9. PDF RAPOR PENDIDIKAN EXTRACTION & CONFIRMATION
    // =========================================================================
    window._pendingPdfItems = [];

    async function handlePdfUpload(event) {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (!file) return;

      if (!window.pdfjsLib) {
        showToast("Pustaka pembaca PDF belum tersedia. Gunakan unggah Excel atau Input Manual.", "warning");
        return;
      }

      showToast("Membaca dokumen PDF Rapor Pendidikan...", "info");
      try {
        const buffer = await file.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
        let fullText = "";
        for (let p = 1; p <= Math.min(pdf.numPages, 15); p++) {
          const page = await pdf.getPage(p);
          const content = await page.getTextContent();
          fullText += content.items.map(it => it.str).join(" ") + "\n";
        }

        // Match standard indicators mentioned in PDF text and extract nearby numeric score
        const detected = [];
        KATALOG_INDIKATOR_STANDAR.forEach(std => {
          const idx = fullText.toLowerCase().indexOf(std.nama.toLowerCase());
          if (idx !== -1) {
            const snippet = fullText.slice(idx, idx + 140);
            const scoreMatch = snippet.match(/\b(\d{1,2}[.,]\d{1,2}|100(?:[.,]0+)?)\b/);
            const parsedVal = scoreMatch ? parseFloat(scoreMatch[1].replace(",", ".")) : "";
            const valid = parsedVal !== "" && !isNaN(parsedVal) && parsedVal >= 0 && parsedVal <= 100;
            detected.push({
              id: uid(),
              kode: std.kode,
              nama: std.nama,
              dimensi: std.dimensi,
              nilai: valid ? parsedVal : "",
              kategori: valid ? classifyCategory(parsedVal) : "Belum Terverifikasi",
              perubahan: "",
              peringkat: "",
              rekomendasiRapor: "",
              terverifikasi: valid
            });
          }
        });

        window._pendingPdfItems = detected.length > 0 ? detected : KATALOG_INDIKATOR_STANDAR.slice(0, 6).map(std => ({
          id: uid(),
          kode: std.kode,
          nama: std.nama,
          dimensi: std.dimensi,
          nilai: "",
          kategori: "Belum Terverifikasi",
          perubahan: "",
          peringkat: "",
          rekomendasiRapor: "",
          terverifikasi: false
        }));

        openCustomModal(`
          <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
            <div class="flex items-center justify-between pb-3 border-b border-slate-200">
              <div>
                <h3 class="font-bold text-slate-900 text-base">Konfirmasi Ekstraksi PDF Rapor Pendidikan</h3>
                <p class="text-xs text-slate-500">Periksa dan lengkapi <strong>Nilai Capaian (0–100)</strong> sebelum disimpan (aplikasi tidak mengarang angka yang tidak terbaca).</p>
              </div>
              <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
            </div>
            <div class="mt-4 border border-slate-200 rounded-xl overflow-hidden">
              <table class="w-full text-left text-xs">
                <thead class="bg-slate-100 border-b border-slate-200">
                  <tr>
                    <th class="py-2 px-3">Indikator</th>
                    <th class="py-2 px-3 text-center">Nilai Capaian (0–100)</th>
                    <th class="py-2 px-3 text-center">Kategori</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                  ${window._pendingPdfItems.map((it, idx) => `
                    <tr>
                      <td class="py-2 px-3 font-semibold text-slate-900">[${escapeHtml(it.kode)}] ${escapeHtml(it.nama)}</td>
                      <td class="py-2 px-3 text-center">
                        <input type="number" step="0.01" min="0" max="100" value="${it.nilai}" placeholder="Isi nilai"
                          oninput="window._pendingPdfItems[${idx}].nilai = this.value !== '' ? parseFloat(this.value) : ''; window._pendingPdfItems[${idx}].kategori = classifyCategory(this.value); window._pendingPdfItems[${idx}].terverifikasi = this.value !== ''; document.getElementById('pdfKat_${idx}').textContent = window._pendingPdfItems[${idx}].kategori;"
                          class="w-24 text-center font-mono font-bold rounded-lg border border-slate-300 px-2 py-1" />
                      </td>
                      <td id="pdfKat_${idx}" class="py-2 px-3 text-center font-semibold">${escapeHtml(it.kategori)}</td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
            <div class="mt-4 flex justify-end gap-2 text-xs">
              <button onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button onclick="confirmPdfImport()" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Hasil PDF</button>
            </div>
          </div>
        `, "max-w-2xl");
      } catch (err) {
        console.error(err);
        showToast("Dokumen PDF tidak dapat diekstrak secara otomatis (kemungkinan hasil pindaian gambar). Silakan gunakan Excel atau Input Manual.", "warning");
      }
    }

    function confirmPdfImport() {
      const data = getCurrentData();
      window._pendingPdfItems.forEach(item => data.indicators.push(item));
      closeCustomModal();
      triggerAutoSave("Input Data", "Impor PDF Rapor", `Menambahkan ${window._pendingPdfItems.length} indikator dari PDF.`);
      renderAll();
      showToast("Data indikator dari PDF berhasil disimpan.", "success");
    }

    // =========================================================================
    // 10. MODULE 2: ANALISIS CAPAIAN RAPOR PENDIDIKAN
    // =========================================================================
    function setAnalisisFilter(filterId) {
      APP_STATE.analisisFilter = filterId;
      renderAll();
    }

    function renderAnalisisModule() {
      const data = getCurrentData();
      const all = data.indicators;

      const baik = all.filter(i => classifyCategory(i.nilai, i.kategori) === "Baik");
      const sedang = all.filter(i => classifyCategory(i.nilai, i.kategori) === "Sedang");
      const kurang = all.filter(i => classifyCategory(i.nilai, i.kategori) === "Kurang");
      const turun = all.filter(i => i.perubahan !== "" && parseFloat(i.perubahan) < 0);
      const naik = all.filter(i => i.perubahan !== "" && parseFloat(i.perubahan) > 0);
      const prioritas = all.filter(i => {
        const k = classifyCategory(i.nilai, i.kategori);
        return k === "Kurang" || k === "Sedang" || (i.perubahan !== "" && parseFloat(i.perubahan) < 0);
      });

      const filters = [
        { id: "all", label: `Semua (${all.length})` },
        { id: "prioritas", label: `Prioritas Pembenahan (${prioritas.length})` },
        { id: "kurang", label: `Capaian Kurang (${kurang.length})` },
        { id: "sedang", label: `Capaian Sedang (${sedang.length})` },
        { id: "baik", label: `Capaian Baik (${baik.length})` },
        { id: "turun", label: `Mengalami Penurunan (${turun.length})` },
        { id: "naik", label: `Mengalami Peningkatan (${naik.length})` }
      ];

      let displayed = all;
      if (APP_STATE.analisisFilter === "prioritas") displayed = prioritas;
      if (APP_STATE.analisisFilter === "kurang") displayed = kurang;
      if (APP_STATE.analisisFilter === "sedang") displayed = sedang;
      if (APP_STATE.analisisFilter === "baik") displayed = baik;
      if (APP_STATE.analisisFilter === "turun") displayed = turun;
      if (APP_STATE.analisisFilter === "naik") displayed = naik;

      return `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 2 • Pemetaan Capaian</span>
            <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Analisis Capaian Rapor Pendidikan</h2>
            <p class="text-xs text-slate-600 mt-0.5">Pengelompokan otomatis berdasarkan <strong>Nilai Capaian</strong>, kategori, dan tren perubahan dibanding tahun lalu.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button onclick="exportPrintReport('analisis')" class="px-3.5 py-2.5 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-slate-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="download" class="w-4 h-4 text-rose-700"></i>
              <span>Unduh PDF Analisis</span>
            </button>
            <button onclick="generateRootCausesAndPriorities()" class="px-4 py-2.5 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center gap-2 shadow">
              <i data-lucide="sparkles" class="w-4 h-4 text-amber-300"></i>
              <span>Lanjut Analisis Akar Masalah &amp; Prioritas</span>
            </button>
          </div>
        </div>

        <!-- SUMMARY CARDS -->
        <section class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 text-xs">
          <div class="card-elevated p-3.5 border-t-4 border-t-emerald-600">
            <div class="text-slate-500 font-semibold">Capaian Baik</div>
            <div class="font-mono text-xl font-extrabold text-emerald-800 mt-1">${baik.length}</div>
          </div>
          <div class="card-elevated p-3.5 border-t-4 border-t-amber-500">
            <div class="text-slate-500 font-semibold">Capaian Sedang</div>
            <div class="font-mono text-xl font-extrabold text-amber-700 mt-1">${sedang.length}</div>
          </div>
          <div class="card-elevated p-3.5 border-t-4 border-t-rose-600">
            <div class="text-slate-500 font-semibold">Capaian Kurang</div>
            <div class="font-mono text-xl font-extrabold text-rose-700 mt-1">${kurang.length}</div>
          </div>
          <div class="card-elevated p-3.5 border-t-4 border-t-rose-400">
            <div class="text-slate-500 font-semibold">Menurun</div>
            <div class="font-mono text-xl font-extrabold text-rose-700 mt-1">${turun.length}</div>
          </div>
          <div class="card-elevated p-3.5 border-t-4 border-t-teal-500">
            <div class="text-slate-500 font-semibold">Meningkat</div>
            <div class="font-mono text-xl font-extrabold text-teal-700 mt-1">${naik.length}</div>
          </div>
          <div class="card-elevated p-3.5 border-t-4 border-t-slate-900 bg-amber-50/40">
            <div class="text-slate-700 font-bold">Prioritas Benahi</div>
            <div class="font-mono text-xl font-extrabold text-slate-900 mt-1">${prioritas.length}</div>
          </div>
        </section>

        <!-- FILTER PILLS -->
        <div class="flex flex-wrap gap-1.5">
          ${filters.map(f => `
            <button onclick="setAnalisisFilter('${f.id}')"
              class="px-3 py-1.5 rounded-xl text-xs font-bold transition border ${APP_STATE.analisisFilter === f.id ? 'emerald-gradient text-white border-emerald-900 shadow-xs' : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'}">
              ${f.label}
            </button>
          `).join("")}
        </div>

        <!-- ANALISIS TABLE -->
        <section class="card-elevated overflow-hidden">
          ${displayed.length === 0 ? `
            <div class="p-10 text-center text-xs text-slate-500">
              Tidak ada indikator pada kelompok filter ini. Silakan pilih filter lain atau masukkan data Rapor Pendidikan terlebih dahulu.
            </div>
          ` : `
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs">
                <thead class="bg-slate-50 border-b border-slate-200 text-slate-600 uppercase text-[10px]">
                  <tr>
                    <th class="py-3 px-4">Nama Indikator</th>
                    <th class="py-3 px-3 text-center">Nilai Capaian</th>
                    <th class="py-3 px-3 text-center">Kategori</th>
                    <th class="py-3 px-3 text-center">Perubahan</th>
                    <th class="py-3 px-3 text-center">Status</th>
                    <th class="py-3 px-4 text-center">Prioritas Pembenahan</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                  ${displayed.map(ind => {
                    const kat = classifyCategory(ind.nilai, ind.kategori);
                    const isPrio = kat === "Kurang" || kat === "Sedang" || (ind.perubahan !== "" && parseFloat(ind.perubahan) < 0);
                    return `
                      <tr class="hover:bg-slate-50">
                        <td class="py-3 px-4">
                          <div class="font-bold text-slate-900">${ind.kode ? `<span class="font-mono text-emerald-800 mr-1">[${escapeHtml(ind.kode)}]</span>` : ""}${escapeHtml(ind.nama)}</div>
                          <div class="text-[11px] text-slate-500">${escapeHtml(ind.dimensi || "-")}</div>
                        </td>
                        <td class="py-3 px-3 text-center font-mono font-extrabold text-sm text-slate-900">${formatScore(ind.nilai)}</td>
                        <td class="py-3 px-3 text-center">${getCategoryBadgeHtml(kat)}</td>
                        <td class="py-3 px-3 text-center">${getDeltaBadgeHtml(ind.perubahan)}</td>
                        <td class="py-3 px-3 text-center">
                          ${ind.terverifikasi
                            ? `<span class="text-emerald-800 font-bold text-[11px]">Terkonfirmasi</span>`
                            : `<span class="text-amber-700 font-semibold text-[11px]">Perlu Verifikasi</span>`}
                        </td>
                        <td class="py-3 px-4 text-center">
                          ${isPrio
                            ? `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full bg-rose-100 text-rose-900 font-bold text-[11px] border border-rose-200">Prioritas Dibenahi</span>`
                            : `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-800 font-semibold text-[11px]">Pertahankan</span>`}
                        </td>
                      </tr>
                    `;
                  }).join("")}
                </tbody>
              </table>
            </div>
          `}
        </section>
      `;
    }

    // =========================================================================
    // 11. MODULE 3: ANALISIS AKAR MASALAH (5-WHY)
    // =========================================================================
    function generateRootCausesAndPriorities() {
      const data = getCurrentData();
      if (data.indicators.length === 0) {
        showToast("Belum ada data indikator Rapor Pendidikan. Silakan input data terlebih dahulu.", "warning");
        setActiveTab("input");
        return;
      }

      const targets = data.indicators.filter(i => {
        const k = classifyCategory(i.nilai, i.kategori);
        return k === "Kurang" || k === "Sedang" || (i.perubahan !== "" && parseFloat(i.perubahan) < 0);
      });

      const listToProcess = targets.length > 0 ? targets : data.indicators;

      listToProcess.forEach((ind, idx) => {
        const kat = classifyCategory(ind.nilai, ind.kategori);
        const delta = parseFloat(ind.perubahan);

        if (!data.rootCauses.some(r => r.indicatorId === ind.id)) {
          data.rootCauses.push({
            id: uid(),
            indicatorId: ind.id,
            masalahUtama: `Capaian indikator ${ind.nama} berada pada nilai ${formatScore(ind.nilai)} (${kat})${!isNaN(delta) ? ` dengan perubahan ${delta > 0 ? "+" + delta : delta} poin` : ""}.`,
            dugaanPenyebab: [
              {
                id: uid(),
                teks: `Strategi pembelajaran dan pendampingan di kelas pada aspek ${ind.nama} belum berjalan merata dan kontekstual.`,
                terkonfirmasi: false
              },
              {
                id: uid(),
                teks: `Optimalisasi Komunitas Belajar (Kombel) sekolah untuk refleksi perangkat ajar dan asesmen diagnostik ${ind.nama} masih perlu ditingkatkan.`,
                terkonfirmasi: false
              }
            ],
            whyChain: [
              `Why 1: Nilai capaian ${ind.nama} (${formatScore(ind.nilai)}) belum mencapai kategori optimal.`,
              `Why 2: Praktik pembelajaran sehari-hari belum sepenuhnya menyesuaikan kebutuhan belajar peserta didik berdasarkan asesmen awal.`,
              `Why 3: Berbagi praktik baik antar guru melalui observasi kelas dan diskusi Kombel belum terjadwal rutin.`,
              `Why 4: Ketersediaan perangkat ajar, bahan bacaan/peraga pendukung, serta supervisi akademik perlu diperkuat.`
            ],
            buktiPendukung: `Nilai Rapor Pendidikan ${ind.kode || ""} ${ind.nama} = ${formatScore(ind.nilai)} (${kat}).${ind.rekomendasiRapor ? " Catatan Rapor: " + ind.rekomendasiRapor : ""}`,
            pertanyaanRefleksi: [
              `Apa tantangan nyata yang dihadapi guru di kelas dalam meningkatkan ${ind.nama}?`,
              `Sumber daya atau media belajar apa yang paling dibutuhkan siswa untuk mengejar ketertinggalan pada aspek ini?`,
              `Bagaimana Komunitas Belajar sekolah dapat memfasilitasi perbaikan praktik mengajar secara berkala?`
            ],
            dataTambahan: "Hasil asesmen diagnostik kelas, catatan supervisi akademik kepala sekolah, dan rekap hasil belajar siswa.",
            dampakPesertaDidik: `Peserta didik berisiko belum menguasai kompetensi esensial pada ${ind.nama} yang berdampak pada kesiapan belajar di jenjang berikutnya.`,
            alternatifSolusi: ind.rekomendasiRapor || `1) Lokakarya/IHT peningkatan kompetensi guru pada ${ind.nama}; 2) Penguatan Komunitas Belajar & observasi rekan sejawat; 3) Penyediaan sarana/bahan ajar pendukung.`
          });
        }

        if (!data.priorities.some(p => p.indicatorId === ind.id)) {
          const skorCapaian = kat === "Kurang" ? 5 : kat === "Sedang" ? 4 : 2;
          const skorTren = !isNaN(delta) && delta < 0 ? 5 : 3;
          const skorDampak = (ind.kode === "A.1" || ind.kode === "A.2" || ind.kode === "D.1") ? 5 : 4;
          const skorUrgensi = kat === "Kurang" ? 5 : 4;
          const skorKemampuan = 4;
          const totalSkor = skorCapaian + skorTren + skorDampak + skorUrgensi + skorKemampuan;

          data.priorities.push({
            id: uid(),
            indicatorId: ind.id,
            skorCapaian,
            skorTren,
            skorDampak,
            skorUrgensi,
            skorKemampuan,
            totalSkor,
            peringkatManual: data.priorities.length + 1,
            alasanPrioritas: `Nilai capaian ${formatScore(ind.nilai)} (${kat})${!isNaN(delta) && delta < 0 ? ` dan mengalami penurunan (${delta})` : ""} sehingga berdampak langsung pada mutu layanan belajar peserta didik.`
          });
        }
      });

      triggerAutoSave("Akar Masalah & Prioritas", "Analisis Otomatis", `Menyusun analisis akar masalah dan peringkat prioritas untuk ${listToProcess.length} indikator.`);
      setActiveTab("akar");
      showToast("Analisis akar masalah (5-Why) dan peringkat prioritas berhasil disusun!", "success");
    }

    function renderAkarMasalahModule() {
      const data = getCurrentData();

      return `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 3 • Diagnosis Mendalam (5-Why)</span>
            <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Analisis Masalah dan Akar Masalah</h2>
            <p class="text-xs text-slate-600 mt-0.5">Bedakan secara tegas antara <strong>Dugaan AI</strong> dan <strong>Fakta Terkonfirmasi Sekolah</strong> melalui diskusi dewan guru.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button onclick="exportWordDocument('akar')" class="px-3.5 py-2 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-blue-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="download" class="w-3.5 h-3.5"></i>
              <span>Unduh Word</span>
            </button>
            <button onclick="generateRootCausesAndPriorities()" class="px-3.5 py-2 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-slate-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="refresh-cw" class="w-3.5 h-3.5 text-emerald-700"></i>
              <span>Generate dari Indikator</span>
            </button>
            <button onclick="setActiveTab('prioritas')" class="px-4 py-2 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center gap-1.5 shadow">
              <span>Lanjut Penentuan Prioritas</span>
              <i data-lucide="arrow-right" class="w-4 h-4"></i>
            </button>
          </div>
        </div>

        ${data.rootCauses.length === 0 ? `
          <div class="card-elevated p-10 text-center space-y-3">
            <p class="text-xs text-slate-500">Belum ada analisis akar masalah. Klik tombol <strong>Generate dari Indikator</strong> untuk memproses indikator Rapor Pendidikan Anda.</p>
            <button onclick="generateRootCausesAndPriorities()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Mulai Analisis Akar Masalah</button>
          </div>
        ` : `
          <div class="space-y-5">
            ${data.rootCauses.map(rc => {
              const ind = data.indicators.find(i => i.id === rc.indicatorId) || { nama: "Indikator", kode: "", nilai: "", kategori: "" };
              return `
                <article class="card-elevated p-5 space-y-4">
                  <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-3 border-b border-slate-200">
                    <div class="flex items-center gap-2 flex-wrap">
                      <h3 class="font-bold text-slate-900 text-base">${ind.kode ? `[${escapeHtml(ind.kode)}] ` : ""}${escapeHtml(ind.nama)}</h3>
                      <span class="px-2.5 py-0.5 rounded-full bg-slate-100 border border-slate-200 font-mono text-xs font-bold text-slate-800">Nilai Capaian: ${formatScore(ind.nilai)}</span>
                      ${getCategoryBadgeHtml(classifyCategory(ind.nilai, ind.kategori))}
                    </div>
                    <button onclick="openEditRootCauseModal('${rc.id}')" class="px-3 py-1.5 rounded-xl bg-emerald-50 text-emerald-900 border border-emerald-200 hover:bg-emerald-100 text-xs font-bold flex items-center gap-1.5 self-start">
                      <i data-lucide="pencil" class="w-3.5 h-3.5"></i>
                      <span>Edit &amp; Konfirmasi Sekolah</span>
                    </button>
                  </div>

                  <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 text-xs">
                    <div class="space-y-3">
                      <div class="p-3.5 rounded-xl bg-rose-50/70 border border-rose-200">
                        <div class="font-bold text-rose-900 mb-1">1. Masalah Utama</div>
                        <p class="text-slate-800 leading-relaxed">${escapeHtml(rc.masalahUtama)}</p>
                      </div>

                      <div class="p-3.5 rounded-xl bg-slate-50 border border-slate-200">
                        <div class="font-bold text-slate-900 mb-2">2. Kemungkinan Akar Penyebab (Klik untuk Konfirmasi Fakta Sekolah)</div>
                        <div class="space-y-2">
                          ${(rc.dugaanPenyebab || []).map(c => `
                            <label class="flex items-start gap-2.5 p-2 rounded-lg bg-white border ${c.terkonfirmasi ? "border-emerald-300" : "border-amber-200"} cursor-pointer">
                              <input type="checkbox" ${c.terkonfirmasi ? "checked" : ""} onchange="toggleCauseConfirmed('${rc.id}', '${c.id}')" class="mt-0.5 rounded text-emerald-700" />
                              <div class="flex-1">
                                <span class="inline-block px-1.5 py-0.5 rounded text-[10px] font-bold mb-1 ${c.terkonfirmasi ? "bg-emerald-100 text-emerald-900" : "bg-amber-100 text-amber-900"}">
                                  ${c.terkonfirmasi ? "FAKTA TERKONFIRMASI SEKOLAH" : "DUGAAN / HIPOTESIS AI (PERLU VERIFIKASI)"}
                                </span>
                                <p class="text-slate-800 leading-relaxed">${escapeHtml(c.teks)}</p>
                              </div>
                            </label>
                          `).join("")}
                        </div>
                      </div>

                      <div class="p-3.5 rounded-xl bg-emerald-50/50 border border-emerald-200">
                        <div class="font-bold text-emerald-950 mb-1.5">Penelusuran 5-Why</div>
                        <ul class="space-y-1 text-slate-700 list-disc list-inside">
                          ${(rc.whyChain || []).map(w => `<li>${escapeHtml(w)}</li>`).join("")}
                        </ul>
                      </div>
                    </div>

                    <div class="space-y-3">
                      <div class="p-3.5 rounded-xl bg-white border border-slate-200">
                        <div class="font-bold text-slate-900 mb-1">3. Bukti / Data Pendukung</div>
                        <p class="text-slate-700 leading-relaxed">${escapeHtml(rc.buktiPendukung)}</p>
                      </div>

                      <div class="p-3.5 rounded-xl bg-amber-50/60 border border-amber-200">
                        <div class="font-bold text-amber-950 mb-1">4. Pertanyaan Refleksi untuk Diskusi Guru</div>
                        <ul class="space-y-1 text-slate-800 list-decimal list-inside">
                          ${(rc.pertanyaanRefleksi || []).map(q => `<li>${escapeHtml(q)}</li>`).join("")}
                        </ul>
                      </div>

                      <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                          <div class="font-bold text-slate-900 mb-1">5. Data Tambahan Perlu Dikumpulkan</div>
                          <p class="text-slate-700">${escapeHtml(rc.dataTambahan)}</p>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                          <div class="font-bold text-slate-900 mb-1">6. Dampak Terhadap Peserta Didik</div>
                          <p class="text-slate-700">${escapeHtml(rc.dampakPesertaDidik)}</p>
                        </div>
                      </div>

                      <div class="p-3.5 rounded-xl bg-emerald-900 text-white">
                        <div class="font-bold text-amber-300 mb-1">7. Alternatif Solusi Pembenahan</div>
                        <p class="text-emerald-50 leading-relaxed">${escapeHtml(rc.alternatifSolusi)}</p>
                      </div>
                    </div>
                  </div>
                </article>
              `;
            }).join("")}
          </div>
        `}
      `;
    }

    function toggleCauseConfirmed(rcId, causeId) {
      const data = getCurrentData();
      const rc = data.rootCauses.find(r => r.id === rcId);
      const cause = rc?.dugaanPenyebab?.find(c => c.id === causeId);
      if (!cause) return;
      cause.terkonfirmasi = !cause.terkonfirmasi;
      triggerAutoSave("Akar Masalah", "Konfirmasi Akar Penyebab", `Mengubah status penyebab menjadi ${cause.terkonfirmasi ? "Fakta Terkonfirmasi" : "Dugaan AI"}.`);
      renderAll();
    }

    function openEditRootCauseModal(rcId) {
      const data = getCurrentData();
      const rc = data.rootCauses.find(r => r.id === rcId);
      if (!rc) return;

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Sunting Analisis Akar Masalah</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRootCauseEdit(event, '${rc.id}')" class="mt-4 space-y-3 text-xs">
            <div>
              <label class="block font-bold text-slate-700 mb-1">1. Masalah Utama</label>
              <textarea name="masalahUtama" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rc.masalahUtama)}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">2. Daftar Akar Penyebab (Satu penyebab per baris)</label>
              <textarea name="penyebabLines" rows="3" class="w-full rounded-xl border border-slate-300 px-3 py-2">${(rc.dugaanPenyebab || []).map(c => c.teks).join("\n")}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">3. Bukti / Data Pendukung</label>
              <textarea name="buktiPendukung" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rc.buktiPendukung)}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">6. Dampak Terhadap Peserta Didik</label>
              <textarea name="dampakPesertaDidik" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rc.dampakPesertaDidik)}</textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">7. Alternatif Solusi</label>
              <textarea name="alternatifSolusi" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rc.alternatifSolusi)}</textarea>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Perubahan</button>
            </div>
          </form>
        </div>
      `);
    }

    function saveRootCauseEdit(e, rcId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const rc = data.rootCauses.find(r => r.id === rcId);
      if (!rc) return;

      rc.masalahUtama = String(fd.get("masalahUtama") || "").trim();
      const lines = String(fd.get("penyebabLines") || "").split("\n").map(s => s.trim()).filter(Boolean);
      rc.dugaanPenyebab = lines.map((txt, idx) => ({
        id: rc.dugaanPenyebab[idx]?.id || uid(),
        teks: txt,
        terkonfirmasi: rc.dugaanPenyebab[idx]?.terkonfirmasi || false
      }));
      rc.buktiPendukung = String(fd.get("buktiPendukung") || "").trim();
      rc.dampakPesertaDidik = String(fd.get("dampakPesertaDidik") || "").trim();
      rc.alternatifSolusi = String(fd.get("alternatifSolusi") || "").trim();

      closeCustomModal();
      triggerAutoSave("Akar Masalah", "Edit Akar Masalah", "Menyunting analisis akar masalah sekolah.");
      renderAll();
      showToast("Analisis akar masalah berhasil diperbarui.", "success");
    }

    // =========================================================================
    // 12. MODULE 4: PENENTUAN PRIORITAS PEMBENAHAN
    // =========================================================================
    function renderPrioritasModule() {
      const data = getCurrentData();
      const sorted = [...data.priorities].sort((a, b) => (a.peringkatManual || 999) - (b.peringkatManual || 999) || b.totalSkor - a.totalSkor);

      return `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 4 • Pembobotan &amp; Urutan Fokus</span>
            <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Penentuan Prioritas Pembenahan</h2>
            <p class="text-xs text-slate-600 mt-0.5">Pemeringkatan berdasarkan <strong>Nilai Capaian</strong>, tren perubahan, dampak ke siswa, urgensi, dan kapasitas sekolah (dapat diatur manual).</p>
          </div>
          <button onclick="generateRecommendationsFromPriorities()" class="px-4 py-2.5 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center gap-2 shadow">
            <i data-lucide="lightbulb" class="w-4 h-4 text-amber-300"></i>
            <span>Buat Rekomendasi Program</span>
          </button>
        </div>

        ${sorted.length === 0 ? `
          <div class="card-elevated p-10 text-center space-y-3">
            <p class="text-xs text-slate-500">Belum ada daftar prioritas pembenahan.</p>
            <button onclick="generateRootCausesAndPriorities()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Hitung Prioritas dari Rapor</button>
          </div>
        ` : `
          <div class="space-y-3">
            ${sorted.map((prio, index) => {
              const ind = data.indicators.find(i => i.id === prio.indicatorId) || { nama: "Indikator", kode: "", nilai: "", kategori: "" };
              return `
                <div class="card-elevated p-4 sm:p-5 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                  <div class="flex items-start gap-3.5">
                    <div class="w-10 h-10 rounded-xl emerald-gradient text-amber-300 font-heading font-extrabold text-base flex items-center justify-center shrink-0 shadow">
                      #${index + 1}
                    </div>
                    <div>
                      <div class="flex items-center gap-2 flex-wrap">
                        <h3 class="font-bold text-slate-900 text-sm sm:text-base">${ind.kode ? `[${escapeHtml(ind.kode)}] ` : ""}${escapeHtml(ind.nama)}</h3>
                        <span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-mono font-bold bg-slate-100 text-slate-800 border border-slate-200">
                          Nilai Capaian: ${formatScore(ind.nilai)}
                        </span>
                        ${getCategoryBadgeHtml(classifyCategory(ind.nilai, ind.kategori))}
                      </div>
                      <p class="text-xs text-slate-600 mt-1 leading-relaxed"><strong>Alasan Prioritas:</strong> ${escapeHtml(prio.alasanPrioritas)}</p>
                      <div class="flex flex-wrap gap-2 mt-2 text-[11px] font-mono">
                        <span class="px-2 py-0.5 rounded bg-slate-100">Capaian: ${prio.skorCapaian}/5</span>
                        <span class="px-2 py-0.5 rounded bg-slate-100">Tren: ${prio.skorTren}/5</span>
                        <span class="px-2 py-0.5 rounded bg-slate-100">Dampak: ${prio.skorDampak}/5</span>
                        <span class="px-2 py-0.5 rounded bg-slate-100">Urgensi: ${prio.skorUrgensi}/5</span>
                        <span class="px-2 py-0.5 rounded bg-slate-100">Kapasitas: ${prio.skorKemampuan}/5</span>
                        <span class="px-2 py-0.5 rounded bg-amber-100 text-amber-950 font-bold">Total Skor: ${prio.totalSkor}/25</span>
                      </div>
                    </div>
                  </div>

                  <div class="flex items-center gap-1.5 self-end lg:self-center shrink-0">
                    <button onclick="movePriorityOrder('${prio.id}', -1)" ${index === 0 ? "disabled" : ""} class="p-2 rounded-xl border border-slate-200 hover:bg-slate-50 disabled:opacity-40" title="Naikkan Urutan Manual">
                      <i data-lucide="arrow-up" class="w-4 h-4"></i>
                    </button>
                    <button onclick="movePriorityOrder('${prio.id}', 1)" ${index === sorted.length - 1 ? "disabled" : ""} class="p-2 rounded-xl border border-slate-200 hover:bg-slate-50 disabled:opacity-40" title="Turunkan Urutan Manual">
                      <i data-lucide="arrow-down" class="w-4 h-4"></i>
                    </button>
                    <button onclick="openEditPriorityModal('${prio.id}')" class="px-3 py-2 rounded-xl bg-emerald-50 text-emerald-900 border border-emerald-200 text-xs font-bold hover:bg-emerald-100">
                      Ubah Bobot / Alasan
                    </button>
                  </div>
                </div>
              `;
            }).join("")}
          </div>
        `}
      `;
    }

    function movePriorityOrder(prioId, direction) {
      const data = getCurrentData();
      const sorted = [...data.priorities].sort((a, b) => (a.peringkatManual || 999) - (b.peringkatManual || 999));
      const idx = sorted.findIndex(p => p.id === prioId);
      const targetIdx = idx + direction;
      if (idx < 0 || targetIdx < 0 || targetIdx >= sorted.length) return;

      const temp = sorted[idx];
      sorted[idx] = sorted[targetIdx];
      sorted[targetIdx] = temp;

      sorted.forEach((p, i) => {
        p.peringkatManual = i + 1;
      });
      triggerAutoSave("Prioritas", "Ubah Urutan Prioritas", "Menyesuaikan peringkat prioritas secara manual.");
      renderAll();
    }

    function openEditPriorityModal(prioId) {
      const data = getCurrentData();
      const prio = data.priorities.find(p => p.id === prioId);
      if (!prio) return;

      openCustomModal(`
        <div class="p-5 sm:p-6">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Sesuaikan Skor &amp; Alasan Prioritas</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="savePriorityEdit(event, '${prio.id}')" class="mt-4 space-y-3 text-xs">
            <div class="grid grid-cols-2 sm:grid-cols-5 gap-2">
              <div><label class="block font-bold mb-1">Capaian (1-5)</label><input type="number" name="skorCapaian" min="1" max="5" value="${prio.skorCapaian}" class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5 font-mono" /></div>
              <div><label class="block font-bold mb-1">Tren (1-5)</label><input type="number" name="skorTren" min="1" max="5" value="${prio.skorTren}" class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5 font-mono" /></div>
              <div><label class="block font-bold mb-1">Dampak (1-5)</label><input type="number" name="skorDampak" min="1" max="5" value="${prio.skorDampak}" class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5 font-mono" /></div>
              <div><label class="block font-bold mb-1">Urgensi (1-5)</label><input type="number" name="skorUrgensi" min="1" max="5" value="${prio.skorUrgensi}" class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5 font-mono" /></div>
              <div><label class="block font-bold mb-1">Kapasitas (1-5)</label><input type="number" name="skorKemampuan" min="1" max="5" value="${prio.skorKemampuan}" class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5 font-mono" /></div>
            </div>
            <div>
              <label class="block font-bold text-slate-700 mb-1">Alasan Penetapan Prioritas</label>
              <textarea name="alasanPrioritas" rows="3" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(prio.alasanPrioritas)}</textarea>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Prioritas</button>
            </div>
          </form>
        </div>
      `);
    }

    function savePriorityEdit(e, prioId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const prio = data.priorities.find(p => p.id === prioId);
      if (!prio) return;

      prio.skorCapaian = Number(fd.get("skorCapaian")) || 3;
      prio.skorTren = Number(fd.get("skorTren")) || 3;
      prio.skorDampak = Number(fd.get("skorDampak")) || 3;
      prio.skorUrgensi = Number(fd.get("skorUrgensi")) || 3;
      prio.skorKemampuan = Number(fd.get("skorKemampuan")) || 3;
      prio.totalSkor = prio.skorCapaian + prio.skorTren + prio.skorDampak + prio.skorUrgensi + prio.skorKemampuan;
      prio.alasanPrioritas = String(fd.get("alasanPrioritas") || "").trim();

      closeCustomModal();
      triggerAutoSave("Prioritas", "Edit Prioritas", "Memperbarui bobot dan alasan prioritas.");
      renderAll();
      showToast("Prioritas pembenahan diperbarui.", "success");
    }

    // =========================================================================
    // 13. MODULE 5: REKOMENDASI PROGRAM
    // =========================================================================
    function generateRecommendationsFromPriorities() {
      const data = getCurrentData();
      if (data.priorities.length === 0) {
        generateRootCausesAndPriorities();
      }

      const sorted = [...data.priorities].sort((a, b) => (a.peringkatManual || 999) - (b.peringkatManual || 999));
      sorted.forEach(prio => {
        if (!data.recommendations.some(r => r.indicatorId === prio.indicatorId)) {
          const ind = data.indicators.find(i => i.id === prio.indicatorId);
          const rc = data.rootCauses.find(r => r.indicatorId === prio.indicatorId);
          if (!ind) return;

          const currentScore = parseFloat(ind.nilai);
          const targetScore = !isNaN(currentScore) ? Math.min(100, currentScore + 8).toFixed(2) : "Kategori Baik";

          data.recommendations.push({
            id: uid(),
            indicatorId: ind.id,
            permasalahan: rc ? rc.masalahUtama : `Capaian ${ind.nama} berada pada nilai ${formatScore(ind.nilai)}.`,
            tujuanPerbaikan: `Meningkatkan Nilai Capaian ${ind.nama} dari ${formatScore(ind.nilai)} menjadi ${targetScore} melalui penguatan kompetensi pendidik dan layanan belajar.`,
            sasaranProgram: "Seluruh pendidik, tenaga kependidikan, dan peserta didik satuan pendidikan.",
            bentukKegiatan: `Penguatan Komunitas Belajar (Kombel), Lokakarya Praktik Baik, dan Penyediaan Perangkat Pembelajaran untuk ${ind.nama}.`,
            langkahPelaksanaan: "1. Pemetaan kebutuhan awal; 2. Pelaksanaan lokakarya/IHT di Kombel; 3. Implementasi di kelas; 4. Supervisi dan refleksi berkala.",
            penanggungJawab: "Kepala Sekolah & Koordinator Komunitas Belajar",
            indikatorKeberhasilan: `Terlaksananya kegiatan penguatan dan meningkatnya skor capaian ${ind.nama} pada Rapor Pendidikan.`,
            targetCapaian: `Nilai Capaian meningkat menuju ${targetScore}.`,
            waktuPelaksanaan: `Semester Ganjil – Genap ${data.school.tahunPerencanaan}`,
            kebutuhanSumberDaya: "Narasumber/fasilitator internal, perangkat ajar, ATK kegiatan Kombel.",
            perkiraanAnggaran: 3500000,
            disetujui: false
          });
        }
      });

      triggerAutoSave("Rekomendasi", "Susun Rekomendasi", "Menyusun rekomendasi program berdasarkan prioritas Rapor Pendidikan.");
      setActiveTab("rekomendasi");
      showToast("Rekomendasi program berhasil disusun. Silakan tinjau, sunting, dan setujui.", "success");
    }

    function renderRekomendasiModule() {
      const data = getCurrentData();

      return `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 5 • Desain Intervensi Sekolah</span>
            <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Rekomendasi Program Pembenahan</h2>
            <p class="text-xs text-slate-600 mt-0.5">Setiap rekomendasi terhubung dengan <strong>Nilai Capaian</strong> indikator dan dapat disunting serta disetujui tim sekolah.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button onclick="generateRecommendationsFromPriorities()" class="px-3.5 py-2 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-slate-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="refresh-cw" class="w-3.5 h-3.5 text-emerald-700"></i>
              <span>Generate Rekomendasi</span>
            </button>
            <button onclick="generateRkjmFromRecommendations()" class="px-4 py-2 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center gap-1.5 shadow">
              <span>Susun RKJM Otomatis</span>
              <i data-lucide="arrow-right" class="w-4 h-4"></i>
            </button>
          </div>
        </div>

        ${data.recommendations.length === 0 ? `
          <div class="card-elevated p-10 text-center space-y-3">
            <p class="text-xs text-slate-500">Belum ada rekomendasi program. Klik tombol <strong>Generate Rekomendasi</strong> dari hasil prioritas pembenahan.</p>
            <button onclick="generateRecommendationsFromPriorities()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Buat Rekomendasi Sekarang</button>
          </div>
        ` : `
          <div class="space-y-4">
            ${data.recommendations.map(rec => {
              const ind = data.indicators.find(i => i.id === rec.indicatorId) || { nama: "Indikator", kode: "", nilai: "", kategori: "" };
              return `
                <article class="card-elevated p-5 space-y-4">
                  <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-3 border-b border-slate-200">
                    <div class="flex items-center gap-2 flex-wrap">
                      <span class="px-2.5 py-0.5 rounded-lg bg-emerald-900 text-amber-300 font-bold text-xs">
                        ${ind.kode ? `[${escapeHtml(ind.kode)}] ` : ""}${escapeHtml(ind.nama)}
                      </span>
                      <span class="px-2.5 py-0.5 rounded-full bg-slate-100 border border-slate-200 font-mono text-xs font-bold text-slate-800">
                        Nilai Capaian: ${formatScore(ind.nilai)}
                      </span>
                    </div>
                    <div class="flex items-center gap-2">
                      <button onclick="toggleApproveRecommendation('${rec.id}')"
                        class="px-3 py-1.5 rounded-xl text-xs font-bold flex items-center gap-1.5 border ${rec.disetujui ? 'bg-emerald-100 text-emerald-900 border-emerald-300' : 'bg-amber-50 text-amber-900 border-amber-300'}">
                        <i data-lucide="${rec.disetujui ? 'check-circle-2' : 'clock'}" class="w-3.5 h-3.5"></i>
                        <span>${rec.disetujui ? "Disetujui Tim Sekolah" : "Klik Setujui Program"}</span>
                      </button>
                      <button onclick="openEditRecommendationModal('${rec.id}')" class="p-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-700" title="Edit Rekomendasi">
                        <i data-lucide="pencil" class="w-3.5 h-3.5"></i>
                      </button>
                    </div>
                  </div>

                  <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 text-xs">
                    <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                      <div class="font-bold text-slate-700 mb-1">Permasalahan</div>
                      <p class="text-slate-800">${escapeHtml(rec.permasalahan)}</p>
                    </div>
                    <div class="p-3 rounded-xl bg-emerald-50/60 border border-emerald-200">
                      <div class="font-bold text-emerald-950 mb-1">Tujuan Perbaikan</div>
                      <p class="text-slate-800">${escapeHtml(rec.tujuanPerbaikan)}</p>
                    </div>
                    <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                      <div class="font-bold text-slate-700 mb-1">Bentuk Kegiatan</div>
                      <p class="text-slate-800 font-semibold">${escapeHtml(rec.bentukKegiatan)}</p>
                    </div>
                    <div class="p-3 rounded-xl bg-white border border-slate-200">
                      <div class="font-bold text-slate-700 mb-1">Langkah Pelaksanaan</div>
                      <p class="text-slate-700">${escapeHtml(rec.langkahPelaksanaan)}</p>
                    </div>
                    <div class="p-3 rounded-xl bg-white border border-slate-200">
                      <div class="font-bold text-slate-700 mb-1">Sasaran &amp; Penanggung Jawab</div>
                      <p class="text-slate-700"><strong>Sasaran:</strong> ${escapeHtml(rec.sasaranProgram)}<br/><strong>PJ:</strong> ${escapeHtml(rec.penanggungJawab)}</p>
                    </div>
                    <div class="p-3 rounded-xl bg-amber-50/70 border border-amber-200">
                      <div class="font-bold text-amber-950 mb-1">Target, Waktu &amp; Estimasi Biaya</div>
                      <p class="text-slate-800">
                        <strong>Target:</strong> ${escapeHtml(rec.targetCapaian)}<br/>
                        <strong>Waktu:</strong> ${escapeHtml(rec.waktuPelaksanaan)}<br/>
                        <strong>Perkiraan Anggaran:</strong> <span class="font-mono font-bold text-emerald-900">${formatRupiah(rec.perkiraanAnggaran)}</span>
                      </p>
                    </div>
                  </div>
                </article>
              `;
            }).join("")}
          </div>
        `}
      `;
    }

    function toggleApproveRecommendation(recId) {
      const data = getCurrentData();
      const rec = data.recommendations.find(r => r.id === recId);
      if (!rec) return;
      rec.disetujui = !rec.disetujui;
      triggerAutoSave("Rekomendasi", "Persetujuan Program", `${rec.disetujui ? "Menyetujui" : "Membatalkan persetujuan"} program rekomendasi.`);
      renderAll();
    }

    function openEditRecommendationModal(recId) {
      const data = getCurrentData();
      const rec = data.recommendations.find(r => r.id === recId);
      if (!rec) return;

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Edit Rekomendasi Program</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRecommendationEdit(event, '${rec.id}')" class="mt-4 space-y-3 text-xs">
            <div><label class="block font-bold mb-1">Bentuk Kegiatan</label><input type="text" name="bentukKegiatan" value="${escapeHtml(rec.bentukKegiatan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            <div><label class="block font-bold mb-1">Tujuan Perbaikan</label><textarea name="tujuanPerbaikan" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rec.tujuanPerbaikan)}</textarea></div>
            <div><label class="block font-bold mb-1">Langkah Pelaksanaan</label><textarea name="langkahPelaksanaan" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(rec.langkahPelaksanaan)}</textarea></div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div><label class="block font-bold mb-1">Sasaran Program</label><input type="text" name="sasaranProgram" value="${escapeHtml(rec.sasaranProgram)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Penanggung Jawab</label><input type="text" name="penanggungJawab" value="${escapeHtml(rec.penanggungJawab)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div><label class="block font-bold mb-1">Target Capaian</label><input type="text" name="targetCapaian" value="${escapeHtml(rec.targetCapaian)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Waktu Pelaksanaan</label><input type="text" name="waktuPelaksanaan" value="${escapeHtml(rec.waktuPelaksanaan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Perkiraan Anggaran (Rp)</label><input type="number" name="perkiraanAnggaran" value="${Number(rec.perkiraanAnggaran) || 0}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" /></div>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Rekomendasi</button>
            </div>
          </form>
        </div>
      `);
    }

    function saveRecommendationEdit(e, recId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const rec = data.recommendations.find(r => r.id === recId);
      if (!rec) return;

      rec.bentukKegiatan = String(fd.get("bentukKegiatan") || "").trim();
      rec.tujuanPerbaikan = String(fd.get("tujuanPerbaikan") || "").trim();
      rec.langkahPelaksanaan = String(fd.get("langkahPelaksanaan") || "").trim();
      rec.sasaranProgram = String(fd.get("sasaranProgram") || "").trim();
      rec.penanggungJawab = String(fd.get("penanggungJawab") || "").trim();
      rec.targetCapaian = String(fd.get("targetCapaian") || "").trim();
      rec.waktuPelaksanaan = String(fd.get("waktuPelaksanaan") || "").trim();
      rec.perkiraanAnggaran = Number(fd.get("perkiraanAnggaran")) || 0;

      closeCustomModal();
      triggerAutoSave("Rekomendasi", "Edit Rekomendasi", "Menyunting detail rekomendasi program.");
      renderAll();
      showToast("Rekomendasi program diperbarui.", "success");
    }

    // =========================================================================
    // 14. MODULE 6: PENYUSUNAN RKJM (4-Year Strategic Plan)
    // =========================================================================
    function generateRkjmFromRecommendations() {
      const data = getCurrentData();
      if (data.recommendations.length === 0) {
        generateRecommendationsFromPriorities();
      }

      data.recommendations.forEach(rec => {
        if (!data.rkjm.some(p => p.indicatorId === rec.indicatorId)) {
          const ind = data.indicators.find(i => i.id === rec.indicatorId);
          const rc = data.rootCauses.find(r => r.indicatorId === rec.indicatorId);
          if (!ind) return;

          const baseScore = parseFloat(ind.nilai);
          const t1 = !isNaN(baseScore) ? `Skor ${Math.min(100, baseScore + 5).toFixed(1)}` : "Meningkat Th-1";
          const t2 = !isNaN(baseScore) ? `Skor ${Math.min(100, baseScore + 10).toFixed(1)}` : "Meningkat Th-2";
          const t3 = !isNaN(baseScore) ? `Skor ${Math.min(100, baseScore + 15).toFixed(1)}` : "Kategori Baik";
          const t4 = !isNaN(baseScore) ? `Skor ${Math.min(100, baseScore + 20).toFixed(1)}` : "Optimal";

          data.rkjm.push({
            id: uid(),
            indicatorId: ind.id,
            dimensi: ind.dimensi || "Mutu Pendidikan",
            indikatorNama: `${ind.kode ? "[" + ind.kode + "] " : ""}${ind.nama}`,
            akarMasalah: rc ? rc.masalahUtama : rec.permasalahan,
            tujuanStrategis: rec.tujuanPerbaikan,
            programPengembangan: `Program Peningkatan ${ind.nama}`,
            targetTahun1: t1,
            targetTahun2: t2,
            targetTahun3: t3,
            targetTahun4: t4,
            indikatorKeberhasilan: rec.indikatorKeberhasilan,
            penanggungJawab: rec.penanggungJawab,
            sumberPendanaan: "BOSP Reguler"
          });
        }
      });

      triggerAutoSave("RKJM", "Susun RKJM Otomatis", "Menurunkan program prioritas ke dokumen RKJM 4 tahunan.");
      setActiveTab("rkjm");
      showToast("Dokumen RKJM 4 tahunan berhasil disusun!", "success");
    }

    function renderRkjmModule() {
      const data = getCurrentData();
      const yStart = parseInt(data.school.periodeRkjmAwal || data.school.tahunPerencanaan || "2026", 10);
      const yEnd = parseInt(data.school.periodeRkjmAkhir || String(yStart + 3), 10);

      return `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 6 • Perencanaan Strategis 4 Tahunan</span>
            <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Penyusunan Rencana Kerja Jangka Menengah (RKJM)</h2>
            <p class="text-xs text-slate-600 mt-0.5">Hubungkan program prioritas Rapor Pendidikan dengan target jangka menengah dan sumber pendanaan sekolah.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button onclick="exportWordDocument('rkjm')" class="px-3 py-2 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-blue-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="file-text" class="w-4 h-4"></i>
              <span>Unduh Word</span>
            </button>
            <button onclick="exportExcelDocument('rkjm')" class="px-3 py-2 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-emerald-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
              <span>Unduh Excel</span>
            </button>
            <button onclick="openAddRkjmModal()" class="px-3.5 py-2 rounded-xl bg-white border border-slate-300 text-slate-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="plus" class="w-4 h-4 text-emerald-700"></i>
              <span>Tambah Program RKJM</span>
            </button>
            <button onclick="generateRktFromRkjm()" class="px-4 py-2 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center gap-1.5 shadow">
              <span>Turunkan ke RKT</span>
              <i data-lucide="arrow-right" class="w-4 h-4"></i>
            </button>
          </div>
        </div>

        <!-- PERIOD CONFIGURATOR -->
        <section class="card-elevated p-4 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div class="flex flex-wrap items-center gap-3">
            <span class="font-bold text-slate-800">Periode Perencanaan RKJM:</span>
            <input type="number" id="rkjmStartInput" value="${yStart}" class="w-24 rounded-lg border border-slate-300 px-2.5 py-1.5 font-mono font-bold" />
            <span>s.d.</span>
            <input type="number" id="rkjmEndInput" value="${yEnd}" class="w-24 rounded-lg border border-slate-300 px-2.5 py-1.5 font-mono font-bold" />
            <button onclick="updateRkjmPeriod()" class="px-3 py-1.5 rounded-lg bg-emerald-900 text-white font-bold">Terapkan Periode</button>
          </div>
          <span class="text-[11px] text-slate-500">Total Program Strategis: <strong>${data.rkjm.length}</strong></span>
        </section>

        <!-- RKJM TABLE -->
        <section class="card-elevated overflow-hidden">
          ${data.rkjm.length === 0 ? `
            <div class="p-10 text-center space-y-3">
              <p class="text-xs text-slate-500">Belum ada program RKJM. Klik tombol di bawah untuk menyusun RKJM dari hasil rekomendasi program.</p>
              <button onclick="generateRkjmFromRecommendations()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Susun RKJM Otomatis</button>
            </div>
          ` : `
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs">
                <thead class="bg-slate-50 border-b border-slate-200 text-slate-600 uppercase text-[10px]">
                  <tr>
                    <th class="py-3 px-3">Indikator &amp; Capaian Awal</th>
                    <th class="py-3 px-3">Akar Masalah</th>
                    <th class="py-3 px-3">Program Pengembangan &amp; Tujuan</th>
                    <th class="py-3 px-2 text-center">Th 1 (${yStart})</th>
                    <th class="py-3 px-2 text-center">Th 2 (${yStart + 1})</th>
                    <th class="py-3 px-2 text-center">Th 3 (${yStart + 2})</th>
                    <th class="py-3 px-2 text-center">Th 4 (${yEnd})</th>
                    <th class="py-3 px-3">PJ &amp; Sumber Dana</th>
                    <th class="py-3 px-3 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                  ${data.rkjm.map(p => {
                    const linkedInd = data.indicators.find(i => i.id === p.indicatorId);
                    return `
                      <tr class="hover:bg-slate-50 align-top">
                        <td class="py-3 px-3 min-w-[180px]">
                          <span class="inline-block px-2 py-0.5 rounded bg-emerald-50 text-emerald-900 font-bold text-[11px] border border-emerald-200">${escapeHtml(p.indikatorNama)}</span>
                          ${linkedInd && linkedInd.nilai !== "" ? `<div class="text-[11px] font-mono font-bold text-slate-800 mt-1">Capaian Awal: ${formatScore(linkedInd.nilai)} (${classifyCategory(linkedInd.nilai, linkedInd.kategori)})</div>` : ""}
                          <div class="text-[10px] text-slate-500 mt-1">${escapeHtml(p.dimensi)}</div>
                        </td>
                        <td class="py-3 px-3 min-w-[170px] text-slate-700">${escapeHtml(p.akarMasalah)}</td>
                        <td class="py-3 px-3 min-w-[200px]">
                          <div class="font-bold text-slate-900">${escapeHtml(p.programPengembangan)}</div>
                          <div class="text-[11px] text-slate-600 mt-1">${escapeHtml(p.tujuanStrategis)}</div>
                        </td>
                        <td class="py-3 px-2 text-center font-mono font-semibold text-emerald-900">${escapeHtml(p.targetTahun1)}</td>
                        <td class="py-3 px-2 text-center font-mono font-semibold text-emerald-900">${escapeHtml(p.targetTahun2)}</td>
                        <td class="py-3 px-2 text-center font-mono font-semibold text-emerald-900">${escapeHtml(p.targetTahun3)}</td>
                        <td class="py-3 px-2 text-center font-mono font-semibold text-emerald-900">${escapeHtml(p.targetTahun4)}</td>
                        <td class="py-3 px-3 min-w-[130px]">
                          <div class="font-semibold text-slate-800">${escapeHtml(p.penanggungJawab)}</div>
                          <div class="text-[11px] text-amber-800 font-semibold mt-0.5">${escapeHtml(p.sumberPendanaan)}</div>
                        </td>
                        <td class="py-3 px-3 text-right whitespace-nowrap">
                          <button onclick="openEditRkjmModal('${p.id}')" class="p-1.5 rounded-lg hover:bg-emerald-50 text-emerald-800" title="Edit RKJM"><i data-lucide="pencil" class="w-3.5 h-3.5"></i></button>
                          <button onclick="deleteRkjm('${p.id}')" class="p-1.5 rounded-lg hover:bg-rose-50 text-rose-700" title="Hapus"><i data-lucide="trash-2" class="w-3.5 h-3.5"></i></button>
                        </td>
                      </tr>
                    `;
                  }).join("")}
                </tbody>
              </table>
            </div>
          `}
        </section>
      `;
    }

    function updateRkjmPeriod() {
      const s = document.getElementById("rkjmStartInput")?.value;
      const e = document.getElementById("rkjmEndInput")?.value;
      const data = getCurrentData();
      if (s) data.school.periodeRkjmAwal = s;
      if (e) data.school.periodeRkjmAkhir = e;
      triggerAutoSave("RKJM", "Ubah Periode RKJM", `Mengatur periode RKJM menjadi ${s} - ${e}.`);
      renderAll();
      showToast("Periode RKJM berhasil diperbarui.", "success");
    }

    function openAddRkjmModal() {
      const data = getCurrentData();
      openCustomModal(`
        <div class="p-5 sm:p-6">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Tambah Program RKJM</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRkjmModalForm(event, null)" class="mt-4 space-y-3 text-xs">
            <div>
              <label class="block font-bold mb-1">Indikator Rapor Terkait</label>
              <select name="indicatorId" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${data.indicators.map(i => `<option value="${i.id}">${i.kode ? "[" + i.kode + "] " : ""}${escapeHtml(i.nama)} (Nilai: ${formatScore(i.nilai)})</option>`).join("")}
              </select>
            </div>
            <div><label class="block font-bold mb-1">Nama Program Pengembangan Sekolah *</label><input type="text" name="programPengembangan" required class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            <div><label class="block font-bold mb-1">Tujuan Strategis</label><textarea name="tujuanStrategis" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2"></textarea></div>
            <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
              <div><label class="block font-bold mb-1">Target Th 1</label><input type="text" name="targetTahun1" placeholder="Skor ..." class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5" /></div>
              <div><label class="block font-bold mb-1">Target Th 2</label><input type="text" name="targetTahun2" placeholder="Skor ..." class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5" /></div>
              <div><label class="block font-bold mb-1">Target Th 3</label><input type="text" name="targetTahun3" placeholder="Skor ..." class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5" /></div>
              <div><label class="block font-bold mb-1">Target Th 4</label><input type="text" name="targetTahun4" placeholder="Skor ..." class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div><label class="block font-bold mb-1">Penanggung Jawab</label><input type="text" name="penanggungJawab" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Sumber Pendanaan</label><input type="text" name="sumberPendanaan" value="BOSP Reguler" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan RKJM</button>
            </div>
          </form>
        </div>
      `);
    }

    function openEditRkjmModal(rkjmId) {
      const data = getCurrentData();
      const p = data.rkjm.find(r => r.id === rkjmId);
      if (!p) return;

      openCustomModal(`
        <div class="p-5 sm:p-6">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Edit Program RKJM</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRkjmModalForm(event, '${p.id}')" class="mt-4 space-y-3 text-xs">
            <input type="hidden" name="indicatorId" value="${escapeHtml(p.indicatorId)}" />
            <div><label class="block font-bold mb-1">Nama Program Pengembangan Sekolah *</label><input type="text" name="programPengembangan" required value="${escapeHtml(p.programPengembangan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            <div><label class="block font-bold mb-1">Tujuan Strategis</label><textarea name="tujuanStrategis" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(p.tujuanStrategis)}</textarea></div>
            <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
              <div><label class="block font-bold mb-1">Target Th 1</label><input type="text" name="targetTahun1" value="${escapeHtml(p.targetTahun1)}" class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5" /></div>
              <div><label class="block font-bold mb-1">Target Th 2</label><input type="text" name="targetTahun2" value="${escapeHtml(p.targetTahun2)}" class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5" /></div>
              <div><label class="block font-bold mb-1">Target Th 3</label><input type="text" name="targetTahun3" value="${escapeHtml(p.targetTahun3)}" class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5" /></div>
              <div><label class="block font-bold mb-1">Target Th 4</label><input type="text" name="targetTahun4" value="${escapeHtml(p.targetTahun4)}" class="w-full rounded-lg border border-slate-300 px-2.5 py-1.5" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div><label class="block font-bold mb-1">Penanggung Jawab</label><input type="text" name="penanggungJawab" value="${escapeHtml(p.penanggungJawab)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Sumber Pendanaan</label><input type="text" name="sumberPendanaan" value="${escapeHtml(p.sumberPendanaan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Perubahan</button>
            </div>
          </form>
        </div>
      `);
    }

    function saveRkjmModalForm(e, rkjmId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const indId = String(fd.get("indicatorId") || "");
      const ind = data.indicators.find(i => i.id === indId);

      const obj = {
        id: rkjmId || uid(),
        indicatorId: indId,
        dimensi: ind ? ind.dimensi : "Mutu Pendidikan",
        indikatorNama: ind ? `${ind.kode ? "[" + ind.kode + "] " : ""}${ind.nama}` : "Program Sekolah",
        akarMasalah: data.rootCauses.find(r => r.indicatorId === indId)?.masalahUtama || "-",
        tujuanStrategis: String(fd.get("tujuanStrategis") || "").trim(),
        programPengembangan: String(fd.get("programPengembangan") || "").trim(),
        targetTahun1: String(fd.get("targetTahun1") || "").trim(),
        targetTahun2: String(fd.get("targetTahun2") || "").trim(),
        targetTahun3: String(fd.get("targetTahun3") || "").trim(),
        targetTahun4: String(fd.get("targetTahun4") || "").trim(),
        indikatorKeberhasilan: "Tercapainya target peningkatan Rapor Pendidikan",
        penanggungJawab: String(fd.get("penanggungJawab") || "").trim(),
        sumberPendanaan: String(fd.get("sumberPendanaan") || "BOSP Reguler").trim()
      };

      if (rkjmId) {
        const idx = data.rkjm.findIndex(r => r.id === rkjmId);
        if (idx >= 0) data.rkjm[idx] = { ...data.rkjm[idx], ...obj };
      } else {
        data.rkjm.push(obj);
      }

      closeCustomModal();
      triggerAutoSave("RKJM", rkjmId ? "Edit RKJM" : "Tambah RKJM", `Menyimpan program "${obj.programPengembangan}".`);
      renderAll();
      showToast("Program RKJM berhasil disimpan.", "success");
    }

    function deleteRkjm(rkjmId) {
      const data = getCurrentData();
      data.rkjm = data.rkjm.filter(r => r.id !== rkjmId);
      triggerAutoSave("RKJM", "Hapus RKJM", "Menghapus program dari tabel RKJM.");
      renderAll();
      showToast("Program RKJM dihapus.", "info");
    }

    // =========================================================================
    // 15. MODULE 7: PENYUSUNAN RKT & KALENDER KEGIATAN TAHUNAN
    // =========================================================================
    function generateRktFromRkjm() {
      const data = getCurrentData();
      if (data.rkjm.length === 0) {
        generateRkjmFromRecommendations();
      }

      data.rkjm.forEach(prog => {
        if (!data.rkt.some(k => k.rkjmId === prog.id)) {
          const rec = data.recommendations.find(r => r.indicatorId === prog.indicatorId);
          data.rkt.push({
            id: uid(),
            rkjmId: prog.id,
            indicatorId: prog.indicatorId,
            namaProgram: prog.programPengembangan,
            namaKegiatan: rec ? rec.bentukKegiatan : `Pelaksanaan ${prog.programPengembangan}`,
            latarBelakang: prog.akarMasalah,
            tujuan: prog.tujuanStrategis,
            sasaran: rec ? rec.sasaranProgram : "Pendidik dan Peserta Didik",
            indikatorKeberhasilan: prog.indikatorKeberhasilan,
            target: prog.targetTahun1 || "100% Terlaksana",
            bulanPelaksanaan: [7, 8, 9],
            jadwalTeks: rec ? rec.waktuPelaksanaan : `Juli – September ${data.school.tahunPerencanaan}`,
            penanggungJawab: prog.penanggungJawab,
            kebutuhanAnggaran: rec ? Number(rec.perkiraanAnggaran) || 2500000 : 2500000,
            sumberDana: prog.sumberPendanaan || "BOSP Reguler",
            buktiPelaksanaan: "Daftar hadir, notulen, dokumentasi foto, dan laporan kegiatan."
          });
        }
      });

      triggerAutoSave("RKT", "Turunkan RKJM ke RKT", "Menyusun kegiatan RKT tahunan dari program RKJM.");
      setActiveTab("rkt");
      showToast("Rencana Kerja Tahunan (RKT) berhasil diturunkan dari RKJM!", "success");
    }

    function renderRktModule() {
      const data = getCurrentData();
      const months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];

      return `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 7 • Operasional Tahunan</span>
            <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Rencana Kerja Tahunan (RKT ${escapeHtml(data.school.tahunPerencanaan)})</h2>
            <p class="text-xs text-slate-600 mt-0.5">Penjabaran program RKJM ke dalam kegiatan tahunan, sasaran, target, serta kalender pelaksanaan 12 bulan.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button onclick="exportWordDocument('rkt')" class="px-3 py-2 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-blue-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="file-text" class="w-4 h-4"></i>
              <span>Unduh Word</span>
            </button>
            <button onclick="exportExcelDocument('rkt')" class="px-3 py-2 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-emerald-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
              <span>Unduh Excel</span>
            </button>
            <button onclick="openAddRktModal()" class="px-3.5 py-2 rounded-xl bg-white border border-slate-300 text-slate-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="plus" class="w-4 h-4 text-emerald-700"></i>
              <span>Tambah Kegiatan RKT</span>
            </button>
            <button onclick="generateArkasFromRkt()" class="px-4 py-2 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center gap-1.5 shadow">
              <span>Buat Rancangan ARKAS</span>
              <i data-lucide="arrow-right" class="w-4 h-4"></i>
            </button>
          </div>
        </div>

        <!-- RKT ACTIVITIES TABLE -->
        <section class="card-elevated overflow-hidden">
          ${data.rkt.length === 0 ? `
            <div class="p-10 text-center space-y-3">
              <p class="text-xs text-slate-500">Belum ada kegiatan RKT. Klik tombol di bawah untuk menurunkan kegiatan dari program RKJM.</p>
              <button onclick="generateRktFromRkjm()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Susun RKT dari RKJM</button>
            </div>
          ` : `
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs">
                <thead class="bg-slate-50 border-b border-slate-200 text-slate-600 uppercase text-[10px]">
                  <tr>
                    <th class="py-3 px-3">Program RKJM &amp; Kegiatan RKT</th>
                    <th class="py-3 px-3">Tujuan &amp; Sasaran</th>
                    <th class="py-3 px-3">Indikator &amp; Target</th>
                    <th class="py-3 px-3">Jadwal &amp; PJ</th>
                    <th class="py-3 px-3 text-right">Kebutuhan Anggaran</th>
                    <th class="py-3 px-3">Bukti Pelaksanaan</th>
                    <th class="py-3 px-3 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                  ${data.rkt.map(k => {
                    const ind = data.indicators.find(i => i.id === k.indicatorId);
                    return `
                      <tr class="hover:bg-slate-50 align-top">
                        <td class="py-3 px-3 min-w-[200px]">
                          <div class="text-[10px] font-bold text-emerald-800 uppercase">${escapeHtml(k.namaProgram)}</div>
                          <div class="font-bold text-slate-900 mt-0.5">${escapeHtml(k.namaKegiatan)}</div>
                          ${ind ? `<div class="text-[10px] font-mono text-slate-500 mt-1">Indikator: ${escapeHtml(ind.nama)} (Nilai: ${formatScore(ind.nilai)})</div>` : ""}
                        </td>
                        <td class="py-3 px-3 min-w-[180px]">
                          <div class="text-slate-700">${escapeHtml(k.tujuan)}</div>
                          <div class="text-[11px] font-semibold text-slate-900 mt-1">Sasaran: ${escapeHtml(k.sasaran)}</div>
                        </td>
                        <td class="py-3 px-3 min-w-[160px]">
                          <div class="text-slate-700">${escapeHtml(k.indikatorKeberhasilan)}</div>
                          <div class="text-[11px] font-bold text-emerald-800 mt-1">Target: ${escapeHtml(k.target)}</div>
                        </td>
                        <td class="py-3 px-3 min-w-[140px]">
                          <div class="font-semibold text-slate-900">${escapeHtml(k.jadwalTeks)}</div>
                          <div class="text-[11px] text-slate-500 mt-0.5">PJ: ${escapeHtml(k.penanggungJawab)}</div>
                        </td>
                        <td class="py-3 px-3 text-right whitespace-nowrap">
                          <div class="font-mono font-extrabold text-slate-900">${formatRupiah(k.kebutuhanAnggaran)}</div>
                          <div class="text-[10px] font-semibold text-amber-800">${escapeHtml(k.sumberDana)}</div>
                        </td>
                        <td class="py-3 px-3 min-w-[150px] text-slate-600">${escapeHtml(k.buktiPelaksanaan)}</td>
                        <td class="py-3 px-3 text-right whitespace-nowrap">
                          <button onclick="openEditRktModal('${k.id}')" class="p-1.5 rounded-lg hover:bg-emerald-50 text-emerald-800"><i data-lucide="pencil" class="w-3.5 h-3.5"></i></button>
                          <button onclick="deleteRkt('${k.id}')" class="p-1.5 rounded-lg hover:bg-rose-50 text-rose-700"><i data-lucide="trash-2" class="w-3.5 h-3.5"></i></button>
                        </td>
                      </tr>
                    `;
                  }).join("")}
                </tbody>
              </table>
            </div>
          `}
        </section>

        <!-- ANNUAL 12-MONTH ACTIVITY CALENDAR -->
        ${data.rkt.length > 0 ? `
          <section class="card-elevated p-5 overflow-hidden">
            <div class="pb-3 border-b border-slate-100">
              <h3 class="font-bold text-slate-900 text-sm">Kalender Pelaksanaan Kegiatan Tahunan (${escapeHtml(data.school.tahunPerencanaan)})</h3>
              <p class="text-xs text-slate-500">Klik kotak bulan untuk menandai atau mengubah jadwal pelaksanaan kegiatan.</p>
            </div>
            <div class="mt-3 overflow-x-auto">
              <table class="w-full text-left text-xs border-collapse">
                <thead>
                  <tr class="bg-emerald-950 text-white text-[10px]">
                    <th class="py-2.5 px-3">Nama Kegiatan RKT</th>
                    ${months.map(m => `<th class="py-2.5 px-1.5 text-center w-10">${m}</th>`).join("")}
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-200">
                  ${data.rkt.map(k => `
                    <tr>
                      <td class="py-2.5 px-3 font-semibold text-slate-800 min-w-[220px]">${escapeHtml(k.namaKegiatan)}</td>
                      ${months.map((_, mIdx) => {
                        const mNum = mIdx + 1;
                        const active = (k.bulanPelaksanaan || []).includes(mNum);
                        return `
                          <td class="p-1 text-center border-l border-slate-100">
                            <button onclick="toggleRktMonth('${k.id}', ${mNum})"
                              class="w-7 h-7 rounded-lg font-bold text-[10px] transition ${active ? 'emerald-gradient text-amber-300 shadow-xs' : 'bg-slate-100 text-slate-400 hover:bg-slate-200'}">
                              ${active ? "✓" : ""}
                            </button>
                          </td>
                        `;
                      }).join("")}
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </section>
        ` : ""}
      `;
    }

    function toggleRktMonth(rktId, monthNum) {
      const data = getCurrentData();
      const k = data.rkt.find(item => item.id === rktId);
      if (!k) return;
      if (!Array.isArray(k.bulanPelaksanaan)) k.bulanPelaksanaan = [];
      if (k.bulanPelaksanaan.includes(monthNum)) {
        k.bulanPelaksanaan = k.bulanPelaksanaan.filter(m => m !== monthNum);
      } else {
        k.bulanPelaksanaan.push(monthNum);
        k.bulanPelaksanaan.sort((a, b) => a - b);
      }
      triggerAutoSave("RKT", "Ubah Jadwal Kalender", "Memperbarui bulan pelaksanaan kegiatan RKT.");
      renderAll();
    }

    function openAddRktModal() {
      const data = getCurrentData();
      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Tambah Kegiatan RKT</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRktModalForm(event, null)" class="mt-4 space-y-3 text-xs">
            <div>
              <label class="block font-bold mb-1">Program RKJM Induk</label>
              <select name="rkjmId" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                <option value="">-- Pilih Program RKJM --</option>
                ${data.rkjm.map(p => `<option value="${p.id}">${escapeHtml(p.programPengembangan)} (${escapeHtml(p.indikatorNama)})</option>`).join("")}
              </select>
            </div>
            <div><label class="block font-bold mb-1">Nama Kegiatan Tahunan *</label><input type="text" name="namaKegiatan" required class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            <div><label class="block font-bold mb-1">Latar Belakang</label><textarea name="latarBelakang" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2"></textarea></div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div><label class="block font-bold mb-1">Tujuan Kegiatan</label><input type="text" name="tujuan" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Sasaran</label><input type="text" name="sasaran" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div><label class="block font-bold mb-1">Indikator Keberhasilan</label><input type="text" name="indikatorKeberhasilan" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Target Capaian</label><input type="text" name="target" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div><label class="block font-bold mb-1">Jadwal Pelaksanaan</label><input type="text" name="jadwalTeks" placeholder="Juli - Sept 2026" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Penanggung Jawab</label><input type="text" name="penanggungJawab" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Kebutuhan Anggaran (Rp)</label><input type="number" name="kebutuhanAnggaran" value="0" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold mb-1">Sumber Dana</label>
                <select name="sumberDana" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  ${SUMBER_DANA_STANDAR.map(sd => `<option value="${sd}">${sd}</option>`).join("")}
                </select>
              </div>
              <div><label class="block font-bold mb-1">Bukti Pelaksanaan</label><input type="text" name="buktiPelaksanaan" value="Daftar hadir, notulen, foto kegiatan" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Kegiatan</button>
            </div>
          </form>
        </div>
      `);
    }

    function openEditRktModal(rktId) {
      const data = getCurrentData();
      const k = data.rkt.find(item => item.id === rktId);
      if (!k) return;

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Edit Kegiatan RKT</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveRktModalForm(event, '${k.id}')" class="mt-4 space-y-3 text-xs">
            <input type="hidden" name="rkjmId" value="${escapeHtml(k.rkjmId)}" />
            <div><label class="block font-bold mb-1">Nama Kegiatan Tahunan *</label><input type="text" name="namaKegiatan" required value="${escapeHtml(k.namaKegiatan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            <div><label class="block font-bold mb-1">Latar Belakang</label><textarea name="latarBelakang" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(k.latarBelakang)}</textarea></div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div><label class="block font-bold mb-1">Tujuan Kegiatan</label><input type="text" name="tujuan" value="${escapeHtml(k.tujuan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Sasaran</label><input type="text" name="sasaran" value="${escapeHtml(k.sasaran)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div><label class="block font-bold mb-1">Indikator Keberhasilan</label><input type="text" name="indikatorKeberhasilan" value="${escapeHtml(k.indikatorKeberhasilan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Target Capaian</label><input type="text" name="target" value="${escapeHtml(k.target)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div><label class="block font-bold mb-1">Jadwal Pelaksanaan</label><input type="text" name="jadwalTeks" value="${escapeHtml(k.jadwalTeks)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Penanggung Jawab</label><input type="text" name="penanggungJawab" value="${escapeHtml(k.penanggungJawab)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Kebutuhan Anggaran (Rp)</label><input type="number" name="kebutuhanAnggaran" value="${Number(k.kebutuhanAnggaran) || 0}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold mb-1">Sumber Dana</label>
                <select name="sumberDana" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  ${SUMBER_DANA_STANDAR.map(sd => `<option value="${sd}" ${k.sumberDana === sd ? "selected" : ""}>${sd}</option>`).join("")}
                </select>
              </div>
              <div><label class="block font-bold mb-1">Bukti Pelaksanaan</label><input type="text" name="buktiPelaksanaan" value="${escapeHtml(k.buktiPelaksanaan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Perubahan</button>
            </div>
          </form>
        </div>
      `);
    }

    function saveRktModalForm(e, rktId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const rkjmId = String(fd.get("rkjmId") || "");
      const prog = data.rkjm.find(p => p.id === rkjmId);

      const existing = rktId ? data.rkt.find(k => k.id === rktId) : null;
      const obj = {
        id: rktId || uid(),
        rkjmId,
        indicatorId: prog ? prog.indicatorId : (existing?.indicatorId || ""),
        namaProgram: prog ? prog.programPengembangan : (existing?.namaProgram || "Program Pengembangan Sekolah"),
        namaKegiatan: String(fd.get("namaKegiatan") || "").trim(),
        latarBelakang: String(fd.get("latarBelakang") || "").trim(),
        tujuan: String(fd.get("tujuan") || "").trim(),
        sasaran: String(fd.get("sasaran") || "").trim(),
        indikatorKeberhasilan: String(fd.get("indikatorKeberhasilan") || "").trim(),
        target: String(fd.get("target") || "").trim(),
        bulanPelaksanaan: existing?.bulanPelaksanaan || [7, 8],
        jadwalTeks: String(fd.get("jadwalTeks") || "").trim(),
        penanggungJawab: String(fd.get("penanggungJawab") || "").trim(),
        kebutuhanAnggaran: Number(fd.get("kebutuhanAnggaran")) || 0,
        sumberDana: String(fd.get("sumberDana") || "BOSP Reguler").trim(),
        buktiPelaksanaan: String(fd.get("buktiPelaksanaan") || "").trim()
      };

      if (rktId) {
        const idx = data.rkt.findIndex(k => k.id === rktId);
        if (idx >= 0) data.rkt[idx] = obj;
      } else {
        data.rkt.push(obj);
      }

      closeCustomModal();
      triggerAutoSave("RKT", rktId ? "Edit Kegiatan RKT" : "Tambah Kegiatan RKT", `Menyimpan kegiatan "${obj.namaKegiatan}".`);
      renderAll();
      showToast("Kegiatan RKT berhasil disimpan.", "success");
    }

    function deleteRkt(rktId) {
      const data = getCurrentData();
      data.rkt = data.rkt.filter(k => k.id !== rktId);
      triggerAutoSave("RKT", "Hapus Kegiatan RKT", "Menghapus kegiatan dari tabel RKT.");
      renderAll();
      showToast("Kegiatan RKT dihapus.", "info");
    }

    // =========================================================================
    // 16. MODULE 8: RANCANGAN ANGGARAN ARKAS (Volume x Harga Satuan + Rekap)
    // =========================================================================
    function generateArkasFromRkt() {
      const data = getCurrentData();
      if (data.rkt.length === 0) {
        generateRktFromRkjm();
      }

      data.rkt.forEach(keg => {
        if (!data.arkas.some(a => a.rktId === keg.id)) {
          data.arkas.push({
            id: uid(),
            rktId: keg.id,
            rkjmId: keg.rkjmId,
            indicatorId: keg.indicatorId,
            program: keg.namaProgram,
            kegiatan: keg.namaKegiatan,
            komponen: "06. Pengembangan Profesi Pendidik dan Tenaga Kependidikan",
            uraianBelanja: `Konsumsi & Bahan Pelaksanaan ${keg.namaKegiatan} (Perlu Verifikasi Rincian)`,
            volume: 20,
            satuan: "Paket",
            hargaSatuan: Math.max(25000, Math.round((Number(keg.kebutuhanAnggaran) || 1000000) / 20)),
            sumberDana: keg.sumberDana || "BOSP Reguler",
            waktuPelaksanaan: keg.jadwalTeks || `Tahun ${data.school.tahunPerencanaan}`,
            statusVerifikasi: "Perlu Verifikasi Komponen & Harga Satuan"
          });
        }
      });

      triggerAutoSave("ARKAS", "Susun Rancangan ARKAS", "Menyusun draf rincian belanja ARKAS dari kegiatan RKT.");
      setActiveTab("arkas");
      showToast("Rancangan anggaran ARKAS berhasil disusun. Silakan verifikasi rincian belanja.", "success");
    }

    function renderArkasModule() {
      const data = getCurrentData();
      const totalAll = data.arkas.reduce((s, a) => s + ((Number(a.volume) || 0) * (Number(a.hargaSatuan) || 0)), 0);

      // Recap Maps
      const byProgram = {};
      const byKegiatan = {};
      const byKomponen = {};

      data.arkas.forEach(a => {
        const sub = (Number(a.volume) || 0) * (Number(a.hargaSatuan) || 0);
        byProgram[a.program] = (byProgram[a.program] || 0) + sub;
        byKegiatan[a.kegiatan] = (byKegiatan[a.kegiatan] || 0) + sub;
        byKomponen[a.komponen] = (byKomponen[a.komponen] || 0) + sub;
      });

      return `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 8 • Rencana Kerja &amp; Anggaran Sekolah</span>
            <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Rancangan Anggaran ARKAS (${escapeHtml(data.school.tahunPerencanaan)})</h2>
            <p class="text-xs text-slate-600 mt-0.5">Terhubung langsung dengan kegiatan RKT. Rumus otomatis: <strong>Jumlah = Volume × Harga Satuan</strong>.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button onclick="exportExcelDocument('arkas')" class="px-3.5 py-2.5 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-emerald-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
              <span>Unduh Excel ARKAS</span>
            </button>
            <button onclick="generateArkasFromRkt()" class="px-3.5 py-2.5 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-slate-800 font-bold text-xs flex items-center gap-1.5">
              <i data-lucide="refresh-cw" class="w-4 h-4 text-emerald-700"></i>
              <span>Susun dari RKT</span>
            </button>
            <button onclick="openAddArkasModal()" class="px-4 py-2.5 rounded-xl emerald-gradient text-white font-bold text-xs flex items-center gap-1.5 shadow">
              <i data-lucide="plus" class="w-4 h-4 text-amber-300"></i>
              <span>Tambah Uraian Belanja</span>
            </button>
          </div>
        </div>

        <!-- REKAPITULASI CARDS -->
        <section class="grid grid-cols-1 lg:grid-cols-3 gap-4 text-xs">
          <div class="card-elevated p-4">
            <h3 class="font-bold text-slate-900 pb-2 border-b border-slate-100">Rekap per Program RKJM</h3>
            <div class="mt-2 space-y-2 max-h-40 overflow-y-auto">
              ${Object.keys(byProgram).length === 0 ? `<p class="text-slate-400">Belum ada anggaran.</p>` : Object.entries(byProgram).map(([k, v]) => `
                <div class="flex justify-between gap-2">
                  <span class="text-slate-700 truncate">${escapeHtml(k)}</span>
                  <span class="font-mono font-bold text-slate-900 shrink-0">${formatRupiah(v)}</span>
                </div>
              `).join("")}
            </div>
          </div>

          <div class="card-elevated p-4">
            <h3 class="font-bold text-slate-900 pb-2 border-b border-slate-100">Rekap per Kegiatan RKT</h3>
            <div class="mt-2 space-y-2 max-h-40 overflow-y-auto">
              ${Object.keys(byKegiatan).length === 0 ? `<p class="text-slate-400">Belum ada anggaran.</p>` : Object.entries(byKegiatan).map(([k, v]) => `
                <div class="flex justify-between gap-2">
                  <span class="text-slate-700 truncate">${escapeHtml(k)}</span>
                  <span class="font-mono font-bold text-slate-900 shrink-0">${formatRupiah(v)}</span>
                </div>
              `).join("")}
            </div>
          </div>

          <div class="card-elevated p-4">
            <h3 class="font-bold text-slate-900 pb-2 border-b border-slate-100">Rekap per Komponen Pembiayaan BOSP</h3>
            <div class="mt-2 space-y-2 max-h-40 overflow-y-auto">
              ${Object.keys(byKomponen).length === 0 ? `<p class="text-slate-400">Belum ada anggaran.</p>` : Object.entries(byKomponen).map(([k, v]) => `
                <div class="flex justify-between gap-2">
                  <span class="text-slate-700 truncate">${escapeHtml(k)}</span>
                  <span class="font-mono font-bold text-emerald-900 shrink-0">${formatRupiah(v)}</span>
                </div>
              `).join("")}
            </div>
          </div>
        </section>

        <!-- ARKAS ITEMS TABLE -->
        <section class="card-elevated overflow-hidden">
          <div class="p-4 bg-emerald-950 text-white flex flex-wrap items-center justify-between gap-2">
            <span class="font-bold text-xs sm:text-sm">Rincian Kertas Kerja Belanja (${data.arkas.length} Item)</span>
            <div class="text-xs sm:text-sm">
              Total Rancangan Anggaran: <span class="font-mono font-extrabold text-amber-300 text-base ml-1">${formatRupiah(totalAll)}</span>
            </div>
          </div>

          ${data.arkas.length === 0 ? `
            <div class="p-10 text-center space-y-3">
              <p class="text-xs text-slate-500">Belum ada rincian belanja ARKAS.</p>
              <button onclick="generateArkasFromRkt()" class="px-4 py-2 rounded-xl emerald-gradient text-white text-xs font-bold">Susun Rancangan dari RKT</button>
            </div>
          ` : `
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs">
                <thead class="bg-slate-50 border-b border-slate-200 text-slate-600 uppercase text-[10px]">
                  <tr>
                    <th class="py-3 px-3">Program &amp; Kegiatan RKT</th>
                    <th class="py-3 px-3">Komponen Pembiayaan &amp; Uraian Belanja</th>
                    <th class="py-3 px-2 text-center">Vol</th>
                    <th class="py-3 px-2 text-center">Satuan</th>
                    <th class="py-3 px-3 text-right">Harga Satuan</th>
                    <th class="py-3 px-3 text-right">Jumlah Anggaran</th>
                    <th class="py-3 px-3">Sumber &amp; Waktu</th>
                    <th class="py-3 px-3 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                  ${data.arkas.map(a => {
                    const subtotal = (Number(a.volume) || 0) * (Number(a.hargaSatuan) || 0);
                    return `
                      <tr class="hover:bg-slate-50 align-top">
                        <td class="py-3 px-3 min-w-[180px]">
                          <div class="text-[10px] font-bold text-emerald-800">${escapeHtml(a.program)}</div>
                          <div class="font-semibold text-slate-900 mt-0.5">${escapeHtml(a.kegiatan)}</div>
                        </td>
                        <td class="py-3 px-3 min-w-[210px]">
                          <div class="text-[10px] font-semibold text-amber-800">${escapeHtml(a.komponen)}</div>
                          <div class="font-bold text-slate-900 mt-0.5">${escapeHtml(a.uraianBelanja)}</div>
                          <div class="text-[10px] text-slate-500 mt-0.5">Status: ${escapeHtml(a.statusVerifikasi || "Perlu Verifikasi")}</div>
                        </td>
                        <td class="py-3 px-2 text-center font-mono font-bold">${Number(a.volume) || 0}</td>
                        <td class="py-3 px-2 text-center">${escapeHtml(a.satuan)}</td>
                        <td class="py-3 px-3 text-right font-mono">${formatRupiah(a.hargaSatuan)}</td>
                        <td class="py-3 px-3 text-right font-mono font-extrabold text-emerald-950 whitespace-nowrap">${formatRupiah(subtotal)}</td>
                        <td class="py-3 px-3 min-w-[130px]">
                          <div class="font-semibold text-slate-800">${escapeHtml(a.sumberDana)}</div>
                          <div class="text-[11px] text-slate-500">${escapeHtml(a.waktuPelaksanaan)}</div>
                        </td>
                        <td class="py-3 px-3 text-right whitespace-nowrap">
                          <button onclick="openEditArkasModal('${a.id}')" class="p-1.5 rounded-lg hover:bg-emerald-50 text-emerald-800"><i data-lucide="pencil" class="w-3.5 h-3.5"></i></button>
                          <button onclick="deleteArkas('${a.id}')" class="p-1.5 rounded-lg hover:bg-rose-50 text-rose-700"><i data-lucide="trash-2" class="w-3.5 h-3.5"></i></button>
                        </td>
                      </tr>
                    `;
                  }).join("")}
                </tbody>
              </table>
            </div>
          `}
        </section>
      `;
    }

    function openAddArkasModal() {
      const data = getCurrentData();
      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Tambah Uraian Belanja ARKAS</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveArkasModalForm(event, null)" class="mt-4 space-y-3 text-xs">
            <div>
              <label class="block font-bold mb-1">Kegiatan RKT Terkait *</label>
              <select name="rktId" required class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${data.rkt.length === 0 ? `<option value="">(Buat kegiatan RKT terlebih dahulu atau isi manual)</option>` : ""}
                ${data.rkt.map(k => `<option value="${k.id}">${escapeHtml(k.namaKegiatan)} — [${escapeHtml(k.namaProgram)}]</option>`).join("")}
              </select>
            </div>
            <div>
              <label class="block font-bold mb-1">Komponen Pembiayaan Standar BOSP</label>
              <select name="komponen" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${KOMPONEN_BOSP_STANDAR.map(c => `<option value="${c}">${c}</option>`).join("")}
              </select>
            </div>
            <div><label class="block font-bold mb-1">Uraian Belanja *</label><input type="text" name="uraianBelanja" required placeholder="Contoh: Pengadaan Buku Pengayaan Literasi Nonteks" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            <div class="grid grid-cols-3 gap-3">
              <div><label class="block font-bold mb-1">Volume *</label><input type="number" step="0.01" name="volume" required value="1" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" /></div>
              <div><label class="block font-bold mb-1">Satuan *</label><input type="text" name="satuan" required placeholder="Paket / Eks / OJ" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Harga Satuan (Rp) *</label><input type="number" name="hargaSatuan" required value="0" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold mb-1">Sumber Dana</label>
                <select name="sumberDana" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  ${SUMBER_DANA_STANDAR.map(sd => `<option value="${sd}">${sd}</option>`).join("")}
                </select>
              </div>
              <div><label class="block font-bold mb-1">Waktu Pelaksanaan</label><input type="text" name="waktuPelaksanaan" placeholder="Agustus 2026" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Belanja</button>
            </div>
          </form>
        </div>
      `);
    }

    function openEditArkasModal(arkasId) {
      const data = getCurrentData();
      const a = data.arkas.find(item => item.id === arkasId);
      if (!a) return;

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[88vh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Edit Uraian Belanja ARKAS</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveArkasModalForm(event, '${a.id}')" class="mt-4 space-y-3 text-xs">
            <input type="hidden" name="rktId" value="${escapeHtml(a.rktId)}" />
            <div>
              <label class="block font-bold mb-1">Komponen Pembiayaan Standar BOSP</label>
              <select name="komponen" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                ${KOMPONEN_BOSP_STANDAR.map(c => `<option value="${c}" ${a.komponen === c ? "selected" : ""}>${c}</option>`).join("")}
              </select>
            </div>
            <div><label class="block font-bold mb-1">Uraian Belanja *</label><input type="text" name="uraianBelanja" required value="${escapeHtml(a.uraianBelanja)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            <div class="grid grid-cols-3 gap-3">
              <div><label class="block font-bold mb-1">Volume *</label><input type="number" step="0.01" name="volume" required value="${Number(a.volume) || 0}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" /></div>
              <div><label class="block font-bold mb-1">Satuan *</label><input type="text" name="satuan" required value="${escapeHtml(a.satuan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
              <div><label class="block font-bold mb-1">Harga Satuan (Rp) *</label><input type="number" name="hargaSatuan" required value="${Number(a.hargaSatuan) || 0}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" /></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold mb-1">Sumber Dana</label>
                <select name="sumberDana" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  ${SUMBER_DANA_STANDAR.map(sd => `<option value="${sd}" ${a.sumberDana === sd ? "selected" : ""}>${sd}</option>`).join("")}
                </select>
              </div>
              <div><label class="block font-bold mb-1">Waktu Pelaksanaan</label><input type="text" name="waktuPelaksanaan" value="${escapeHtml(a.waktuPelaksanaan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            </div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Perubahan</button>
            </div>
          </form>
        </div>
      `);
    }

    function saveArkasModalForm(e, arkasId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const rktId = String(fd.get("rktId") || "");
      const keg = data.rkt.find(k => k.id === rktId);
      const existing = arkasId ? data.arkas.find(a => a.id === arkasId) : null;

      const obj = {
        id: arkasId || uid(),
        rktId,
        rkjmId: keg ? keg.rkjmId : (existing?.rkjmId || ""),
        indicatorId: keg ? keg.indicatorId : (existing?.indicatorId || ""),
        program: keg ? keg.namaProgram : (existing?.program || "Program Pengembangan Sekolah"),
        kegiatan: keg ? keg.namaKegiatan : (existing?.kegiatan || "Kegiatan Sekolah"),
        komponen: String(fd.get("komponen") || "").trim(),
        uraianBelanja: String(fd.get("uraianBelanja") || "").trim(),
        volume: Number(fd.get("volume")) || 0,
        satuan: String(fd.get("satuan") || "").trim(),
        hargaSatuan: Number(fd.get("hargaSatuan")) || 0,
        sumberDana: String(fd.get("sumberDana") || "BOSP Reguler").trim(),
        waktuPelaksanaan: String(fd.get("waktuPelaksanaan") || "").trim(),
        statusVerifikasi: "Terverifikasi Tim BOS Sekolah"
      };

      if (arkasId) {
        const idx = data.arkas.findIndex(a => a.id === arkasId);
        if (idx >= 0) data.arkas[idx] = obj;
      } else {
        data.arkas.push(obj);
      }

      closeCustomModal();
      triggerAutoSave("ARKAS", arkasId ? "Edit Belanja ARKAS" : "Tambah Belanja ARKAS", `Menyimpan belanja "${obj.uraianBelanja}".`);
      renderAll();
      showToast("Rincian anggaran ARKAS berhasil disimpan.", "success");
    }

    function deleteArkas(arkasId) {
      const data = getCurrentData();
      data.arkas = data.arkas.filter(a => a.id !== arkasId);
      triggerAutoSave("ARKAS", "Hapus Belanja ARKAS", "Menghapus 1 item belanja ARKAS.");
      renderAll();
      showToast("Item belanja ARKAS dihapus.", "info");
    }

    // =========================================================================
    // 17. MODULE 9: BENANG MERAH ANTARDOCUMEN (END-TO-END TRACEABILITY)
    // =========================================================================
    function renderKoneksiModule() {
      const data = getCurrentData();

      return `
        <div>
          <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 9 • Audit Keterhubungan PBD</span>
          <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Benang Merah Antardokumen</h2>
          <p class="text-xs text-slate-600 mt-0.5">Penelusuran utuh dari <strong>Nilai Capaian Rapor Pendidikan → Akar Masalah → Prioritas → Rekomendasi → RKJM → RKT → Anggaran ARKAS</strong>.</p>
        </div>

        ${data.indicators.length === 0 ? `
          <div class="card-elevated p-10 text-center text-xs text-slate-500">
            Belum ada indikator Rapor Pendidikan untuk ditelusuri keterhubungannya.
          </div>
        ` : `
          <div class="space-y-4">
            ${data.indicators.map(ind => {
              const rc = data.rootCauses.find(r => r.indicatorId === ind.id);
              const prio = data.priorities.find(p => p.indicatorId === ind.id);
              const rec = data.recommendations.find(r => r.indicatorId === ind.id);
              const progs = data.rkjm.filter(p => p.indicatorId === ind.id);
              const kegs = data.rkt.filter(k => k.indicatorId === ind.id);
              const budgets = data.arkas.filter(a => a.indicatorId === ind.id);
              const totalBudget = budgets.reduce((s, b) => s + ((Number(b.volume) || 0) * (Number(b.hargaSatuan) || 0)), 0);

              return `
                <div class="card-elevated p-5 space-y-3">
                  <div class="flex flex-wrap items-center justify-between gap-2 pb-2.5 border-b border-slate-100">
                    <div class="flex items-center gap-2 flex-wrap">
                      <span class="font-bold text-slate-900 text-sm">${ind.kode ? `[${escapeHtml(ind.kode)}] ` : ""}${escapeHtml(ind.nama)}</span>
                      <span class="px-2.5 py-0.5 rounded-full bg-slate-100 font-mono text-xs font-bold text-slate-800">Nilai Capaian: ${formatScore(ind.nilai)}</span>
                      ${getCategoryBadgeHtml(classifyCategory(ind.nilai, ind.kategori))}
                    </div>
                    <span class="font-mono text-xs font-extrabold text-emerald-900">Alokasi ARKAS: ${formatRupiah(totalBudget)}</span>
                  </div>

                  <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-2 text-[11px]">
                    <div class="p-2.5 rounded-xl border ${rc ? 'bg-emerald-50/60 border-emerald-200' : 'bg-slate-50 border-slate-200 text-slate-400'}">
                      <div class="font-bold text-slate-900">1. Akar Masalah</div>
                      <div class="mt-1 line-clamp-2">${rc ? escapeHtml(rc.masalahUtama) : "Belum dianalisis"}</div>
                    </div>
                    <div class="p-2.5 rounded-xl border ${prio ? 'bg-emerald-50/60 border-emerald-200' : 'bg-slate-50 border-slate-200 text-slate-400'}">
                      <div class="font-bold text-slate-900">2. Prioritas</div>
                      <div class="mt-1">${prio ? `Peringkat #${prio.peringkatManual} (Skor ${prio.totalSkor}/25)` : "Belum diperingkat"}</div>
                    </div>
                    <div class="p-2.5 rounded-xl border ${rec ? 'bg-emerald-50/60 border-emerald-200' : 'bg-slate-50 border-slate-200 text-slate-400'}">
                      <div class="font-bold text-slate-900">3. Rekomendasi</div>
                      <div class="mt-1 line-clamp-2">${rec ? escapeHtml(rec.bentukKegiatan) : "Belum disusun"}</div>
                    </div>
                    <div class="p-2.5 rounded-xl border ${progs.length > 0 ? 'bg-emerald-50/60 border-emerald-200' : 'bg-slate-50 border-slate-200 text-slate-400'}">
                      <div class="font-bold text-slate-900">4. RKJM (${progs.length})</div>
                      <div class="mt-1 line-clamp-2">${progs.length > 0 ? progs.map(p => escapeHtml(p.programPengembangan)).join("; ") : "Belum masuk RKJM"}</div>
                    </div>
                    <div class="p-2.5 rounded-xl border ${kegs.length > 0 ? 'bg-emerald-50/60 border-emerald-200' : 'bg-slate-50 border-slate-200 text-slate-400'}">
                      <div class="font-bold text-slate-900">5. RKT (${kegs.length})</div>
                      <div class="mt-1 line-clamp-2">${kegs.length > 0 ? kegs.map(k => escapeHtml(k.namaKegiatan)).join("; ") : "Belum masuk RKT"}</div>
                    </div>
                    <div class="p-2.5 rounded-xl border ${budgets.length > 0 ? 'bg-amber-50 border-amber-300' : 'bg-slate-50 border-slate-200 text-slate-400'}">
                      <div class="font-bold text-slate-900">6. ARKAS (${budgets.length} Item)</div>
                      <div class="mt-1 font-mono font-bold text-emerald-950">${formatRupiah(totalBudget)}</div>
                    </div>
                  </div>
                </div>
              `;
            }).join("")}
          </div>
        `}
      `;
    }

    // =========================================================================
    // 18. MODULE 10: MONITORING & EVALUASI (MONEV)
    // =========================================================================
    function renderMonevModule() {
      const data = getCurrentData();

      // Sync monev entries with RKT activities
      data.rkt.forEach(keg => {
        if (!data.monev.some(m => m.rktId === keg.id)) {
          data.monev.push({
            id: uid(),
            rktId: keg.id,
            status: "Belum Dilaksanakan",
            realisasiFisik: 0,
            realisasiAnggaran: 0,
            hambatan: "-",
            buktiKegiatan: "-",
            evaluasiHasil: "-",
            tindakLanjut: "-"
          });
        }
      });

      const selesai = data.monev.filter(m => m.status === "Selesai").length;
      const berjalan = data.monev.filter(m => m.status === "Sedang Berjalan").length;
      const belum = data.monev.filter(m => m.status === "Belum Dilaksanakan").length;

      return `
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 10 • Pengendalian &amp; Evaluasi</span>
            <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Monitoring dan Evaluasi Program (Monev)</h2>
            <p class="text-xs text-slate-600 mt-0.5">Pantau keterlaksanaan kegiatan RKT, serapan anggaran, hambatan lapangan, serta tindak lanjut perbaikan.</p>
          </div>
          <button onclick="exportPrintReport('monev')" class="px-3.5 py-2 rounded-xl bg-white border border-slate-300 hover:bg-slate-50 text-rose-700 font-bold text-xs flex items-center gap-1.5 self-start sm:self-auto">
            <i data-lucide="download" class="w-4 h-4"></i>
            <span>Unduh Laporan Monev</span>
          </button>
        </div>

        <section class="grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs">
          <div class="card-elevated p-4 border-l-4 border-l-emerald-600">
            <div class="font-bold text-emerald-800">Program Selesai</div>
            <div class="font-mono text-2xl font-extrabold text-slate-900 mt-1">${selesai}</div>
          </div>
          <div class="card-elevated p-4 border-l-4 border-l-amber-500">
            <div class="font-bold text-amber-800">Sedang Berjalan</div>
            <div class="font-mono text-2xl font-extrabold text-slate-900 mt-1">${berjalan}</div>
          </div>
          <div class="card-elevated p-4 border-l-4 border-l-slate-400">
            <div class="font-bold text-slate-600">Belum Dilaksanakan</div>
            <div class="font-mono text-2xl font-extrabold text-slate-900 mt-1">${belum}</div>
          </div>
        </section>

        ${data.monev.length === 0 ? `
          <div class="card-elevated p-10 text-center text-xs text-slate-500">
            Belum ada kegiatan RKT yang dapat dimonitor. Susun RKT terlebih dahulu.
          </div>
        ` : `
          <div class="space-y-4">
            ${data.monev.map(m => {
              const keg = data.rkt.find(k => k.id === m.rktId);
              if (!keg) return "";
              const pagu = Number(keg.kebutuhanAnggaran) || 0;
              const pctAnggaran = pagu > 0 ? Math.min(100, Math.round(((Number(m.realisasiAnggaran) || 0) / pagu) * 100)) : 0;

              return `
                <div class="card-elevated p-5 space-y-3 text-xs">
                  <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-2.5 border-b border-slate-100">
                    <div>
                      <div class="text-[10px] font-bold text-emerald-800 uppercase">${escapeHtml(keg.namaProgram)}</div>
                      <h3 class="font-bold text-slate-900 text-sm">${escapeHtml(keg.namaKegiatan)}</h3>
                    </div>
                    <div class="flex items-center gap-2">
                      <span class="px-2.5 py-1 rounded-full font-bold text-[11px] ${m.status === 'Selesai' ? 'bg-emerald-100 text-emerald-900' : m.status === 'Sedang Berjalan' ? 'bg-amber-100 text-amber-900' : 'bg-slate-100 text-slate-700'}">
                        ${escapeHtml(m.status)}
                      </span>
                      <button onclick="openEditMonevModal('${m.id}')" class="px-3 py-1.5 rounded-xl bg-emerald-900 text-white font-bold">Update Monev</button>
                    </div>
                  </div>

                  <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <div class="flex justify-between mb-1">
                        <span class="font-semibold text-slate-700">Realisasi Kegiatan Fisik</span>
                        <span class="font-mono font-bold text-emerald-900">${m.realisasiFisik}%</span>
                      </div>
                      <div class="w-full h-2.5 bg-slate-100 rounded-full overflow-hidden">
                        <div class="h-full bg-emerald-600 rounded-full" style="width: ${m.realisasiFisik}%"></div>
                      </div>
                    </div>
                    <div>
                      <div class="flex justify-between mb-1">
                        <span class="font-semibold text-slate-700">Realisasi Anggaran (${formatRupiah(m.realisasiAnggaran)} / ${formatRupiah(pagu)})</span>
                        <span class="font-mono font-bold text-amber-800">${pctAnggaran}%</span>
                      </div>
                      <div class="w-full h-2.5 bg-slate-100 rounded-full overflow-hidden">
                        <div class="h-full bg-amber-500 rounded-full" style="width: ${pctAnggaran}%"></div>
                      </div>
                    </div>
                  </div>

                  <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 pt-1">
                    <div class="p-2.5 rounded-xl bg-slate-50"><strong>Hambatan:</strong> ${escapeHtml(m.hambatan)}</div>
                    <div class="p-2.5 rounded-xl bg-slate-50"><strong>Bukti Kegiatan:</strong> ${escapeHtml(m.buktiKegiatan)}</div>
                    <div class="p-2.5 rounded-xl bg-slate-50"><strong>Evaluasi Hasil:</strong> ${escapeHtml(m.evaluasiHasil)}</div>
                    <div class="p-2.5 rounded-xl bg-emerald-50/70"><strong>Tindak Lanjut:</strong> ${escapeHtml(m.tindakLanjut)}</div>
                  </div>
                </div>
              `;
            }).join("")}
          </div>
        `}
      `;
    }

    function openEditMonevModal(monevId) {
      const data = getCurrentData();
      const m = data.monev.find(item => item.id === monevId);
      if (!m) return;

      openCustomModal(`
        <div class="p-5 sm:p-6">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Perbarui Capaian Monitoring &amp; Evaluasi</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <form onsubmit="saveMonevEdit(event, '${m.id}')" class="mt-4 space-y-3 text-xs">
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block font-bold mb-1">Status Pelaksanaan</label>
                <select name="status" class="w-full rounded-xl border border-slate-300 px-3 py-2">
                  ${["Belum Dilaksanakan", "Sedang Berjalan", "Selesai"].map(st => `<option value="${st}" ${m.status === st ? "selected" : ""}>${st}</option>`).join("")}
                </select>
              </div>
              <div><label class="block font-bold mb-1">Realisasi Fisik (%)</label><input type="number" min="0" max="100" name="realisasiFisik" value="${m.realisasiFisik}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" /></div>
              <div><label class="block font-bold mb-1">Realisasi Anggaran (Rp)</label><input type="number" min="0" name="realisasiAnggaran" value="${m.realisasiAnggaran}" class="w-full rounded-xl border border-slate-300 px-3 py-2 font-mono" /></div>
            </div>
            <div><label class="block font-bold mb-1">Hambatan Pelaksanaan</label><textarea name="hambatan" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(m.hambatan)}</textarea></div>
            <div><label class="block font-bold mb-1">Bukti Kegiatan</label><input type="text" name="buktiKegiatan" value="${escapeHtml(m.buktiKegiatan)}" class="w-full rounded-xl border border-slate-300 px-3 py-2" /></div>
            <div><label class="block font-bold mb-1">Evaluasi Hasil</label><textarea name="evaluasiHasil" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(m.evaluasiHasil)}</textarea></div>
            <div><label class="block font-bold mb-1">Rekomendasi Tindak Lanjut</label><textarea name="tindakLanjut" rows="2" class="w-full rounded-xl border border-slate-300 px-3 py-2">${escapeHtml(m.tindakLanjut)}</textarea></div>
            <div class="flex justify-end gap-2 pt-2">
              <button type="button" onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Batal</button>
              <button type="submit" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Simpan Monev</button>
            </div>
          </form>
        </div>
      `);
    }

    function saveMonevEdit(e, monevId) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const data = getCurrentData();
      const m = data.monev.find(item => item.id === monevId);
      if (!m) return;

      m.status = String(fd.get("status") || "Belum Dilaksanakan");
      m.realisasiFisik = Number(fd.get("realisasiFisik")) || 0;
      m.realisasiAnggaran = Number(fd.get("realisasiAnggaran")) || 0;
      m.hambatan = String(fd.get("hambatan") || "-").trim();
      m.buktiKegiatan = String(fd.get("buktiKegiatan") || "-").trim();
      m.evaluasiHasil = String(fd.get("evaluasiHasil") || "-").trim();
      m.tindakLanjut = String(fd.get("tindakLanjut") || "-").trim();

      closeCustomModal();
      triggerAutoSave("Monev", "Update Monev", "Memperbarui laporan monitoring dan evaluasi kegiatan.");
      renderAll();
      showToast("Data monitoring dan evaluasi berhasil disimpan.", "success");
    }

    // =========================================================================
    // 19. MODULE 11: LAPORAN & EKSPOR DOKUMEN (PDF, WORD .DOC, EXCEL .XLSX)
    // =========================================================================
    function renderEksporModule() {
      const data = getCurrentData();
      const exportsList = [
        { title: "1. Hasil Analisis Rapor Pendidikan & Nilai Capaian", desc: "Laporan lengkap skor capaian, kategori, tren, dan indikator prioritas.", format: "PDF / Cetak", action: "exportPrintReport('analisis')", btnClass: "bg-rose-700 text-white" },
        { title: "2. Dokumen Analisis Akar Masalah (5-Why)", desc: "Masalah utama, dugaan AI vs fakta terkonfirmasi sekolah, dan refleksi guru.", format: "Word (.doc)", action: "exportWordDocument('akar')", btnClass: "bg-blue-700 text-white" },
        { title: "3. Daftar Prioritas Pembenahan", desc: "Tabel pembobotan 5 kriteria dan urutan prioritas pembenahan sekolah.", format: "Excel (.xlsx)", action: "exportExcelDocument('prioritas')", btnClass: "bg-emerald-800 text-white" },
        { title: "4. Dokumen RKJM (4 Tahunan)", desc: "Rencana Kerja Jangka Menengah lengkap dengan capaian awal dan target tahunan.", format: "Word & Excel", dual: true, wordAction: "exportWordDocument('rkjm')", excelAction: "exportExcelDocument('rkjm')" },
        { title: "5. Dokumen RKT (Tahunan)", desc: "Rencana Kerja Tahunan lengkap dengan sasaran, jadwal, dan kebutuhan anggaran.", format: "Word & Excel", dual: true, wordAction: "exportWordDocument('rkt')", excelAction: "exportExcelDocument('rkt')" },
        { title: "6. Rancangan Anggaran ARKAS", desc: "Kertas kerja rincian belanja Volume x Harga Satuan & rekap komponen BOSP.", format: "Excel (.xlsx)", action: "exportExcelDocument('arkas')", btnClass: "bg-emerald-800 text-white" },
        { title: "7. Laporan Monitoring & Evaluasi", desc: "Realisasi fisik, serapan anggaran, hambatan, dan tindak lanjut program.", format: "PDF / Cetak", action: "exportPrintReport('monev')", btnClass: "bg-rose-700 text-white" }
      ];

      return `
        <div>
          <span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">Tahap 11 • Dokumentasi Resmi Sekolah</span>
          <h2 class="font-heading text-xl sm:text-2xl font-extrabold text-slate-900">Laporan &amp; Unduh Dokumen Perencanaan</h2>
          <p class="text-xs text-slate-600 mt-0.5">Seluruh dokumen dilengkapi kop identitas satuan pendidikan, tahun perencanaan, dan kolom pengesahan Kepala Sekolah.</p>
        </div>

        <section class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-xs">
          ${exportsList.map(item => `
            <div class="card-elevated p-5 flex flex-col justify-between">
              <div>
                <span class="inline-block px-2.5 py-0.5 rounded-full bg-amber-100 text-amber-950 font-bold text-[10px] mb-2">${item.format}</span>
                <h3 class="font-bold text-slate-900 text-sm">${item.title}</h3>
                <p class="text-slate-600 mt-1 leading-relaxed">${item.desc}</p>
              </div>
              <div class="mt-4 pt-3 border-t border-slate-100">
                ${item.dual ? `
                  <div class="grid grid-cols-2 gap-2">
                    <button onclick="${item.wordAction}" class="py-2 px-3 rounded-xl bg-blue-700 hover:bg-blue-800 text-white font-bold flex items-center justify-center gap-1.5">
                      <i data-lucide="file-text" class="w-3.5 h-3.5"></i><span>Unduh Word</span>
                    </button>
                    <button onclick="${item.excelAction}" class="py-2 px-3 rounded-xl bg-emerald-800 hover:bg-emerald-900 text-white font-bold flex items-center justify-center gap-1.5">
                      <i data-lucide="file-spreadsheet" class="w-3.5 h-3.5"></i><span>Unduh Excel</span>
                    </button>
                  </div>
                ` : `
                  <button onclick="${item.action}" class="w-full py-2 px-4 rounded-xl ${item.btnClass} hover:opacity-95 font-bold flex items-center justify-center gap-1.5">
                    <i data-lucide="download" class="w-3.5 h-3.5"></i><span>Unduh ${item.format}</span>
                  </button>
                `}
              </div>
            </div>
          `).join("")}
        </section>

        <!-- AUDIT HISTORY LOG -->
        <section class="card-elevated p-5">
          <h3 class="font-bold text-slate-900 text-sm pb-2 border-b border-slate-100">Riwayat Perubahan Data (${data.history.length})</h3>
          <div class="mt-3 max-h-52 overflow-y-auto space-y-2 text-xs">
            ${data.history.map(h => `
              <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-1 py-1.5 border-b border-slate-50">
                <div>
                  <span class="font-bold text-emerald-900">[${escapeHtml(h.modul)}]</span>
                  <span class="font-semibold text-slate-800">${escapeHtml(h.aksi)}:</span>
                  <span class="text-slate-600">${escapeHtml(h.detail)}</span>
                </div>
                <span class="font-mono text-[11px] text-slate-400 shrink-0">${escapeHtml(h.waktu)}</span>
              </div>
            `).join("")}
          </div>
        </section>
      `;
    }

    function buildFormalReportHeaderHtml(docTitle) {
      const s = getCurrentData().school;
      return `
        <div style="text-align:center; border-bottom: 2px solid #064e3b; padding-bottom: 10px; margin-bottom: 16px;">
          <h2 style="margin:0; font-size:16px; color:#064e3b; text-transform:uppercase;">${escapeHtml(docTitle)}</h2>
          <h3 style="margin:4px 0 0; font-size:14px;">${escapeHtml(s.nama || "SATUAN PENDIDIKAN")} (NPSN: ${escapeHtml(s.npsn || "-")})</h3>
          <p style="margin:3px 0 0; font-size:11px; color:#475569;">${escapeHtml([s.alamat, s.kecamatan, s.kabupaten, s.provinsi].filter(Boolean).join(", "))} • Tahun Perencanaan: ${escapeHtml(s.tahunPerencanaan)}</p>
        </div>
      `;
    }

    function exportWordDocument(type) {
      const data = getCurrentData();
      let title = "";
      let bodyHtml = "";

      if (type === "akar") {
        title = "DOKUMEN ANALISIS MASALAH DAN AKAR MASALAH (5-WHY)";
        bodyHtml = data.rootCauses.map((rc, idx) => {
          const ind = data.indicators.find(i => i.id === rc.indicatorId) || { nama: "Indikator", nilai: "" };
          return `
            <h4>${idx + 1}. ${escapeHtml(ind.nama)} (Nilai Capaian: ${formatScore(ind.nilai)})</h4>
            <p><strong>Masalah Utama:</strong> ${escapeHtml(rc.masalahUtama)}</p>
            <p><strong>Akar Penyebab:</strong></p>
            <ul>${(rc.dugaanPenyebab || []).map(c => `<li>[${c.terkonfirmasi ? "Fakta Terkonfirmasi" : "Dugaan AI"}] ${escapeHtml(c.teks)}</li>`).join("")}</ul>
            <p><strong>Dampak Peserta Didik:</strong> ${escapeHtml(rc.dampakPesertaDidik)}</p>
            <p><strong>Alternatif Solusi:</strong> ${escapeHtml(rc.alternatifSolusi)}</p>
          `;
        }).join("");
      } else if (type === "rkjm") {
        title = `RENCANA KERJA JANGKA MENENGAH (RKJM) ${data.school.periodeRkjmAwal} - ${data.school.periodeRkjmAkhir}`;
        bodyHtml = `
          <table border="1" cellspacing="0" cellpadding="6" style="width:100%; border-collapse:collapse; font-size:11px;">
            <tr style="background:#064e3b; color:#fff;">
              <th>No</th><th>Indikator</th><th>Program Pengembangan</th><th>Tujuan Strategis</th><th>Target Th 1</th><th>Target Th 2</th><th>Target Th 3</th><th>Target Th 4</th><th>Sumber Dana</th>
            </tr>
            ${data.rkjm.map((p, i) => `
              <tr>
                <td>${i + 1}</td><td>${escapeHtml(p.indikatorNama)}</td><td>${escapeHtml(p.programPengembangan)}</td><td>${escapeHtml(p.tujuanStrategis)}</td>
                <td>${escapeHtml(p.targetTahun1)}</td><td>${escapeHtml(p.targetTahun2)}</td><td>${escapeHtml(p.targetTahun3)}</td><td>${escapeHtml(p.targetTahun4)}</td><td>${escapeHtml(p.sumberPendanaan)}</td>
              </tr>
            `).join("")}
          </table>
        `;
      } else if (type === "rkt") {
        title = `RENCANA KERJA TAHUNAN (RKT) TAHUN ${data.school.tahunPerencanaan}`;
        bodyHtml = `
          <table border="1" cellspacing="0" cellpadding="6" style="width:100%; border-collapse:collapse; font-size:11px;">
            <tr style="background:#064e3b; color:#fff;">
              <th>No</th><th>Program</th><th>Kegiatan</th><th>Sasaran</th><th>Target</th><th>Jadwal</th><th>PJ</th><th>Anggaran</th>
            </tr>
            ${data.rkt.map((k, i) => `
              <tr>
                <td>${i + 1}</td><td>${escapeHtml(k.namaProgram)}</td><td>${escapeHtml(k.namaKegiatan)}</td><td>${escapeHtml(k.sasaran)}</td>
                <td>${escapeHtml(k.target)}</td><td>${escapeHtml(k.jadwalTeks)}</td><td>${escapeHtml(k.penanggungJawab)}</td><td>${formatRupiah(k.kebutuhanAnggaran)}</td>
              </tr>
            `).join("")}
          </table>
        `;
      }

      const fullHtml = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word'>
        <head><meta charset='utf-8'><title>${title}</title></head>
        <body style="font-family: Arial, sans-serif; font-size:12px;">
          ${buildFormalReportHeaderHtml(title)}
          ${bodyHtml || "<p>Data belum tersedia.</p>"}
          <br/><br/>
          <table style="width:100%; font-size:11px;">
            <tr>
              <td style="width:60%;"></td>
              <td style="text-align:center;">Kepala Sekolah,<br/><br/><br/><br/><strong>${escapeHtml(data.school.kepalaSekolah || "______________________")}</strong></td>
            </tr>
          </table>
        </body></html>
      `;

      const blob = new Blob(["\ufeff", fullHtml], { type: "application/msword" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${title.replace(/[^a-zA-Z0-9]+/g, "_")}.doc`;
      a.click();
      URL.revokeObjectURL(url);
      showToast("Dokumen Word (.doc) berhasil diunduh.", "success");
    }

    function exportExcelDocument(type) {
      if (!window.XLSX) {
        showToast("Pustaka Excel belum siap.", "warning");
        return;
      }
      const data = getCurrentData();
      const s = data.school;
      let sheetName = "Laporan";
      let fileName = "Laporan_RAPOR_SMART_AI.xlsx";
      let aoa = [];

      if (type === "prioritas") {
        sheetName = "Prioritas Pembenahan";
        fileName = `Daftar_Prioritas_${s.tahunPerencanaan}.xlsx`;
        aoa = [
          ["DAFTAR PRIORITAS PEMBENAHAN RAPOR PENDIDIKAN"],
          [`Sekolah: ${s.nama || "-"} | NPSN: ${s.npsn || "-"} | Tahun: ${s.tahunPerencanaan}`],
          [],
          ["Peringkat", "Kode", "Nama Indikator", "Nilai Capaian", "Kategori", "Total Skor Prioritas", "Alasan Prioritas"],
          ...data.priorities.map((p, i) => {
            const ind = data.indicators.find(x => x.id === p.indicatorId) || {};
            return [i + 1, ind.kode || "-", ind.nama || "-", ind.nilai !== "" ? ind.nilai : "-", classifyCategory(ind.nilai, ind.kategori), p.totalSkor, p.alasanPrioritas];
          })
        ];
      } else if (type === "rkjm") {
        sheetName = "RKJM";
        fileName = `RKJM_${s.periodeRkjmAwal}_${s.periodeRkjmAkhir}.xlsx`;
        aoa = [
          [`RENCANA KERJA JANGKA MENENGAH (RKJM) ${s.periodeRkjmAwal} - ${s.periodeRkjmAkhir}`],
          [`Sekolah: ${s.nama || "-"} | NPSN: ${s.npsn || "-"}`],
          [],
          ["No", "Dimensi", "Indikator Rapor", "Akar Masalah", "Program Pengembangan", "Target Th 1", "Target Th 2", "Target Th 3", "Target Th 4", "Penanggung Jawab", "Sumber Dana"],
          ...data.rkjm.map((p, i) => [i + 1, p.dimensi, p.indikatorNama, p.akarMasalah, p.programPengembangan, p.targetTahun1, p.targetTahun2, p.targetTahun3, p.targetTahun4, p.penanggungJawab, p.sumberPendanaan])
        ];
      } else if (type === "rkt") {
        sheetName = "RKT";
        fileName = `RKT_${s.tahunPerencanaan}.xlsx`;
        aoa = [
          [`RENCANA KERJA TAHUNAN (RKT) TAHUN ${s.tahunPerencanaan}`],
          [`Sekolah: ${s.nama || "-"} | NPSN: ${s.npsn || "-"}`],
          [],
          ["No", "Program RKJM", "Kegiatan RKT", "Tujuan", "Sasaran", "Target", "Jadwal", "Penanggung Jawab", "Kebutuhan Anggaran", "Sumber Dana"],
          ...data.rkt.map((k, i) => [i + 1, k.namaProgram, k.namaKegiatan, k.tujuan, k.sasaran, k.target, k.jadwalTeks, k.penanggungJawab, Number(k.kebutuhanAnggaran) || 0, k.sumberDana])
        ];
      } else if (type === "arkas") {
        sheetName = "Rancangan ARKAS";
        fileName = `Rancangan_ARKAS_${s.tahunPerencanaan}.xlsx`;
        aoa = [
          [`RANCANGAN ANGGARAN ARKAS (KERTAS KERJA INTERNAL) TAHUN ${s.tahunPerencanaan}`],
          [`Sekolah: ${s.nama || "-"} | NPSN: ${s.npsn || "-"}`],
          [],
          ["No", "Program", "Kegiatan", "Komponen Pembiayaan", "Uraian Belanja", "Volume", "Satuan", "Harga Satuan", "Jumlah Anggaran", "Sumber Dana", "Waktu"],
          ...data.arkas.map((a, i) => {
            const jml = (Number(a.volume) || 0) * (Number(a.hargaSatuan) || 0);
            return [i + 1, a.program, a.kegiatan, a.komponen, a.uraianBelanja, Number(a.volume) || 0, a.satuan, Number(a.hargaSatuan) || 0, jml, a.sumberDana, a.waktuPelaksanaan];
          })
        ];
      }

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      XLSX.writeFile(wb, fileName);
      showToast(`File Excel "${fileName}" berhasil diunduh.`, "success");
    }

    function exportPrintReport(type) {
      const data = getCurrentData();
      let title = "";
      let tableHtml = "";

      if (type === "analisis") {
        title = `LAPORAN HASIL ANALISIS RAPOR PENDIDIKAN & NILAI CAPAIAN (${data.school.tahunRapor})`;
        tableHtml = `
          <table border="1" cellspacing="0" cellpadding="6" style="width:100%; border-collapse:collapse; font-size:11px;">
            <tr style="background:#064e3b; color:#fff;">
              <th>No</th><th>Kode</th><th>Nama Indikator</th><th>Dimensi</th><th>Nilai Capaian</th><th>Kategori</th><th>Perubahan</th>
            </tr>
            ${data.indicators.map((ind, i) => `
              <tr>
                <td>${i + 1}</td><td>${escapeHtml(ind.kode || "-")}</td><td>${escapeHtml(ind.nama)}</td><td>${escapeHtml(ind.dimensi || "-")}</td>
                <td style="text-align:center; font-weight:bold;">${formatScore(ind.nilai)}</td><td>${classifyCategory(ind.nilai, ind.kategori)}</td><td style="text-align:center;">${ind.perubahan !== "" ? ind.perubahan : "-"}</td>
              </tr>
            `).join("")}
          </table>
        `;
      } else {
        title = `LAPORAN MONITORING DAN EVALUASI PROGRAM (${data.school.tahunPerencanaan})`;
        tableHtml = `
          <table border="1" cellspacing="0" cellpadding="6" style="width:100%; border-collapse:collapse; font-size:11px;">
            <tr style="background:#064e3b; color:#fff;">
              <th>No</th><th>Kegiatan RKT</th><th>Status</th><th>Realisasi Fisik</th><th>Realisasi Anggaran</th><th>Hambatan</th><th>Tindak Lanjut</th>
            </tr>
            ${data.monev.map((m, i) => {
              const keg = data.rkt.find(k => k.id === m.rktId) || {};
              return `
                <tr>
                  <td>${i + 1}</td><td>${escapeHtml(keg.namaKegiatan || "-")}</td><td>${escapeHtml(m.status)}</td>
                  <td>${m.realisasiFisik}%</td><td>${formatRupiah(m.realisasiAnggaran)}</td><td>${escapeHtml(m.hambatan)}</td><td>${escapeHtml(m.tindakLanjut)}</td>
                </tr>
              `;
            }).join("")}
          </table>
        `;
      }

      const fullDocHtml = `
        <html><head><meta charset="utf-8"><title>${title}</title></head>
        <body style="font-family: Arial, sans-serif; padding: 24px;">
          ${buildFormalReportHeaderHtml(title)}
          ${tableHtml}
          <br/><br/>
          <div style="text-align:right; font-size:11px; margin-top:24px;">
            Kepala Sekolah,<br/><br/><br/><br/>
            <strong>${escapeHtml(data.school.kepalaSekolah || "______________________")}</strong>
          </div>
        </body></html>
      `;

      window._lastPrintableReport = { title, html: fullDocHtml };

      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[90vh] overflow-y-auto space-y-4">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <div>
              <h3 class="font-bold text-slate-900 text-base">Pratinjau &amp; Unduh Laporan</h3>
              <p class="text-xs text-slate-500">Pilih Cetak / Simpan PDF melalui dialog browser atau langsung unduh file dokumen siap buka.</p>
            </div>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <div class="border border-slate-200 rounded-xl p-4 bg-white overflow-x-auto max-h-80">
            ${buildFormalReportHeaderHtml(title)}
            ${tableHtml}
          </div>
          <div class="flex flex-wrap justify-end gap-2 text-xs pt-2">
            <button onclick="closeCustomModal()" class="px-4 py-2 rounded-xl border border-slate-200 font-semibold">Tutup</button>
            <button onclick="downloadLastReportAsWord()" class="px-4 py-2 rounded-xl bg-blue-700 hover:bg-blue-800 text-white font-bold flex items-center gap-1.5">
              <i data-lucide="file-text" class="w-4 h-4"></i>
              <span>Unduh File Dokumen (.doc)</span>
            </button>
            <button onclick="printLastReportWindow()" class="px-4 py-2 rounded-xl emerald-gradient text-white font-bold flex items-center gap-1.5 shadow">
              <i data-lucide="printer" class="w-4 h-4 text-amber-300"></i>
              <span>Cetak / Simpan PDF</span>
            </button>
          </div>
        </div>
      `, "max-w-3xl");
    }

    function downloadLastReportAsWord() {
      const rep = window._lastPrintableReport;
      if (!rep) return;
      const blob = new Blob(["\ufeff", rep.html], { type: "application/msword" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${rep.title.replace(/[^a-zA-Z0-9]+/g, "_")}.doc`;
      a.click();
      URL.revokeObjectURL(url);
      showToast("File laporan (.doc) berhasil diunduh.", "success");
    }

    function printLastReportWindow() {
      const rep = window._lastPrintableReport;
      if (!rep) return;
      const printWin = window.open("", "_blank");
      if (printWin) {
        printWin.document.write(rep.html + `<script>window.onload = function(){ window.print(); };<\/script>`);
        printWin.document.close();
      } else {
        window.print();
      }
    }

    function renderAssistantMessages() {
      const box = document.getElementById("assistantMessages");
      if (!box) return;
      box.innerHTML = APP_STATE.chatHistory.map(m => `
        <div class="flex ${m.role === 'user' ? 'justify-end' : 'justify-start'}">
          <div class="max-w-[85%] rounded-2xl px-3.5 py-2.5 leading-relaxed ${m.role === 'user' ? 'emerald-gradient text-white' : 'bg-white border border-slate-200 text-slate-800 shadow-2xs'}">
            ${escapeHtml(m.text).replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>").replace(/\n/g, "<br/>")}
          </div>
        </div>
      `).join("");
      box.scrollTop = box.scrollHeight;
    }

    function askAssistantQuick(promptText) {
      APP_STATE.chatHistory.push({ role: "user", text: promptText });
      APP_STATE.chatHistory.push({ role: "assistant", text: generateContextAwareAssistantReply(promptText) });
      renderAssistantMessages();
    }

    function handleAssistantSubmit(e) {
      e.preventDefault();
      const input = document.getElementById("assistantInput");
      if (!input) return;
      const q = input.value.trim();
      if (!q) return;
      input.value = "";
      askAssistantQuick(q);
    }

    function generateContextAwareAssistantReply(query) {
      const data = getCurrentData();
      const s = data.school;
      const low = query.toLowerCase();

      if (data.indicators.length === 0) {
        return `Saat ini belum ada data indikator Rapor Pendidikan maupun **Nilai Capaian** yang dimasukkan untuk ${s.nama || "sekolah Anda"}.\n\nSilakan unggah file Excel/PDF Rapor Pendidikan atau pilih indikator dari **Katalog PBD** pada menu **1. Input Rapor & Nilai Capaian** terlebih dahulu agar saya dapat menganalisis berdasarkan data nyata.`;
      }

      if (low.includes("ringkas") || low.includes("nilai") || low.includes("rapor")) {
        const lines = data.indicators.map(i => `• ${i.kode ? "[" + i.kode + "] " : ""}${i.nama}: **${formatScore(i.nilai)}** (${classifyCategory(i.nilai, i.kategori)})`).join("\n");
        return `Berikut ringkasan **Nilai Capaian** aktual ${s.nama || "Sekolah"} (Tahun ${s.tahunRapor}):\n\n${lines}\n\nApakah Bapak/Ibu ingin saya bantu menyusun analisis akar masalah atau menurunkannya ke RKJM?`;
      }

      if (low.includes("prioritas") || low.includes("akar")) {
        const prioList = data.priorities.length > 0
          ? data.priorities.map((p, idx) => {
              const ind = data.indicators.find(i => i.id === p.indicatorId);
              return `${idx + 1}. **${ind?.nama || "Indikator"}** (Nilai: ${formatScore(ind?.nilai)}) — Skor Prioritas ${p.totalSkor}/25`;
            }).join("\n")
          : "Daftar prioritas belum dihitung. Klik tombol **Lanjut Analisis Akar Masalah & Prioritas** pada menu Analisis.";
        return `Berikut pemetaan prioritas pembenahan berdasarkan data sekolah Anda:\n\n${prioList}\n\nPastikan setiap dugaan penyebab pada modul **3. Akar Masalah** sudah dicentang konfirmasinya bersama dewan guru.`;
      }

      if (low.includes("anggaran") || low.includes("arkas")) {
        const total = data.arkas.reduce((sum, a) => sum + ((Number(a.volume) || 0) * (Number(a.hargaSatuan) || 0)), 0);
        return `Total rancangan anggaran ARKAS saat ini sebesar **${formatRupiah(total)}** yang terbagi ke dalam **${data.arkas.length} uraian belanja**.\n\nIngat: Rancangan ini adalah kertas kerja internal sekolah dan perlu diverifikasi kesesuaian komponen BOSP serta standar harga satuan daerah sebelum diinput ke aplikasi ARKAS resmi.`;
      }

      if (low.includes("benang merah") || low.includes("keterhubungan")) {
        return `Status keterhubungan dokumen ${s.nama || "sekolah"}:\n• Indikator Rapor: **${data.indicators.length}**\n• Analisis Akar Masalah: **${data.rootCauses.length}**\n• Program RKJM: **${data.rkjm.length}**\n• Kegiatan RKT: **${data.rkt.length}**\n• Item Belanja ARKAS: **${data.arkas.length}**\n\nGunakan menu **Benang Merah Dokumen** untuk melihat penelusuran setiap indikator dari hulu ke hilir.`;
      }

      return `Berdasarkan data aktual ${s.nama || "sekolah Anda"}, tercatat **${data.indicators.length} indikator** Rapor Pendidikan, **${data.rkjm.length} program RKJM**, **${data.rkt.length} kegiatan RKT**, dan total rancangan anggaran **${formatRupiah(data.arkas.reduce((sum, a) => sum + ((Number(a.volume) || 0) * (Number(a.hargaSatuan) || 0)), 0))}**. Silakan sebutkan bagian spesifik yang ingin Anda diskusikan (misalnya: *Ringkas Nilai Capaian*, *Cek Prioritas*, atau *Rekap Anggaran ARKAS*).`;
    }

    function openGuideModal() {
      openCustomModal(`
        <div class="p-5 sm:p-6 max-h-[85vh] overflow-y-auto text-xs space-y-3">
          <div class="flex items-center justify-between pb-3 border-b border-slate-200">
            <h3 class="font-bold text-slate-900 text-base">Panduan Penggunaan RAPOR SMART AI</h3>
            <button onclick="closeCustomModal()" class="p-1.5 rounded-lg hover:bg-slate-100"><i data-lucide="x" class="w-4 h-4"></i></button>
          </div>
          <ol class="list-decimal list-inside space-y-2 text-slate-700 leading-relaxed">
            <li><strong>Input Rapor &amp; Nilai Capaian:</strong> Unggah langsung file Excel Rapor Pendidikan (.xlsx), PDF, atau pilih indikator dari Katalog PBD lalu ketik <strong>Nilai Capaian (0–100)</strong>.</li>
            <li><strong>Analisis &amp; Akar Masalah (5-Why):</strong> Tinjau pengelompokan capaian dan centang akar penyebab yang merupakan fakta terkonfirmasi diskusi guru.</li>
            <li><strong>Prioritas &amp; Rekomendasi:</strong> Sesuaikan bobot prioritas dan setujui rekomendasi program sekolah.</li>
            <li><strong>RKJM → RKT → ARKAS:</strong> Turunkan program jangka menengah (4 tahun) menjadi kegiatan tahunan dan rincian anggaran belanja.</li>
            <li><strong>Monitoring &amp; Unduh Laporan:</strong> Pantau realisasi program dan unduh dokumen ke format PDF, Word, atau Excel.</li>
          </ol>
          <div class="pt-2 flex justify-end">
            <button onclick="closeCustomModal()" class="px-5 py-2 rounded-xl emerald-gradient text-white font-bold">Mengerti</button>
          </div>
        </div>
      `);
    }

    // =========================================================================
    // 22. MASTER RENDER LOOP
    // =========================================================================
    function renderAll() {
      renderNavigationAndHeader();
      const container = document.getElementById("mainModuleContainer");
      if (!container) return;

      switch (APP_STATE.activeTab) {
        case "dashboard":
          container.innerHTML = renderDashboard();
          break;
        case "input":
          container.innerHTML = renderInputModule();
          break;
        case "analisis":
          container.innerHTML = renderAnalisisModule();
          break;
        case "akar":
          container.innerHTML = renderAkarMasalahModule();
          break;
        case "prioritas":
          container.innerHTML = renderPrioritasModule();
          break;
        case "rekomendasi":
          container.innerHTML = renderRekomendasiModule();
          break;
        case "rkjm":
          container.innerHTML = renderRkjmModule();
          break;
        case "rkt":
          container.innerHTML = renderRktModule();
          break;
        case "arkas":
          container.innerHTML = renderArkasModule();
          break;
        case "koneksi":
          container.innerHTML = renderKoneksiModule();
          break;
        case "monev":
          container.innerHTML = renderMonevModule();
          break;
        case "ekspor":
          container.innerHTML = renderEksporModule();
          break;
        default:
          container.innerHTML = renderDashboard();
      }

      if (window.lucide) {
        window.lucide.createIcons();
      }
    }

    window.addEventListener("DOMContentLoaded", () => {
      renderAll();
    });
  </script>
</body>
</html>
