/*
 * Konfigurasi frontend.
 * Isi APP_SCRIPT_URL dengan URL Web App Google Apps Script berakhiran /exec.
 */
const APP_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby-mdsr_0Z9xBJnR0votH2k9ceh11vbGkjghmpBZGwzXRtFKx_S8gZWcLKzWZIOKAJOZA/exec';
const MAX_PAYLOAD_CHARS = 7 * 1024 * 1024;

// Penanda versi. Setelah memilih file Excel, teks di bawah nama file menampilkan versi ini. Bila tidak muncul,
// browser masih menjalankan app.js lama (cache / belum ter-deploy di GitHub Pages).
const APP_VERSION = 'v4-withdraw-2026-10-18';
console.info('app.js', APP_VERSION);

// Tujuan impor. Kuncinya dikirim ke server; nama tab di sini hanya untuk tampilan (server punya daftar sendiri).
const IMPORT_TARGETS = {
  anggota: 'Update_Anggota',
  disbursement: 'Update_Disbursement',
  withdraw: 'Update_Withdraw'
};

/*
 * Tata letak khusus untuk Excel hasil laporan (ada judul di atas tabel, kolom A kosong, sel digabung/merge).
 * Tujuan yang TIDAK ada di daftar ini tetap memakai aturan lama: baris 1 = header, kolom A = kolom pertama.
 *
 * - headerMarkers: nama kolom yang harus ada pada baris header. Baris header dicari otomatis di scanRows baris pertama
 *   (huruf besar/kecil, spasi, baris baru, dan tanda baca diabaikan), jadi tidak peduli header ada di baris 5 atau 6.
 * - Tabel dibaca mulai dari sel header paling kiri (mis. kolom B), jadi kolom A yang kosong otomatis dilewati.
 * - Sel gabungan (merge): nilainya diletakkan di kolom PALING KIRI gabungan, kolom sisanya dibiarkan kosong.
 *   Contoh: CLIENTID (Excel C:E) -> Sheet kolom B, kolom C dan D kosong. OFFICERNAME (Excel O:P) -> Sheet kolom N, O kosong.
 * - Header Excel di baris 5 menjadi baris 1 di Google Sheets, data baris 6 menjadi baris 2.
 */
const TARGET_LAYOUTS = {
  withdraw: { headerMarkers: ['CLIENTID', 'ACCOUNTNO', 'OFFICERNAME'], scanRows: 50 }
};

const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');
const fileName = document.getElementById('fileName');
const fileHelp = document.getElementById('fileHelp');
const parseError = document.getElementById('parseError');
const previewBlock = document.getElementById('previewBlock');
const sheetPickerWrap = document.getElementById('sheetPickerWrap');
const sheetPicker = document.getElementById('sheetPicker');
const sheetNameLabel = document.getElementById('sheetNameLabel');
const rowCountLabel = document.getElementById('rowCountLabel');
const colCountLabel = document.getElementById('colCountLabel');
const previewTable = document.getElementById('previewTable');
const headerStatus = document.getElementById('headerStatus');
const importPanel = document.getElementById('importPanel');
const confirmCheckbox = document.getElementById('confirmCheckbox');
const importPassword = document.getElementById('importPassword');
const importButton = document.getElementById('importButton');
const resultMessage = document.getElementById('resultMessage');
const targetSwitch = document.querySelector('.target-switch');
const targetRadios = document.querySelectorAll('input[name="importTarget"]');

let workbook = null;
let selectedData = null;
let isSubmitting = false;
let currentTarget = 'anggota';

function setHidden(element, hidden) {
  element.classList.toggle('hidden', hidden);
}

function setParseError(message) {
  parseError.textContent = message;
  setHidden(parseError, !message);
}

function setResult(message, type) {
  resultMessage.textContent = message;
  resultMessage.className = 'result-message ' + type;
  setHidden(resultMessage, false);
}

function updateButtonState() {
  importButton.disabled = !selectedData || !confirmCheckbox.checked || !importPassword.value || isSubmitting || !isEndpointConfigured();
}

function isEndpointConfigured() {
  return /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/.test(APP_SCRIPT_URL);
}

function resetPreview() {
  selectedData = null;
  setHidden(previewBlock, true);
  setHidden(importPanel, true);
  confirmCheckbox.checked = false;
  importPassword.value = '';
  importButton.disabled = true;
  resultMessage.textContent = '';
  setHidden(resultMessage, true);
}

function resetFileSelection() {
  workbook = null;
  fileInput.value = '';
  fileName.textContent = 'Klik untuk memilih file Excel';
  fileHelp.textContent = 'atau seret file ke area ini';
  setParseError('');
  resetPreview();
}

function setTarget(key) {
  if (!IMPORT_TARGETS[key] || key === currentTarget) return;
  currentTarget = key;
  targetRadios.forEach(radio => {
    radio.checked = radio.value === key;
    radio.closest('.target-option').classList.toggle('active', radio.checked);
  });
  document.querySelectorAll('.js-target-tab').forEach(el => { el.textContent = IMPORT_TARGETS[key]; });
  // File yang sudah dipilih milik tujuan sebelumnya: kosongkan agar tidak terkirim ke tab yang salah.
  resetFileSelection();
}

function lockTarget(locked) {
  targetRadios.forEach(radio => { radio.disabled = locked; });
  targetSwitch.classList.toggle('locked', locked);
}

function isBlank(value) {
  return value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

/*
 * TANGGAL EXCEL
 * Excel menyimpan tanggal sebagai angka serial (jumlah hari sejak 1899-12-30, pecahan = jam).
 * Angka itu langsung diubah menjadi teks kalender "yyyy-MM-dd HH:mm:ss" (jam dinding Excel).
 * Tidak ada objek Date lokal, tidak ada ISO/UTC, dan tidak ada zona waktu yang ikut campur,
 * jadi 19/10/2026 di Excel selalu menjadi 2026-10-19 di browser mana pun.
 *
 * Opsi cellDates milik SheetJS sengaja TIDAK dipakai: ia membuat Date lokal dengan koreksi offset zona waktu
 * tahun 1899 (yang punya selisih detik), sehingga 00:00:00 bisa terbaca 23:59:56 pada hari sebelumnya.
 */
let workbookDate1904 = false;

// Teks tampilan Excel yang berbentuk tanggal, mis. "19/10/2026", "19-10-26", "2026.10.19 07:30:00".
const DATE_TEXT_RE = /^\s*\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}(\s+\d{1,2}:\d{2}(:\d{2})?)?\s*$/;

// Sel angka dengan format tanggal/jam (mis. dd/mm/yyyy, d-mmm-yy, hh:mm) dianggap tanggal.
function isDateCell(cell) {
  if (!cell || cell.t !== 'n' || typeof cell.v !== 'number') return false;
  if (cell.z && cell.z !== 'General') {
    if (XLSX.SSF && typeof XLSX.SSF.is_date === 'function') return XLSX.SSF.is_date(cell.z);
    // Cadangan bila SSF tidak tersedia: buang teks dalam kutip/kurung, lalu cari kode tanggal/jam.
    return /[dmyhs]/i.test(String(cell.z).replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''));
  }
  // Format tidak diketahui: andalkan teks yang tampil di Excel (cell.w), misalnya "19/10/2026".
  return typeof cell.w === 'string' && DATE_TEXT_RE.test(cell.w) && cell.v > 0 && cell.v < 2958466;
}

// Statistik pembacaan terakhir, ditampilkan di bawah nama file sebagai bukti tanggal terbaca benar.
let lastDateStats = { count: 0, sample: null };

function describeDateStats() {
  if (!lastDateStats.count) return 'tanggal terdeteksi: 0 sel';
  const s = lastDateStats.sample;
  return 'tanggal terdeteksi: ' + lastDateStats.count.toLocaleString('id-ID') + ' sel (contoh: serial ' + s.serial + ' → ' + wallToDisplay({ __excelWall: s.wall }) + ')';
}

function excelSerialToValue(serial, date1904) {
  if (!Number.isFinite(serial) || serial < 0) return '';
  // Dibulatkan ke detik terdekat untuk membuang galat floating-point (mis. 46314.99999994 = 19/10/2026 00:00:00).
  const totalSeconds = Math.round(serial * 86400);
  const days = Math.floor(totalSeconds / 86400);
  const secs = totalSeconds - days * 86400;
  const time = pad2(Math.floor(secs / 3600)) + ':' + pad2(Math.floor((secs % 3600) / 60)) + ':' + pad2(secs % 60);
  if (days === 0) return time; // sel berisi jam saja (tanpa tanggal)
  // Date dipakai hanya sebagai kalkulator kalender UTC murni: hanya getUTC* yang dibaca.
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const d = new Date(base + days * 86400000);
  return {
    __excelWall: d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate()) + ' ' + time
  };
}

function cellToValue(cell) {
  if (!cell || cell.t === 'z' || cell.t === 'e') return '';
  if (isDateCell(cell)) {
    const value = excelSerialToValue(cell.v, workbookDate1904);
    if (value && typeof value === 'object') {
      lastDateStats.count++;
      if (!lastDateStats.sample) lastDateStats.sample = { serial: cell.v, wall: value.__excelWall };
    }
    return value;
  }
  return cell.v === null || cell.v === undefined ? '' : cell.v;
}

// Pengganti sheet_to_json: membaca sel satu per satu agar angka serial tanggal tidak berubah jadi Date.
function sheetToMatrix(worksheet) {
  const matrix = [];
  lastDateStats = { count: 0, sample: null };
  if (!worksheet['!ref']) return matrix;
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      row.push(cellToValue(worksheet[XLSX.utils.encode_cell({ r: r, c: c })]));
    }
    matrix.push(row);
  }
  return matrix;
}

// Tampilan untuk preview/header: tanggal murni dd/MM/yyyy, ditambah jam bila bukan 00:00:00.
function wallToDisplay(value) {
  if (value && typeof value === 'object' && typeof value.__excelWall === 'string') {
    const parts = value.__excelWall.split(' ');
    const ymd = parts[0].split('-');
    const time = parts[1] || '';
    return ymd[2] + '/' + ymd[1] + '/' + ymd[0] + (time && time !== '00:00:00' ? ' ' + time : '');
  }
  return value === null || value === undefined ? '' : String(value);
}

function makeCellValue(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') return value; // penanda tanggal { __excelWall }
  if (typeof value === 'number' && !Number.isFinite(value)) return '';
  if (typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number') return value;
  return String(value);
}

// Huruf kecil tanpa spasi/baris baru/tanda baca, supaya "CENTER\nID", "Center ID", dan "CENTER_ID" dianggap sama.
function normalizeHeaderText(value) {
  return String(value === null || value === undefined ? '' : value).toLowerCase().replace(/[^a-z0-9]/g, '');
}

// Pembaca khusus untuk Excel laporan: cari baris header otomatis, mulai dari kolom header paling kiri,
// dan lebarkan sampai ujung sel gabungan terakhir di baris header. Hasilnya matriks dengan header di indeks 0.
function sheetToMatrixByLayout(worksheet, layout) {
  lastDateStats = { count: 0, sample: null };
  if (!worksheet['!ref']) return { error: 'Worksheet kosong.' };
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const markers = layout.headerMarkers.map(normalizeHeaderText);
  const lastScanRow = Math.min(range.e.r, range.s.r + (layout.scanRows || 50) - 1);

  let headerRow = -1;
  let firstCol = -1;
  let lastCol = -1;
  for (let r = range.s.r; r <= lastScanRow && headerRow < 0; r++) {
    const seen = new Set();
    let first = -1;
    let last = -1;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: r, c: c })];
      const text = normalizeHeaderText(cell && cell.v);
      if (!text) continue;
      seen.add(text);
      if (first < 0) first = c;
      last = c;
    }
    if (markers.every(m => seen.has(m))) {
      headerRow = r;
      firstCol = first;
      lastCol = last;
    }
  }
  if (headerRow < 0) {
    return {
      error: 'Baris header tidak ditemukan. Untuk tujuan ini Excel harus punya kolom ' + layout.headerMarkers.join(', ') +
        ' pada ' + (layout.scanRows || 50) + ' baris pertama. Periksa juga pilihan worksheet.'
    };
  }

  // Header terakhir yang digabung (mis. OFFICERNAME di O:P) dihitung sampai ujung gabungannya.
  (worksheet['!merges'] || []).forEach(m => {
    if (m.s.r <= headerRow && headerRow <= m.e.r && m.s.c >= firstCol && m.s.c <= lastCol && m.e.c > lastCol) {
      lastCol = m.e.c;
    }
  });

  const matrix = [];
  for (let r = headerRow; r <= range.e.r; r++) {
    const row = [];
    for (let c = firstCol; c <= lastCol; c++) {
      // Sel gabungan: hanya sel kiri-atas yang berisi nilai di Excel, sel lainnya otomatis kosong.
      row.push(cellToValue(worksheet[XLSX.utils.encode_cell({ r: r, c: c })]));
    }
    matrix.push(row);
  }
  return {
    matrix: matrix,
    info: {
      headerRow: headerRow + 1,
      firstCol: XLSX.utils.encode_col(firstCol),
      lastCol: XLSX.utils.encode_col(lastCol)
    }
  };
}

function parseWorksheet(sheetName) {
  if (!workbook || !workbook.Sheets[sheetName]) return;
  const worksheet = workbook.Sheets[sheetName];
  const layout = TARGET_LAYOUTS[currentTarget] || null;
  let matrix;
  let layoutInfo = null;
  try {
    if (layout) {
      const result = sheetToMatrixByLayout(worksheet, layout);
      if (result.error) {
        setParseError(result.error);
        resetPreview();
        return;
      }
      matrix = result.matrix;
      layoutInfo = result.info;
    } else {
      matrix = sheetToMatrix(worksheet);
    }
  } catch (error) {
    console.error(error);
    setParseError('Worksheet tidak dapat dibaca. Pastikan file Excel tidak rusak atau dilindungi sandi.');
    resetPreview();
    return;
  }

  while (matrix.length && matrix[matrix.length - 1].every(isBlank)) matrix.pop();
  if (matrix.length < 2 || !Array.isArray(matrix[0])) {
    setParseError('Worksheet harus memiliki baris header dan minimal satu baris data.');
    resetPreview();
    return;
  }

  const headerRow = matrix[0].map(v => {
    const text = wallToDisplay(v).trim();
    return layout ? text.replace(/\s+/g, ' ') : text; // "CENTER\nID" -> "CENTER ID"
  });
  let lastHeaderIndex = headerRow.length - 1;
  // Tata letak khusus: kolom kosong di ujung kanan adalah bagian dari sel gabungan, jangan dibuang.
  while (!layout && lastHeaderIndex >= 0 && headerRow[lastHeaderIndex] === '') lastHeaderIndex--;
  if (lastHeaderIndex < 0) {
    setParseError('Baris pertama Excel tidak memiliki nama kolom.');
    resetPreview();
    return;
  }
  const headers = headerRow.slice(0, lastHeaderIndex + 1);
  const rows = matrix.slice(1).map(row => {
    const r = Array.isArray(row) ? row : [];
    return headers.map((_, index) => makeCellValue(r[index]));
  });
  while (rows.length && rows[rows.length - 1].every(isBlank)) rows.pop();

  if (!rows.length) {
    setParseError('Excel tidak memiliki baris data setelah header. Tidak ada data yang akan diimpor.');
    resetPreview();
    return;
  }
  if (rows.length > 25000) {
    setParseError('Jumlah baris data terlalu banyak. Batas yang disiapkan adalah 25.000 baris.');
    resetPreview();
    return;
  }
  if (headers.length > 100) {
    setParseError('Jumlah kolom melebihi batas 100 kolom.');
    resetPreview();
    return;
  }

  selectedData = {
    sourceSheet: sheetName,
    headers,
    rows,
    rawRows: matrix.slice(0, Math.min(matrix.length, 9)),
    layoutInfo
  };
  setParseError('');
  renderPreview();
  setHidden(previewBlock, false);
  setHidden(importPanel, false);
  updateButtonState();
}

function renderPreview() {
  if (!selectedData) return;
  sheetNameLabel.textContent = selectedData.sourceSheet;
  rowCountLabel.textContent = selectedData.rows.length.toLocaleString('id-ID');
  colCountLabel.textContent = selectedData.headers.length.toLocaleString('id-ID');

  const tbody = document.createElement('tbody');
  const headerTr = document.createElement('tr');
  selectedData.headers.forEach(value => {
    const th = document.createElement('th');
    if (value === '' && selectedData.layoutInfo) {
      // Kolom sisa dari sel gabungan (merge): sengaja kosong, nilainya ada di kolom kiri.
      th.textContent = '↳';
      th.title = 'Bagian dari sel yang digabung (merge); nilainya ada di kolom sebelah kiri';
    } else {
      th.textContent = value;
    }
    headerTr.appendChild(th);
  });
  tbody.appendChild(headerTr);

  const shownRows = selectedData.rawRows.slice(1, 9);
  shownRows.forEach((row, index) => {
    const tr = document.createElement('tr');
    selectedData.headers.forEach((_, colIndex) => {
      const td = document.createElement('td');
      td.textContent = wallToDisplay(Array.isArray(row) ? row[colIndex] : '');
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  previewTable.replaceChildren(tbody);
  if (selectedData.layoutInfo) {
    const li = selectedData.layoutInfo;
    headerStatus.textContent = 'Header baris ' + li.headerRow + ' · kolom ' + li.firstCol + '–' + li.lastCol + ' → baris 1 Sheet';
  } else {
    headerStatus.textContent = 'Validasi saat impor';
  }
  headerStatus.className = 'small-tag neutral-tag';
}

async function handleFile(file) {
  resetPreview();
  setParseError('');
  if (!file) return;
  const allowed = /\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name);
  if (!allowed) {
    setParseError('Pilih file Excel berformat .xlsx, .xls, .xlsm, atau .xlsb.');
    return;
  }
  if (typeof XLSX === 'undefined') {
    setParseError('Pustaka pembaca Excel tidak berhasil dimuat. Periksa koneksi internet, lalu muat ulang halaman.');
    return;
  }
  fileName.textContent = file.name;
  fileHelp.textContent = (file.size / 1024).toLocaleString('id-ID', { maximumFractionDigits: 0 }) + ' KB · membaca file…';
  try {
    const buffer = await file.arrayBuffer();
    // cellNF: true -> setiap sel menyimpan format angkanya (cell.z) agar sel tanggal bisa dikenali.
    // cellDates sengaja tidak dipakai (lihat catatan "TANGGAL EXCEL" di atas).
    workbook = XLSX.read(buffer, { type: 'array', cellNF: true, cellFormula: false });
    workbookDate1904 = !!(workbook.Workbook && workbook.Workbook.WBProps && workbook.Workbook.WBProps.date1904);
    if (!workbook.SheetNames || !workbook.SheetNames.length) throw new Error('no-sheets');
    sheetPicker.replaceChildren();
    workbook.SheetNames.forEach(name => {
      const option = document.createElement('option');
      option.value = name;
      option.textContent = name;
      sheetPicker.appendChild(option);
    });
    setHidden(sheetPickerWrap, workbook.SheetNames.length < 2);
    sheetPicker.value = workbook.SheetNames[0];
    parseWorksheet(sheetPicker.value);
    fileHelp.textContent = (file.size / 1024).toLocaleString('id-ID', { maximumFractionDigits: 0 }) + ' KB · file siap diperiksa · ' +
      describeDateStats() + ' · app.js ' + APP_VERSION;
    if (!isEndpointConfigured()) {
      setResult('Satu konfigurasi belum diisi: edit APP_SCRIPT_URL di app.js dengan URL Web App Apps Script yang berakhiran /exec.', 'error');
    }
  } catch (error) {
    console.error(error);
    workbook = null;
    resetPreview();
    setParseError('File Excel tidak dapat dibaca. Pastikan file valid, bukan file yang rusak atau terkunci dengan sandi.');
    fileHelp.textContent = 'File gagal dibaca';
  }
}

function buildPayload() {
  const payload = {
    sourceSheet: selectedData.sourceSheet,
    headers: selectedData.headers,
    rows: selectedData.rows
  };
  const serialized = JSON.stringify(payload);
  if (serialized.length > MAX_PAYLOAD_CHARS) {
    throw new Error('Ukuran data setelah dibaca terlalu besar untuk satu kali pengiriman. Kurangi jumlah kolom/baris atau ukuran isi sel, lalu coba kembali.');
  }
  return serialized;
}

function submitImport() {
  if (!selectedData || isSubmitting) return;
  if (!isEndpointConfigured()) {
    setResult('URL Apps Script belum dikonfigurasi. Isi APP_SCRIPT_URL di app.js terlebih dahulu.', 'error');
    return;
  }
  if (!confirmCheckbox.checked) {
    setResult('Centang konfirmasi penggantian data sebelum melanjutkan.', 'error');
    return;
  }
  if (!importPassword.value) {
    setResult('Masukkan kata sandi impor.', 'error');
    return;
  }

  let serialized;
  try {
    serialized = buildPayload();
  } catch (error) {
    setResult(error.message, 'error');
    return;
  }

  const form = document.createElement('form');
  form.method = 'POST';
  form.action = APP_SCRIPT_URL;
  form.target = 'gasResultFrame';
  form.acceptCharset = 'UTF-8';
  form.style.display = 'none';
  const fields = {
    action: 'importExcel',
    target: currentTarget,
    confirm: 'REPLACE',
    password: importPassword.value,
    payload: serialized
  };
  Object.entries(fields).forEach(([key, value]) => {
    const input = document.createElement('input');
    input.type = 'hidden';
    input.name = key;
    input.value = value;
    form.appendChild(input);
  });

  document.body.appendChild(form);
  isSubmitting = true;
  lockTarget(true);
  importButton.textContent = 'Mengirim dan mengimpor…';
  updateButtonState();
  setResult('Permintaan sudah dikirim. Jangan tutup halaman sampai hasil impor ditampilkan.', 'loading');
  form.submit();
  window.setTimeout(() => form.remove(), 5000);
  // Jika respons tidak tiba (misalnya deployment/izin salah), jangan biarkan UI macet permanen.
  window.setTimeout(() => {
    if (isSubmitting) {
      isSubmitting = false;
      lockTarget(false);
      importButton.innerHTML = '<span class="button-icon">↥</span> Impor ke Google Sheets';
      updateButtonState();
      setResult('Belum menerima konfirmasi dari Apps Script. Periksa URL /exec, izin deployment, dan menu Executions di Apps Script sebelum mencoba ulang.', 'error');
    }
  }, 45000);
}

fileInput.addEventListener('change', event => handleFile(event.target.files && event.target.files[0]));
sheetPicker.addEventListener('change', () => parseWorksheet(sheetPicker.value));
targetRadios.forEach(radio => radio.addEventListener('change', () => setTarget(radio.value)));
confirmCheckbox.addEventListener('change', updateButtonState);
importPassword.addEventListener('input', updateButtonState);
importButton.addEventListener('click', submitImport);

dropZone.addEventListener('keydown', event => {
  if (event.key === 'Enter' || event.key === ' ') {
    event.preventDefault();
    fileInput.click();
  }
});
dropZone.addEventListener('dragover', event => {
  event.preventDefault();
  dropZone.classList.add('dragging');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragging'));
dropZone.addEventListener('drop', event => {
  event.preventDefault();
  dropZone.classList.remove('dragging');
  const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
  if (file) handleFile(file);
});

window.addEventListener('message', event => {
  const data = event.data;
  if (!data || data.type !== 'EXCEL_IMPORT_RESULT' || !data.payload) return;
  const result = data.payload;
  isSubmitting = false;
  lockTarget(false);
  importButton.innerHTML = '<span class="button-icon">↥</span> Impor ke Google Sheets';
  if (result.ok) {
    importPassword.value = '';
    confirmCheckbox.checked = false;
    updateButtonState();
    const details = ' ' + Number(result.rows || 0).toLocaleString('id-ID') + ' baris dan ' + Number(result.columns || 0).toLocaleString('id-ID') + ' kolom berhasil diproses.';
    setResult((result.message || 'Impor berhasil.') + details, 'success');
  } else {
    updateButtonState();
    setResult(result.message || 'Impor gagal. Data tidak dapat dipastikan telah diperbarui.', 'error');
  }
});

updateButtonState();
