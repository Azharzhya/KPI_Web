/*
 * Konfigurasi frontend.
 * Isi APP_SCRIPT_URL dengan URL Web App Google Apps Script berakhiran /exec.
 */
const APP_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby-mdsr_0Z9xBJnR0votH2k9ceh11vbGkjghmpBZGwzXRtFKx_S8gZWcLKzWZIOKAJOZA/exec';
const MAX_PAYLOAD_CHARS = 7 * 1024 * 1024;

// Tujuan impor. Kuncinya dikirim ke server; nama tab di sini hanya untuk tampilan (server punya daftar sendiri).
const IMPORT_TARGETS = {
  anggota: 'Update_Anggota',
  disbursement: 'Update_Disbursement'
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

function makeCellValue(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return { __excelDate: value.toISOString() };
  if (typeof value === 'number' && !Number.isFinite(value)) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'boolean' || typeof value === 'number') return value;
  return String(value);
}

function parseWorksheet(sheetName) {
  if (!workbook || !workbook.Sheets[sheetName]) return;
  const worksheet = workbook.Sheets[sheetName];
  let matrix;
  try {
    matrix = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: '',
      raw: true,
      blankrows: true
    });
  } catch (error) {
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

  const headerRow = matrix[0].map(v => v == null ? '' : String(v).trim());
  let lastHeaderIndex = headerRow.length - 1;
  while (lastHeaderIndex >= 0 && headerRow[lastHeaderIndex] === '') lastHeaderIndex--;
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
    rawRows: matrix.slice(0, Math.min(matrix.length, 9))
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
    th.textContent = value;
    headerTr.appendChild(th);
  });
  tbody.appendChild(headerTr);

  const shownRows = selectedData.rawRows.slice(1, 9);
  shownRows.forEach((row, index) => {
    const tr = document.createElement('tr');
    selectedData.headers.forEach((_, colIndex) => {
      const td = document.createElement('td');
      let value = Array.isArray(row) ? row[colIndex] : '';
      if (value instanceof Date) value = value.toLocaleDateString('id-ID');
      if (value && typeof value === 'object' && value.__excelDate) {
        const parsed = new Date(value.__excelDate);
        value = Number.isNaN(parsed.getTime()) ? '' : parsed.toLocaleDateString('id-ID');
      }
      td.textContent = value == null ? '' : String(value);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  previewTable.replaceChildren(tbody);
  headerStatus.textContent = 'Validasi saat impor';
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
    workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellFormula: false });
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
    fileHelp.textContent = (file.size / 1024).toLocaleString('id-ID', { maximumFractionDigits: 0 }) + ' KB · file siap diperiksa';
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
