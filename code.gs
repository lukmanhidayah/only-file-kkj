function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Upload File UKKJ');
}

// Helper function to get columns that need single quote prefix
function getSingleQuotePrefixColumns() {
  // Kolom yang membutuhkan tanda kutip satu (0-based index):
  // No. Sertifikat, Tgl. Sertifikat, NIP, TMT Kenaikan Pangkat Terakhir, Nilai PAK Konversi Terakhir
  return new Set([4, 5, 6, 10, 15]);
}

// Helper function to process value with single quote prefix if needed
function processValueWithQuotePrefix(value, columnIndex) {
  const SINGLE_QUOTE_PREFIX_COLUMNS = getSingleQuotePrefixColumns();
  
  if (!SINGLE_QUOTE_PREFIX_COLUMNS.has(columnIndex)) {
    return value;
  }
  
  // Handle null or undefined
  if (value === null || value === undefined) {
    return '';
  }
  
  // Convert to string
  let stringValue = value.toString().trim();
  
  // If empty, return as is
  if (!stringValue) {
    return stringValue;
  }
  
  // Add single quote prefix if it doesn't already exist
  if (!stringValue.startsWith("'")) {
    stringValue = "'" + stringValue;
  }
  
  return stringValue;
}

function addRow(data) {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');
  
  // Process data to add single quote prefix where needed
  const processedData = data.map((value, index) => {
    return processValueWithQuotePrefix(value, index);
  });
  
  sh.appendRow(processedData);
  return 'Row added';
}

function updateRow(rowIndex, values) {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');
  
  // Process values to add single quote prefix where needed
  const processedValues = values.map((value, index) => {
    return processValueWithQuotePrefix(value, index);
  });
  
  sh.getRange(rowIndex, 1, 1, processedValues.length).setValues([processedValues]);
  return 'Row updated';
}

function deleteRow(rowIndex) {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');
  sh.deleteRow(rowIndex);
  return 'Row deleted';
}

function updateLastColumnCheckbox(rowIndex, isChecked) {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');
  
  // Dapatkan indeks kolom terakhir secara dinamis
  const lastColumn = sh.getLastColumn();
  
  // Dapatkan sel di baris dan kolom yang tepat
  const cell = sh.getRange(rowIndex, lastColumn);
  
  // Set nilai sel menjadi true atau false
  cell.setValue(isChecked);
  
  return 'Status updated';
}

function getData() {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');
  const values = sh.getDataRange().getValues();
  return values;
}



function updateKeterangan(rowIndex, keteranganText) {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');
  
  // Mendapatkan indeks kolom terakhir secara dinamis
  const lastColumn = sh.getLastColumn();
  
  // Menargetkan sel di baris yang diberikan dan kolom terakhir
  const cell = sh.getRange(rowIndex, lastColumn);
  cell.setValue(keteranganText);
  
  return 'Status updated';
}


function updateRowFields(rowIndex, updates) {
  if (!updates || typeof updates !== 'object') {
    return 'No updates';
  }

  const rowNumber = Number(rowIndex);
  if (!rowNumber) {
    return 'No updates';
  }

  const entries = Object.entries(updates);
  if (!entries.length) {
    return 'No updates';
  }

  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');

  // First, process the direct updates
  entries.forEach(([colIndexStr, rawValue]) => {
    const colIndex = Number(colIndexStr);
    if (isNaN(colIndex)) {
      return;
    }

    let finalValue = rawValue;
    if (finalValue === null || finalValue === undefined) {
      finalValue = '';
    } else if (typeof finalValue === 'string') {
      const trimmed = finalValue.trim();
      if (trimmed === '') {
        finalValue = '';
      } else {
        const parsedDate = tryParseISODate(trimmed);
        finalValue = parsedDate || trimmed;
      }
    }

    // Apply single quote prefix using helper function
    finalValue = processValueWithQuotePrefix(finalValue, colIndex);

    sh.getRange(rowNumber, colIndex + 1).setValue(finalValue);
  });

  // Additional step: Ensure ALL columns that need single quotes have them
  // Get current row data
  const lastColumn = sh.getLastColumn();
  const rowData = sh.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
  
  // Check and fix single quote prefix for all relevant columns
  const quotePrefixColumns = getSingleQuotePrefixColumns();
  let hasFixedColumns = false;
  
  quotePrefixColumns.forEach(colIndex => {
    if (colIndex < rowData.length) {
      const currentValue = rowData[colIndex];
      const processedValue = processValueWithQuotePrefix(currentValue, colIndex);
      
      // Only update if the value changed (to avoid unnecessary writes)
      if (currentValue !== processedValue) {
        sh.getRange(rowNumber, colIndex + 1).setValue(processedValue);
        hasFixedColumns = true;
      }
    }
  });

  return hasFixedColumns ? 'Row fields updated and quote prefixes fixed' : 'Row fields updated';
}

function tryParseISODate(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  const date = new Date(year, month, day);
  if (isNaN(date.getTime())) {
    return null;
  }
  return date;
}

/**
 * Cek apakah file Google Drive dengan fileId ada & dapat diakses oleh skrip ini.
 * Return: true jika ada, false kalau tidak atau tidak punya akses.
 */
function fileExists(fileId) {
  try {
    // Jika tidak ada/akses ditolak, getFileById akan throw
    var file = DriveApp.getFileById(fileId);
    // Bisa tambahkan cek tambahan (mis. ukuran > 0), tapi cukup return true
    return !!file && !!file.getId();
  } catch (e) {
    return false;
  }
}

function getKeteranganData(rowIndex) {
  const sheet = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA').getSheetByName('DBClear');
  const keteranganValue = sheet.getRange(rowIndex, 32).getValue(); // Column 32 is Keterangan (index 31)
  
  return keteranganValue;
}

const COMMON_FORM_PREFILL_MAPPING = [
  { entry: 'entry.1754826825', columnIndex: 5 },
  { entry: 'entry.1866181824', columnIndex: 6 },
  { entry: 'entry.1857305381', columnIndex: 7 },
  { entry: 'entry.1787474135', columnIndex: 8 },
  { entry: 'entry.1638657637', columnIndex: 9 },
  { entry: 'entry.2111946046', columnIndex: 10 },
  { entry: 'entry.706142243', columnIndex: 11 },
  { entry: 'entry.1023101296', columnIndex: 12 },
  { entry: 'entry.1046224612', columnIndex: 13 },
  { entry: 'entry.644211122', columnIndex: 14 },
  { entry: 'entry.1001936770', columnIndex: 15 },
  { entry: 'entry.364450287', columnIndex: 16 },
  { entry: 'entry.1952150778', columnIndex: 17 }
];

const HINDU_PREFILL_MAPPING = [
  { entry: 'entry.234611057', columnIndex: 5 },
  { entry: 'entry.1301652810', columnIndex: 6, type: 'date' },
  { entry: 'entry.544587430', columnIndex: 7 },
  { entry: 'entry.251716686', columnIndex: 8 },
  { entry: 'entry.1262554733', columnIndex: 9 },
  { entry: 'entry.763422945', columnIndex: 10 },
  { entry: 'entry.2109867146', columnIndex: 11, type: 'date' },
  { entry: 'entry.1586430140', columnIndex: 12 },
  { entry: 'entry.1513596288', columnIndex: 13 },
  { entry: 'entry.363414788', columnIndex: 14 },
  { entry: 'entry.707422534', columnIndex: 15 },
  { entry: 'entry.244763738', columnIndex: 16 },
  { entry: 'entry.1516891442', columnIndex: 17 },
];

const BUDDHA_PREFILL_MAPPING = [
  { entry: 'entry.1257097526', columnIndex: 5 },
  { entry: 'entry.1504592986', columnIndex: 6 },
  { entry: 'entry.2066527703', columnIndex: 7 },
  { entry: 'entry.952330150', columnIndex: 8 },
  { entry: 'entry.1072933178', columnIndex: 9 },
  { entry: 'entry.1811816817', columnIndex: 10 },
  { entry: 'entry.1419661682', columnIndex: 11 },
  { entry: 'entry.297392157', columnIndex: 12 },
  { entry: 'entry.249993333', columnIndex: 13 },
  { entry: 'entry.2137003292', columnIndex: 14 },
  { entry: 'entry.124255052', columnIndex: 15 },
  { entry: 'entry.480417322', columnIndex: 16 },
  { entry: 'entry.6783364', columnIndex: 17 },
];

const KRISTEN_PREFILL_MAPPING = [
  { entry: 'entry.399672719', headers: ['Nama Lengkap (Berikut Gelar)\nPenulisan Nama memperhatikan huruf kecil dan huruf kapital, Contoh:  Farah Khoirunnisaa, S.Pd'], type: 'uppercase' },
  { entry: 'entry.651742535', headers: ['NIP\nPengisian diawali dengan tanda kutip, Contoh:\n\'197605022005011005'] },
  { entry: 'entry.1379892291', headers: ['Status Sertifikat LULUS UKKJ'] },
  { entry: 'entry.1923577468', headers: ['Nomor Sertifikat Uji Kompetensi Kenaikan Jenjang Jabatan Fungsional Guru.\nPengisian diawali dengan tanda kutip, Contoh:\n\'00033/04.01/MM/XII/2024'] },
  { entry: 'entry.227278811', headers: ['Unit Kerja\n(Penulisan unit kerja tidak boleh disingkat), Contoh:\nKantor Kementerian Agama Kota Bandung'] },
  { entry: 'entry.216579854', headers: ['Unit Kerja\n(Penulisan unit kerja tidak boleh disingkat), Contoh:\nKantor Kementerian Agama Kota Bandung'], type: 'extractKabKota' },
  { entry: 'entry.811694245', headers: ['Provinsi\nPenulisan diawali dengan kata provinsi, Contoh:\nProvinsi Jawa Barat'] },
  { entry: 'entry.1176655894', headers: ['Satuan Kerja\n(Penulisan satuan kerja tidak boleh disingkat), Contoh:\nMadrasah Aliyah Negeri 1 Kota Bandung'] },
];


const KATOLIK_PREFILL_MAPPING = [
  { entry: 'entry.1182555051', headers: ['Nomor Sertifikat Uji Kompetensi Kenaikan Jenjang Jabatan Fungsional Guru.\nPengisian diawali dengan tanda kutip, Contoh:\n\'00033/04.01/MM/XII/2024'] },
  { entry: 'entry.1117137670', headers: ['Tanggal  Sertifikat Uji Kompetensi Kenaikan Jenjang  Jabatan Fungsional Guru.\nPengisian diawali dengan tanda kutip, Contoh: \n\'17 Desember 2024'] },
  { entry: 'entry.2074483170', headers: ['NIP\nPengisian diawali dengan tanda kutip, Contoh:\n\'197605022005011005'] },
  { entry: 'entry.123245688', headers: ['Nama Lengkap (Berikut Gelar)\nPenulisan Nama memperhatikan huruf kecil dan huruf kapital, Contoh:  Farah Khoirunnisaa, S.Pd'] },
  { entry: 'entry.2105303303', headers: ['Pangkat'] },
  { entry: 'entry.1753706123', headers: ['Golongan/Ruang'] },
  { entry: 'entry.754677027', headers: ['TMT Kenaikan Pangkat Terakhir\nPengisian diawali dengan tanda kutip, Contoh:\n\'1 April 2018'] },
  { entry: 'entry.1453227642', headers: ['Satuan Kerja\n(Penulisan satuan kerja tidak boleh disingkat), Contoh:\nMadrasah Aliyah Negeri 1 Kota Bandung'] },
  { entry: 'entry.1275958833', headers: ['Unit Kerja\n(Penulisan unit kerja tidak boleh disingkat), Contoh:\nKantor Kementerian Agama Kota Bandung'] },
  { entry: 'entry.1240111494', headers: ['Provinsi\nPenulisan diawali dengan kata provinsi, Contoh:\nProvinsi Jawa Barat'] },
  { entry: 'entry.1688447553', headers: ['Jabatan Lama'] },
  { entry: 'entry.1481269566', headers: ['Nilai PAK  Konversi Terakhir  (Penetapan Angka Kredit) \nPengisian diawali dengan tanda kutip, Contoh: \n\'284,923'] },
  { entry: 'entry.71817823', headers: ['Jabatan Baru'] },
];

const PAI_PREFILL_MAPPING = [
  { entry: 'entry.2140303337', columnIndex: 5 },
  { entry: 'entry.1669319240', columnIndex: 6 },
  { entry: 'entry.1959212525', columnIndex: 7 },
  { entry: 'entry.811377439', columnIndex: 8 },
  { entry: 'entry.811554715', columnIndex: 9 },
  { entry: 'entry.441346566', columnIndex: 10 },
  { entry: 'entry.1359414804', columnIndex: 11 },
  { entry: 'entry.1236549644', columnIndex: 12 },
  { entry: 'entry.2102927956', columnIndex: 13 },
  { entry: 'entry.933730319', columnIndex: 14 },
  { entry: 'entry.1608900927', columnIndex: 15 },
  { entry: 'entry.5802051', columnIndex: 16 },
  { entry: 'entry.2146081769', columnIndex: 17 },
];

const FORM_CONFIGS = {
  pendisGuruMadrasah: {
    label: 'Link Pendis Guru Madrasah',
    baseUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSd3eHdh50pX7awV7tJ9pRj84ik7MwRXlVPWEmxwSLLOjjrrNQ/viewform',
    mapping: COMMON_FORM_PREFILL_MAPPING,
  },
  pendisGuruPai: {
    label: 'Link Pendis Guru PAI',
    baseUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSeHlMDmPURIXer6BDzvDjWi2d7tLxS36ja2nJyF831M5SVL5w/viewform',
    mapping: PAI_PREFILL_MAPPING,
  },
  kristen: {
    label: 'Link Kristen',
    baseUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSegMFQYpluo8ephxSeSnmBQs1I8e2MtAVE6eQGFfRj4XeQQaA/viewform',
    mapping: KRISTEN_PREFILL_MAPPING,
  },
  katolik: {
    label: 'Link Katolik',
    baseUrl: 'https://docs.google.com/forms/d/1oPGwdvnYktOaeKQroPKX9STipnlNL52UvqU2s_N9r_0/viewform',
    mapping: KATOLIK_PREFILL_MAPPING,
  },
  hindu: {
    label: 'Link Hindu',
    baseUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSdcicvfjkyWFSNRfJF1B2eLwBP2OEF3IlUYzDNkVciFuMDI4Q/viewform',
    mapping: HINDU_PREFILL_MAPPING,
  },
  buddha: {
    label: 'Link Buddha',
    baseUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSc9Eu2pW1KhpPxgLeA5G7sMWHYJyJ1Fvr9u706-QNPlL__H6Q/viewform',
    mapping: BUDDHA_PREFILL_MAPPING,
  },
};

const DEFAULT_FORM_KEY = 'pendisGuruMadrasah';
const FORM_PREFILL_OUTPUT_HEADER = 'Prefilled Form URL';

function generatePrefilledFormUrl(rowNumber, formKey) {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');
  const lastColumn = sh.getLastColumn();
  if (rowNumber < 2 || rowNumber > sh.getLastRow()) {
    throw new Error('Row number out of range.');
  }

  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  const rowValues = sh.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
  return buildPrefilledFormUrl(headers, rowValues, formKey);
}

function syncPrefilledFormUrls(formKey) {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) {
    return 'No data rows to sync';
  }

  const lastColumn = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  const data = sh.getRange(2, 1, lastRow - 1, lastColumn).getValues();

  let outputColumnIndex = headers.findIndex((header) => normalizeHeader(header) === normalizeHeader(FORM_PREFILL_OUTPUT_HEADER));
  let outputColumn = outputColumnIndex + 1;
  if (outputColumnIndex === -1) {
    outputColumn = lastColumn + 1;
    sh.getRange(1, outputColumn).setValue(FORM_PREFILL_OUTPUT_HEADER);
  }

  const urls = data.map((rowValues) => [buildPrefilledFormUrl(headers, rowValues, formKey)]);
  sh.getRange(2, outputColumn, urls.length, 1).setValues(urls);
  return 'Prefilled URLs updated';
}

function syncPrefillForRow(rowNumber, formKey) {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');
  const lastRow = sh.getLastRow();
  if (rowNumber < 2 || rowNumber > lastRow) {
    return { url: '', message: 'Baris tidak ditemukan.' };
  }

  const formConfig = getFormConfig(formKey);
  if (!formConfig) {
    return { url: '', message: 'Jenis form tidak dikenali.' };
  }

  const lastColumn = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  const rowValues = sh.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
  const url = buildPrefilledFormUrl(headers, rowValues, formConfig.key);

  let outputColumnIndex = headers.findIndex((header) => normalizeHeader(header) === normalizeHeader(FORM_PREFILL_OUTPUT_HEADER));
  let outputColumn = outputColumnIndex + 1;
  if (outputColumnIndex === -1) {
    outputColumn = lastColumn + 1;
    sh.getRange(1, outputColumn).setValue(FORM_PREFILL_OUTPUT_HEADER);
  }

  sh.getRange(rowNumber, outputColumn).setValue(url);

  const message = url ? 'Prefilled URL berhasil dibuat.' : 'Data tidak lengkap untuk membuat tautan prefill.';
  return { url, message, formKey: formConfig.key };
}

function buildPrefilledFormUrl(headers, rowValues, formKey) {
  const formConfig = getFormConfig(formKey);
  if (!formConfig || !formConfig.baseUrl) {
    return '';
  }

  const mapping = formConfig.mapping || [];
  const params = [];

  mapping.forEach((field) => {
    const columnIndex = resolveColumnIndex(field, headers);
    if (columnIndex < 0 || columnIndex >= rowValues.length) {
      return;
    }

    const rawValue = rowValues[columnIndex];

    if (field.type === 'date') {
      const parts = extractDateParts(rawValue);
      if (!parts) {
        return;
      }

      params.push(`${field.entry}_year=${encodeURIComponent(parts.year)}`);
      params.push(`${field.entry}_month=${encodeURIComponent(parts.month)}`);
      params.push(`${field.entry}_day=${encodeURIComponent(parts.day)}`);

      if (field.includeFormattedValue) {
        const formatted = formatValueForPrefillWithPrefix(rawValue, columnIndex, headers);
        if (formatted) {
          params.push(`${field.entry}=${encodeURIComponent(formatted)}`);
        }
      }
      return;
    }

    // Handle uppercase
    if(field.type === 'uppercase') {
      const value = formatValueForPrefillWithPrefix(rawValue, columnIndex, headers);
      if (!value) {
        return;
      }
      
      // Convert to uppercase (preserving single quote prefix if present)
      let uppercaseValue = value.toUpperCase();
      params.push(`${field.entry}=${encodeURIComponent(uppercaseValue)}`);
      return;
    }

    // Handle ekstraksi kabupaten/kota dari unit kerja
    if (field.type === 'extractKabKota') {
      const extractedValue = extractKabKotaFromUnitKerja(rawValue);
      if (!extractedValue) {
        return;
      }
      params.push(`${field.entry}=${encodeURIComponent(extractedValue)}`);
      return;
    }

    // Use the new function that handles single quote prefix
    const value = formatValueForPrefillWithPrefix(rawValue, columnIndex, headers);
    if (!value) {
      return;
    }

    params.push(`${field.entry}=${encodeURIComponent(value)}`);
  });

  if (!params.length) {
    return '';
  }

  const baseUrl = normalizeFormBaseUrl(formConfig.baseUrl);
  if (!baseUrl) {
    return '';
  }

  const separator = baseUrl.indexOf('?') === -1 ? '?' : '&';
  return `${baseUrl}${separator}usp=pp_url&${params.join('&')}`;
}

function getFormConfig(formKey) {
  const isMissingKey = formKey === undefined || formKey === null || formKey === '';
  let selectedKey = DEFAULT_FORM_KEY;

  if (!isMissingKey) {
    if (Object.prototype.hasOwnProperty.call(FORM_CONFIGS, formKey)) {
      selectedKey = formKey;
    } else {
      return null;
    }
  }

  const config = FORM_CONFIGS[selectedKey];
  if (!config) {
    return null;
  }

  return {
    key: selectedKey,
    label: config.label,
    baseUrl: config.baseUrl,
    mapping: config.mapping,
  };
}

function normalizeFormBaseUrl(rawUrl) {
  if (!rawUrl) {
    return '';
  }

  let url = rawUrl.trim();
  url = url.replace(/(\?|&)edit_requested=true/gi, '');
  url = url.replace(/(\?|&)usp=pp_url/gi, '');
  url = url.replace(/[?&]+$/, '');
  return url;
}

function resolveColumnIndex(field, headers) {
  if (field.columnIndex) {
    return field.columnIndex - 1;
  }

  const candidates = [];

  if (Array.isArray(field.headers)) {
    candidates.push(...field.headers);
  }

  if (field.header !== undefined && field.header !== null) {
    if (Array.isArray(field.header)) {
      candidates.push(...field.header);
    } else {
      candidates.push(field.header);
    }
  }

  if (!candidates.length) {
    return -1;
  }

  for (let i = 0; i < candidates.length; i += 1) {
    const candidate = candidates[i];
    const normalizedTarget = normalizeHeader(candidate);
    if (!normalizedTarget) {
      continue;
    }

    const resolvedIndex = headers.findIndex(
      (header) => normalizeHeader(header) === normalizedTarget,
    );
    if (resolvedIndex !== -1) {
      return resolvedIndex;
    }
  }

  return -1;
}


function normalizeHeader(value) {
  return value ? value.toString().replace(/\s+/g, ' ').trim().toLowerCase() : '';
}

function extractDateParts(rawValue) {
  const date = coerceDateValue(rawValue);
  if (!date) {
    return null;
  }

  return {
    year: date.getFullYear(),
    month: date.getMonth() + 1,
    day: date.getDate(),
  };
}

function coerceDateValue(rawValue) {
  if (rawValue instanceof Date) {
    return Number.isNaN(rawValue.getTime()) ? null : rawValue;
  }

  if (typeof rawValue === 'number' && !Number.isNaN(rawValue)) {
    const numericDate = new Date(rawValue);
    return Number.isNaN(numericDate.getTime()) ? null : numericDate;
  }

  if (typeof rawValue === 'string') {
    const trimmed = rawValue.trim();
    if (!trimmed) {
      return null;
    }

    const iso = tryParseISODate(trimmed);
    if (iso) {
      return iso;
    }

    const parsed = new Date(trimmed);
    if (!Number.isNaN(parsed.getTime())) {
      return parsed;
    }

    const dmyMatch = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/.exec(trimmed);
    if (dmyMatch) {
      const day = Number(dmyMatch[1]);
      const month = Number(dmyMatch[2]) - 1;
      let year = Number(dmyMatch[3]);

      if (Number.isNaN(day) || Number.isNaN(month) || Number.isNaN(year)) {
        return null;
      }

      if (year < 100) {
        year += year >= 70 ? 1900 : 2000;
      }

      const dmyDate = new Date(year, month, day);
      if (!Number.isNaN(dmyDate.getTime())) {
        return dmyDate;
      }
    }
  }

  return null;
}

function formatValueForPrefill(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'd MMMM yyyy');
  }

  if (typeof value === 'number') {
    return value.toString();
  }

  const str = value.toString().trim();
  return str;
}

// New function to handle single quote prefix for specific columns
function formatValueForPrefillWithPrefix(value, columnIndex, headers) {
  let formattedValue = formatValueForPrefill(value);
  
  // Apply single quote prefix using the centralized helper function
  formattedValue = processValueWithQuotePrefix(formattedValue, columnIndex);
  
  return formattedValue;
}

/**
 * Function to fix single quote prefix for all existing data in the spreadsheet
 * This ensures consistency across all rows
 */
function fixAllSingleQuotePrefixes() {
  const ss = SpreadsheetApp.openById('1DZbPotFtHoPWVnG6N3kk1J8XeOmwUo_K1v-ZXBpbVrA');
  const sh = ss.getSheetByName('DBClear');
  
  const lastRow = sh.getLastRow();
  const lastColumn = sh.getLastColumn();
  
  if (lastRow <= 1) {
    return 'No data rows to fix';
  }
  
  // Get all data (excluding header row)
  const data = sh.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const quotePrefixColumns = getSingleQuotePrefixColumns();
  
  let fixedCount = 0;
  
  // Process each row
  data.forEach((row, rowIndex) => {
    let rowNeedsUpdate = false;
    const updatedRow = row.map((value, colIndex) => {
      const processedValue = processValueWithQuotePrefix(value, colIndex);
      if (value !== processedValue) {
        rowNeedsUpdate = true;
      }
      return processedValue;
    });
    
    // Update the row if any changes were made
    if (rowNeedsUpdate) {
      sh.getRange(rowIndex + 2, 1, 1, lastColumn).setValues([updatedRow]);
      fixedCount++;
    }
  });
  
  return `Fixed single quote prefixes in ${fixedCount} rows`;
}

function extractKabKotaFromUnitKerja(unitKerjaText) {
  if (!unitKerjaText || typeof unitKerjaText !== 'string') {
    return '';
  }
  
  const text = unitKerjaText.trim();
  
  // Pattern untuk mengekstrak nama kabupaten/kota dari format:
  // "Kantor Kementerian Agama Kota [Nama]" atau "Kantor Kementerian Agama Kabupaten [Nama]"
  const kotaPattern = /Kantor Kementerian Agama (Kota\s+.+)/i;
  const kabupatenPattern = /Kantor Kementerian Agama (Kabupaten\s+.+)/i;
  
  let match = text.match(kotaPattern);
  if (match) {
    return match[1].trim();
  }
  
  match = text.match(kabupatenPattern);
  if (match) {
    return match[1].trim();
  }
  
  // Jika tidak match dengan pattern, return empty string
  return '';
}

/**
 * Upload file to Google Drive - FIXED VERSION
 * This function should be added to your Google Apps Script project (code.gs)
 */
function uploadFileToDrive(fileData) {
  try {
    console.log('Starting file upload:', fileData.name);
    
    // Validate input
    if (!fileData || !fileData.content || !fileData.folderId) {
      throw new Error('Invalid file data provided');
    }
    
    // Decode base64 content
    const fileBlob = Utilities.newBlob(
      Utilities.base64Decode(fileData.content),
      fileData.mimeType || 'application/octet-stream',
      fileData.name
    );
    
    console.log('File blob created, size:', fileBlob.getBytes().length);
    
    // Get the target folder
    const folder = DriveApp.getFolderById(fileData.folderId);
    console.log('Target folder found:', folder.getName());
    
    // Create the file in the folder
    const file = folder.createFile(fileBlob);
    console.log('File created with ID:', file.getId());
    
    // Set file name (in case it needs to be cleaned up)
    file.setName(fileData.name);
    
    // Make file viewable by anyone with the link (optional)
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      console.log('File sharing set to view-only');
    } catch (shareError) {
      console.warn('Could not set file sharing:', shareError.toString());
      // Continue without sharing settings
    }
    
    const fileUrl = file.getUrl();
    console.log('File uploaded successfully. URL:', fileUrl);
    
    // Return the file URL - IMPORTANT: Return just the URL string for simplicity
    return fileUrl;
    
  } catch (error) {
    console.error('Error uploading file:', error.toString());
    console.error('Stack trace:', error.stack);
    
    // Re-throw the error so it's caught by the frontend error handler
    throw new Error('Upload failed: ' + error.toString());
  }
}

/**
 * Alternative version that returns detailed object
 * Use this if you need more detailed response
 */
function uploadFileToDriveDetailed(fileData) {
  try {
    console.log('Starting detailed file upload:', fileData.name);
    
    if (!fileData || !fileData.content || !fileData.folderId) {
      throw new Error('Invalid file data provided');
    }
    
    const fileBlob = Utilities.newBlob(
      Utilities.base64Decode(fileData.content),
      fileData.mimeType || 'application/octet-stream',
      fileData.name
    );
    
    const folder = DriveApp.getFolderById(fileData.folderId);
    const file = folder.createFile(fileBlob);
    file.setName(fileData.name);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareError) {
      console.warn('Could not set file sharing:', shareError.toString());
    }
    
    const result = {
      success: true,
      url: file.getUrl(),
      fileId: file.getId(),
      fileName: file.getName(),
      fileSize: file.getSize(),
      mimeType: file.getBlob().getContentType(),
      createdDate: file.getDateCreated(),
      message: 'File uploaded successfully'
    };
    
    console.log('Detailed upload result:', JSON.stringify(result));
    return result;
    
  } catch (error) {
    console.error('Detailed upload error:', error.toString());
    
    const errorResult = {
      success: false,
      error: error.toString(),
      message: 'Failed to upload file',
      timestamp: new Date().toISOString()
    };
    
    console.log('Error result:', JSON.stringify(errorResult));
    return errorResult;
  }
}

/**
 * Test function to verify folder IDs are correct
 * Run this function in Apps Script to test each folder
 */
function testFolderAccess() {
  const folderIds = {
    'Sertifikat UKKJ': '1_EuHAesCw0pQe2ANpBhuC0uWWPLs_yhGKII6tBhtBAQw-3Q-IzCJpiarIjb6wR_oyln1TDdF',
    'Surat Pengantar': '1-purrtZPyKDSqmXNmG1-RKq__PNkJQCZBjWX4kI_-UQmuHeIZL-iExXLTTB9Mdhj8GOmnRpv',
    'SK CPNS': '1DaJ1kmf2UBNkmBNjrDrdHy3ivkbcYRey1WctZIht7MFebootADwrIt_hkm31KHH49v-8Snoe',
    'PAK Konversi': '1CKKIj3cOZ7i-Vl6H-LbzZDALW2b7ys8LHZdJb2DkpVcukoXTeiPh17HUS48zNnvDE6pCpejU',
    'Ijazah': '1oBy2pyteJC8akIrqZiidG0fmxCmraNujhDr1zKU0p3y5Z0MKh8C7QXg8oUocu2GgSgXBvMvC',
    'SKP': '1d5xENqACl4W84_0yvW244Ci80XL1d_imi573ywI_Zi0Wb3xiTzsUmuuzpmTALvNmveeJASwg',
    'TUBEL': '1jlOYqTuvZtTuLWlDKGcnoYA4spy13v8lg4M18sx2J0EDBDr164YT_YLlt4JHnyrRBQ6gnadS',
    'HUKDIS': '1KMNa_4iN4U8ZVcAKjWrREkjvFiciK8BJ66XQbO5_-zR0X2bujUgb_HSlze8tAVhvBbWsCkR8'
  };
  
  for (const [folderName, folderId] of Object.entries(folderIds)) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      console.log(`✅ ${folderName}: ${folder.getName()}`);
    } catch (error) {
      console.log(`❌ ${folderName}: ${error.toString()}`);
    }
  }
}

/**
 * Simple test upload function
 * Run this to test the upload functionality manually
 */
function testUpload() {
  const testFileData = {
    content: Utilities.base64Encode('This is a test file content'),
    name: 'test-upload.txt',
    mimeType: 'text/plain',
    folderId: '1_EuHAesCw0pQe2ANpBhuC0uWWPLs_yhGKII6tBhtBAQw-3Q-IzCJpiarIjb6wR_oyln1TDdF' // Sertifikat folder
  };
  
  try {
    const result = uploadFileToDrive(testFileData);
    console.log('Test upload successful:', result);
    return result;
  } catch (error) {
    console.error('Test upload failed:', error.toString());
    return null;
  }
}