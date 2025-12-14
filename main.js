const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');
const ExcelJS = require('exceljs');
const { exec } = require('child_process');

let mainWindow;
let db;
const DATABASE_FILE_NAME = 'ferramental_database.db';
const REPLACEMENT_COLUMN_NAME = 'replacement_tooling_id';
const REPLACEMENT_COLUMN_TYPE = 'INTEGER';
let replacementColumnEnsured = false;
let replacementColumnEnsuringPromise = null;
let supplierMetadataEnsured = false;
let supplierMetadataPromise = null;

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const TOOLING_DATA_HEADERS = [
  'ID',
  'PN *',
  'PN Description *',
  'Tooling Description *',
  'Tooling Life (quantity) *',
  'Produced (quantity) *',
  'Production Date',
  'Annual Volume',
  'Annual Volume Date',
  'Expiration Date',
  "Supplier's Comments"
];
const VERIFICATION_SHEET_NAME = '_verification';
const VERIFICATION_KEY_VALUE = '123456';
const SUPPLIER_INFO_SHEET_NAME = 'Info & Instructions';
const SUPPLIER_INFO_TIMESTAMP_LABEL = 'Last Import Timestamp';
const SUPPLIER_METADATA_TABLE = 'supplier_metadata';
const CHANGE_TRACKING_FIELDS = {
  produced: { label: 'Produced (qty)', type: 'number' },
  tooling_life_qty: { label: 'Tooling Life (qty)', type: 'number' },
  annual_volume_forecast: { label: 'Forecast (qty)', type: 'number' },
  date_remaining_tooling_life: { label: 'Production Date', type: 'date' },
  date_annual_volume: { label: 'Annual Volume Date', type: 'date' }
};
const CHANGE_TRACKING_IGNORED_FIELDS = new Set([
  'comments',
  'remaining_tooling_life_pcs',
  'percent_tooling_life',
  'last_update'
]);
const CHANGE_NUMBER_FORMATTER = new Intl.NumberFormat('pt-BR', {
  minimumFractionDigits: 0,
  maximumFractionDigits: 2
});

function sanitizeReplacementId(value) {
  if (value === null || value === undefined) {
    return '';
  }
  const digitsOnly = String(value).trim().replace(/\D+/g, '');
  if (!digitsOnly) {
    return '';
  }
  const numeric = parseInt(digitsOnly, 10);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return '';
  }
  return String(numeric);
}

function parseLocalizedNumber(value) {
  if (value === null || value === undefined) {
    return 0;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  let s = String(value).trim();
  if (s === '') {
    return 0;
  }

  s = s.replace(/\s+/g, '').replace(/\u00A0/g, '');
  s = s.replace(/[^\d.,-]/g, '');
  if (s === '' || s === '-' || s === '+') {
    return 0;
  }

  const thousandsPattern = /^-?\d{1,3}(?:[.,]\d{3})+$/;
  if (thousandsPattern.test(s)) {
    const normalized = Number(s.replace(/[.,]/g, ''));
    return Number.isNaN(normalized) ? 0 : normalized;
  }

  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');
  const separatorIndex = Math.max(lastComma, lastDot);

  if (separatorIndex !== -1) {
    const integerPartRaw = s.slice(0, separatorIndex);
    const fractionalRaw = s.slice(separatorIndex + 1);
    const sign = integerPartRaw.startsWith('-') ? '-' : '';
    const integerPart = integerPartRaw.replace(/[^\d]/g, '') || '0';
    const fractionalPart = fractionalRaw.replace(/[^\d]/g, '');
    if (fractionalPart.length > 0) {
      const normalized = Number(`${sign}${integerPart}.${fractionalPart}`);
      if (!Number.isNaN(normalized)) {
        return normalized;
      }
    }
  }

  const fallback = Number(s.replace(/[^\d-]/g, ''));
  return Number.isNaN(fallback) ? 0 : fallback;
}

function normalizeExpirationDate(rawValue) {
  if (rawValue === null || rawValue === undefined) {
    return null;
  }

  const rawString = String(rawValue).trim();
  if (rawString === '') {
    return null;
  }

  const numericValue = Number(rawString);
  const isNumeric = !Number.isNaN(numericValue) && /^-?\d+(\.\d+)?$/.test(rawString);

  if (isNumeric) {
    const excelEpoch = new Date(EXCEL_EPOCH_MS);
    const days = Math.floor(numericValue);
    const milliseconds = Math.round((numericValue - days) * MS_PER_DAY);
    const date = new Date(excelEpoch.getTime() + days * MS_PER_DAY + milliseconds);
    if (!Number.isNaN(date.getTime())) {
      return date.toISOString().split('T')[0];
    }
  } else {
    const parsed = new Date(rawString);
    if (!Number.isNaN(parsed.getTime())) {
      return parsed.toISOString().split('T')[0];
    }
  }
  return null;
}

function calculateExpirationFromFormula({ remaining, forecast, productionDate }) {
  const baseDate = productionDate ? new Date(productionDate) : new Date();
  if (Number.isNaN(baseDate.getTime())) {
    return null;
  }

  // Se remaining <= 0, já expirou - retorna a data base
  if (remaining <= 0) {
    return baseDate.toISOString().split('T')[0];
  }

  const totalDays = forecast <= 0
    ? Math.round(remaining)
    : Math.round((remaining / forecast) * 365);

  if (Number.isNaN(totalDays)) {
    return null;
  }

  const expirationDate = new Date(baseDate);
  expirationDate.setDate(expirationDate.getDate() + totalDays);
  if (Number.isNaN(expirationDate.getTime())) {
    return null;
  }
  return expirationDate.toISOString().split('T')[0];
}

function resolveToolingExpirationDate(item) {
  if (!item) {
    return '';
  }

  let expirationDateValue = normalizeExpirationDate(item.expiration_date);
  if (!expirationDateValue) {
    const toolingLife = Number(parseLocalizedNumber(item.tooling_life_qty)) || Number(item.tooling_life_qty) || 0;
    const produced = Number(parseLocalizedNumber(item.produced)) || Number(item.produced) || 0;
    const remaining = toolingLife - produced;
    const forecast = Number(parseLocalizedNumber(item.annual_volume_forecast)) || Number(item.annual_volume_forecast) || 0;
    const productionDateValue = item.date_remaining_tooling_life || '';
    const calculatedExpiration = calculateExpirationFromFormula({
      remaining,
      forecast,
      productionDate: productionDateValue
    });
    if (calculatedExpiration) {
      expirationDateValue = calculatedExpiration;
    }
  }
  return expirationDateValue || '';
}

function getExpirationDiffDays(expirationDate) {
  if (!expirationDate) {
    return null;
  }
  const expDate = new Date(expirationDate);
  if (Number.isNaN(expDate.getTime())) {
    return null;
  }
  const now = new Date();
  return Math.ceil((expDate - now) / MS_PER_DAY);
}

function classifyToolingExpirationState(item) {
  if (!item) {
    return { state: 'ok', expirationDate: '', diffDays: null };
  }

  const normalizedStatus = String(item.status || '').trim().toLowerCase();
  const hasReplacementLink = Boolean(sanitizeReplacementId(item.replacement_tooling_id));
  if (normalizedStatus === 'obsolete') {
    return {
      state: hasReplacementLink ? 'obsolete-replaced' : 'obsolete',
      expirationDate: '',
      diffDays: null
    };
  }

  const toolingLife = Number(parseLocalizedNumber(item.tooling_life_qty)) || Number(item.tooling_life_qty) || 0;
  const produced = Number(parseLocalizedNumber(item.produced)) || Number(item.produced) || 0;
  const percentUsedValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
  const expirationDateValue = resolveToolingExpirationDate(item) || '';
  const diffDays = getExpirationDiffDays(expirationDateValue);

  if (percentUsedValue >= 100 || (typeof diffDays === 'number' && diffDays < 0)) {
    return { state: 'expired', expirationDate: expirationDateValue, diffDays };
  }

  if (typeof diffDays === 'number' && diffDays <= 730) {
    return { state: 'warning', expirationDate: expirationDateValue, diffDays };
  }

  return { state: 'ok', expirationDate: expirationDateValue, diffDays };
}

function dbGet(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => {
      if (err) {
        reject(err);
        return;
      }
      resolve(row);
    });
  });
}

function dbRun(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function onRun(err) {
      if (err) {
        reject(err);
        return;
      }
      resolve(this.changes || 0);
    });
  });
}

function cellValueToString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value.trim();
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return value.toString();
  }
  if (typeof value === 'object') {
    if (value.richText && Array.isArray(value.richText)) {
      return value.richText.map(part => part.text).join('').trim();
    }
    if (value.text) {
      return String(value.text).trim();
    }
    if (value.result !== undefined && value.result !== null) {
      return cellValueToString(value.result);
    }
  }
  return '';
}

function parseNumericCell(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (typeof value === 'number') {
    return Number.isNaN(value) ? null : value;
  }
  const text = cellValueToString(value);
  if (!text) {
    return null;
  }
  const sanitized = text.replace(/\s+/g, '');
  let normalized = sanitized;

  if (sanitized.includes(',') && sanitized.includes('.')) {
    normalized = sanitized.replace(/\./g, '').replace(',', '.');
  } else if (sanitized.includes(',') && !sanitized.includes('.')) {
    normalized = sanitized.replace(',', '.');
  }
  const parsed = Number(normalized);
  return Number.isNaN(parsed) ? null : parsed;
}

function parseExcelDate(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  
  // Se é um objeto de célula do Excel, extrair o valor real
  if (typeof value === 'object' && value !== null) {
    // Verificar se tem result (fórmula)
    if (value.result !== undefined && value.result !== null) {
      return parseExcelDate(value.result);
    }
    // Verificar se é um objeto Date
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return value;
    }
    // Verificar se tem propriedade text ou richText
    const textValue = cellValueToString(value);
    if (textValue) {
      return parseExcelDate(textValue);
    }
    return null;
  }
  
  // Se é um Date válido
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }
  
  // Se é um número serial do Excel
  if (typeof value === 'number' && !Number.isNaN(value)) {
    // Verificar se é um número serial válido (entre 1 e ~50000 para datas razoáveis)
    if (value >= 1 && value <= 100000) {
      return new Date(EXCEL_EPOCH_MS + value * MS_PER_DAY);
    }
    return null;
  }
  
  // Se é uma string, tentar parsear diferentes formatos
  const text = String(value).trim();
  if (!text) {
    return null;
  }

  // Valores numéricos em texto (datas salvas como serial do Excel formatado como texto)
  const numericText = text.replace(',', '.');
  if (/^-?\d+(?:\.\d+)?$/.test(numericText)) {
    const serialValue = parseFloat(numericText);
    if (!Number.isNaN(serialValue) && serialValue >= 1 && serialValue <= 100000) {
      return new Date(EXCEL_EPOCH_MS + serialValue * MS_PER_DAY);
    }
  }
  
  // Formato ISO: YYYY-MM-DD
  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    const [_, year, month, day] = isoMatch;
    return new Date(`${year}-${month}-${day}T00:00:00`);
  }
  
  // Formato com barras (prioriza padrão brasileiro DD/MM/YYYY)
  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const [_, firstPart, secondPart, year] = slashMatch;
    const first = parseInt(firstPart, 10);
    const second = parseInt(secondPart, 10);
    const bothValid = first >= 1 && first <= 31 && second >= 1 && second <= 31;

    if (bothValid) {
      const day = firstPart.padStart(2, '0');
      const month = secondPart.padStart(2, '0');
      return new Date(`${year}-${month}-${day}T00:00:00`);
    }
  }
  
  // Tentar parsear com Date nativo como último recurso
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatDateToISO(date) {
  if (!date || Number.isNaN(date.getTime())) {
    return null;
  }
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatDateToBR(date) {
  if (!date || Number.isNaN(date.getTime())) {
    return '';
  }
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

function sanitizeDateLikeInput(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value.trim().replace(/^\+/, '');
  }
  return value;
}

function normalizeDateInputToISO(value) {
  const sanitized = sanitizeDateLikeInput(value);
  if (sanitized === '' || sanitized === null || sanitized === undefined) {
    return '';
  }
  const parsed = parseExcelDate(sanitized);
  if (parsed && !Number.isNaN(parsed.getTime())) {
    return formatDateToISO(parsed);
  }
  const iso = toISODateString(sanitized);
  return iso || '';
}

function normalizeDateInputToBR(value) {
  const iso = normalizeDateInputToISO(value);
  if (!iso) {
    const sanitized = sanitizeDateLikeInput(value);
    return typeof sanitized === 'string' ? sanitized : '';
  }
  const [year, month, day] = iso.split('-');
  return `${day}/${month}/${year}`;
}

function humanizeFieldLabel(field) {
  if (!field || typeof field !== 'string') {
    return 'Field';
  }
  return field
    .split('_')
    .filter(Boolean)
    .map(segment => segment.charAt(0).toUpperCase() + segment.slice(1))
    .join(' ')
    || field;
}

function inferChangeTypeFromField(field) {
  const normalized = (field || '').toLowerCase();
  if (normalized.includes('date')) {
    return 'date';
  }
  if (/(qty|quantity|amount|produced|forecast|life|percent|value)/.test(normalized)) {
    return 'number';
  }
  return 'string';
}

function resolveChangeFieldMeta(field) {
  if (field && Object.prototype.hasOwnProperty.call(CHANGE_TRACKING_FIELDS, field)) {
    return CHANGE_TRACKING_FIELDS[field];
  }
  return {
    label: humanizeFieldLabel(field),
    type: inferChangeTypeFromField(field)
  };
}

function shouldTrackChangeField(field) {
  if (!field || typeof field !== 'string') {
    return false;
  }
  const normalized = field.trim();
  if (normalized === '' || normalized.startsWith('_') || normalized.startsWith('$')) {
    return false;
  }
  return !CHANGE_TRACKING_IGNORED_FIELDS.has(normalized);
}

function formatChangeCommentText(changeEntries = []) {
  if (!Array.isArray(changeEntries) || changeEntries.length === 0) {
    return '';
  }
  return changeEntries
    .map(entry => `${entry.label}, ${entry.oldFormatted} --> ${entry.newFormatted}`)
    .join('\n');
}

function buildUpdatedComments(existing, supplierComment, currentDateStr, toolingLife = null, isNewRecord = false, changeEntries = []) {
  let comments = [];
  
  // Tentar parsear comentários existentes
  if (existing) {
    try {
      const parsed = JSON.parse(existing);
      if (Array.isArray(parsed)) {
        comments = parsed;
      }
    } catch (e) {
      // Se falhar, é formato antigo de texto - converter para array vazio
      comments = [];
    }
  }
  
  // Se for novo registro e tiver tooling life, adicionar comentário inicial
  if (isNewRecord && toolingLife !== null && toolingLife !== undefined) {
    const formattedLife = toolingLife.toLocaleString('pt-BR');
    comments.push({
      date: currentDateStr,
      text: `Created with Tooling Life: ${formattedLife} pcs`,
      initial: true
    });
  }
  
  // Se há comentário do supplier, adicionar
  if (supplierComment && supplierComment.trim()) {
    comments.push({
      date: currentDateStr,
      text: supplierComment.trim(),
      initial: false
    });
  }

  const changeCommentText = formatChangeCommentText(changeEntries);
  if (changeCommentText) {
    comments.push({
      date: currentDateStr,
      text: changeCommentText,
      initial: false,
      system: true,
      origin: 'import'
    });
  }
  
  return JSON.stringify(comments);
}

function ensureSupplierMetadataTable() {
  if (supplierMetadataEnsured) {
    return supplierMetadataPromise || Promise.resolve();
  }

  supplierMetadataPromise = new Promise((resolve, reject) => {
    db.run(
      `CREATE TABLE IF NOT EXISTS ${SUPPLIER_METADATA_TABLE} (
        supplier TEXT PRIMARY KEY,
        last_import_timestamp TEXT
      )`,
      err => {
        if (err) {
          reject(err);
        } else {
          supplierMetadataEnsured = true;
          resolve();
        }
      }
    );
  });

  return supplierMetadataPromise;
}

async function getSupplierImportTimestamp(supplierName) {
  if (!supplierName) {
    return null;
  }
  await ensureSupplierMetadataTable();
  return new Promise((resolve, reject) => {
    db.get(
      `SELECT last_import_timestamp FROM ${SUPPLIER_METADATA_TABLE} WHERE supplier = ?`,
      [supplierName],
      (err, row) => {
        if (err) {
          reject(err);
        } else {
          resolve(row?.last_import_timestamp || null);
        }
      }
    );
  });
}

async function setSupplierImportTimestamp(supplierName, timestamp) {
  if (!supplierName) {
    return;
  }
  await ensureSupplierMetadataTable();
  return new Promise((resolve, reject) => {
    db.run(
      `INSERT INTO ${SUPPLIER_METADATA_TABLE} (supplier, last_import_timestamp)
       VALUES (?, ?)
       ON CONFLICT(supplier) DO UPDATE SET last_import_timestamp = excluded.last_import_timestamp`,
      [supplierName, timestamp],
      err => {
        if (err) {
          reject(err);
        } else {
          resolve();
        }
      }
    );
  });
}

function upsertSupplierInfoRow(sheet, label, value = '') {
  if (!sheet) {
    return;
  }

  let targetRow = null;
  sheet.eachRow(row => {
    const currentLabel = cellValueToString(row.getCell(1).value).toLowerCase();
    if (currentLabel === label.toLowerCase()) {
      targetRow = row;
    }
  });

  if (!targetRow) {
    targetRow = sheet.addRow({ field: label, value });
  } else {
    targetRow.getCell(2).value = value;
  }

  return targetRow;
}

function extractSupplierNameFromInfoSheet(sheet) {
  if (!sheet) {
    return '';
  }

  let supplierName = '';
  sheet.eachRow(row => {
    const label = cellValueToString(row.getCell(1).value).toLowerCase();
    if (label === 'supplier name') {
      supplierName = cellValueToString(row.getCell(2).value);
    }
  });

  return supplierName.trim();
}

function ensureToolingHeaderOrder(worksheet) {
  const headerRow = worksheet.getRow(1);
  if (!headerRow) {
    throw new Error('Spreadsheet missing header row.');
  }

  // Aliases para retrocompatibilidade (planilhas antigas podem ter nomes antigos)
  const headerAliases = {
    'Annual Volume': ['Annual Volume', 'Forecast'],
    'Annual Volume Date': ['Annual Volume Date', 'Forecast Date'],
    // Aceita headers com ou sem asterisco para campos obrigatórios
    'PN *': ['PN *', 'PN'],
    'PN Description *': ['PN Description *', 'PN Description'],
    'Tooling Description *': ['Tooling Description *', 'Tooling Description'],
    'Tooling Life (quantity) *': ['Tooling Life (quantity) *', 'Tooling Life (quantity)'],
    'Produced (quantity) *': ['Produced (quantity) *', 'Produced (quantity)']
  };

  const mismatches = [];
  TOOLING_DATA_HEADERS.forEach((expected, index) => {
    const actual = cellValueToString(headerRow.getCell(index + 1).value);
    // Verifica se o header atual corresponde ao esperado ou a um alias aceito
    const acceptedValues = headerAliases[expected] || [expected];
    if (!acceptedValues.includes(actual)) {
      mismatches.push({ expected, actual: actual || '(empty)', position: index + 1 });
    }
  });

  if (mismatches.length > 0) {
    const details = mismatches
      .map(m => `Col ${m.position}: expected "${m.expected}" but found "${m.actual}"`)
      .join('; ');
    throw new Error(`Spreadsheet header mismatch. Please use the official export template. ${details}`);
  }
}

function validateVerificationSheet(workbook) {
  const metaSheet = workbook.getWorksheet(VERIFICATION_SHEET_NAME);
  if (!metaSheet) {
    throw new Error('Verification sheet missing. Please re-export the template.');
  }

  const keyLabel = cellValueToString(metaSheet.getCell('A1').value).toLowerCase();
  const keyValue = cellValueToString(metaSheet.getCell('B1').value);

  if (keyLabel !== 'key' || keyValue !== VERIFICATION_KEY_VALUE) {
    throw new Error('Verification key mismatch. Please re-export the template.');
  }
}

function updateSupplierInfoTimestamp(workbook, timestampStr) {
  const sheet = workbook.getWorksheet(SUPPLIER_INFO_SHEET_NAME);
  if (!sheet) {
    return;
  }

  upsertSupplierInfoRow(sheet, SUPPLIER_INFO_TIMESTAMP_LABEL, timestampStr);

  sheet.protect('30625629', {
    selectLockedCells: true,
    selectUnlockedCells: false,
    formatCells: false,
    formatColumns: false,
    formatRows: false,
    insertRows: false,
    insertColumns: false,
    deleteRows: false,
    deleteColumns: false,
    sort: false,
    autoFilter: false,
    pivotTables: false
  });
}

function getCurrentTimestampBR() {
  const now = new Date();
  const date = formatDateToBR(now);
  const time = now.toLocaleTimeString('pt-BR', { hour12: false });
  return `${date} ${time}`;
}

function parseCommentsBlob(source) {
  if (!source) {
    return [];
  }
  if (Array.isArray(source)) {
    return [...source];
  }
  if (typeof source !== 'string') {
    return [];
  }
  try {
    const parsed = JSON.parse(source);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function parseNumericValueForChange(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  const text = String(value).trim();
  if (text === '') {
    return null;
  }
  const compact = text.replace(/\s+/g, '').replace(/\u00A0/g, '');
  let normalized = compact;
  const hasDot = compact.includes('.');
  const hasComma = compact.includes(',');
  if (hasDot && hasComma) {
    normalized = compact.replace(/\./g, '').replace(',', '.');
  } else if (hasComma && !hasDot) {
    normalized = compact.replace(',', '.');
  }
  const parsed = Number(normalized);
  return Number.isNaN(parsed) ? null : parsed;
}

function toISODateString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().split('T')[0];
  }
  const text = String(value).trim();
  if (text === '') {
    return '';
  }
  
  // Formato ISO: YYYY-MM-DD
  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    const year = parseInt(isoMatch[1], 10);
    // Validar que o ano é razoável (entre 1900 e 2200)
    if (year >= 1900 && year <= 2200) {
      return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
    }
    return '';
  }
  
  // Formato BR: DD/MM/YYYY
  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const year = parseInt(slashMatch[3], 10);
    if (year >= 1900 && year <= 2200) {
      const day = slashMatch[1].padStart(2, '0');
      const month = slashMatch[2].padStart(2, '0');
      return `${slashMatch[3]}-${month}-${day}`;
    }
    return '';
  }
  
  // Se o texto for apenas números (possivelmente um serial date do Excel), ignorar
  if (/^\d+$/.test(text)) {
    return '';
  }
  
  // Tentar parsear como data, mas validar o resultado
  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) {
    return '';
  }
  
  // Validar que a data parseada está em um range razoável
  const year = parsed.getFullYear();
  if (year < 1900 || year > 2200) {
    return '';
  }
  
  return parsed.toISOString().split('T')[0];
}
function normalizeComparableValueForChange(value, type) {
  if (type === 'number') {
    const numeric = parseNumericValueForChange(value);
    return numeric === null ? '' : numeric;
  }
  if (type === 'date') {
    return toISODateString(value) || '';
  }
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function formatChangeValueForDisplay(value, type) {
  if (type === 'number') {
    const numeric = parseNumericValueForChange(value);
    return numeric === null ? 'N/A' : CHANGE_NUMBER_FORMATTER.format(numeric);
  }
  if (type === 'date') {
    const iso = toISODateString(value);
    if (!iso) {
      return 'N/A';
    }
    const [year, month, day] = iso.split('-');
    return `${day}/${month}/${year}`;
  }
  if (value === null || value === undefined) {
    return 'N/A';
  }
  const trimmed = String(value).trim();
  return trimmed === '' ? 'N/A' : trimmed;
}

function collectChangeEntries(existingRecord, newValues) {
  if (!existingRecord || !newValues) {
    return [];
  }
  const entries = [];
  Object.keys(newValues).forEach((field) => {
    if (!shouldTrackChangeField(field)) {
      return;
    }
    const meta = resolveChangeFieldMeta(field);
    const before = normalizeComparableValueForChange(existingRecord[field], meta.type);
    const after = normalizeComparableValueForChange(newValues[field], meta.type);
    if (before === after) {
      return;
    }
    entries.push({
      field,
      label: meta.label,
      oldFormatted: formatChangeValueForDisplay(existingRecord[field], meta.type),
      newFormatted: formatChangeValueForDisplay(newValues[field], meta.type)
    });
  });
  return entries;
}

function mergeChangeEntriesIntoComments(existingComments, incomingComments, changeEntries, timestampStr = getCurrentTimestampBR()) {
  if (!Array.isArray(changeEntries) || changeEntries.length === 0) {
    return incomingComments !== undefined ? incomingComments : existingComments;
  }
  const base = parseCommentsBlob(incomingComments !== undefined ? incomingComments : existingComments);
  const changeText = formatChangeCommentText(changeEntries);
  if (!changeText) {
    return incomingComments !== undefined ? incomingComments : existingComments;
  }
  base.push({
    date: timestampStr,
    text: changeText,
    initial: false,
    system: true
  });
  try {
    return JSON.stringify(base);
  } catch (error) {
    return incomingComments !== undefined ? incomingComments : existingComments;
  }
}

// Obter diretório base do executável
function getAppBaseDir() {
  // Em produção (empacotado), usa o diretório real do executável gerado
  // Em desenvolvimento, usa o diretório do código fonte
  if (process.env.PORTABLE_EXECUTABLE_DIR) {
    return process.env.PORTABLE_EXECUTABLE_DIR;
  }

  if (app && app.isPackaged) {
    return path.dirname(process.execPath);
  }

  if (process.resourcesPath && !process.resourcesPath.endsWith('app.asar')) {
    return path.join(process.resourcesPath, '..');
  }

  return __dirname;
}

function ensureDatabaseFile(baseDir) {
  const targetPath = path.join(baseDir, DATABASE_FILE_NAME);

  if (fs.existsSync(targetPath)) {
    return targetPath;
  }

  try {
    if (!fs.existsSync(baseDir)) {
      fs.mkdirSync(baseDir, { recursive: true });
    }

    const packagedDbPath = path.join(app.getAppPath(), DATABASE_FILE_NAME);
    if (fs.existsSync(packagedDbPath)) {
      fs.copyFileSync(packagedDbPath, targetPath);
      return targetPath;
    }
    fs.writeFileSync(targetPath, '');
    return targetPath;
  } catch (error) {
    return targetPath;
  }
}

// Conectar ao banco de dados
function connectDatabase() {
  const baseDir = getAppBaseDir();
  const dbPath = ensureDatabaseFile(baseDir);
  return new Promise((resolve, reject) => {
    db = new sqlite3.Database(dbPath, (err) => {
      if (err) {
        reject(err);
        return;
      }
      checkAndAddColumns()
        .then(() => ensureTodosTable())
        .then(() => ensureStepHistoryTable())
        .then(resolve)
        .catch((schemaError) => {
          reject(schemaError);
        });
    });
  });
}

// Verificar e adicionar colunas que possam estar faltando
function checkAndAddColumns() {
  const columnsToCheck = [
    { name: 'pn', type: 'TEXT' },
    { name: 'pn_description', type: 'TEXT' },
    { name: 'supplier', type: 'TEXT' },
    { name: 'tool_description', type: 'TEXT' },
    { name: 'tool_number_arb', type: 'TEXT' },
    { name: 'asset_number', type: 'TEXT' },
    { name: 'tool_ownership', type: 'TEXT' },
    { name: 'customer', type: 'TEXT' },
    { name: 'tooling_life_qty', type: 'TEXT' },
    { name: 'produced', type: 'TEXT' },
    { name: 'remaining_tooling_life_pcs', type: 'TEXT' },
    { name: 'percent_tooling_life', type: 'TEXT' },
    { name: 'annual_volume_forecast', type: 'TEXT' },
    { name: 'date_remaining_tooling_life', type: 'TEXT' },
    { name: 'date_annual_volume', type: 'TEXT' },
    { name: 'expiration_date', type: 'TEXT' },
    { name: 'finish_due_date', type: 'TEXT' },
    { name: 'amount_brl', type: 'TEXT' },
    { name: 'tool_quantity', type: 'TEXT' },
    { name: 'bailment_agreement_signed', type: 'TEXT' },
    { name: 'tooling_book', type: 'TEXT' },
    { name: 'disposition', type: 'TEXT' },
    { name: 'status', type: 'TEXT' },
    { name: 'steps', type: 'TEXT' },
    { name: 'bu', type: 'TEXT' },
    { name: 'category', type: 'TEXT' },
    { name: 'cummins_responsible', type: 'TEXT' },
    { name: 'comments', type: 'TEXT' },
    { name: 'todos', type: 'TEXT' },
    { name: 'stim_tooling_management', type: 'TEXT' },
    { name: 'vpcr', type: 'TEXT' },
    { name: 'replacement_tooling_id', type: 'INTEGER' }
  ];
  
  return new Promise((resolve, reject) => {
    db.all('PRAGMA table_info(ferramental)', [], (err, columns) => {
      if (err) {
        reject(err);
        return;
      }

      const existingColumns = new Set(columns.map(col => col.name));
      if (existingColumns.has(REPLACEMENT_COLUMN_NAME)) {
        replacementColumnEnsured = true;
      }

      const missingColumns = columnsToCheck.filter(({ name }) => !existingColumns.has(name));
      if (missingColumns.length === 0) {
        resolve();
        return;
      }

      let completed = 0;
      let hasFailed = false;

      db.serialize(() => {
        missingColumns.forEach(({ name, type }) => {
          const columnType = type || 'TEXT';
          db.run(`ALTER TABLE ferramental ADD COLUMN ${name} ${columnType}`, (alterErr) => {
            if (hasFailed) {
              return;
            }

            if (alterErr && !/duplicate column/i.test(alterErr.message || '')) {
              hasFailed = true;
              reject(alterErr);
              return;
            }

            if (name === REPLACEMENT_COLUMN_NAME) {
              replacementColumnEnsured = true;
            }

            completed += 1;
            if (completed === missingColumns.length) {
              resolve();
            }
          });
        });
      });
    });
  });
}

function ensureReplacementColumnExists(force = false) {
  if (force) {
    replacementColumnEnsured = false;
    replacementColumnEnsuringPromise = null;
  }

  if (replacementColumnEnsured) {
    return Promise.resolve();
  }

  if (replacementColumnEnsuringPromise) {
    return replacementColumnEnsuringPromise;
  }

  replacementColumnEnsuringPromise = new Promise((resolve, reject) => {
    db.all('PRAGMA table_info(ferramental)', [], (err, columns) => {
      if (err) {
        replacementColumnEnsuringPromise = null;
        reject(err);
        return;
      }

      const hasColumn = Array.isArray(columns) && columns.some(col => col.name === REPLACEMENT_COLUMN_NAME);
      if (hasColumn) {
        replacementColumnEnsured = true;
        replacementColumnEnsuringPromise = null;
        resolve();
        return;
      }

      db.run(`ALTER TABLE ferramental ADD COLUMN ${REPLACEMENT_COLUMN_NAME} ${REPLACEMENT_COLUMN_TYPE}`, (alterErr) => {
        replacementColumnEnsuringPromise = null;
        if (alterErr && !/duplicate column/i.test(alterErr.message || '')) {
          reject(alterErr);
          return;
        }
        replacementColumnEnsured = true;
        resolve();
      });
    });
  });

  return replacementColumnEnsuringPromise;
}

function ensureStepHistoryTable() {
  return new Promise((resolve, reject) => {
    db.run(`
      CREATE TABLE IF NOT EXISTS step_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tooling_id INTEGER NOT NULL,
        old_step TEXT,
        new_step TEXT NOT NULL,
        changed_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (tooling_id) REFERENCES ferramental(id) ON DELETE CASCADE
      )
    `, (err) => {
      if (err) {
        console.error('[StepHistory] Error creating step_history table:', err);
        reject(err);
      } else {
        console.log('[StepHistory] step_history table ensured');
        resolve();
      }
    });
  });
}

function recordStepChange(toolingId, oldStep, newStep) {
  return new Promise((resolve, reject) => {
    if (oldStep === newStep) {
      resolve({ recorded: false, reason: 'no_change' });
      return;
    }
    
    db.run(
      `INSERT INTO step_history (tooling_id, old_step, new_step) VALUES (?, ?, ?)`,
      [toolingId, oldStep || null, newStep],
      function(err) {
        if (err) {
          console.error('[StepHistory] Error recording step change:', err);
          reject(err);
        } else {
          console.log(`[StepHistory] Recorded step change for tooling ${toolingId}: ${oldStep} -> ${newStep}`);
          resolve({ recorded: true, id: this.lastID });
        }
      }
    );
  });
}

function getStepHistory(toolingId) {
  return new Promise((resolve, reject) => {
    db.all(
      `SELECT * FROM step_history WHERE tooling_id = ? ORDER BY changed_at DESC`,
      [toolingId],
      (err, rows) => {
        if (err) {
          console.error('[StepHistory] Error getting step history:', err);
          reject(err);
        } else {
          resolve(rows || []);
        }
      }
    );
  });
}

function clearStepHistory(toolingId) {
  return new Promise((resolve, reject) => {
    db.run(
      `DELETE FROM step_history WHERE tooling_id = ?`,
      [toolingId],
      function(err) {
        if (err) {
          console.error('[StepHistory] Error clearing step history:', err);
          reject(err);
        } else {
          console.log(`[StepHistory] Cleared ${this.changes} entries for tooling ${toolingId}`);
          resolve({ success: true, deleted: this.changes });
        }
      }
    );
  });
}

function clearAllStepHistory() {
  return new Promise((resolve, reject) => {
    // First, ensure the step_history table exists
    db.run(`CREATE TABLE IF NOT EXISTS step_history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      tooling_id INTEGER NOT NULL,
      old_step TEXT,
      new_step TEXT NOT NULL,
      changed_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )`, (createErr) => {
      if (createErr) {
        console.error('[StepHistory] Error ensuring step_history table:', createErr);
      }
      
      // Delete all step history records
      db.run(`DELETE FROM step_history`, function(err) {
        if (err) {
          console.error('[StepHistory] Error clearing step_history table:', err);
          // Continue anyway to clear the steps field
        }
        
        const historyDeleted = this ? this.changes : 0;
        console.log(`[StepHistory] Cleared ${historyDeleted} entries from step_history`);
        
        // Then, set all steps fields to empty in ferramental table
        db.run(`UPDATE ferramental SET steps = '', last_update = datetime('now')`, function(updateErr) {
          if (updateErr) {
            console.error('[StepHistory] Error clearing steps field:', updateErr);
            reject(updateErr);
          } else {
            console.log(`[StepHistory] Cleared steps field from ${this.changes} tooling records`);
            resolve({ success: true, historyDeleted, toolingsUpdated: this.changes });
          }
        });
      });
    });
  });
}

function isMissingReplacementColumnError(err) {
  if (!err || typeof err.message !== 'string') {
    return false;
  }
  return err.message.includes(`no such column: ${REPLACEMENT_COLUMN_NAME}`);
}

function ensureTodosTable() {
  return new Promise((resolve, reject) => {
    // First check if table exists
    db.get("SELECT name FROM sqlite_master WHERE type='table' AND name='todos'", (err, row) => {
      if (err) {
        reject(err);
        return;
      }

      if (!row) {
        // Table doesn't exist, create it
        db.run(`
          CREATE TABLE todos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tooling_id INTEGER NOT NULL,
            text TEXT NOT NULL,
            completed INTEGER DEFAULT 0,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (tooling_id) REFERENCES ferramental(id) ON DELETE CASCADE
          )
        `, (createErr) => {
          if (createErr) {
            reject(createErr);
          } else {
            resolve();
          }
        });
      } else {
        // Table exists, verify columns
        db.all("PRAGMA table_info(todos)", (pragmaErr, columns) => {
          if (pragmaErr) {
            reject(pragmaErr);
            return;
          }

          const columnNames = columns.map(col => col.name);
          const requiredColumns = ['id', 'tooling_id', 'text', 'completed', 'created_at'];
          const missingColumns = requiredColumns.filter(col => !columnNames.includes(col));

          if (missingColumns.length > 0) {
            // Recreate table with all columns
            db.serialize(() => {
              db.run('DROP TABLE IF EXISTS todos_backup', (dropErr) => {
              });
              
              db.run('ALTER TABLE todos RENAME TO todos_backup', (renameErr) => {
                if (renameErr) {
                  reject(renameErr);
                  return;
                }

                db.run(`
                  CREATE TABLE todos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tooling_id INTEGER NOT NULL,
                    text TEXT NOT NULL,
                    completed INTEGER DEFAULT 0,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (tooling_id) REFERENCES ferramental(id) ON DELETE CASCADE
                  )
                `, (createErr) => {
                  if (createErr) {
                    reject(createErr);
                    return;
                  }

                  // Try to copy data if possible
                  db.run(`
                    INSERT INTO todos (id, tooling_id, text, completed, created_at)
                    SELECT id, tooling_id, text, completed, created_at FROM todos_backup
                  `, (copyErr) => {
                    if (copyErr) {
                    }

                    db.run('DROP TABLE todos_backup', (dropErr) => {
                    });
                    resolve();
                  });
                });
              });
            });
          } else {
            resolve();
          }
        });
      }
    });
  });
}

function executeToolingUpdate(id, payload, attempt = 1, options = {}) {
  return new Promise((resolve, reject) => {
    const data = { ...payload };
    const skipWhenNoTrackedChanges = options?.skipWhenNoTrackedChanges === true;

    if (!data || Object.keys(data).length === 0) {
      resolve({ success: true, changes: 0, comments: null, skipped: true });
      return;
    }

    if (data.hasOwnProperty('tooling_life_qty') || data.hasOwnProperty('produced')) {
      const toolingLife = parseFloat(data.tooling_life_qty) || 0;
      const produced = parseFloat(data.produced) || 0;
      const remaining = toolingLife - produced;
      const percentUsed = toolingLife > 0 ? ((produced / toolingLife) * 100).toFixed(1) : 0;

      data.remaining_tooling_life_pcs = remaining >= 0 ? remaining : 0;
      data.percent_tooling_life = percentUsed;
    }

    const trackableFields = Object.keys(data).filter(field => shouldTrackChangeField(field));
    const requiresLookup = trackableFields.length > 0;

    const finalizeUpdate = (existingRecord) => {
      let changeEntries = [];
      if (existingRecord) {
        changeEntries = collectChangeEntries(existingRecord, data);
        if (changeEntries.length > 0) {
          console.debug('[Ferramental][ChangeDebug]', {
            id,
            updates: changeEntries.map((entry) => ({
              field: entry.field,
              before: entry.oldFormatted,
              after: entry.newFormatted
            }))
          });
          const customTimestamp = options?.commentTimestamp;
          data.comments = mergeChangeEntriesIntoComments(
            existingRecord.comments,
            data.comments,
            changeEntries,
            customTimestamp || getCurrentTimestampBR()
          );
        }
      }

      const validFields = Object.keys(data).filter(key => {
        if (!key || typeof key !== 'string' || key.trim() === '') {
          return false;
        }
        const fieldName = key.trim();
        if (fieldName.startsWith('_') || fieldName.startsWith('$')) {
          return false;
        }
        return data[key] !== undefined;
      });

      if (validFields.length === 0) {
        resolve({ success: true, changes: 0, comments: existingRecord?.comments ?? null, skipped: true });
        return;
      }

      const hasNonTrackableUpdates = validFields.some(field => !shouldTrackChangeField(field));
      if (
        skipWhenNoTrackedChanges &&
        existingRecord &&
        changeEntries.length === 0 &&
        !hasNonTrackableUpdates
      ) {
        resolve({
          success: true,
          changes: 0,
          comments: existingRecord?.comments ?? null,
          skipped: true
        });
        return;
      }

      const validValues = validFields.map(key => data[key]);
      const setClause = validFields.map(field => `${field.trim()} = ?`).join(', ');
      const query = `UPDATE ferramental SET ${setClause}, last_update = datetime('now') WHERE id = ?`;

      db.run(
        query,
        [...validValues, id],
        async function(err) {
          if (err && isMissingReplacementColumnError(err) && attempt === 1) {
            try {
              await ensureReplacementColumnExists(true);
              const retryResult = await executeToolingUpdate(id, payload, attempt + 1, options);
              resolve(retryResult);
              return;
            } catch (retryError) {
              reject(retryError);
              return;
            }
          }

          if (err) {
            reject(err);
          } else {
            // Record step change if steps field was modified
            if (data.hasOwnProperty('steps') && existingRecord) {
              const oldStep = existingRecord.steps || null;
              const newStep = data.steps || null;
              if (oldStep !== newStep) {
                try {
                  await recordStepChange(id, oldStep, newStep);
                } catch (stepErr) {
                  console.error('[StepHistory] Failed to record step change:', stepErr);
                }
              }
            }
            
            resolve({
              success: true,
              changes: this.changes,
              comments: data.comments ?? existingRecord?.comments ?? null,
              skipped: false
            });
          }
        }
      );
    };

    if (requiresLookup) {
      const selectFields = new Set(['id', 'comments']);
      trackableFields.forEach(field => selectFields.add(field));
      const columnList = Array.from(selectFields).join(', ');
      db.get(`SELECT ${columnList} FROM ferramental WHERE id = ?`, [id], (selectErr, row) => {
        if (selectErr) {
          reject(selectErr);
          return;
        }
        finalizeUpdate(row || null);
      });
    } else {
      finalizeUpdate(null);
    }
  });
}

// Handlers IPC para operações do banco
ipcMain.handle('load-tooling', async () => {
  return new Promise((resolve, reject) => {
    db.all('SELECT * FROM ferramental ORDER BY id DESC', [], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows);
      }
    });
  });
});

ipcMain.handle('get-suppliers-with-stats', async () => {
  return new Promise((resolve, reject) => {
    db.all(
      `SELECT * FROM ferramental WHERE supplier IS NOT NULL AND supplier != ''`,
      [],
      (err, rows) => {
        if (err) {
          reject(err);
          return;
        }

        const supplierMap = new Map();

        rows.forEach(item => {
          const supplierName = String(item.supplier || '').trim();
          if (!supplierName) {
            return;
          }

          if (!supplierMap.has(supplierName)) {
            supplierMap.set(supplierName, {
              supplier: supplierName,
              total: 0,
              expired: 0,
              warning_1year: 0,
              warning_2years: 0,
              ok_5years: 0,
              ok_plus: 0
            });
          }

          const metrics = supplierMap.get(supplierName);
          metrics.total += 1;

          const classification = classifyToolingExpirationState(item);

          if (classification.state === 'expired') {
            metrics.expired += 1;
          } else if (classification.state === 'warning') {
            metrics.warning_1year += 1;
          } else if (classification.state === 'ok' && typeof classification.diffDays === 'number') {
            if (classification.diffDays > 730 && classification.diffDays <= 1825) {
              metrics.ok_5years += 1;
            } else if (classification.diffDays > 1825) {
              metrics.ok_plus += 1;
            }
          }
        });

        const result = Array.from(supplierMap.values()).sort((a, b) =>
          a.supplier.localeCompare(b.supplier, 'pt-BR', { sensitivity: 'base' })
        );
        resolve(result);
      }
    );
  });
});

ipcMain.handle('get-tooling-by-supplier', async (event, supplier) => {
  return new Promise((resolve, reject) => {
    db.all(`
      SELECT * FROM ferramental 
      WHERE supplier = ?
      ORDER BY 
        CASE 
          WHEN julianday('now') > julianday(expiration_date) THEN 1
          WHEN julianday(expiration_date) - julianday('now') < 365 THEN 2
          ELSE 3
        END,
        expiration_date
    `, [supplier], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows);
      }
    });
  });
});

ipcMain.handle('search-tooling', async (event, term) => {
  return new Promise((resolve, reject) => {
    const searchTerm = `%${term}%`;
    db.all(`
      SELECT * FROM ferramental 
      WHERE pn LIKE ? 
         OR pn_description LIKE ? 
         OR supplier LIKE ?
         OR tool_description LIKE ?
         OR tool_number_arb LIKE ?
         OR asset_number LIKE ?
         OR tool_ownership LIKE ?
         OR customer LIKE ?
         OR tooling_life_qty LIKE ?
         OR produced LIKE ?
         OR annual_volume_forecast LIKE ?
         OR status LIKE ?
         OR steps LIKE ?
         OR bu LIKE ?
         OR category LIKE ?
         OR cummins_responsible LIKE ?
         OR comments LIKE ?
         OR stim_tooling_management LIKE ?
         OR vpcr LIKE ?
         OR disposition LIKE ?
         OR CAST(id AS TEXT) LIKE ?
      ORDER BY id DESC
    `, [
      searchTerm, searchTerm, searchTerm, searchTerm, searchTerm,
      searchTerm, searchTerm, searchTerm, searchTerm, searchTerm,
      searchTerm, searchTerm, searchTerm, searchTerm, searchTerm,
      searchTerm, searchTerm, searchTerm, searchTerm, searchTerm,
      searchTerm
    ], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows);
      }
    });
  });
});

ipcMain.handle('get-tooling-by-id', async (event, id) => {
  return new Promise((resolve, reject) => {
    db.get('SELECT * FROM ferramental WHERE id = ?', [id], (err, row) => {
      if (err) {
        reject(err);
      } else {
        resolve(row);
      }
    });
  });
});

ipcMain.handle('get-tooling-by-replacement-id', async (event, replacementId) => {
  return new Promise((resolve, reject) => {
    const numericReplacementId = Number(replacementId);
    if (!Number.isFinite(numericReplacementId)) {
      resolve([]);
      return;
    }
    db.all(
      'SELECT * FROM ferramental WHERE replacement_tooling_id = ? ORDER BY id ASC',
      [numericReplacementId],
      (err, rows) => {
        if (err) {
          reject(err);
        } else {
          resolve(rows);
        }
      }
    );
  });
});

ipcMain.handle('get-all-tooling-ids', async () => {
  return new Promise((resolve, reject) => {
    db.all(
      'SELECT id, pn, pn_description, supplier, tool_description FROM ferramental ORDER BY id ASC',
      [],
      (err, rows) => {
        if (err) {
          reject(err);
        } else {
          resolve(rows);
        }
      }
    );
  });
});

ipcMain.handle('get-unique-responsibles', async () => {
  return new Promise((resolve, reject) => {
    db.all(
      `SELECT DISTINCT cummins_responsible 
       FROM ferramental 
       WHERE cummins_responsible IS NOT NULL 
         AND TRIM(cummins_responsible) != '' 
       ORDER BY cummins_responsible ASC`,
      [],
      (err, rows) => {
        if (err) {
          reject(err);
        } else {
          const responsibles = rows
            .map(row => String(row.cummins_responsible || '').trim())
            .filter(name => name.length > 0);
          resolve(responsibles);
        }
      }
    );
  });
});

ipcMain.handle('get-analytics', async () => {
  return new Promise((resolve, reject) => {
    db.all(`
      WITH filtered AS (
        SELECT *
        FROM ferramental
        WHERE supplier IS NOT NULL
          AND TRIM(supplier) != ''
      )
      SELECT 
        COUNT(*) as total,
        COUNT(DISTINCT supplier) as suppliers,
        COUNT(DISTINCT cummins_responsible) as responsibles,
        SUM(CASE 
          WHEN UPPER(status) = 'ACTIVE' THEN 1
          ELSE 0
        END) as active,
        SUM(CASE 
          WHEN ((percent_tooling_life IS NOT NULL AND CAST(percent_tooling_life AS REAL) >= 100.0)
            OR (tooling_life_qty > 0 AND produced >= tooling_life_qty))
            AND NOT (UPPER(status) = 'OBSOLETE' AND replacement_tooling_id IS NOT NULL AND replacement_tooling_id != '')
          THEN 1
          ELSE 0
        END) as expired_total,
        SUM(CASE 
          WHEN NOT ((percent_tooling_life IS NOT NULL AND CAST(percent_tooling_life AS REAL) >= 100.0)
                OR (tooling_life_qty > 0 AND produced >= tooling_life_qty))
            AND expiration_date IS NOT NULL 
            AND julianday(expiration_date) - julianday('now') > 0
            AND julianday(expiration_date) - julianday('now') <= 730
            AND NOT (UPPER(status) = 'OBSOLETE' AND replacement_tooling_id IS NOT NULL AND replacement_tooling_id != '')
          THEN 1
          ELSE 0
        END) as expiring_two_years
      FROM filtered
    `, [], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows[0]);
      }
    });
  });
});

ipcMain.handle('get-steps-summary', async () => {
  return new Promise((resolve, reject) => {
    db.all(`
      SELECT 
        steps,
        COUNT(*) as count
      FROM ferramental
      WHERE steps IS NOT NULL 
        AND TRIM(steps) != '' 
        AND TRIM(steps) != '0'
      GROUP BY steps
      ORDER BY CAST(steps AS INTEGER)
    `, [], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows || []);
      }
    });
  });
});

ipcMain.handle('update-tooling', async (event, id, data) => {
  try {
    await ensureReplacementColumnExists();
  } catch (error) {
  }
  return executeToolingUpdate(id, data);
});

ipcMain.handle('create-tooling', async (event, data) => {
  return new Promise((resolve, reject) => {
    try {
      const pn = (data?.pn || '').trim();
      const pnDescription = (data?.pn_description || '').trim();
      const supplier = (data?.supplier || '').trim();
      const toolingLife = parseFloat(data?.tooling_life_qty) || 0;
      const produced = parseFloat(data?.produced) || 0;
      const toolDescription = (data?.tool_description || '').trim();
      const productionDateValue = (data?.date_remaining_tooling_life || '').trim();
      const productionDate = productionDateValue.length > 0 ? productionDateValue : null;
      const forecastRaw = data?.annual_volume_forecast;
      const parsedForecast =
        forecastRaw === null || forecastRaw === undefined
          ? null
          : (typeof forecastRaw === 'number' ? forecastRaw : parseFloat(forecastRaw));
      const annualForecast = Number.isFinite(parsedForecast) && parsedForecast > 0 ? parsedForecast : null;
      const forecastDateValue = (data?.date_annual_volume || '').trim();
      const forecastDate = forecastDateValue.length > 0 ? forecastDateValue : null;
      const comments = data?.comments || null;

      if (!pn || !supplier) {
        resolve({ success: false, error: 'PN e fornecedor são obrigatórios.' });
        return;
      }

      const remaining = toolingLife - produced;
      const percent = toolingLife > 0 ? ((produced / toolingLife) * 100).toFixed(1) : 0;

      db.run(
        `INSERT INTO ferramental (
          pn,
          pn_description,
          supplier,
          tool_description,
          tooling_life_qty,
          produced,
          remaining_tooling_life_pcs,
          percent_tooling_life,
          annual_volume_forecast,
          date_annual_volume,
          status,
          last_update,
          date_remaining_tooling_life,
          comments
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), ?, ?)`,
        [
          pn,
          pnDescription,
          supplier,
          toolDescription,
          toolingLife,
          produced,
          remaining >= 0 ? remaining : 0,
          percent,
          annualForecast,
          forecastDate,
          'ACTIVE',
          productionDate,
          comments
        ],
        function(err) {
          if (err) {
            reject(err);
          } else {
            resolve({ success: true, id: this.lastID });
          }
        }
      );
    } catch (error) {
      reject(error);
    }
  });
});

ipcMain.handle('delete-tooling', async (event, id) => {
  return new Promise((resolve, reject) => {
    db.run('DELETE FROM ferramental WHERE id = ?', [id], function(err) {
      if (err) {
        reject(err);
      } else {
        resolve({ success: true, changes: this.changes });
      }
    });
  });
});

// Exportar dados do supplier para Excel
ipcMain.handle('export-supplier-data', async (event, supplierName, filteredIds = null) => {
  return new Promise((resolve, reject) => {
    // Construir query baseada nos IDs filtrados
    let query;
    let params;
    
    if (filteredIds && Array.isArray(filteredIds) && filteredIds.length > 0) {
      // Exportar apenas os IDs filtrados
      const placeholders = filteredIds.map(() => '?').join(',');
      query = `
        SELECT 
          id,
          pn,
          pn_description,
          tool_description,
          tooling_life_qty,
          produced,
          date_remaining_tooling_life as production_date,
          annual_volume_forecast as forecast,
          date_annual_volume as forecast_date
        FROM ferramental 
        WHERE supplier = ? AND id IN (${placeholders})
        ORDER BY id
      `;
      params = [supplierName, ...filteredIds];
    } else {
      // Exportar todos do supplier (fallback)
      query = `
        SELECT 
          id,
          pn,
          pn_description,
          tool_description,
          tooling_life_qty,
          produced,
          date_remaining_tooling_life as production_date,
          annual_volume_forecast as forecast,
          date_annual_volume as forecast_date
        FROM ferramental 
        WHERE supplier = ?
        ORDER BY id
      `;
      params = [supplierName];
    }
    
    db.all(query, params, async (err, rows) => {
      if (err) {
        reject(err);
        return;
      }

      try {
        const workbook = new ExcelJS.Workbook();
        
        // Criar aba de dados
        const worksheet = workbook.addWorksheet('Tooling Data');

        // Definir colunas com larguras
        worksheet.columns = [
          { header: 'ID', key: 'id', width: 6 },
          { header: 'PN *', key: 'pn', width: 15 },
          { header: 'PN Description *', key: 'pn_description', width: 25 },
          { header: 'Tooling Description *', key: 'tool_description', width: 25 },
          { header: 'Tooling Life (quantity) *', key: 'tooling_life_qty', width: 12 },
          { header: 'Produced (quantity) *', key: 'produced', width: 12 },
          { header: 'Production Date', key: 'production_date', width: 14 },
          { header: 'Annual Volume', key: 'forecast', width: 12 },
          { header: 'Annual Volume Date', key: 'forecast_date', width: 14 },
          { header: 'Expiration Date', key: 'expiration_date', width: 14 },
          { header: "Supplier's Comments", key: 'supplier_comments', width: 50 }
        ];

        // Fixar primeira linha (cabeçalho)
        worksheet.views = [
          { state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }
        ];

        // Aplicar estilo Cummins Red no cabeçalho para destaque visual
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC8102E' } // Cummins Red
          };
          cell.font = {
            color: { argb: 'FFFFFFFF' },
            bold: true
          };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });
        headerRow.commit();

        // Centralizar colunas A, E, F, G, H, I, J para facilitar leitura
        ['A', 'E', 'F', 'G', 'H', 'I', 'J'].forEach(columnKey => {
          const column = worksheet.getColumn(columnKey);
          column.alignment = { horizontal: 'center', vertical: 'middle' };
        });

        // Função auxiliar para converter string de data para objeto Date
        function parseExcelDate(dateStr) {
          if (!dateStr || typeof dateStr !== 'string') return null;
          // Formato YYYY-MM-DD
          if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
            const [year, month, day] = dateStr.split('-').map(Number);
            return new Date(year, month - 1, day); // mês é 0-indexed
          }
          // Formato DD/MM/YYYY
          if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
            const [day, month, year] = dateStr.split('/').map(Number);
            return new Date(year, month - 1, day);
          }
          return null;
        }

        // Adicionar dados
        rows.forEach(row => {
          // Converter datas para objetos Date do JavaScript (ExcelJS converte automaticamente)
          const prodDate = parseExcelDate(row.production_date);
          const foreDate = parseExcelDate(row.forecast_date);
          
          worksheet.addRow({
            id: row.id || '',
            pn: row.pn || '',
            pn_description: row.pn_description || '',
            tool_description: row.tool_description || '',
            tooling_life_qty: row.tooling_life_qty || '',
            produced: row.produced || '',
            production_date: prodDate, // Objeto Date real
            forecast: row.forecast || '',
            forecast_date: foreDate, // Objeto Date real
            expiration_date: '', // Será preenchido com fórmula depois
            supplier_comments: ''
          });
        });

        const numericColumnIndexes = [5, 6, 8]; // Tooling Life, Produced, Forecast
        const dateColumnIndexes = [7, 9, 10]; // Production Date, Annual Volume Date, Expiration Date

        // Calcular primeira linha vazia (para desbloquear PN)
        const firstEmptyRow = rows.length + 2; // +2 porque: +1 para header, +1 para próxima linha
        
        // Adicionar 100 linhas vazias extras
        for (let i = 0; i < 100; i++) {
          worksheet.addRow({
            id: '',
            pn: '',
            pn_description: '',
            tool_description: '',
            tooling_life_qty: '',
            produced: '',
            production_date: '',
            forecast: '',
            forecast_date: '',
            expiration_date: '', // Será preenchido com fórmula depois
            supplier_comments: ''
          });
        }

        // Adicionar fórmula de Expiration Date para todas as linhas (exceto header)
        const totalRows = rows.length + 100; // Total de linhas de dados + vazias
        for (let rowNum = 2; rowNum <= totalRows + 1; rowNum++) {
          const row = worksheet.getRow(rowNum);
          const expirationCell = row.getCell(10); // Coluna J (Expiration Date)
          
          // Fórmula: Segue a mesma lógica do programa
          // 1. Se Production Date vazio ou dados incompletos -> não mostra nada
          // 2. Se remaining (Tooling Life - Produced) <= 0 -> retorna Production Date (já expirou)
          // 3. Se Forecast <= 0 -> não mostra nada (vida útil indeterminada)
          // 4. Caso contrário: Production Date + ROUND((remaining / forecast) * 365, 0)
          expirationCell.value = {
            formula: `IF(OR(G${rowNum}="",E${rowNum}="",F${rowNum}=""),"",IF((E${rowNum}-F${rowNum})<=0,G${rowNum},IF(H${rowNum}<=0,"",G${rowNum}+ROUND((E${rowNum}-F${rowNum})/H${rowNum}*365,0))))`,
            date1904: false
          };
          expirationCell.numFmt = 'dd/mm/yyyy';
        }

        // Proteger planilha com senha
        await worksheet.protect('30625629', {
          selectLockedCells: true,
          selectUnlockedCells: true,
          formatCells: false,
          formatColumns: false,
          formatRows: false,
          insertRows: false,
          insertColumns: false,
          deleteRows: false,
          deleteColumns: false,
          sort: false,
          autoFilter: false,
          pivotTables: false
        });

        // Desbloquear colunas C a I (Description até Supplier Comments) e aplicar validações
        // A partir da primeira linha vazia, desbloquear também a coluna B (PN) para permitir novos ferramentais
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) {
            // Header row - manter bloqueado
            return;
          }
          
          // Verificar se é linha vazia ou além dos dados existentes
          const isEmptyOrNew = rowNumber >= firstEmptyRow;
          
          // Desbloquear colunas C até I e K (índices 3 a 9 e 11) - excluindo J (Expiration Date)
          for (let colIndex = 3; colIndex <= 11; colIndex++) {
            if (colIndex === 10) continue; // Pular coluna J (Expiration Date) - mantém bloqueada
            const cell = row.getCell(colIndex);
            cell.protection = { locked: false };
          }
          
          // Se for linha vazia ou nova, desbloquear também coluna B (PN)
          if (isEmptyOrNew) {
            const pnCell = row.getCell(2);
            pnCell.protection = { locked: false };
          }

          // Aplicar restrições numéricas
          numericColumnIndexes.forEach(colIndex => {
            const cell = row.getCell(colIndex);
            cell.numFmt = '#,##0';
            cell.dataValidation = {
              type: 'decimal',
              operator: 'between',
              allowBlank: true,
              formulae: [-9999999999, 9999999999],
              showErrorMessage: true,
              errorTitle: 'Valor inválido',
              error: 'Use apenas números nestes campos.'
            };
          });

          // Aplicar restrições de data
          dateColumnIndexes.forEach(colIndex => {
            const cell = row.getCell(colIndex);
            cell.numFmt = 'dd/mm/yyyy';
            cell.dataValidation = {
              type: 'date',
              operator: 'between',
              allowBlank: true,
              formulae: ['DATE(2000,1,1)', 'DATE(2100,12,31)'],
              showErrorMessage: true,
              errorTitle: 'Data inválida',
              error: 'Use apenas datas nestes campos.'
            };

            if (cell.value) {
              const parsedDate = new Date(cell.value);
              if (!Number.isNaN(parsedDate.getTime())) {
                cell.value = parsedDate;
              }
            }
          });
        });

        // Criar segunda aba com informações do supplier
        await ensureSupplierMetadataTable();
        const lastImportTimestamp = await getSupplierImportTimestamp(supplierName);
        const supplierSheet = workbook.addWorksheet(SUPPLIER_INFO_SHEET_NAME);
        
        // Remover gridlines da aba Instructions
        supplierSheet.views = [
          { showGridLines: false }
        ];
        
        supplierSheet.columns = [
          { header: 'Field', key: 'field', width: 28 },
          { header: 'Value', key: 'value', width: 55 }
        ];
        upsertSupplierInfoRow(supplierSheet, 'Supplier Name', supplierName);
        upsertSupplierInfoRow(
          supplierSheet,
          SUPPLIER_INFO_TIMESTAMP_LABEL,
          lastImportTimestamp || ''
        );
        // Adicionar instruções para fornecedores
        supplierSheet.addRow({ field: '', value: '' }); // Linha em branco
        const instrRow = supplierSheet.addRow({ field: 'INSTRUÇÕES', value: '' });
        instrRow.getCell(1).font = { bold: true, size: 12 };
        supplierSheet.addRow({ field: '', value: '' });
        const addNewRow = supplierSheet.addRow({ field: 'Como Adicionar Novo Ferramental:', value: '' });
        addNewRow.getCell(1).font = { bold: true };
        supplierSheet.addRow({ field: '', value: 'Os itens com * são obrigatórios para adição de novo ferramental' });
        supplierSheet.addRow({ field: '', value: 'Adicione todos os dados necessários e deixe o ID vazio.' });
        supplierSheet.addRow({ field: '', value: 'O sistema atribuirá o ID automaticamente ao importar.' });
        supplierSheet.addRow({ field: '', value: '' });
        const reqFieldsRow = supplierSheet.addRow({ field: 'Campos Obrigatórios:', value: '' });
        reqFieldsRow.getCell(1).font = { bold: true };
        supplierSheet.addRow({ field: '→ PN', value: 'Part Number do ferramental' });
        supplierSheet.addRow({ field: '→ PN Description', value: 'Descrição do Part Number' });
        supplierSheet.addRow({ field: '→ Tooling Description', value: 'Descrição do ferramental' });
        supplierSheet.addRow({ field: '→ Tooling Life', value: 'Vida útil do ferramental (quantidade)' });
        supplierSheet.addRow({ field: '→ Produced', value: 'Quantidade já produzida' });
        supplierSheet.addRow({ field: '', value: '' });
        const optFieldsRow = supplierSheet.addRow({ field: 'Campos Opcionais:', value: '' });
        optFieldsRow.getCell(1).font = { bold: true };
        supplierSheet.addRow({ field: '→ Production Date', value: 'Data em que a quantidade "Produced" foi medida' });
        supplierSheet.addRow({ field: '→ Annual Volume', value: 'Volume anual previsto' });
        supplierSheet.addRow({ field: '→ Annual Volume Date', value: 'Data em que o "Annual Volume" foi calculado' });
        supplierSheet.addRow({ field: '→ Supplier\'s Comments', value: 'Comentários adicionais (carimbados com data/hora ao importar)' });

        const supplierHeaderRow = supplierSheet.getRow(1);
        supplierHeaderRow.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC8102E' }
          };
          cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        supplierHeaderRow.commit();

        supplierSheet.protect('30625629', {
          selectLockedCells: true,
          selectUnlockedCells: false,
          formatCells: false,
          formatColumns: false,
          formatRows: false,
          insertRows: false,
          insertColumns: false,
          deleteRows: false,
          deleteColumns: false,
          sort: false,
          autoFilter: false,
          pivotTables: false
        });

        const verificationSheet = workbook.addWorksheet(VERIFICATION_SHEET_NAME);
        verificationSheet.state = 'veryHidden';
        verificationSheet.getCell('A1').value = 'key';
        verificationSheet.getCell('B1').value = VERIFICATION_KEY_VALUE;

        // Auto-fit nas colunas da aba Tooling Data - calcular largura baseada no conteúdo real
        worksheet.columns.forEach((column, index) => {
          let maxContentLength = 0;
          
          column.eachCell({ includeEmpty: false }, cell => {
            // Pular o cabeçalho (linha 1) - vamos tratar separadamente
            if (cell.row === 1) return;
            
            let cellLength = 0;
            // Tratar datas especialmente
            if (cell.value instanceof Date) {
              cellLength = 10; // DD/MM/YYYY = 10 caracteres
            } else if (typeof cell.value === 'number') {
              cellLength = cell.value.toString().length;
            } else if (cell.value) {
              cellLength = cell.value.toString().length;
            }
            
            if (cellLength > maxContentLength) {
              maxContentLength = cellLength;
            }
          });
          
          // Calcular largura do header
          const headerLength = column.header ? column.header.toString().length : 0;
          
          // Largura final = maior entre header e conteúdo + padding de 2
          // Largura mínima de 8, máxima de 50 para acomodar colunas de comentários
          const calculatedWidth = Math.max(headerLength, maxContentLength) + 2;
          column.width = Math.min(Math.max(calculatedWidth, 8), 50);
        });

        // Auto-fit nas colunas da aba Info & Instructions
        supplierSheet.columns.forEach(column => {
          let maxLength = 0;
          column.eachCell({ includeEmpty: false }, cell => {
            const cellLength = cell.value ? cell.value.toString().length : 0;
            if (cellLength > maxLength) {
              maxLength = cellLength;
            }
          });
          // Largura baseada no conteúdo com padding moderado
          column.width = Math.min(Math.max(maxLength + 2, 10), 50);
        });

        // Escolher onde salvar o arquivo
        const result = await dialog.showSaveDialog(mainWindow, {
          title: 'Export Supplier Data',
          defaultPath: `Tooling-Data-${supplierName.replace(/[^a-z0-9]/gi, '_')}.xlsx`,
          filters: [
            { name: 'Excel Files', extensions: ['xlsx'] }
          ]
        });

        if (result.canceled) {
          resolve({ success: false, message: 'Export cancelled' });
          return;
        }

        // Salvar o arquivo
        await workbook.xlsx.writeFile(result.filePath);

        resolve({ 
          success: true, 
          message: 'File exported successfully',
          filePath: result.filePath 
        });

      } catch (error) {
        reject(error);
      }
    });
  });
});

// Importar dados do supplier a partir de Excel
ipcMain.handle('import-supplier-data', async (event, supplierName) => {
  if (!supplierName) {
    return { success: false, message: 'Supplier not provided' };
  }

  const dialogResult = await dialog.showOpenDialog(mainWindow, {
    title: 'Import Supplier Data',
    buttonLabel: 'Import',
    properties: ['openFile'],
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
  });

  if (dialogResult.canceled || !dialogResult.filePaths || dialogResult.filePaths.length === 0) {
    return { success: false, message: 'Import cancelled' };
  }

  const filePath = dialogResult.filePaths[0];
  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile(filePath);

  validateVerificationSheet(workbook);
  const supplierInfoSheet = workbook.getWorksheet(SUPPLIER_INFO_SHEET_NAME);
  if (!supplierInfoSheet) {
    throw new Error('Aba "Info & Instructions" não encontrada. Por favor, re-exporte o template.');
  }

  const supplierNameInFile = extractSupplierNameFromInfoSheet(supplierInfoSheet);
  const normalizedFileSupplier = supplierNameInFile.trim().toLowerCase();
  const normalizedCurrentSupplier = supplierName.trim().toLowerCase();
  if (
    supplierNameInFile &&
    normalizedFileSupplier &&
    normalizedFileSupplier !== normalizedCurrentSupplier
  ) {
    return {
      success: false,
      message: `Spreadsheet belongs to "${supplierNameInFile}" but you selected "${supplierName}".`
    };
  }

  const worksheet = workbook.getWorksheet('Tooling Data') || workbook.worksheets[0];

  if (!worksheet) {
    throw new Error('Could not locate "Tooling Data" worksheet in the provided file.');
  }

  if (worksheet.rowCount < 2) {
    return { success: false, message: 'Spreadsheet does not contain data rows.' };
  }

  ensureToolingHeaderOrder(worksheet);

  let updated = 0;
  let created = 0;
  let skipped = 0;
  const todayStr = formatDateToBR(new Date());

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);

    const idValue = cellValueToString(row.getCell(1).value);
    const id = parseInt(idValue, 10);
    const pn = cellValueToString(row.getCell(2).value);
    const pnDescription = cellValueToString(row.getCell(3).value);
    const toolDescription = cellValueToString(row.getCell(4).value);
    const toolingLifeQty = parseNumericCell(row.getCell(5).value);
    const producedQty = parseNumericCell(row.getCell(6).value);
    const productionDateISO = formatDateToISO(parseExcelDate(row.getCell(7).value));
    const forecastQty = parseNumericCell(row.getCell(8).value);
    const forecastDateISO = formatDateToISO(parseExcelDate(row.getCell(9).value));
    // Coluna 10 (Expiration Date) é ignorada - calculada automaticamente pelo programa
    const supplierCommentRaw = cellValueToString(row.getCell(11).value);
    const supplierComment = supplierCommentRaw?.trim() || '';

    // Se não há ID mas há PN, é um novo ferramental
    if (!id) {
      if (!pn || pn.trim() === '') {
        // Linha vazia ou sem dados relevantes
        skipped += 1;
        continue;
      }

      // Criar novo registro
      const mergedComments = buildUpdatedComments('', supplierComment, todayStr, toolingLifeQty, true);
      
      await dbRun(
        `INSERT INTO ferramental (
          supplier, pn, pn_description, tool_description, tooling_life_qty, produced,
          date_remaining_tooling_life, annual_volume_forecast,
          date_annual_volume, comments, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          supplierName,
          pn || '',
          pnDescription || '',
          toolDescription || '',
          toolingLifeQty,
          producedQty,
          productionDateISO,
          forecastQty,
          forecastDateISO,
          mergedComments,
          'ACTIVE'
        ]
      );

      created += 1;
      continue;
    }

    // Se há ID, atualizar registro existente
    const record = await dbGet(
      `SELECT pn,
              pn_description,
              tool_description,
              tooling_life_qty,
              produced,
              date_remaining_tooling_life,
              annual_volume_forecast,
              date_annual_volume,
              comments
         FROM ferramental
        WHERE id = ? AND supplier = ?`,
      [id, supplierName]
    );

    if (!record) {
      skipped += 1;
      continue;
    }

    const changeEntries = collectChangeEntries(record, {
      pn: pn || '',
      pn_description: pnDescription || '',
      tool_description: toolDescription || '',
      tooling_life_qty: toolingLifeQty,
      produced: producedQty,
      date_remaining_tooling_life: productionDateISO,
      annual_volume_forecast: forecastQty,
      date_annual_volume: forecastDateISO
    });

    if (changeEntries.length === 0 && !supplierComment) {
      skipped += 1;
      continue;
    }

    const mergedComments = buildUpdatedComments(
      record.comments || '',
      supplierComment,
      todayStr,
      null,
      false,
      changeEntries
    );

    await dbRun(
      `UPDATE ferramental 
         SET pn = ?,
             pn_description = ?,
             tool_description = ?,
             tooling_life_qty = ?,
             produced = ?,
             date_remaining_tooling_life = ?,
             annual_volume_forecast = ?,
             date_annual_volume = ?,
             comments = ?
       WHERE id = ? AND supplier = ?`,
      [
        pn || '',
        pnDescription || '',
        toolDescription || '',
        toolingLifeQty,
        producedQty,
        productionDateISO,
        forecastQty,
        forecastDateISO,
        mergedComments,
        id,
        supplierName
      ]
    );

    updated += 1;
  }

  const timestampStr = getCurrentTimestampBR();
  updateSupplierInfoTimestamp(workbook, timestampStr);
  await setSupplierImportTimestamp(supplierName, timestampStr);
  const verificationSheet = workbook.getWorksheet(VERIFICATION_SHEET_NAME);
  if (verificationSheet) {
    verificationSheet.state = 'veryHidden';
  }
  await workbook.xlsx.writeFile(filePath);
  return {
    success: true,
    updated,
    created,
    skipped,
    message: `Updated ${updated} item(s), created ${created} new item(s). Skipped ${skipped}.`
  };
});

// ===== GERENCIAMENTO DE ANEXOS =====

// Função para obter diretório de anexos
function getAttachmentsDir() {
  const baseDir = getAppBaseDir();
  const attachmentsPath = path.join(baseDir, 'attachments');
  // Garante que o diretório de anexos existe
  if (!fs.existsSync(attachmentsPath)) {
    try {
      fs.mkdirSync(attachmentsPath, { recursive: true });
    } catch (error) {
    }
  }

  return attachmentsPath;
}

// Lista anexos de um fornecedor
ipcMain.handle('get-attachments', async (event, supplierName, itemId = null) => {
  const attachmentsDir = getAttachmentsDir();
  let targetDir;
  
  if (itemId) {
    targetDir = path.join(attachmentsDir, sanitizeFileName(supplierName), String(itemId));
  } else {
    targetDir = path.join(attachmentsDir, sanitizeFileName(supplierName));
  }
  
  // Usar caminho longo para Windows se necessário
  const longTargetDir = getLongPath(targetDir);
  
  if (!fs.existsSync(longTargetDir)) {
    return [];
  }
  
  try {
    const allItems = fs.readdirSync(longTargetDir);
    const files = allItems.filter(itemName => {
      const fullPath = path.join(targetDir, itemName);
      const longFullPath = getLongPath(fullPath);
      const stats = fs.statSync(longFullPath);
      return stats.isFile();
    });
    
    return files.map(fileName => {
      const filePath = path.join(targetDir, fileName);
      const longFilePath = getLongPath(filePath);
      const stats = fs.statSync(longFilePath);
      return {
        fileName,
        supplierName,
        itemId,
        fileSize: stats.size,
        uploadDate: stats.birthtime
      };
    });
  } catch (error) {
    return [];
  }
});

// Busca contagem de anexos para múltiplos IDs em lote (otimização)
ipcMain.handle('get-attachments-count-batch', async (event, supplierName, itemIds) => {
  if (!supplierName || !Array.isArray(itemIds) || itemIds.length === 0) {
    return {};
  }

  const attachmentsDir = getAttachmentsDir();
  const supplierDir = path.join(attachmentsDir, sanitizeFileName(supplierName));
  const longSupplierDir = getLongPath(supplierDir);
  const counts = {};

  if (!fs.existsSync(longSupplierDir)) {
    itemIds.forEach(id => { counts[id] = 0; });
    return counts;
  }

  itemIds.forEach(itemId => {
    const targetDir = path.join(supplierDir, String(itemId));
    const longTargetDir = getLongPath(targetDir);
    if (!fs.existsSync(longTargetDir)) {
      counts[itemId] = 0;
      return;
    }

    try {
      const allItems = fs.readdirSync(longTargetDir);
      const fileCount = allItems.filter(itemName => {
        try {
          const fullPath = path.join(targetDir, itemName);
          const longFullPath = getLongPath(fullPath);
          const stats = fs.statSync(longFullPath);
          return stats.isFile();
        } catch {
          return false;
        }
      }).length;
      counts[itemId] = fileCount;
    } catch (error) {
      counts[itemId] = 0;
    }
  });

  return counts;
});

// Upload de arquivo
ipcMain.handle('upload-attachment', async (event, supplierName, itemId = null) => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'All Files', extensions: ['*'] }
    ]
  });
  
  if (result.canceled || result.filePaths.length === 0) {
    return { success: false, cancelled: true };
  }
  
  try {
    const attachmentsDir = getAttachmentsDir();
    let targetDir;
    
    const supplierDir = path.join(attachmentsDir, sanitizeFileName(supplierName));
    const hasCardScope = itemId !== null && itemId !== undefined;
    targetDir = hasCardScope
      ? path.join(supplierDir, String(itemId))
      : supplierDir;
    
    // Usar caminho longo para Windows se necessário
    const longTargetDir = getLongPath(targetDir);
    
    if (!fs.existsSync(longTargetDir)) {
      try {
        fs.mkdirSync(longTargetDir, { recursive: true });
      } catch (error) {
        return { success: false, error: 'Não foi possível criar diretório de anexos.' };
      }
    }
    
    const results = result.filePaths.map(sourcePath => {
      try {
        const fileName = path.basename(sourcePath);
        const destPath = path.join(targetDir, fileName);
        const longDestPath = getLongPath(destPath);
        const longSourcePath = getLongPath(sourcePath);
        fs.copyFileSync(longSourcePath, longDestPath);
        const existsAfterCopy = fs.existsSync(longDestPath);
        return { success: true, fileName };
      } catch (error) {
        return { success: false, fileName: path.basename(sourcePath), error: error.message };
      }
    });
    
    const hasFailure = results.some(r => r.success !== true);
    const fileCount = results.filter(r => r.success === true).length;
    
    return { 
      success: !hasFailure, 
      results,
      message: fileCount > 1 ? `${fileCount} arquivos anexados` : 'Arquivo anexado'
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('upload-attachment-from-paths', async (event, supplierName, filePaths, itemId = null) => {
  if (!supplierName || !Array.isArray(filePaths) || filePaths.length === 0) {
    return { success: false, error: 'Nenhum arquivo informado.' };
  }

  try {
    const attachmentsDir = getAttachmentsDir();
    const supplierDir = path.join(attachmentsDir, sanitizeFileName(supplierName));
    const hasCardScope = itemId !== null && itemId !== undefined;
    const targetDir = hasCardScope
      ? path.join(supplierDir, String(itemId))
      : supplierDir;

    // Usar caminho longo para Windows se necessário
    const longTargetDir = getLongPath(targetDir);

    if (!fs.existsSync(longTargetDir)) {
      try {
        fs.mkdirSync(longTargetDir, { recursive: true });
      } catch (error) {
        return { success: false, error: 'Não foi possível criar diretório de anexos.' };
      }
    }

    const results = filePaths.map(sourcePath => {
      const longSourcePath = getLongPath(sourcePath);
      if (!sourcePath || !fs.existsSync(longSourcePath)) {
        return { success: false, error: 'Arquivo não encontrado.' };
      }

      const stats = fs.statSync(longSourcePath);
      if (!stats.isFile()) {
        return { success: false, error: 'Apenas arquivos podem ser anexados.' };
      }

      const fileName = path.basename(sourcePath);
      const destPath = path.join(targetDir, fileName);
      const longDestPath = getLongPath(destPath);

      try {
        fs.copyFileSync(longSourcePath, longDestPath);
        return { success: true, fileName };
      } catch (error) {
        return { success: false, fileName, error: error.message };
      }
    });

    const hasFailure = results.some(item => item.success !== true);
    const fileCount = results.filter(item => item.success === true).length;

    return {
      success: !hasFailure,
      results,
      message: fileCount > 1 ? `${fileCount} arquivos anexados` : 'Arquivo anexado'
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Abre arquivo anexado
ipcMain.handle('open-attachment', async (event, supplierName, fileName, itemId = null) => {
  const attachmentsDir = getAttachmentsDir();
  let filePath;
  
  if (itemId) {
    filePath = path.join(attachmentsDir, sanitizeFileName(supplierName), String(itemId), fileName);
  } else {
    filePath = path.join(attachmentsDir, sanitizeFileName(supplierName), fileName);
  }
  
  // Usar caminho longo para Windows se necessário
  const longPath = getLongPath(filePath);
  
  if (!fs.existsSync(longPath)) {
    throw new Error('Arquivo não encontrado');
  }
  
  try {
    // Usar função que lida com caminhos longos no Windows
    return await openFileWithLongPath(filePath);
  } catch (error) {
    throw error;
  }
});

// Abre a pasta de anexos do supplier
ipcMain.handle('open-attachments-folder', async (event, supplierName) => {
  const attachmentsDir = getAttachmentsDir();
  const supplierDir = path.join(attachmentsDir, sanitizeFileName(supplierName));
  
  // Criar diretório se não existir
  const longSupplierDir = getLongPath(supplierDir);
  if (!fs.existsSync(longSupplierDir)) {
    fs.mkdirSync(longSupplierDir, { recursive: true });
  }
  
  try {
    await shell.openPath(supplierDir);
    return { success: true };
  } catch (error) {
    throw new Error('Error opening folder: ' + error.message);
  }
});

// Exclui arquivo anexado
ipcMain.handle('delete-attachment', async (event, supplierName, fileName, itemId = null) => {
  const attachmentsDir = getAttachmentsDir();
  let filePath;
  
  if (itemId) {
    filePath = path.join(attachmentsDir, sanitizeFileName(supplierName), String(itemId), fileName);
  } else {
    filePath = path.join(attachmentsDir, sanitizeFileName(supplierName), fileName);
  }
  
  // Usar caminho longo para Windows se necessário
  const longPath = getLongPath(filePath);
  
  if (!fs.existsSync(longPath)) {
    return { success: false, error: 'Arquivo não encontrado' };
  }
  
  try {
    fs.unlinkSync(longPath);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Função auxiliar para sanitizar nomes de arquivo/pasta
function sanitizeFileName(name) {
  return name.replace(/[<>:"/\\|?*]/g, '_').trim();
}

// Função auxiliar para lidar com caminhos longos no Windows (>260 caracteres)
function getLongPath(filePath) {
  // No Windows, usar sempre o prefixo \\?\ para caminhos longos ou próximos do limite
  if (process.platform === 'win32') {
    // Normaliza o caminho
    const normalizedPath = path.resolve(filePath);
    // Se já tem o prefixo ou é muito curto, retorna normalizado
    if (normalizedPath.startsWith('\\\\?\\')) {
      return normalizedPath;
    }
    // Adiciona prefixo para caminhos maiores que 200 caracteres (margem de segurança antes do limite de 260)
    if (normalizedPath.length > 200) {
      return '\\\\?\\' + normalizedPath;
    }
  }
  return filePath;
}

// Função para abrir arquivo com suporte a caminhos longos no Windows
function openFileWithLongPath(filePath) {
  return new Promise((resolve, reject) => {
    if (process.platform === 'win32') {
      // No Windows, usar cmd /c start para abrir arquivos com caminhos longos
      // Funciona melhor que PowerShell para caminhos extensos
      const normalizedPath = path.resolve(filePath);
      const command = `cmd /c start "" "${normalizedPath}"`;  
      
      exec(command, { shell: true, windowsHide: true }, (error) => {
        if (error) {
          console.error('[OpenFile] cmd /c start failed, trying shell.openPath:', error);
          // Fallback para shell.openPath
          shell.openPath(filePath)
            .then(() => resolve({ success: true }))
            .catch((err) => reject(err));
        } else {
          resolve({ success: true });
        }
      });
    } else {
      // Em outros sistemas, usar shell.openPath normalmente
      shell.openPath(filePath)
        .then(() => resolve({ success: true }))
        .catch((err) => reject(err));
    }
  });
}

function createWindow () {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    frame: false,
    icon: path.join(__dirname, 'ferramentas.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.removeMenu();
  mainWindow.loadFile('index.html');
  
  mainWindow.once('ready-to-show', () => {
    mainWindow.maximize();
  });
}

app.whenReady().then(async () => {
  try {
    await connectDatabase();
  } catch (error) {
  }

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

// Todos handlers
ipcMain.handle('get-todos', async (event, toolingId) => {
  return new Promise((resolve, reject) => {
    db.all('SELECT * FROM todos WHERE tooling_id = ? ORDER BY created_at ASC', [toolingId], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows);
      }
    });
  });
});

ipcMain.handle('add-todo', async (event, toolingId, text) => {
  return new Promise((resolve, reject) => {
    db.run('INSERT INTO todos (tooling_id, text, completed) VALUES (?, ?, 0)', [toolingId, text], function(err) {
      if (err) {
        reject(err);
      } else {
        resolve({ id: this.lastID, tooling_id: toolingId, text, completed: 0 });
      }
    });
  });
});

ipcMain.handle('update-todo', async (event, todoId, text, completed) => {
  return new Promise((resolve, reject) => {
    db.run('UPDATE todos SET text = ?, completed = ? WHERE id = ?', [text, completed, todoId], (err) => {
      if (err) {
        reject(err);
      } else {
        resolve({ success: true });
      }
    });
  });
});

ipcMain.handle('delete-todo', async (event, todoId) => {
  return new Promise((resolve, reject) => {
    db.run('DELETE FROM todos WHERE id = ?', [todoId], (err) => {
      if (err) {
        reject(err);
      } else {
        resolve({ success: true });
      }
    });
  });
});

// Step History handlers
ipcMain.handle('get-step-history', async (event, toolingId) => {
  try {
    const history = await getStepHistory(toolingId);
    return history;
  } catch (err) {
    console.error('[IPC] Error getting step history:', err);
    throw err;
  }
});

ipcMain.handle('clear-step-history', async (event, toolingId) => {
  try {
    const result = await clearStepHistory(toolingId);
    return result;
  } catch (err) {
    console.error('[IPC] Error clearing step history:', err);
    throw err;
  }
});

ipcMain.handle('clear-all-step-history', async () => {
  try {
    const result = await clearAllStepHistory();
    return result;
  } catch (err) {
    console.error('[IPC] Error clearing all step history:', err);
    throw err;
  }
});

// Window control handlers
ipcMain.on('minimize-window', () => {
  if (mainWindow) {
    mainWindow.minimize();
  }
});

ipcMain.on('maximize-window', () => {
  if (mainWindow) {
    if (mainWindow.isMaximized()) {
      mainWindow.unmaximize();
    } else {
      mainWindow.maximize();
    }
  }
});

ipcMain.on('close-window', () => {
  if (mainWindow) {
    mainWindow.close();
  }
});

// Export Forecast for Supplier (ID, PN, Supplier, Forecast, Annual Volume Date, Expiration)
ipcMain.handle('export-forecast-supplier', async () => {
  return new Promise((resolve, reject) => {
    db.all(`
      SELECT 
        id,
        pn,
        supplier,
        annual_volume_forecast as forecast,
        date_annual_volume as forecast_date,
        expiration_date
      FROM ferramental 
      ORDER BY id
    `, [], async (err, rows) => {
      if (err) {
        reject(err);
        return;
      }

      try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Forecast Data');

        worksheet.columns = [
          { header: 'ID', key: 'id', width: 10 },
          { header: 'PN', key: 'pn', width: 25 },
          { header: 'Supplier', key: 'supplier', width: 30 },
          { header: 'Annual Forecast', key: 'forecast', width: 20 },
          { header: 'Annual Volume Date', key: 'forecast_date', width: 20 },
          { header: 'Expiration', key: 'expiration_date', width: 20 }
        ];

        // Header styling
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC8102E' }
          };
          cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        headerRow.commit();

        // Add data rows
        rows.forEach(row => {
          const foreDate = normalizeDateInputToBR(row.forecast_date);
          const expDate = normalizeDateInputToBR(row.expiration_date);
          worksheet.addRow({
            id: row.id || '',
            pn: row.pn || '',
            supplier: row.supplier || '',
            forecast: row.forecast || '',
            forecast_date: foreDate,
            expiration_date: expDate
          });
        });

        // Center align ID and date columns
        ['A', 'D', 'E', 'F'].forEach(columnKey => {
          const column = worksheet.getColumn(columnKey);
          column.alignment = { horizontal: 'center', vertical: 'middle' };
        });

        // Add data validations
        const lastRow = worksheet.rowCount;
        
        // Annual Forecast validation (column D) - numeric only
        for (let i = 2; i <= lastRow; i++) {
          worksheet.getCell(`D${i}`).dataValidation = {
            type: 'decimal',
            operator: 'greaterThanOrEqual',
            allowBlank: true,
            formulae: [-999999999],
            showErrorMessage: true,
            errorStyle: 'warning',
            errorTitle: 'Invalid Number',
            error: 'Please enter a valid number for Annual Forecast'
          };
        }
        
        // Annual Volume Date validation (column E) - date only
        for (let i = 2; i <= lastRow; i++) {
          worksheet.getCell(`E${i}`).dataValidation = {
            type: 'date',
            operator: 'greaterThan',
            showErrorMessage: true,
            allowBlank: true,
            formulae: [new Date(1900, 0, 1)],
            errorStyle: 'warning',
            errorTitle: 'Invalid Date',
            error: 'Please enter a valid date in format dd/mm/yyyy',
            promptTitle: 'Date Format',
            prompt: 'Enter date in format: dd/mm/yyyy'
          };
        }

        // Lock ID, PN, Supplier, Expiration columns
        ['A', 'B', 'C', 'F'].forEach(columnKey => {
          worksheet.getColumn(columnKey).eachCell({ includeEmpty: true }, (cell, rowNumber) => {
            if (rowNumber > 1) {
              cell.protection = { locked: true };
            }
          });
        });

        // Unlock Forecast and Annual Volume Date columns
        ['D', 'E'].forEach(columnKey => {
          worksheet.getColumn(columnKey).eachCell({ includeEmpty: true }, (cell, rowNumber) => {
            if (rowNumber > 1) {
              cell.protection = { locked: false };
            }
          });
        });

        // Protect worksheet with password
        await worksheet.protect('30625629', {
          selectLockedCells: true,
          selectUnlockedCells: true,
          formatCells: false,
          formatColumns: false,
          formatRows: false,
          insertRows: false,
          insertColumns: false,
          deleteRows: false,
          deleteColumns: false,
          sort: false,
          autoFilter: false,
          pivotTables: false
        });

        const result = await dialog.showSaveDialog(mainWindow, {
          title: 'Export Forecast Data (Supplier)',
          defaultPath: `Forecast-Supplier-${new Date().toISOString().split('T')[0]}.xlsx`,
          filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
        });

        if (result.canceled) {
          resolve({ success: false, cancelled: true });
          return;
        }

        await workbook.xlsx.writeFile(result.filePath);
        resolve({ success: true, filePath: result.filePath });

      } catch (error) {
        reject(error);
      }
    });
  });
});

// Import Forecast for Supplier (updates Forecast and Annual Volume Date by ID)
ipcMain.handle('import-forecast-supplier', async () => {
  const dialogResult = await dialog.showOpenDialog(mainWindow, {
    title: 'Import Forecast Data (Supplier)',
    buttonLabel: 'Import',
    properties: ['openFile'],
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
  });

  if (dialogResult.canceled || !dialogResult.filePaths || dialogResult.filePaths.length === 0) {
    return { success: false, cancelled: true };
  }

  try {
    const filePath = dialogResult.filePaths[0];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      return { success: false, error: 'No worksheet found in file' };
    }

    // Validate header row - accept both old (5 columns) and new (6 columns) format
    const headerRow = worksheet.getRow(1);
    const firstHeader = headerRow.getCell(1).value?.toString().trim();
    const fourthHeader = headerRow.getCell(4).value?.toString().trim();
    const fifthHeader = headerRow.getCell(5).value?.toString().trim();

    if (firstHeader !== 'ID' || fourthHeader !== 'Annual Forecast' || (fifthHeader !== 'Annual Volume Date' && fifthHeader !== 'Forecast Date')) {
      return { 
        success: false, 
        error: 'Invalid file format. Expected columns: ID, PN, Supplier, Annual Forecast, Annual Volume Date' 
      };
    }

    let updatedCount = 0;
    const errors = [];
    const importDateStamp = formatDateToBR(new Date());

    // Skip header row (row 1)
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      
      const id = row.getCell(1).value; // Column A
      const forecast = row.getCell(4).value; // Column D
      const forecastDate = row.getCell(5).value; // Column E

      // Skip if no ID
      if (!id) continue;

      try {
        // Prepare update data
        const updateData = {};
        
        const normalizedForecast = cellValueToString(forecast);
        if (normalizedForecast !== '') {
          // Validate that forecast is a number
          const forecastNum = parseFloat(normalizedForecast.replace(/,/g, '.'));
          if (isNaN(forecastNum)) {
            errors.push(`Row ${rowNumber} (ID ${id}): Annual Forecast must be a number`);
            continue;
          }
          updateData.annual_volume_forecast = normalizedForecast;
        }

        const normalizedForecastDate = normalizeDateInputToISO(forecastDate);
        if (forecastDate && !normalizedForecastDate) {
          errors.push(`Row ${rowNumber} (ID ${id}): Invalid date format for Annual Volume Date`);
          continue;
        }
        if (normalizedForecastDate) {
          updateData.date_annual_volume = normalizedForecastDate;
        }

        // Only update if there's data to update
        if (Object.keys(updateData).length > 0) {
          const result = await executeToolingUpdate(id, updateData, 1, {
            commentTimestamp: importDateStamp,
            skipWhenNoTrackedChanges: true
          });
          if (result?.changes > 0) {
            updatedCount++;
          }
        }

      } catch (error) {
        errors.push(`Row ${rowNumber} (ID ${id}): ${error.message}`);
      }
    }

    return { 
      success: true, 
      updatedCount,
      errors: errors.length > 0 ? errors : undefined
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Export Full Database for Manager
ipcMain.handle('export-forecast-manager', async () => {
  return new Promise((resolve, reject) => {
    db.all(`SELECT * FROM ferramental ORDER BY id`, [], async (err, rows) => {
      if (err) {
        reject(err);
        return;
      }

      try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Full Database');

        // Define all columns from database
        const columns = [
          { header: 'ID', key: 'id', width: 10 },
          { header: 'PN', key: 'pn', width: 20 },
          { header: 'PN Description', key: 'pn_description', width: 30 },
          { header: 'Supplier', key: 'supplier', width: 30 },
          { header: 'Tool Description', key: 'tool_description', width: 30 },
          { header: 'Tool Number ARB', key: 'tool_number_arb', width: 20 },
          { header: 'Asset Number', key: 'asset_number', width: 20 },
          { header: 'Tool Ownership', key: 'tool_ownership', width: 15 },
          { header: 'Customer', key: 'customer', width: 20 },
          { header: 'Tooling Life Qty', key: 'tooling_life_qty', width: 15 },
          { header: 'Produced', key: 'produced', width: 15 },
          { header: 'Remaining Life', key: 'remaining_tooling_life_pcs', width: 15 },
          { header: '% Life', key: 'percent_tooling_life', width: 10 },
          { header: 'Annual Forecast', key: 'annual_volume_forecast', width: 15 },
          { header: 'Annual Volume Date', key: 'date_annual_volume', width: 15 },
          { header: 'Expiration Date', key: 'expiration_date', width: 15 },
          { header: 'Finish Due Date', key: 'finish_due_date', width: 15 },
          { header: 'Amount BRL', key: 'amount_brl', width: 15 },
          { header: 'Tool Quantity', key: 'tool_quantity', width: 12 },
          { header: 'Bailment Signed', key: 'bailment_agreement_signed', width: 15 },
          { header: 'Tooling Book', key: 'tooling_book', width: 15 },
          { header: 'Disposition', key: 'disposition', width: 15 },
          { header: 'Status', key: 'status', width: 15 },
          { header: 'Steps', key: 'steps', width: 30 },
          { header: 'BU', key: 'bu', width: 15 },
          { header: 'Category', key: 'category', width: 15 },
          { header: 'Owner', key: 'cummins_responsible', width: 20 },
          { header: 'STIM', key: 'stim_tooling_management', width: 15 },
          { header: 'VPCR', key: 'vpcr', width: 15 }
        ];

        worksheet.columns = columns;

        // Header styling
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF2563EB' } // Blue for manager
          };
          cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        headerRow.commit();

        // Add data rows
        rows.forEach(row => {
          const rowData = {
            id: row.id || '',
            pn: row.pn || '',
            pn_description: row.pn_description || '',
            supplier: row.supplier || '',
            tool_description: row.tool_description || '',
            tool_number_arb: row.tool_number_arb || '',
            asset_number: row.asset_number || '',
            tool_ownership: row.tool_ownership || '',
            customer: row.customer || '',
            tooling_life_qty: row.tooling_life_qty || '',
            produced: row.produced || '',
            remaining_tooling_life_pcs: row.remaining_tooling_life_pcs || '',
            percent_tooling_life: row.percent_tooling_life || '',
            annual_volume_forecast: row.annual_volume_forecast || '',
            date_annual_volume: normalizeDateInputToBR(row.date_annual_volume),
            expiration_date: normalizeDateInputToBR(row.expiration_date),
            finish_due_date: normalizeDateInputToBR(row.finish_due_date),
            amount_brl: row.amount_brl || '',
            tool_quantity: row.tool_quantity || '',
            bailment_agreement_signed: row.bailment_agreement_signed || '',
            tooling_book: row.tooling_book || '',
            disposition: row.disposition || '',
            status: row.status || '',
            steps: row.steps || '',
            bu: row.bu || '',
            category: row.category || '',
            cummins_responsible: row.cummins_responsible || '',
            stim_tooling_management: row.stim_tooling_management || '',
            vpcr: row.vpcr || ''
          };
          worksheet.addRow(rowData);
        });

        // Center align ID column
        worksheet.getColumn('A').alignment = { horizontal: 'center', vertical: 'middle' };

        // Add data validations
        const lastRow = worksheet.rowCount;
        
        // Status column validation (column W = 23) - list validation
        const statusOptions = ['ACTIVE', 'UNDER CONSTRUCTION', 'OBSOLETE', 'INACTIVE'];
        for (let i = 2; i <= lastRow; i++) {
          worksheet.getCell(`W${i}`).dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [`"${statusOptions.join(',')}"`],
            showErrorMessage: true,
            errorStyle: 'warning',
            errorTitle: 'Invalid Status',
            error: 'Please select a valid status from the list'
          };
        }

        // Numeric columns validation
        const numericColumns = ['J', 'K', 'L', 'M', 'N', 'R', 'S']; // Tooling Life, Produced, Remaining, %, Forecast, Amount, Qty
        for (let i = 2; i <= lastRow; i++) {
          numericColumns.forEach(col => {
            worksheet.getCell(`${col}${i}`).dataValidation = {
              type: 'decimal',
              operator: 'greaterThanOrEqual',
              allowBlank: true,
              formulae: [-999999999],
              showErrorMessage: true,
              errorStyle: 'warning',
              errorTitle: 'Invalid Number',
              error: 'Please enter a valid number'
            };
          });
        }

        // Date columns validation
        const dateColumns = ['O', 'P', 'Q']; // Annual Volume Date, Expiration Date, Finish Due Date
        for (let i = 2; i <= lastRow; i++) {
          dateColumns.forEach(col => {
            worksheet.getCell(`${col}${i}`).dataValidation = {
              type: 'date',
              operator: 'greaterThan',
              allowBlank: true,
              formulae: [new Date(1900, 0, 1)],
              showErrorMessage: true,
              errorStyle: 'warning',
              errorTitle: 'Invalid Date',
              error: 'Please enter a valid date (dd/mm/yyyy)',
              promptTitle: 'Date Format',
              prompt: 'Enter date in format: dd/mm/yyyy'
            };
          });
        }

        // Lock ID column
        worksheet.getColumn('A').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          if (rowNumber > 1) {
            cell.protection = { locked: true };
          }
        });

        // Unlock all other columns (use column index instead of letter for >26 columns)
        for (let colIdx = 2; colIdx <= columns.length; colIdx++) {
          worksheet.getColumn(colIdx).eachCell({ includeEmpty: true }, (cell, rowNumber) => {
            if (rowNumber > 1) {
              cell.protection = { locked: false };
            }
          });
        }

        // Protect worksheet
        await worksheet.protect('30625629', {
          selectLockedCells: true,
          selectUnlockedCells: true,
          formatCells: false,
          formatColumns: false,
          formatRows: false,
          insertRows: false,
          insertColumns: false,
          deleteRows: false,
          deleteColumns: false,
          sort: false,
          autoFilter: false,
          pivotTables: false
        });

        const result = await dialog.showSaveDialog(mainWindow, {
          title: 'Export Full Database (Manager)',
          defaultPath: `Full-Database-${new Date().toISOString().split('T')[0]}.xlsx`,
          filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
        });

        if (result.canceled) {
          resolve({ success: false, cancelled: true });
          return;
        }

        await workbook.xlsx.writeFile(result.filePath);
        resolve({ success: true, filePath: result.filePath });

      } catch (error) {
        reject(error);
      }
    });
  });
});

// Import Full Database for Manager
ipcMain.handle('import-forecast-manager', async () => {
  const dialogResult = await dialog.showOpenDialog(mainWindow, {
    title: 'Import Full Database (Manager)',
    buttonLabel: 'Import',
    properties: ['openFile'],
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
  });

  if (dialogResult.canceled || !dialogResult.filePaths || dialogResult.filePaths.length === 0) {
    return { success: false, cancelled: true };
  }

  try {
    const filePath = dialogResult.filePaths[0];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      return { success: false, error: 'No worksheet found in file' };
    }

    // Validate that first column is ID
    const headerRow = worksheet.getRow(1);
    const firstHeader = headerRow.getCell(1).value?.toString().trim();
    
    if (firstHeader !== 'ID') {
      return { 
        success: false, 
        error: 'Invalid file format. First column must be ID' 
      };
    }

    // Map headers to column indices
    const headerMap = {};
    headerRow.eachCell((cell, colNumber) => {
      const headerValue = cell.value?.toString().trim();
      if (headerValue) {
        headerMap[headerValue] = colNumber;
      }
    });

    // Define mapping from Excel headers to database columns
    const columnMapping = {
      'PN': 'pn',
      'PN Description': 'pn_description',
      'Supplier': 'supplier',
      'Tool Description': 'tool_description',
      'Tool Number ARB': 'tool_number_arb',
      'Asset Number': 'asset_number',
      'Tool Ownership': 'tool_ownership',
      'Customer': 'customer',
      'Tooling Life Qty': 'tooling_life_qty',
      'Produced': 'produced',
      'Remaining Life': 'remaining_tooling_life_pcs',
      '% Life': 'percent_tooling_life',
      'Annual Forecast': 'annual_volume_forecast',
      'Annual Volume Date': 'date_annual_volume',
      'Forecast Date': 'date_annual_volume',
      'Expiration Date': 'expiration_date',
      'Finish Due Date': 'finish_due_date',
      'Amount BRL': 'amount_brl',
      'Tool Quantity': 'tool_quantity',
      'Bailment Signed': 'bailment_agreement_signed',
      'Tooling Book': 'tooling_book',
      'Disposition': 'disposition',
      'Status': 'status',
      'Steps': 'steps',
      'BU': 'bu',
      'Category': 'category',
      'Owner': 'cummins_responsible',
      'STIM': 'stim_tooling_management',
      'VPCR': 'vpcr'
    };

    // Date columns for validation
    const dateColumns = ['Annual Volume Date', 'Forecast Date', 'Expiration Date', 'Finish Due Date'];
    
    // Numeric columns for validation
    const numericColumns = ['Tooling Life Qty', 'Produced', 'Remaining Life', '% Life', 'Annual Forecast', 'Amount BRL', 'Tool Quantity'];

    let updatedCount = 0;
    const errors = [];
    const importDateStamp = formatDateToBR(new Date());

    // Skip header row (row 1)
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      
      const id = row.getCell(headerMap['ID']).value;

      // Skip if no ID
      if (!id) continue;

      try {
        const updateData = {};

        // Process each mapped column
        for (const [excelHeader, dbColumn] of Object.entries(columnMapping)) {
          const colIndex = headerMap[excelHeader];
          if (!colIndex) continue; // Column not found in file
          
          const cellValue = row.getCell(colIndex).value;
          
          // Skip empty cells
          if (cellValue === null || cellValue === undefined || cellValue === '') continue;

          // Validate and process date columns
          if (dateColumns.includes(excelHeader)) {
            const normalizedDate = normalizeDateInputToISO(cellValue);
            if (cellValue && !normalizedDate) {
              errors.push(`Row ${rowNumber} (ID ${id}): Invalid date format for ${excelHeader}`);
              continue;
            }
            if (normalizedDate) {
              updateData[dbColumn] = normalizedDate;
            }
          }
          // Validate numeric columns
          else if (numericColumns.includes(excelHeader)) {
            const strValue = cellValueToString(cellValue);
            if (strValue !== '') {
              const numValue = parseFloat(strValue.replace(/,/g, '.'));
              if (isNaN(numValue)) {
                errors.push(`Row ${rowNumber} (ID ${id}): ${excelHeader} must be a number`);
                continue;
              }
              updateData[dbColumn] = strValue;
            }
          }
          // Text columns
          else {
            updateData[dbColumn] = cellValueToString(cellValue);
          }
        }

        // Only update if there's data to update
        if (Object.keys(updateData).length > 0) {
          const result = await executeToolingUpdate(id, updateData, 1, {
            commentTimestamp: importDateStamp,
            skipWhenNoTrackedChanges: true
          });
          if (result?.changes > 0) {
            updatedCount++;
          }
        }

      } catch (error) {
        errors.push(`Row ${rowNumber} (ID ${id}): ${error.message}`);
      }
    }

    return { 
      success: true, 
      updatedCount,
      errors: errors.length > 0 ? errors : undefined
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
});

app.on('window-all-closed', () => {
  if (db) {
    db.close();
  }
  if (process.platform !== 'darwin') {
    app.quit();
  }

});
