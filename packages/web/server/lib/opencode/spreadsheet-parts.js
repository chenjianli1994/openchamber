import { read, utils } from '../../../../ui/node_modules/xlsx/xlsx.mjs';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const SPREADSHEET_MIME_TYPES = new Set([
  'application/vnd.ms-excel',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.oasis.opendocument.spreadsheet',
  'text/csv',
]);

const SPREADSHEET_EXTENSIONS = new Set(['.xls', '.xlsx', '.ods', '.csv']);

const MAX_SHEETS = 3;
const MAX_ROWS = 200;
const MAX_COLUMNS = 20;
const MAX_CELL_LENGTH = 500;

const getFileExtension = (filename = '') => {
  const lastDot = filename.lastIndexOf('.');
  return lastDot >= 0 ? filename.slice(lastDot).toLowerCase() : '';
};

const escapeMarkdownCell = (value) =>
  value.replace(/\|/g, '\\|').replace(/\r?\n/g, '<br />');

const truncateCell = (value) => {
  const normalized = String(value ?? '').trim();
  if (normalized.length <= MAX_CELL_LENGTH) {
    return normalized;
  }
  return `${normalized.slice(0, MAX_CELL_LENGTH)}...`;
};

const toTextDataUrl = (text) =>
  `data:text/plain;base64,${Buffer.from(text, 'utf8').toString('base64')}`;

const renderWorksheet = (sheetName, worksheet) => {
  const rows = utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: '',
  });

  if (rows.length === 0) {
    return `## Sheet: ${sheetName}\n\n(Empty sheet)\n`;
  }

  const limitedRows = rows.slice(0, MAX_ROWS).map((row) =>
    Array.isArray(row)
      ? row.slice(0, MAX_COLUMNS).map((cell) => truncateCell(cell))
      : [],
  );

  const columnCount = limitedRows.reduce((max, row) => Math.max(max, row.length), 0);
  if (columnCount === 0) {
    return `## Sheet: ${sheetName}\n\n(Empty sheet)\n`;
  }

  const normalizedRows = limitedRows.map((row) => {
    const padded = [...row];
    while (padded.length < columnCount) padded.push('');
    return padded;
  });

  const headerRow = normalizedRows[0].map((value, index) => value || `Column ${index + 1}`);
  const separatorRow = headerRow.map(() => '---');
  const bodyRows = normalizedRows.slice(1);

  const markdownRows = [
    `| ${headerRow.map(escapeMarkdownCell).join(' | ')} |`,
    `| ${separatorRow.join(' | ')} |`,
    ...bodyRows.map((row) => `| ${row.map(escapeMarkdownCell).join(' | ')} |`),
  ];

  const truncationNotes = [];
  if (rows.length > MAX_ROWS) {
    truncationNotes.push(`kept first ${MAX_ROWS} rows`);
  }
  if (rows.some((row) => Array.isArray(row) && row.length > MAX_COLUMNS)) {
    truncationNotes.push(`kept first ${MAX_COLUMNS} columns`);
  }

  const note = truncationNotes.length > 0
    ? `\n\n_Truncated: ${truncationNotes.join(', ')}._\n`
    : '\n';

  return `## Sheet: ${sheetName}\n\n${markdownRows.join('\n')}${note}`;
};

const buildSpreadsheetText = (filename, buffer) => {
  const workbook = read(buffer, {
    type: 'buffer',
    dense: true,
  });

  const selectedSheetNames = workbook.SheetNames.slice(0, MAX_SHEETS);
  const sections = selectedSheetNames.map((sheetName) =>
    renderWorksheet(sheetName, workbook.Sheets[sheetName]),
  );

  const notes = [];
  if (workbook.SheetNames.length > MAX_SHEETS) {
    notes.push(`kept first ${MAX_SHEETS} sheets out of ${workbook.SheetNames.length}`);
  }

  return [
    '# Spreadsheet attachment',
    '',
    `Original filename: ${filename}`,
    'Converted to plain text before sending because the current model route does not support raw spreadsheet file parts.',
    notes.length > 0 ? `Notes: ${notes.join('; ')}.` : '',
    '',
    ...sections,
  ].filter(Boolean).join('\n');
};

const buildFallbackText = (filename, reason) => [
  '# Spreadsheet attachment',
  '',
  `Original filename: ${filename}`,
  'The server intercepted this spreadsheet attachment before forwarding it to the model route.',
  `Conversion note: ${reason}`,
].join('\n');

const decodeDataUrl = (url) => {
  if (typeof url !== 'string' || !url.startsWith('data:')) return null;
  const commaIndex = url.indexOf(',');
  if (commaIndex < 0) return null;

  const metadata = url.slice(5, commaIndex);
  const payload = url.slice(commaIndex + 1);
  const [mime = ''] = metadata.split(';', 1);
  const isBase64 = metadata.toLowerCase().includes(';base64');
  const buffer = isBase64
    ? Buffer.from(payload, 'base64')
    : Buffer.from(decodeURIComponent(payload), 'utf8');

  return {
    mime: mime.toLowerCase(),
    buffer,
  };
};

const shouldConvertSpreadsheetPart = (part) => {
  if (!part || part.type !== 'file') return false;
  const mime = String(part.mime || '').trim().toLowerCase();
  const filename = typeof part.filename === 'string' ? part.filename : '';
  const url = typeof part.url === 'string' ? part.url : '';
  const dataMime = url.startsWith('data:') ? decodeDataUrl(url)?.mime : '';
  if (mime === 'text/plain' && dataMime === 'text/plain') return false;
  return SPREADSHEET_MIME_TYPES.has(mime)
    || SPREADSHEET_MIME_TYPES.has(dataMime)
    || SPREADSHEET_EXTENSIONS.has(getFileExtension(filename));
};

const convertSpreadsheetPart = (part) => {
  const filename = typeof part.filename === 'string' && part.filename.trim().length > 0
    ? part.filename.trim()
    : `spreadsheet${getFileExtension(part.filename || '') || '.xlsx'}`;

  const decoded = decodeDataUrl(part.url);
  if (!decoded) {
    if (typeof part.url === 'string' && part.url.startsWith('file:')) {
      try {
        return {
          ...part,
          mime: 'text/plain',
          url: toTextDataUrl(buildSpreadsheetText(filename, readFileSync(fileURLToPath(part.url)))),
        };
      } catch (error) {
        const reason = error instanceof Error && error.message
          ? error.message
          : 'Spreadsheet file could not be read.';
        return {
          ...part,
          mime: 'text/plain',
          url: toTextDataUrl(buildFallbackText(filename, reason)),
        };
      }
    }

    return {
      ...part,
      mime: 'text/plain',
      url: toTextDataUrl(buildFallbackText(filename, 'Spreadsheet content was not available as a data URL.')),
    };
  }

  if (decoded.mime === 'text/plain') {
    return part;
  }

  try {
    return {
      ...part,
      mime: 'text/plain',
      url: toTextDataUrl(buildSpreadsheetText(filename, decoded.buffer)),
    };
  } catch (error) {
    const reason = error instanceof Error && error.message
      ? error.message
      : 'Spreadsheet parsing failed.';
    return {
      ...part,
      mime: 'text/plain',
      url: toTextDataUrl(buildFallbackText(filename, reason)),
    };
  }
};

const transformFileParts = (parts) => {
  if (!Array.isArray(parts)) return parts;
  let changed = false;
  const nextParts = parts.map((part) => {
    if (!shouldConvertSpreadsheetPart(part)) return part;
    const converted = convertSpreadsheetPart(part);
    changed = changed || converted !== part || converted.mime !== part.mime || converted.url !== part.url;
    return converted;
  });
  return changed ? nextParts : parts;
};

export const rewriteSpreadsheetPartsInPromptBody = (body) => {
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    return body;
  }

  let changed = false;
  const nextBody = { ...body };

  const rewrittenParts = transformFileParts(body.parts);
  if (rewrittenParts !== body.parts) {
    nextBody.parts = rewrittenParts;
    changed = true;
  }

  if (Array.isArray(body.messages)) {
    const rewrittenMessages = body.messages.map((message) => {
      if (!message || typeof message !== 'object') return message;
      const nextPartsForMessage = transformFileParts(message.parts);
      if (nextPartsForMessage === message.parts) return message;
      changed = true;
      return {
        ...message,
        parts: nextPartsForMessage,
      };
    });
    if (changed) {
      nextBody.messages = rewrittenMessages;
    }
  }

  if (changed) {
    console.log('[proxy] converted spreadsheet file parts to text/plain before forwarding prompt');
    return nextBody;
  }

  return body;
};
