import { read, utils, type WorkSheet } from "xlsx"

const SPREADSHEET_MIME_TYPES = new Set([
  "application/vnd.ms-excel",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.oasis.opendocument.spreadsheet",
  "text/csv",
])

const SPREADSHEET_EXTENSIONS = new Set([".xls", ".xlsx", ".ods", ".csv"])
const MAX_SHEETS = 3
const MAX_ROWS = 200
const MAX_COLUMNS = 20
const MAX_CELL_LENGTH = 500

const getFileExtension = (filename: string): string => {
  const lastDot = filename.lastIndexOf(".")
  return lastDot >= 0 ? filename.slice(lastDot).toLowerCase() : ""
}

const escapeMarkdownCell = (value: string): string =>
  value.replace(/\|/g, "\\|").replace(/\r?\n/g, "<br />")

const truncateCell = (value: unknown): string => {
  const normalized = String(value ?? "").trim()
  if (normalized.length <= MAX_CELL_LENGTH) {
    return normalized
  }
  return `${normalized.slice(0, MAX_CELL_LENGTH)}...`
}

const toTextDataUrl = (text: string): string => {
  const bytes = new TextEncoder().encode(text)
  let binary = ""
  const chunkSize = 0x8000
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.slice(index, index + chunkSize))
  }
  return `data:text/plain;base64,${btoa(binary)}`
}

const renderWorksheet = (sheetName: string, worksheet: WorkSheet): string => {
  const rows = utils.sheet_to_json<(string | number | boolean | null)[]>(worksheet, {
    header: 1,
    raw: false,
    defval: "",
  })

  if (rows.length === 0) {
    return `## Sheet: ${sheetName}\n\n(Empty sheet)\n`
  }

  const limitedRows = rows.slice(0, MAX_ROWS).map((row) =>
    Array.isArray(row)
      ? row.slice(0, MAX_COLUMNS).map((cell) => truncateCell(cell))
      : [],
  )
  const columnCount = limitedRows.reduce((max, row) => Math.max(max, row.length), 0)

  if (columnCount === 0) {
    return `## Sheet: ${sheetName}\n\n(Empty sheet)\n`
  }

  const normalizedRows = limitedRows.map((row) => {
    const padded = [...row]
    while (padded.length < columnCount) {
      padded.push("")
    }
    return padded
  })

  const firstRow = normalizedRows[0]
  const headerRow = firstRow.map((value, index) => value || `Column ${index + 1}`)
  const separatorRow = headerRow.map(() => "---")
  const bodyRows = normalizedRows.slice(1)

  const markdownRows = [
    `| ${headerRow.map(escapeMarkdownCell).join(" | ")} |`,
    `| ${separatorRow.join(" | ")} |`,
    ...bodyRows.map((row) => `| ${row.map(escapeMarkdownCell).join(" | ")} |`),
  ]

  const truncationNotes: string[] = []
  if (rows.length > MAX_ROWS) {
    truncationNotes.push(`kept first ${MAX_ROWS} rows`)
  }
  if (rows.some((row) => Array.isArray(row) && row.length > MAX_COLUMNS)) {
    truncationNotes.push(`kept first ${MAX_COLUMNS} columns`)
  }

  const note = truncationNotes.length > 0
    ? `\n\n_Truncated: ${truncationNotes.join(", ")}._\n`
    : "\n"

  return `## Sheet: ${sheetName}\n\n${markdownRows.join("\n")}${note}`
}

type SpreadsheetLike = {
  name?: string
  filename?: string
  type?: string
  mime?: string
}

const getSpreadsheetName = (file: SpreadsheetLike): string => file.name ?? file.filename ?? ""
const getSpreadsheetMime = (file: SpreadsheetLike): string => file.type ?? file.mime ?? ""

const buildSpreadsheetText = (filename: string, buffer: ArrayBuffer): string => {
  const workbook = read(buffer, {
    type: "array",
    dense: true,
  })

  const selectedSheetNames = workbook.SheetNames.slice(0, MAX_SHEETS)
  const sections = selectedSheetNames.map((sheetName) =>
    renderWorksheet(sheetName, workbook.Sheets[sheetName]),
  )

  const notes: string[] = []
  if (workbook.SheetNames.length > MAX_SHEETS) {
    notes.push(`kept first ${MAX_SHEETS} sheets out of ${workbook.SheetNames.length}`)
  }

  return [
    "# Spreadsheet attachment",
    "",
    `Original filename: ${filename}`,
    "Converted to plain text before sending because the current model route does not support raw spreadsheet file parts.",
    notes.length > 0 ? `Notes: ${notes.join("; ")}.` : "",
    "",
    ...sections,
  ].filter(Boolean).join("\n")
}

const dataUrlToArrayBuffer = (dataUrl: string): ArrayBuffer | null => {
  if (!dataUrl.startsWith("data:")) {
    return null
  }

  const commaIndex = dataUrl.indexOf(",")
  if (commaIndex < 0) {
    return null
  }

  const meta = dataUrl.slice(0, commaIndex).toLowerCase()
  const payload = dataUrl.slice(commaIndex + 1)
  const binary = meta.endsWith(";base64")
    ? atob(payload)
    : decodeURIComponent(payload)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes.buffer
}

export const isSpreadsheetAttachment = (file: SpreadsheetLike): boolean => {
  const name = getSpreadsheetName(file)
  const type = getSpreadsheetMime(file)
  const normalizedType = type.trim().toLowerCase()
  if (normalizedType && SPREADSHEET_MIME_TYPES.has(normalizedType)) {
    return true
  }
  return SPREADSHEET_EXTENSIONS.has(getFileExtension(name))
}

export const convertSpreadsheetFileToTextAttachment = async (file: File): Promise<{
  file: File
  dataUrl: string
  mimeType: string
  filename: string
  size: number
}> => {
  const content = buildSpreadsheetText(file.name, await file.arrayBuffer())

  const normalizedFile = new File([content], file.name, { type: "text/plain" })

  return {
    file: normalizedFile,
    dataUrl: toTextDataUrl(content),
    mimeType: "text/plain",
    filename: file.name,
    size: normalizedFile.size,
  }
}

export const convertSpreadsheetDataUrlToTextFilePart = (file: {
  mime: string
  filename?: string
  url: string
}): { mime: string; filename?: string; url: string } | null => {
  const buffer = dataUrlToArrayBuffer(file.url)
  if (!buffer) {
    return null
  }

  const filename = file.filename ?? "spreadsheet"
  const content = buildSpreadsheetText(filename, buffer)
  return {
    mime: "text/plain",
    filename,
    url: toTextDataUrl(content),
  }
}
