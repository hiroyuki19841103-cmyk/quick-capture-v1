const WORKBOOK_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const WORKBOOK_FILE_NAME = "Digital_NOLTY_Memos.xlsx";
const WORKBOOK_FORMAT_VERSION = "editable-v1";
const WORKBOOK_TABLE_NAME = "DigitalNoltyMemoTable";
const MEMO_TEMPLATE_URL = new URL("./Digital_NOLTY_Memos_EDITABLE_TEMPLATE.xlsx", import.meta.url).href;
const HEADERS = ["メモID", "種類", "週開始日", "対象日", "メモ内容（編集可）", "更新日時", "更新元", "版"];

function requireJsZip() {
  if (!globalThis.JSZip) throw new Error("Excel同期部品を読み込めませんでした");
  return globalThis.JSZip;
}

function escapeXml(value = "") {
  return String(value).replace(/[&<>"']/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&apos;"
  }[character]));
}

function normalizeRecord(record) {
  const memoId = String(record?.memoId || "").trim();
  const memoType = /^WEEK_LEFT-\d{4}-\d{2}-\d{2}$/.test(memoId)
    ? "WEEK_LEFT"
    : /^WEEK_RIGHT-\d{4}-\d{2}-\d{2}$/.test(memoId)
      ? "WEEK_RIGHT"
      : /^WEEK-\d{4}-\d{2}-\d{2}$/.test(memoId) ? "WEEK"
        : /^DAILY_VIEW-\d{4}-\d{2}-\d{2}$/.test(memoId) ? "DAILY_VIEW" : "DAY";
  return {
    memoId,
    memoType,
    weekStart: String(record?.weekStart || "").slice(0, 10),
    targetDate: ["DAY", "DAILY_VIEW"].includes(memoType) ? String(record?.targetDate || "").slice(0, 10) : "",
    memoText: String(record?.memoText || ""),
    updatedAt: String(record?.updatedAt || ""),
    updatedBy: String(record?.updatedBy || ""),
    revision: Math.max(1, Number.parseInt(record?.revision, 10) || 1)
  };
}

function sortedRecords(records) {
  const typeOrder = { WEEK_LEFT: 0, WEEK_RIGHT: 1, WEEK: 2, DAY: 3, DAILY_VIEW: 4 };
  return [...records]
    .map(normalizeRecord)
    .filter((record) => record.memoId)
    .sort((a, b) => a.weekStart.localeCompare(b.weekStart)
      || (typeOrder[a.memoType] - typeOrder[b.memoType])
      || a.targetDate.localeCompare(b.targetDate));
}

function inlineCell(reference, value, style = 0) {
  const preserve = /^\s|\s$|\n/.test(String(value));
  return `<c r="${reference}" s="${style}" t="inlineStr"><is><t${preserve ? ' xml:space="preserve"' : ""}>${escapeXml(value)}</t></is></c>`;
}

function numericCell(reference, value, style = 0) {
  return `<c r="${reference}" s="${style}"><v>${Number(value) || 0}</v></c>`;
}

function sheetXml(records, tableRelationshipId) {
  const rows = [
    `<row r="1" ht="28" customHeight="1">${HEADERS.map((header, index) => inlineCell(`${String.fromCharCode(65 + index)}1`, header, 3)).join("")}</row>`
  ];

  sortedRecords(records).forEach((record, index) => {
    const row = index + 2;
    rows.push(`<row r="${row}" ht="42" customHeight="1">${[
      inlineCell(`A${row}`, record.memoId, 6),
      inlineCell(`B${row}`, record.memoType, 6),
      inlineCell(`C${row}`, record.weekStart, 6),
      inlineCell(`D${row}`, record.targetDate, 6),
      inlineCell(`E${row}`, record.memoText, 10),
      inlineCell(`F${row}`, record.updatedAt, 13),
      inlineCell(`G${row}`, record.updatedBy, 13),
      numericCell(`H${row}`, record.revision, 13)
    ].join("")}</row>`);
  });

  if (!records.length) {
    rows.push(`<row r="2" ht="42" customHeight="1">${[
      inlineCell("A2", "", 6),
      inlineCell("B2", "", 6),
      inlineCell("C2", "", 6),
      inlineCell("D2", "", 6),
      inlineCell("E2", "", 10),
      inlineCell("F2", "", 13),
      inlineCell("G2", "", 13),
      numericCell("H2", 0, 13)
    ].join("")}</row>`);
  }

  const lastRow = Math.max(2, records.length + 1);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:H${lastRow}"/>
  <sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="24" customWidth="1"/>
    <col min="2" max="2" width="14" customWidth="1"/>
    <col min="3" max="4" width="14" customWidth="1"/>
    <col min="5" max="5" width="64" customWidth="1"/>
    <col min="6" max="6" width="25" customWidth="1"/>
    <col min="7" max="7" width="18" customWidth="1"/>
    <col min="8" max="8" width="8" customWidth="1"/>
  </cols>
  <sheetData>${rows.join("")}</sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <tableParts count="1"><tablePart r:id="${escapeXml(tableRelationshipId)}"/></tableParts>
</worksheet>`;
}

function tableXml(recordCount) {
  const lastRow = Math.max(2, recordCount + 1);
  const columns = HEADERS.map((header, index) => `<tableColumn id="${index + 1}" name="${escapeXml(header)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="${WORKBOOK_TABLE_NAME}" displayName="${WORKBOOK_TABLE_NAME}" ref="A1:H${lastRow}" headerRowCount="1" totalsRowCount="0" totalsRowShown="0">
  <tableColumns count="8">${columns}</tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`;
}

async function workbookTemplateBytes() {
  const response = await fetch(MEMO_TEMPLATE_URL);
  if (!response.ok) throw new Error(`編集用Excelテンプレートを読み込めませんでした（HTTP ${response.status}）`);
  return response.arrayBuffer();
}

function tableRelationshipId(relationshipsXml) {
  const relationship = String(relationshipsXml || "").match(/<Relationship\b[^>]*\bType=["'][^"']*\/table["'][^>]*>/i)?.[0] || "";
  const id = relationship.match(/\bId=["']([^"']+)["']/i)?.[1] || "";
  if (!id) throw new Error("編集用Excelテンプレートのテーブル関連付けが見つかりません");
  return id;
}

export async function buildMemoWorkbook(records, templateBytesOverride = null) {
  const JSZip = requireJsZip();
  const templateBytes = templateBytesOverride || await workbookTemplateBytes();
  const zip = await JSZip.loadAsync(templateBytes);
  const normalized = sortedRecords(records);
  const relationshipsFile = zip.file("xl/worksheets/_rels/sheet1.xml.rels");
  if (!relationshipsFile) throw new Error("編集用Excelテンプレートの構成が不正です");
  const relationshipId = tableRelationshipId(await relationshipsFile.async("text"));
  zip.file("xl/worksheets/sheet1.xml", sheetXml(normalized, relationshipId));
  zip.file("xl/tables/table1.xml", tableXml(normalized.length));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

function parseXml(xmlText, label) {
  const document = new DOMParser().parseFromString(xmlText, "application/xml");
  if (document.getElementsByTagName("parsererror").length) throw new Error(`${label}のXMLを解析できません`);
  return document;
}

function sharedStringsFrom(document) {
  if (!document) return [];
  return [...document.getElementsByTagName("si")].map((node) => [...node.getElementsByTagName("t")].map((text) => text.textContent || "").join(""));
}

function cellValue(cell, sharedStrings) {
  const type = cell.getAttribute("t") || "";
  if (type === "inlineStr") return [...cell.getElementsByTagName("t")].map((node) => node.textContent || "").join("");
  const raw = cell.getElementsByTagName("v")[0]?.textContent || "";
  if (type === "s") return sharedStrings[Number.parseInt(raw, 10)] || "";
  return raw;
}

export async function parseMemoWorkbook(bytes) {
  const JSZip = requireJsZip();
  const zip = await JSZip.loadAsync(bytes);
  const sheetFile = zip.file("xl/worksheets/sheet1.xml");
  if (!sheetFile) throw new Error("NOLTY_MEMOSシートが見つかりません");
  const sheetDocument = parseXml(await sheetFile.async("text"), "NOLTY_MEMOS");
  const sharedFile = zip.file("xl/sharedStrings.xml");
  const sharedDocument = sharedFile ? parseXml(await sharedFile.async("text"), "sharedStrings") : null;
  const sharedStrings = sharedStringsFrom(sharedDocument);
  const records = [];

  [...sheetDocument.getElementsByTagName("row")].forEach((row) => {
    if (Number(row.getAttribute("r") || 0) <= 1) return;
    const values = Array(8).fill("");
    [...row.getElementsByTagName("c")].forEach((cell) => {
      const reference = cell.getAttribute("r") || "";
      const column = reference.charCodeAt(0) - 65;
      if (column >= 0 && column < values.length) values[column] = cellValue(cell, sharedStrings);
    });
    const [memoId, memoType, weekStart, targetDate, memoText, updatedAt, updatedBy, revision] = values;
    if (!memoId || !/^(WEEK|WEEK_LEFT|WEEK_RIGHT|DAY|DAILY_VIEW)-\d{4}-\d{2}-\d{2}$/.test(memoId)) return;
    records.push(normalizeRecord({ memoId, memoType, weekStart, targetDate, memoText, updatedAt, updatedBy, revision }));
  });
  const tableFile = zip.file("xl/tables/table1.xml");
  const tableText = tableFile ? await tableFile.async("text") : "";
  Object.defineProperty(records, "workbookFormatVersion", {
    value: new RegExp(`\\b(?:name|displayName)=["']${WORKBOOK_TABLE_NAME}["']`).test(tableText)
      ? WORKBOOK_FORMAT_VERSION
      : "legacy",
    enumerable: false
  });
  return records;
}

export { WORKBOOK_FILE_NAME, WORKBOOK_MIME, WORKBOOK_FORMAT_VERSION };
