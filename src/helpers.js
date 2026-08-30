import path from "node:path";

export const REQUIRED_HEADERS = [
  "Sendungsreferenz",
  "KUNDENREFERENZ / AUF LABEL ANZEIGEN",
  "VOR-UND NACHNAME",
  "PLZ",
  "ORT",
  "STRASSE",
  "NR",
  "E-MAIL ADRESSE DES KUNDEN",
  "Retouren-Empfänger",
];

export const TRACKING_HEADER = "Sendungsnummer";

export function asText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && "text" in value) return String(value.text ?? "").trim();
  return String(value).trim();
}

export function normalizePostalCode(value) {
  const text = asText(value).replace(/\.0$/, "");
  if (!text) return "";
  return /^\d{4}$/.test(text) ? `0${text}` : text;
}

export const REQUIRED_RECORD_FIELDS = [
  ["shipmentReference", "Sendungsreferenz"],
  ["customerReference", "KUNDENREFERENZ / AUF LABEL ANZEIGEN"],
  ["name1", "VOR-UND NACHNAME"],
  ["postalCode", "PLZ"],
  ["city", "ORT"],
  ["street", "STRASSE"],
  ["streetNumber", "NR"],
  ["email", "E-MAIL ADRESSE DES KUNDEN"],
  ["receiver", "Retouren-Empfänger"],
];

export function validateRecord(record) {
  const errors = REQUIRED_RECORD_FIELDS
    .filter(([field]) => !asText(record[field]))
    .map(([, label]) => `${label} 为空`);
  if (record.postalCode && !/^\d{5}$/.test(record.postalCode)) {
    errors.push(`PLZ 必须是 5 位数字（当前值：${record.postalCode}）`);
  }
  return errors;
}

export function headerMap(headerRow) {
  const map = new Map();
  headerRow.eachCell({ includeEmpty: true }, (cell, column) => {
    const title = asText(cell.value);
    if (title) map.set(title, column);
  });
  const missing = REQUIRED_HEADERS.filter((header) => !map.has(header));
  if (missing.length) throw new Error(`Excel 缺少必需列：${missing.join("、")}`);
  return map;
}

export function rowToRecord(row, columns) {
  const read = (header) => {
    const column = columns.get(header);
    return column ? asText(row.getCell(column).value) : "";
  };
  return {
    excelRow: row.number,
    shipmentReference: read("Sendungsreferenz"),
    customerReference: read("KUNDENREFERENZ / AUF LABEL ANZEIGEN"),
    name1: read("VOR-UND NACHNAME"),
    name2: read("NAMENSZUSATZ 1"),
    name3: read("NAMENSZUSATZ 2"),
    postalCode: normalizePostalCode(read("PLZ")),
    city: read("ORT"),
    street: read("STRASSE"),
    streetNumber: read("NR"),
    email: read("E-MAIL ADRESSE DES KUNDEN"),
    receiver: read("Retouren-Empfänger"),
  };
}

export function outputPathFor(inputPath) {
  const parsed = path.parse(inputPath);
  return path.join(parsed.dir, `${parsed.name}_已填写运单号${parsed.ext || ".xlsx"}`);
}

export function parseShipmentResult(text) {
  const shipmentNumber = text.match(/Sendungsnummer:\s*(\d+)/i)?.[1] ?? "";
  const customerReference = text.match(/Kundenreferenz:\s*([^\n\r]+)/i)?.[1]?.trim() ?? "";
  return { shipmentNumber, customerReference };
}

