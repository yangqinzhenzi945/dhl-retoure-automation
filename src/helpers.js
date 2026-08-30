import path from "node:path";

export const REQUIRED_HEADERS = [
  "Sendungsreferenz",
  "KUNDENREFERENZ / AUF LABEL ANZEIGEN",
  "VOR-UND NACHNAME",
  "NAMENSZUSATZ 1",
  "NAMENSZUSATZ 2",
  "PLZ",
  "ORT",
  "STRASSE",
  "NR",
  "E-MAIL ADRESSE DES KUNDEN",
  "Retouren-Empf盲nger",
];

export const TRACKING_HEADER = "Sendungsnummer";

export function asText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && "text" in value) return String(value.text ?? "").trim();
  return String(value).trim();
}

export function normalizePostalCode(value) {
  const text = asText(value).replace(/\.0$/, "");
  return text.padStart(5, "0");
}

export function headerMap(headerRow) {
  const map = new Map();
  headerRow.eachCell({ includeEmpty: true }, (cell, column) => {
    const title = asText(cell.value);
    if (title) map.set(title, column);
  });
  const missing = REQUIRED_HEADERS.filter((header) => !map.has(header));
  if (missing.length) throw new Error(`Excel 缂哄皯蹇呴渶鍒楋細${missing.join("銆?)}`);
  return map;
}

export function rowToRecord(row, columns) {
  const read = (header) => asText(row.getCell(columns.get(header)).value);
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
    receiver: read("Retouren-Empf盲nger"),
  };
}

export function outputPathFor(inputPath) {
  const parsed = path.parse(inputPath);
  return path.join(parsed.dir, `${parsed.name}_宸插～鍐欒繍鍗曞彿${parsed.ext || ".xlsx"}`);
}

export function parseShipmentResult(text) {
  const shipmentNumber = text.match(/Sendungsnummer:\s*(\d+)/i)?.[1] ?? "";
  const customerReference = text.match(/Kundenreferenz:\s*([^\n\r]+)/i)?.[1]?.trim() ?? "";
  return { shipmentNumber, customerReference };
}

