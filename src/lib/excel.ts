
import ExcelJS from "exceljs";
import { Inspection, MODEL_PREFIX } from "./types";

const cell = (
  ws: ExcelJS.Worksheet,
  address: string,
  value: ExcelJS.CellValue
) => {
  ws.getCell(address).value = value;
};

// Format tanggal: Tanggal Bulan Tahun
const formatDate = (value: string) => {
  if (!value) return "";

  const date = new Date(`${value}T00:00:00`);

  return new Intl.DateTimeFormat("id-ID", {
    day: "2-digit",
    month: "long",
    year: "numeric",
  }).format(date);
};

export async function buildWorkbook(data: Inspection) {
  const res = await fetch("/templates/OQC-rev3-template.xlsx");

  if (!res.ok) {
    throw new Error("Template Excel tidak ditemukan");
  }

  const bytes = await res.arrayBuffer();
  const wb = new ExcelJS.Workbook();

  await wb.xlsx.load(bytes);

  const ws = wb.getWorksheet("OQC rev.3");

  if (!ws) {
    throw new Error("Sheet OQC rev.3 tidak ditemukan");
  }

  // Inspection Information
  cell(ws, "D4", `: ${data.inspector}`);
  cell(ws, "D5", `: ${formatDate(data.inspectionDate)}`);
  cell(ws, "D6", `: O/S Rear View Mirror ${data.line.join(" / ")}`);

  cell(ws, "M4", `: ${formatDate(data.lotDate)}`);
  cell(ws, "M5", `: ${data.shift}`);
  cell(ws, "M6", `: P D I `);

  // Prepared mengikuti Inspector Name
  cell(ws, "J57", ` ${data.inspector}`);

  // Inspection Items
  for (let i = 0; i < 20; i++) {
    const row = 11 + i;
    const item = data.items[i];

    if (!item) {
      ["C", "D", "G", "I", "K", "M"].forEach((col) => {
        cell(ws, `${col}${row}`, null);
      });
      continue;
    }

    const prefix = item.model ? MODEL_PREFIX[item.model] : "";
    const part = prefix + (item.suffix || "").toUpperCase();

    const sample = Number(item.sampleQty) || 0;
    const ng = Number(item.ng) || 0;

    cell(ws, `C${row}`, item.model);
    cell(ws, `D${row}`, part);
    cell(ws, `G${row}`, sample + ng);
    cell(ws, `I${row}`, sample);
    cell(ws, `K${row}`, ng);
    cell(ws, `M${row}`, item.detailNg);
  }

  // Total
  const totalSample = data.items.reduce(
    (s, x) => s + (Number(x.sampleQty) || 0),
    0
  );

  const totalNg = data.items.reduce(
    (s, x) => s + (Number(x.ng) || 0),
    0
  );

  cell(ws, "G31", totalSample + totalNg);
  cell(ws, "I31", totalSample);
  cell(ws, "K31", totalNg);

  const buffer = await wb.xlsx.writeBuffer();

  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export function safeFileName(s: string) {
  return s
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}
