// src/converters/fillOperativeTemplate.ts
import * as XLSX from "xlsx"; 
import type { NormalizedLine, OperativeHeader } from "./types";
//import type { NormalizedLine, OperativeHeader } from "./types";

const HEADER_ROW_INDEX = 8; // row 9 in Excel

function mdy(iso: string) {
  const [y, m, d] = iso.split("-");
  return `${m}/${d}/${y}`;
}

function tf(v: boolean | undefined) {
  return v ? "TRUE" : "FALSE";
}

export function fillOperativeTemplate(
  templateBuf: ArrayBuffer,
  header: OperativeHeader,
  lines: NormalizedLine[]
): XLSX.WorkBook {
  // 1️⃣ Load the native template
  const wb = XLSX.read(templateBuf, { type: "array" });
  const ws = wb.Sheets["SO Template"];
  if (!ws) {
    throw new Error("SO Template sheet not found in template");
  }

  // 2️⃣ Clear existing data rows ONLY
  const range = XLSX.utils.decode_range(ws["!ref"]!);
  for (let r = HEADER_ROW_INDEX + 1; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      delete ws[addr];
    }
  }

  // 3️⃣ Build new data rows (column order MUST match template)
  const dataRows = lines.map((line) => {
    const qty = line.impsUnits ?? 0;
    const costMethod = header.costMethod ?? "CPM";

    const netUnitCost =
      costMethod === "CPM" && qty > 0
        ? line.netInvestment / (qty / 1000)
        : line.netInvestment;

    return [
      header.mediaPlanName ?? "",
      header.sectionName ?? "",
      mdy(line.startDateISO),
      mdy(line.endDateISO),
      line.lineItemName,
      header.groupName ?? "",
      header.productName ?? "",
      0,
      netUnitCost,
      costMethod,
      header.unitType ?? "Impressions",
      qty,
      tf(header.canOutput),
      tf(header.canInvoice),
      header.billableThirdPartyServer ?? "",
      header.dma ?? "",
      header.state ?? "",
      header.congressionalDistrict ?? "",
      "", "", "", "", "",
      "",
      "",
      line.targeting ?? "",
      "",
    ];
  });

  // 4️⃣ Inject rows into existing sheet
  XLSX.utils.sheet_add_aoa(ws, dataRows, {
    origin: { r: HEADER_ROW_INDEX + 1, c: 0 },
  });

  return wb;
}
