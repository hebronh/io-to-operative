// src/converters/convertToOperativeTemplate.ts
import * as XLSX from "xlsx";

export type NormalizedLine = {
  market: string;
  startDateISO: string; // YYYY-MM-DD
  endDateISO: string;
  lineItemName: string;
  netInvestment: number;
  impsUnits: number;
  targeting: string;
};

export type OperativeHeader = {
  orderId: string;
  mediaPlanName?: string;
  sectionName?: string;
  productName?: string;
  groupName?: string;
  costMethod?: string;
  unitType?: string;
  canOutput?: boolean;
  canInvoice?: boolean;
  billableThirdPartyServer?: string;
  dma?: string;
  state?: string;
  congressionalDistrict?: string;
};

function mdy(iso: string) {
  const [y, m, d] = iso.split("-");
  return `${m}/${d}/${y}`;
}

function tf(v: boolean | undefined) {
  return v ? "TRUE" : "FALSE";
}

export function buildOperativeWorkbook(
  templateBuf: ArrayBuffer,           // ✅ pass template in
  header: OperativeHeader,
  lines: NormalizedLine[]
): XLSX.WorkBook {
  const wb = XLSX.read(templateBuf, { type: "array" });
  const ws = wb.Sheets["SO Template"];
  if (!ws) {
    throw new Error("Template missing SO Template sheet");
  }

  // ⚠️ must match the REAL template
  const HEADER_ROW_INDEX = 8;

  const mediaPlan = header.mediaPlanName ?? "";
  const section = header.sectionName ?? "";
  const product = header.productName ?? "";
  const costMethod = header.costMethod ?? "CPM";
  const unitType = header.unitType ?? "Impressions";

  const dataRows = lines.map((line) => {
    const qty = line.impsUnits ?? 0;

    const netUnitCost =
      costMethod.toUpperCase() === "CPM"
        ? qty > 0
          ? line.netInvestment / (qty / 1000)
          : 0
        : line.netInvestment;

    return [
      mediaPlan,
      section,
      mdy(line.startDateISO),
      mdy(line.endDateISO),
      line.lineItemName,
      header.groupName ?? "",
      product,
      0,
      netUnitCost,
      costMethod,
      unitType,
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

  // ✅ SAFE WRITE: append rows only
  XLSX.utils.sheet_add_aoa(ws, dataRows, {
    origin: { r: HEADER_ROW_INDEX + 1, c: 0 },
  });

  return wb;
}
