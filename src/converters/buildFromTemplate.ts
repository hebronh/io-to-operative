import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

export type NormalizedLine = {
  startDateISO: string; // YYYY-MM-DD
  endDateISO: string;   // YYYY-MM-DD
  lineItemName: string;
  netInvestment: number;
  impsUnits: number;
  targeting: string;
  market?: string;
};

export type OperativeHeader = {
  orderId: string;
  mediaPlanName?: string;
  sectionName?: string;
  groupName?: string;
  productName?: string;
  costMethod?: string; // CPM, Flat, etc
  unitType?: string;   // Impressions, Units
  canOutput?: boolean;
  canInvoice?: boolean;
  billableThirdPartyServer?: string;
  dma?: string;
  state?: string;
  congressionalDistrict?: string;
};

const SO_SHEET_NAME = "SO Template";

// The SO Template header row in your template is at row 9 (1-based) => index 8 (0-based)
// Your template has “Order ID” on row 1, then blanks, then headers on row 9
const HEADER_ROW_INDEX = 8;

function toExcelDate(iso: string) {
  return new Date(iso + "T00:00:00");
}

function getSoHeaders(ws: XLSX.WorkSheet): string[] {
  const ref = ws["!ref"];
  if (!ref) return [];
  const range = XLSX.utils.decode_range(ref);

  const headers: string[] = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: HEADER_ROW_INDEX, c });
    const cell = ws[addr];
    headers.push(cell?.v?.toString?.() ?? "");
  }
  return headers;
}

export async function buildAndDownloadOperativeXlsFromTemplate(
  templateUrl: string,
  header: OperativeHeader,
  lines: NormalizedLine[]
) {
  // 1) fetch template
  const res = await fetch(templateUrl);
  const buf = await res.arrayBuffer();

  // 2) read template workbook
  const wb = XLSX.read(buf, { type: "array" });

  // 3) get SO Template sheet
  const ws = wb.Sheets[SO_SHEET_NAME];
  if (!ws) throw new Error(`Template missing sheet: "${SO_SHEET_NAME}"`);

  // 4) read existing header columns from template row
  const headers = getSoHeaders(ws);
  if (!headers.length) throw new Error("Could not read SO Template headers.");

  // 5) build new rows (keep everything above header row intact)
  const existing = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, raw: true }) as any[][];
  const top = existing.slice(0, HEADER_ROW_INDEX + 1); // includes header row

  // Update Order ID (Row 1, Col A/B in template)
  if (top[0]?.length) {
    top[0][0] = "Order ID";
    top[0][1] = header.orderId ?? "";
  }

  const mediaPlan = header.mediaPlanName ?? "Default Media Plan";
  const section = header.sectionName ?? "Default Section";
  const costMethod = header.costMethod ?? "CPM";
  const unitType = header.unitType ?? "Impressions";

  const dataRows: any[][] = lines.map((line) => {
    const qty = line.impsUnits ?? 0;

    const netUnitCost =
      costMethod.toUpperCase() === "CPM"
        ? qty > 0
          ? line.netInvestment / (qty / 1000)
          : 0
        : line.netInvestment;

    // Build a row matching the template’s columns by header name
    const row: Record<string, any> = {
      "Media Plan Name": mediaPlan,
      "Section Name": section,
      "Start Date": toExcelDate(line.startDateISO),
      "End Date": toExcelDate(line.endDateISO),
      "Line Item Name": line.lineItemName,
      "Group Name": header.groupName ?? "",
      "Product Name": header.productName ?? "",
      "Agency Discount": 0,
      "Net Unit Cost": netUnitCost,
      "Cost Method": costMethod,
      "Unit Type": unitType,
      "Quantity": qty,
      "Can Output": header.canOutput ?? true,
      "Can Invoice": header.canInvoice ?? true,
      "Billable Third Party Server": header.billableThirdPartyServer ?? "",
      "DMA": header.dma ?? "",
      "State": header.state ?? "",
      "Congressional District": header.congressionalDistrict ?? "",
      "Targeting": line.targeting ?? "",
    };

    // Output array in the exact order of the template headers
    return headers.map((h) => row[h] ?? "");
  });

  const outAoA = [...top, ...dataRows];

  // 6) replace sheet content
  const newWs = XLSX.utils.aoa_to_sheet(outAoA);
  wb.Sheets[SO_SHEET_NAME] = newWs;

  // 7) write workbook as .xls
  const out = XLSX.write(wb, { bookType: "xls", type: "array" });

  // 8) download
  saveAs(new Blob([out], { type: "application/vnd.ms-excel" }), "Operative_Template_Filled.xls");
}
