// src/converters/parseSourceIo.ts
import * as XLSX from "xlsx";
import type { NormalizedLine, OrderType } from "./types";

function excelSerialToDate(value: number): Date {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  return new Date(excelEpoch.getTime() + value * 86400000);
}

function excelToDate(value: any): Date | null {
  if (value == null) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") return excelSerialToDate(value);

  if (typeof value === "string") {
    const d = new Date(value);
    if (!Number.isNaN(d.getTime())) return d;
  }
  return null;
}

function toISODateOnly(d: Date): string {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function s(v: any): string {
  return String(v ?? "").trim();
}

function up(v: any): string {
  return s(v).toUpperCase();
}

function num(v: any): number {
  if (v == null) return 0;
  const n = Number(String(v).replace(/[$,]/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

type HeaderIndexMap = Record<string, number>;

function buildHeaderIndexMap(headerRow: any[]): HeaderIndexMap {
  // Force every entry to a *real string* (no holes / undefined)
  const cells: string[] = Array.from({ length: (headerRow ?? []).length }, (_, i) =>
    String(headerRow?.[i] ?? "").trim().toUpperCase()
  );

  const findCol = (pred: (cell: string) => boolean): number => {
    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i] ?? ""; // guard
      if (pred(cell)) return i;
    }
    return -1;
  };

  return {
    MARKET: findCol((x) => x === "MARKETS" || x.includes("MARKET")),
    PROPERTY: findCol((x) => x === "PROPERTY"),
    PLACEMENT: findCol((x) => x === "PLACEMENT" || x.includes("PLACEMENT")),
    TARGET: findCol((x) => x === "TARGET" || x.includes("TARGET")),
    START: findCol((x) => x.includes("START") && x.includes("DATE")),
    END: findCol((x) => x.includes("END") && x.includes("DATE")),
    IMPS: findCol((x) => x.includes("IMP") || x.includes("UNITS")),
    NET: findCol((x) => x.includes("NET") && x.includes("INVEST")),
  };
}

function findScheduleHeaderRow(rows: any[][]): number {
  const mustContain = ["MARKETS"];
  const anyOf = ["START", "END", "NET", "INVEST", "IMP", "UNITS", "PLACEMENT", "TARGET", "PROPERTY"];

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const normalized = row.map((c) => up(c)); // always string

    const hasMust = mustContain.every((m) => normalized.includes(m));
    if (!hasMust) continue;

    const hasAny = normalized.some((cell) =>
      anyOf.some((k) => cell.includes(k))
    );

    if (hasAny) return r;
  }
  return -1;
}

function classifyOrderType(args: {
  fileName: string;
  market: string;
  property: string;
  placement: string;
  lineItemName: string;
  targeting: string;
}): OrderType {
  const blob = [
    args.market,
    args.property,
    args.placement,
    args.lineItemName,
    args.targeting,
  ]
    .join(" ")
    .toLowerCase();

  const effectvSignals = ["effectv", "comcast", "xfinity"];
  const spectrumSignals = ["spectrum", "charter", "spectrum reach", "spectrumreach"];

  const hasEffectv = effectvSignals.some((k) => blob.includes(k));
  const hasSpectrum = spectrumSignals.some((k) => blob.includes(k));

  if (hasSpectrum && !hasEffectv) return "Spectrum";
  if (hasEffectv && !hasSpectrum) return "Effectv";

  const f = (args.fileName ?? "").toLowerCase();
  const fileHasSpectrum = spectrumSignals.some((k) => f.includes(k));
  const fileHasEffectv = effectvSignals.some((k) => f.includes(k));

  if (fileHasSpectrum && !fileHasEffectv) return "Spectrum";
  if (fileHasEffectv && !fileHasSpectrum) return "Effectv";

  // Default
  return "Spectrum";
}

export async function parseSourceIo(file: File): Promise<NormalizedLine[]> {
  const buf = await file.arrayBuffer(); 

  const wb = XLSX.read(buf, { type: "array" }); 

  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];

  const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });

  const headerRowIndex = findScheduleHeaderRow(rows);
  if (headerRowIndex === -1) return [];

  const headerMap = buildHeaderIndexMap(rows[headerRowIndex]);
  if (headerMap.MARKET === -1) return [];

  const lines: NormalizedLine[] = [];
  let emptyMarketRun = 0;

  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];

    const market = headerMap.MARKET >= 0 ? s(row[headerMap.MARKET]) : "";
    if (!market) {
      emptyMarketRun++;
      if (emptyMarketRun >= 20) break;
      continue;
    }
    emptyMarketRun = 0;

    const property = headerMap.PROPERTY >= 0 ? s(row[headerMap.PROPERTY]) : "";

    // Only ignore when Property === Ampersand
    if (property.toLowerCase() === "ampersand") continue;

    const placement = headerMap.PLACEMENT >= 0 ? s(row[headerMap.PLACEMENT]) : "";
    const targeting = headerMap.TARGET >= 0 ? s(row[headerMap.TARGET]) : "";

    const startRaw = headerMap.START >= 0 ? row[headerMap.START] : null;
    const endRaw = headerMap.END >= 0 ? row[headerMap.END] : null;

    let start = excelToDate(startRaw);
    let end = excelToDate(endRaw);

    if (!start && end) start = end;
    if (start && !end) end = start;

    if (!start || !end) continue;

    const startDateISO = toISODateOnly(start);
    const endDateISO = toISODateOnly(end);

    const impsUnits = headerMap.IMPS >= 0 ? num(row[headerMap.IMPS]) : 0;
    const netInvestment = headerMap.NET >= 0 ? num(row[headerMap.NET]) : 0;

    const lineItemName = placement || `${market} ${startDateISO}–${endDateISO}`;

    const orderType = classifyOrderType({
      fileName: file.name,
      market,
      property,
      placement,
      lineItemName,
      targeting,
    });

    lines.push({
      orderType,
      market,
      startDateISO,
      endDateISO,
      lineItemName,
      netInvestment,
      impsUnits,
      targeting,
      sourceRowIndex: r,
    });
  }

  return lines;
}
