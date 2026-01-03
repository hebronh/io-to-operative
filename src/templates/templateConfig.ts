export type OrderType = "Spectrum" | "Effectv";

export type TemplateConfig = {
  type: OrderType;
  templateUrl: string;
  sheetName: string;
  headerRowIndex: number; // 0-based
  outputName: (base: string) => string;
};

export const TEMPLATE_CONFIG: Record<OrderType, TemplateConfig> = {
  Spectrum: {
    type: "Spectrum",
    templateUrl: "/templates/operative-spectrum-template.xlsx",
    sheetName: "SO Template",
    headerRowIndex: 8, // adjust if needed
    outputName: (base) => `${base}_Spectrum.xlsx`,
  },
  Effectv: {
    type: "Effectv",
    templateUrl: "/templates/operative-effectv-template.xlsx",
    sheetName: "SO Template",
    headerRowIndex: 8, // adjust if needed
    outputName: (base) => `${base}_Effectv.xlsx`,
  },
};