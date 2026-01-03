// src/converters/types.ts

export type OrderType = "Spectrum" | "Effectv";

export type NormalizedLine = {
  orderType: OrderType;

  market: string;

  // ISO date strings: YYYY-MM-DD
  startDateISO: string;
  endDateISO: string;

  lineItemName: string;

  netInvestment: number;
  impsUnits: number;

  targeting: string;

  // optional, for debugging / traceability
  sourceRowIndex?: number;
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
