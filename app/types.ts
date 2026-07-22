export type EntityType = "ooo" | "ip";
export type CustomerKind = "ooo" | "ip" | "person";
export type PreviewKind = "contract" | "appendix";

export interface CustomerDetails {
  kind: CustomerKind;
  name: string;
  shortName: string;
  representative: string;
  position: string;
  basis: string;
  inn: string;
  kpp: string;
  ogrn: string;
  address: string;
  postalAddress: string;
  bank: string;
  bik: string;
  account: string;
  corrAccount: string;
  email: string;
  phone: string;
}

export interface ServiceItem {
  id: string;
  name: string;
  unit: string;
  quantity: number;
  price: number;
}

export interface UploadedImage {
  id: string;
  name: string;
  url: string;
  file: File;
  width: number;
  height: number;
}

export interface PerformerAssets {
  signature: UploadedImage | null;
  seal: UploadedImage | null;
}

export interface GeneratorForm {
  entity: EntityType;
  contractNumber: string;
  contractDate: string;
  city: string;
  servicePlace: string;
  appendixNumber: string;
  prepaymentPercent: number;
  customer: CustomerDetails;
  items: ServiceItem[];
  includeSignature: boolean;
  includeSeal: boolean;
  additionalConditions: string;
  contractBodies: Record<EntityType, string>;
}

export interface TemplateReport {
  paragraphs: Array<{ text: string }>;
}
