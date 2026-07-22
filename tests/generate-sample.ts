import { readFile, writeFile } from "node:fs/promises";
import ipReport from "../app/data/ip-contract.json";
import oooReport from "../app/data/ooo-contract.json";
import { buildDocx } from "../app/docx-export";
import { buildPdfDefinition } from "../app/pdf-export";
import type {
  GeneratorForm,
  PerformerAssets,
  TemplateReport,
  UploadedImage,
} from "../app/types";
import {
  buildPerformerPreamble,
  entityTaxText,
  extractTemplateBody,
} from "../app/utils";

const form: GeneratorForm = {
  entity: "ooo",
  contractNumber: "TEST-001",
  contractDate: "2026-07-22",
  city: "Москва",
  servicePlace: "г. Москва, тестовый объект",
  appendixNumber: "1",
  prepaymentPercent: 70,
  customer: {
    kind: "ooo",
    name: "ООО «Тестовый заказчик»",
    shortName: "ООО «Тестовый заказчик»",
    representative: "Иванов И.И.",
    position: "генерального директора",
    basis: "Устава",
    inn: "7700000000",
    kpp: "770001001",
    ogrn: "1000000000000",
    address: "г. Москва",
    postalAddress: "",
    bank: "Тестовый банк",
    bik: "044500000",
    account: "40702000000000000000",
    corrAccount: "30101000000000000000",
    email: "",
    phone: "",
  },
  items: [
    {
      id: "sample-1",
      name: "Изготовление и монтаж рекламной конструкции",
      unit: "усл.",
      quantity: 2,
      price: 61000,
    },
  ],
  includeSignature: false,
  includeSeal: false,
  additionalConditions: "Тестовая выгрузка для проверки вёрстки.",
  contractBodies: {
    ooo: extractTemplateBody(oooReport as TemplateReport),
    ip: extractTemplateBody(ipReport as TemplateReport),
  },
};

const logoBytes = await readFile(
  new URL("../public/reklamastroy-logo.png", import.meta.url),
);
const sampleImage: UploadedImage = {
  id: "sample-image",
  name: "Тестовое изображение.png",
  url: "",
  file: new File([logoBytes], "Тестовое изображение.png", { type: "image/png" }),
  width: 1633,
  height: 1633,
};
const performerAssets: PerformerAssets = {
  signature: {
    id: "test-signature",
    name: "Тестовая подпись.png",
    url: "",
    file: new File([logoBytes], "Тестовая подпись.png", {
      type: "image/png",
    }),
    width: 1633,
    height: 1633,
  },
  seal: {
    id: "test-seal",
    name: "Тестовая печать.png",
    url: "",
    file: new File([logoBytes], "Тестовая печать.png", {
      type: "image/png",
    }),
    width: 1633,
    height: 1633,
  },
};

if (!buildPerformerPreamble("ooo").includes("Гассиева Алана Руслановича, действующего на основании Устава")) {
  throw new Error("Преамбула ООО не содержит корректную формулировку об Уставе");
}
if (entityTaxText("ip", 0) !== "Без НДС.") {
  throw new Error("Для ИП должна использоваться формулировка «Без НДС.»");
}

for (const kind of ["contract", "appendix"] as const) {
  const wordForm: GeneratorForm = {
    ...form,
    includeSignature: true,
    includeSeal: false,
  };
  const { blob } = await buildDocx(
    kind,
    wordForm,
    kind === "appendix" ? [sampleImage] : [],
    performerAssets,
  );
  const bytes = new Uint8Array(await blob.arrayBuffer());
  if (bytes.length < 5000 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error(`Некорректный DOCX: ${kind}`);
  }
  await writeFile(`/private/tmp/mr-${kind}-sample.docx`, bytes);
}

const ipForm: GeneratorForm = {
  ...form,
  entity: "ip",
  contractNumber: "IP-TEST-001",
  includeSignature: true,
  includeSeal: false,
};
const { blob: ipAppendixBlob } = await buildDocx(
  "appendix",
  ipForm,
  [],
  performerAssets,
);
const ipAppendixBytes = new Uint8Array(await ipAppendixBlob.arrayBuffer());
if (ipAppendixBytes.length < 5000 || ipAppendixBytes[0] !== 0x50 || ipAppendixBytes[1] !== 0x4b) {
  throw new Error("Некорректный DOCX: приложение ИП");
}
await writeFile("/private/tmp/mr-ip-appendix-sample.docx", ipAppendixBytes);

const [pdfMakeModule, vfsModule] = await Promise.all([
  import("pdfmake/build/pdfmake"),
  import("pdfmake/build/vfs_fonts"),
]);
const pdfMake = pdfMakeModule.default;
const vfs = vfsModule.default;
if (pdfMake.addVirtualFileSystem) pdfMake.addVirtualFileSystem(vfs);
else pdfMake.vfs = vfs;

const dataUrl = (bytes: Uint8Array) => `data:image/png;base64,${Buffer.from(bytes).toString("base64")}`;
const pdfForm: GeneratorForm = { ...form, includeSignature: true, includeSeal: true };
const pdfAssets = { signature: dataUrl(logoBytes), seal: dataUrl(logoBytes) };

for (const kind of ["contract", "appendix"] as const) {
  const definition = buildPdfDefinition(
    kind,
    pdfForm,
    kind === "appendix" ? [dataUrl(logoBytes)] : [],
    pdfAssets,
  );
  const bytes = await new Promise<Uint8Array>((resolve) => {
    pdfMake.createPdf(definition).getBuffer(resolve);
  });
  if (bytes.length < 5000 || String.fromCharCode(...bytes.slice(0, 4)) !== "%PDF") {
    throw new Error(`Некорректный PDF: ${kind}`);
  }
  await writeFile(`/private/tmp/rs-${kind}-sample.pdf`, bytes);
}


const ipPdfForm: GeneratorForm = { ...ipForm, includeSignature: true };
const ipDefinition = buildPdfDefinition("appendix", ipPdfForm, [], {
  signature: pdfAssets.signature,
});
const ipPdfBytes = await new Promise<Uint8Array>((resolve) => {
  pdfMake.createPdf(ipDefinition).getBuffer(resolve);
});
if (ipPdfBytes.length < 5000 || String.fromCharCode(...ipPdfBytes.slice(0, 4)) !== "%PDF") {
  throw new Error("Некорректный PDF: приложение ИП");
}
await writeFile("/private/tmp/rs-ip-appendix-sample.pdf", ipPdfBytes);
