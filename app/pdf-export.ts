import type {
  GeneratorForm,
  PerformerAssets,
  PreviewKind,
  UploadedImage,
} from "./types";
import {
  amountToWords,
  buildAppendixPreamble,
  buildCustomerPreamble,
  buildPerformerPreamble,
  customerRequisites,
  emphasizeTextSegments,
  entityTaxText,
  formatDateLong,
  formatDateNumeric,
  formatDateQuoted,
  formatMoney,
  formatPercent,
  PERFORMERS,
  resolveContractBody,
  roundMoney,
  safeFilePart,
} from "./utils";

type PdfNode = string | Record<string, unknown>;

export interface PdfAssets {
  signature?: string;
  seal?: string;
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error || new Error("Не удалось прочитать изображение"));
    reader.readAsDataURL(file);
  });
}

function performerStack(form: GeneratorForm, assets: PdfAssets): PdfNode[] {
  const performer = PERFORMERS[form.entity];
  const stack: PdfNode[] = [
    { text: "Исполнитель:", bold: true, margin: [0, 0, 0, 7] },
    ...performer.requisites.map((line) => ({ text: line, margin: [0, 0, 0, 2] })),
  ];

  if (form.includeSignature && assets.signature) {
    stack.push({
      image: assets.signature,
      fit: [145, 52],
      alignment: "left",
      margin: [0, 8, 0, 2],
    });
  }
  if (form.entity === "ooo" && form.includeSeal && assets.seal) {
    stack.push({
      image: assets.seal,
      fit: [88, 88],
      alignment: "left",
      margin: [54, -34, 0, -18],
    });
  }
  stack.push({
    text: `________________ /${performer.signature}/`,
    margin: [0, 10, 0, 0],
  });
  return stack;
}

function customerStack(form: GeneratorForm): PdfNode[] {
  const requisites = customerRequisites(form.customer);
  return [
    { text: "Заказчик:", bold: true, margin: [0, 0, 0, 7] },
    ...(requisites.length
      ? requisites.map(([label, value]) => ({ text: `${label}: ${value}`, margin: [0, 0, 0, 2] }))
      : [{ text: "Реквизиты не заполнены", color: "#777777" }]),
    {
      text: `________________ /${form.customer.representative || "________________"}/`,
      margin: [0, 18, 0, 0],
    },
  ];
}

function partiesTable(form: GeneratorForm, assets: PdfAssets): PdfNode {
  return {
    table: {
      widths: ["*", "*"],
      dontBreakRows: true,
      body: [[{ stack: performerStack(form, assets) }, { stack: customerStack(form) }]],
    },
    layout: "noBorders",
    margin: [0, 14, 0, 0],
  };
}

function contractContent(form: GeneratorForm, assets: PdfAssets): PdfNode[] {
  const performer = PERFORMERS[form.entity];
  const body = resolveContractBody(form.contractBodies[form.entity]);
  const bodyParagraphs: PdfNode[] = body
    .split(/\n\s*\n/)
    .map((text) => text.trim())
    .filter(Boolean)
    .map((text) => {
      const heading = /^\d+\.\s+[А-ЯЁ0-9\-–—, :]+$/.test(text);
      return heading
        ? { text, bold: true, alignment: "center", margin: [0, 10, 0, 7], keepWithHeaderRows: 1 }
        : { text, alignment: "justify", margin: [0, 0, 0, 5] };
    });

  const content: PdfNode[] = [
    {
      text: `ДОГОВОР НА ОКАЗАНИЕ УСЛУГ № ${form.contractNumber || "________"}`,
      style: "title",
    },
    {
      text: "(по изготовлению, монтажу и обслуживанию рекламных конструкций)",
      alignment: "center",
      margin: [0, 3, 0, 14],
    },
    {
      columns: [
        { text: `г. ${form.city || "Москва"}` },
        { text: formatDateLong(form.contractDate), alignment: "right" },
      ],
      margin: [0, 0, 0, 14],
    },
    {
      text: emphasizeTextSegments(
        `${buildCustomerPreamble(form.customer)} с одной стороны, и ${buildPerformerPreamble(form.entity)}, с другой стороны, совместно именуемые «Стороны», заключили настоящий Договор о нижеследующем:`,
        [
          form.customer.name,
          form.customer.representative,
          performer.full,
          form.entity === "ooo" ? PERFORMERS.ooo.signerGenitive : PERFORMERS.ip.signer,
        ],
      ),
      alignment: "justify",
      margin: [0, 0, 0, 8],
    },
    ...bodyParagraphs,
  ];

  if (form.additionalConditions.trim()) {
    content.push(
      { text: "ДОПОЛНИТЕЛЬНЫЕ УСЛОВИЯ", bold: true, alignment: "center", margin: [0, 12, 0, 7] },
      { text: form.additionalConditions.trim(), alignment: "justify", margin: [0, 0, 0, 8] },
    );
  }

  content.push(
    {
      text: "12. ЮРИДИЧЕСКИЕ АДРЕСА, БАНКОВСКИЕ РЕКВИЗИТЫ И ПОДПИСИ СТОРОН",
      bold: true,
      alignment: "center",
      margin: [0, 14, 0, 8],
    },
    partiesTable(form, assets),
  );
  return content;
}

function appendixContent(
  form: GeneratorForm,
  uploadedImages: string[],
  assets: PdfAssets,
): PdfNode[] {
  const total = roundMoney(
    form.items.reduce(
      (sum, item) => sum + Number(item.quantity || 0) * Number(item.price || 0),
      0,
    ),
  );
  const vat = form.entity === "ooo" ? roundMoney((total * 22) / 122) : 0;
  const prepayment = roundMoney((total * form.prepaymentPercent) / 100);
  const remainder = roundMoney(total - prepayment);

  const tableBody: PdfNode[][] = [
    [
      { text: "№", style: "tableHeader" },
      { text: "Наименование", style: "tableHeader" },
      { text: "Кол-во", style: "tableHeader" },
      {
        text: form.entity === "ooo" ? "Цена за единицу с НДС, руб." : "Цена за единицу, руб.",
        style: "tableHeader",
      },
      {
        text: form.entity === "ooo" ? "Стоимость с НДС, руб." : "Стоимость, руб.",
        style: "tableHeader",
      },
    ],
    ...form.items.map((item, index) => [
      { text: String(index + 1), alignment: "center" },
      { text: item.name || "—" },
      { text: `${item.quantity || 0} ${item.unit || "усл."}`, alignment: "center" },
      { text: formatMoney(item.price), alignment: "right", noWrap: true },
      {
        text: formatMoney(Number(item.quantity || 0) * Number(item.price || 0)),
        alignment: "right",
        noWrap: true,
      },
    ]),
  ];

  return [
    {
      text: `Приложение № ${form.appendixNumber || "1"} к Договору на оказание услуг № ${form.contractNumber || "________"} от ${formatDateNumeric(form.contractDate)}`,
      style: "appendixTitle",
    },
    {
      columns: [
        { text: `г. ${form.city || "Москва"}` },
        { text: formatDateQuoted(form.contractDate), alignment: "right" },
      ],
      margin: [0, 0, 0, 14],
    },
    {
      text: buildAppendixPreamble(form.customer, form.entity, form.appendixNumber),
      alignment: "justify",
      margin: [0, 0, 0, 9],
    },
    ...uploadedImages.map((image) => ({
      image,
      fit: [480, 250],
      alignment: "center",
      margin: [0, 4, 0, 9],
      pageBreak: undefined,
    })),
    {
      text: `Место оказания услуг: ${form.servicePlace || "не указано"}.`,
      bold: true,
      margin: [0, 5, 0, 8],
    },
    {
      table: {
        headerRows: 1,
        dontBreakRows: true,
        widths: [24, "*", 58, 88, 92],
        body: tableBody,
      },
      layout: {
        fillColor: (rowIndex: number) => (rowIndex === 0 ? "#E8EDF0" : null),
        hLineColor: () => "#222222",
        vLineColor: () => "#222222",
        paddingLeft: () => 5,
        paddingRight: () => 5,
        paddingTop: () => 5,
        paddingBottom: () => 5,
      },
      fontSize: 9,
      margin: [0, 0, 0, 10],
    },
    {
      text: `Итоговая стоимость всех работ ${formatMoney(total)} руб. (${amountToWords(total)}).`,
      bold: true,
      margin: [0, 0, 0, 4],
    },
    { text: entityTaxText(form.entity, vat), margin: [0, 0, 0, 6] },
    {
      text: `Заказчик производит предоплату ${formatPercent(form.prepaymentPercent)}% в размере ${formatMoney(prepayment)} руб. (${amountToWords(prepayment)}). Остаток ${formatPercent(100 - form.prepaymentPercent)}% составляет ${formatMoney(remainder)} руб. (${amountToWords(remainder)}).`,
      alignment: "justify",
      margin: [0, 0, 0, 8],
    },
    partiesTable(form, assets),
  ];
}

export function buildPdfDefinition(
  kind: PreviewKind,
  form: GeneratorForm,
  uploadedImages: string[],
  assets: PdfAssets,
): Record<string, unknown> {
  return {
    info: {
      title:
        kind === "contract"
          ? `Договор № ${form.contractNumber}`
          : `Приложение № ${form.appendixNumber}`,
      author: "Генератор договоров РекламаСтрой",
      subject: "Договорные документы",
    },
    pageSize: "A4",
    pageMargins: [54, 58, 54, 54],
    header: (currentPage: number) => ({
      text: String(currentPage),
      alignment: "center",
      fontSize: 9,
      margin: [0, 22, 0, 0],
    }),
    defaultStyle: {
      font: "Roboto",
      fontSize: 10.5,
      lineHeight: 1.18,
      color: "#111111",
    },
    styles: {
      title: { fontSize: 15, bold: true, alignment: "center", margin: [0, 0, 0, 3] },
      appendixTitle: { fontSize: 13, bold: true, alignment: "center", margin: [0, 0, 0, 14] },
      tableHeader: { bold: true, alignment: "center", fontSize: 8.5 },
    },
    content:
      kind === "contract"
        ? contractContent(form, assets)
        : appendixContent(form, uploadedImages, assets),
  };
}

export async function exportPdf(
  kind: PreviewKind,
  form: GeneratorForm,
  images: UploadedImage[],
  performerAssets: PerformerAssets = { signature: null, seal: null },
) {
  const [pdfMakeModule, vfsModule, uploadedImages, signature, seal] = await Promise.all([
    import("pdfmake/build/pdfmake"),
    import("pdfmake/build/vfs_fonts"),
    Promise.all(images.map((image) => fileToDataUrl(image.file))),
    form.includeSignature && performerAssets.signature
      ? fileToDataUrl(performerAssets.signature.file)
      : Promise.resolve(undefined),
    form.entity === "ooo" && form.includeSeal && performerAssets.seal
      ? fileToDataUrl(performerAssets.seal.file)
      : Promise.resolve(undefined),
  ]);

  const pdfMake = pdfMakeModule.default;
  const vfs = vfsModule.default;
  if (pdfMake.addVirtualFileSystem) pdfMake.addVirtualFileSystem(vfs);
  else pdfMake.vfs = vfs;

  const definition = buildPdfDefinition(kind, form, uploadedImages, { signature, seal });
  const number = safeFilePart(form.contractNumber);
  const filename =
    kind === "contract"
      ? `Договор № ${number} — ${PERFORMERS[form.entity].short}.pdf`
      : `Приложение № ${safeFilePart(form.appendixNumber)} к договору № ${number}.pdf`;
  pdfMake.createPdf(definition).download(filename);
}
