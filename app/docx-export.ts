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

type DocxModule = typeof import("docx");

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function fittedImageDimensions(
  width: number,
  height: number,
  maxWidth: number,
  maxHeight: number,
) {
  const safeWidth = Math.max(1, width);
  const safeHeight = Math.max(1, height);
  const scale = Math.min(maxWidth / safeWidth, maxHeight / safeHeight, 1);
  return {
    width: Math.max(1, Math.round(safeWidth * scale)),
    height: Math.max(1, Math.round(safeHeight * scale)),
  };
}

async function uploadedImageRun(
  docx: DocxModule,
  image: UploadedImage,
  maxWidth = 480,
  maxHeight = 300,
  description = "Изображение к Приложению",
) {
  const data = await image.file.arrayBuffer();
  const transformation = fittedImageDimensions(
    image.width,
    image.height,
    maxWidth,
    maxHeight,
  );
  const type = image.file.type === "image/png" ? "png" : "jpg";
  return new docx.ImageRun({
    data,
    type,
    transformation,
    altText: {
      name: image.name,
      title: image.name,
      description,
    },
  });
}

function allBorders(docx: DocxModule) {
  const border = { style: docx.BorderStyle.SINGLE, size: 4, color: "1B1B1B" };
  return {
    top: border,
    bottom: border,
    left: border,
    right: border,
    insideHorizontal: border,
    insideVertical: border,
  };
}

function noBorders(docx: DocxModule) {
  const border = { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return {
    top: border,
    bottom: border,
    left: border,
    right: border,
    insideHorizontal: border,
    insideVertical: border,
  };
}

function paragraph(
  docx: DocxModule,
  text: string,
  options: {
    bold?: boolean;
    size?: number;
    alignment?: string;
    before?: number;
    after?: number;
    keepNext?: boolean;
  } = {},
) {
  return new docx.Paragraph({
    alignment:
      (options.alignment as (typeof docx.AlignmentType)[keyof typeof docx.AlignmentType]) ||
      docx.AlignmentType.JUSTIFIED,
    spacing: {
      before: options.before ?? 0,
      after: options.after ?? 100,
      line: 276,
    },
    keepNext: options.keepNext,
    children: [
      new docx.TextRun({
        text,
        bold: options.bold,
        size: options.size ?? 22,
        font: "Times New Roman",
      }),
    ],
  });
}

function cell(
  docx: DocxModule,
  children: InstanceType<DocxModule["Paragraph"]>[],
  width?: number,
  shading?: string,
) {
  return new docx.TableCell({
    children,
    width: width ? { size: width, type: docx.WidthType.DXA } : undefined,
    shading: shading ? { fill: shading } : undefined,
    verticalAlign: docx.VerticalAlignTable.CENTER,
    margins: {
      marginUnitType: docx.WidthType.DXA,
      top: 90,
      bottom: 90,
      left: 110,
      right: 110,
    },
  });
}

async function signatureParagraphs(
  docx: DocxModule,
  form: GeneratorForm,
  assets: PerformerAssets,
) {
  const result: InstanceType<DocxModule["Paragraph"]>[] = [];
  if (form.includeSignature && assets.signature) {
    result.push(
      new docx.Paragraph({
        spacing: { before: 80, after: 0 },
        children: [
          await uploadedImageRun(
            docx,
            assets.signature,
            190,
            70,
            "Подпись исполнителя",
          ),
        ],
      }),
    );
  }
  if (form.entity === "ooo" && form.includeSeal && assets.seal) {
    result.push(
      new docx.Paragraph({
        spacing: { before: 0, after: 0 },
        children: [
          await uploadedImageRun(
            docx,
            assets.seal,
            135,
            135,
            "Печать ООО Мастера Рекламы",
          ),
        ],
      }),
    );
  }
  result.push(
    paragraph(
      docx,
      `________________ /${PERFORMERS[form.entity].signature}/`,
      { alignment: docx.AlignmentType.LEFT, before: 80 },
    ),
  );
  return result;
}

async function contractChildren(
  docx: DocxModule,
  form: GeneratorForm,
  assets: PerformerAssets,
) {
  const performer = PERFORMERS[form.entity];
  const customerLines = customerRequisites(form.customer);
  const body = resolveContractBody(form.contractBodies[form.entity]);
  const bodyParagraphs = body
    .split(/\n\s*\n/)
    .map((text) => text.trim())
    .filter(Boolean)
    .map((text) => {
      const heading = /^\d+\.\s+[А-ЯЁ0-9\-–—, :]+$/.test(text);
      return paragraph(docx, text, {
        bold: heading,
        alignment: heading
          ? docx.AlignmentType.CENTER
          : docx.AlignmentType.JUSTIFIED,
        before: heading ? 180 : 0,
        after: heading ? 120 : 70,
        keepNext: heading,
      });
    });

  const dateTable = new docx.Table({
    width: { size: 9360, type: docx.WidthType.DXA },
    columnWidths: [4680, 4680],
    borders: noBorders(docx),
    rows: [
      new docx.TableRow({
        children: [
          cell(
            docx,
            [paragraph(docx, `г. ${form.city || "Москва"}`, { alignment: docx.AlignmentType.LEFT })],
            4680,
          ),
          cell(
            docx,
            [paragraph(docx, formatDateLong(form.contractDate), { alignment: docx.AlignmentType.RIGHT })],
            4680,
          ),
        ],
      }),
    ],
  });

  const performerCell = [
    paragraph(docx, "Исполнитель:", { bold: true, alignment: docx.AlignmentType.LEFT }),
    ...performer.requisites.map((line) =>
      paragraph(docx, line, { alignment: docx.AlignmentType.LEFT, after: 0 }),
    ),
    ...(await signatureParagraphs(docx, form, assets)),
  ];
  const customerCell = [
    paragraph(docx, "Заказчик:", { bold: true, alignment: docx.AlignmentType.LEFT }),
    ...(customerLines.length
      ? customerLines.map(([label, value]) =>
          paragraph(docx, `${label}: ${value}`, {
            alignment: docx.AlignmentType.LEFT,
            after: 0,
          }),
        )
      : [
          paragraph(docx, "Реквизиты не заполнены", {
            alignment: docx.AlignmentType.LEFT,
          }),
        ]),
    paragraph(
      docx,
      `________________ /${form.customer.representative || "________________"}/`,
      { alignment: docx.AlignmentType.LEFT, before: 120 },
    ),
  ];

  const requisitesTable = new docx.Table({
    width: { size: 9360, type: docx.WidthType.DXA },
    columnWidths: [4680, 4680],
    borders: noBorders(docx),
    rows: [
      new docx.TableRow({
        cantSplit: true,
        children: [cell(docx, performerCell, 4680), cell(docx, customerCell, 4680)],
      }),
    ],
  });

  return [
    paragraph(
      docx,
      `ДОГОВОР НА ОКАЗАНИЕ УСЛУГ № ${form.contractNumber || "________"}`,
      {
        bold: true,
        size: 28,
        alignment: docx.AlignmentType.CENTER,
        after: 40,
      },
    ),
    paragraph(
      docx,
      "(по изготовлению, монтажу и обслуживанию рекламных конструкций)",
      { alignment: docx.AlignmentType.CENTER, after: 180 },
    ),
    dateTable,
    paragraph(
      docx,
      `${buildCustomerPreamble(form.customer)} с одной стороны, и ${buildPerformerPreamble(form.entity)}, с другой стороны, совместно именуемые «Стороны», заключили настоящий Договор о нижеследующем:`,
      { before: 150, after: 120 },
    ),
    ...bodyParagraphs,
    ...(form.additionalConditions.trim()
      ? [
          paragraph(docx, "ДОПОЛНИТЕЛЬНЫЕ УСЛОВИЯ", {
            bold: true,
            alignment: docx.AlignmentType.CENTER,
            before: 180,
            keepNext: true,
          }),
          paragraph(docx, form.additionalConditions.trim()),
        ]
      : []),
    paragraph(
      docx,
      "12. ЮРИДИЧЕСКИЕ АДРЕСА, БАНКОВСКИЕ РЕКВИЗИТЫ И ПОДПИСИ СТОРОН",
      {
        bold: true,
        alignment: docx.AlignmentType.CENTER,
        before: 220,
        after: 150,
        keepNext: true,
      },
    ),
    requisitesTable,
  ];
}

async function appendixChildren(
  docx: DocxModule,
  form: GeneratorForm,
  images: UploadedImage[],
  assets: PerformerAssets,
) {
  const performer = PERFORMERS[form.entity];
  const customerLines = customerRequisites(form.customer);
  const total = roundMoney(
    form.items.reduce(
      (sum, item) => sum + Number(item.quantity || 0) * Number(item.price || 0),
      0,
    ),
  );
  const vat = form.entity === "ooo" ? roundMoney((total * 22) / 122) : 0;
  const prepayment = roundMoney((total * form.prepaymentPercent) / 100);
  const remainder = roundMoney(total - prepayment);
  const itemRows = form.items.map(
    (item, index) =>
      new docx.TableRow({
        cantSplit: true,
        children: [
          cell(docx, [paragraph(docx, String(index + 1), { alignment: docx.AlignmentType.CENTER })], 650),
          cell(docx, [paragraph(docx, item.name || "—", { alignment: docx.AlignmentType.LEFT })], 4450),
          cell(
            docx,
            [paragraph(docx, `${item.quantity || 0} ${item.unit || "усл."}`, { alignment: docx.AlignmentType.CENTER })],
            1200,
          ),
          cell(docx, [paragraph(docx, formatMoney(item.price), { alignment: docx.AlignmentType.RIGHT })], 1500),
          cell(
            docx,
            [paragraph(docx, formatMoney(Number(item.quantity || 0) * Number(item.price || 0)), { alignment: docx.AlignmentType.RIGHT })],
            1560,
          ),
        ],
      }),
  );
  const headerCell = (text: string, width: number) =>
    cell(
      docx,
      [paragraph(docx, text, { bold: true, alignment: docx.AlignmentType.CENTER })],
      width,
      "E8EDF0",
    );
  const servicesTable = new docx.Table({
    width: { size: 9360, type: docx.WidthType.DXA },
    columnWidths: [650, 4450, 1200, 1500, 1560],
    borders: allBorders(docx),
    rows: [
      new docx.TableRow({
        tableHeader: true,
        cantSplit: true,
        children: [
          headerCell("№", 650),
          headerCell("Наименование", 4450),
          headerCell("Кол-во", 1200),
          headerCell(form.entity === "ooo" ? "Цена за единицу с НДС, руб." : "Цена за единицу, руб.", 1500),
          headerCell(form.entity === "ooo" ? "Стоимость с НДС, руб." : "Стоимость, руб.", 1560),
        ],
      }),
      ...itemRows,
    ],
  });

  const imageParagraphs: InstanceType<DocxModule["Paragraph"]>[] = [];
  for (const image of images) {
    imageParagraphs.push(
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 100, after: 80 },
        children: [await uploadedImageRun(docx, image)],
      }),
    );
  }

  const appendixMetaTable = new docx.Table({
    width: { size: 9360, type: docx.WidthType.DXA },
    columnWidths: [4680, 4680],
    borders: noBorders(docx),
    rows: [
      new docx.TableRow({
        children: [
          cell(
            docx,
            [paragraph(docx, `г. ${form.city || "Москва"}`, { alignment: docx.AlignmentType.LEFT })],
            4680,
          ),
          cell(
            docx,
            [paragraph(docx, formatDateQuoted(form.contractDate), { alignment: docx.AlignmentType.RIGHT })],
            4680,
          ),
        ],
      }),
    ],
  });

  const signatureTable = new docx.Table({
    width: { size: 9360, type: docx.WidthType.DXA },
    columnWidths: [4680, 4680],
    borders: noBorders(docx),
    rows: [
      new docx.TableRow({
        cantSplit: true,
        children: [
          cell(
            docx,
            [
              paragraph(docx, "Исполнитель:", { bold: true, alignment: docx.AlignmentType.LEFT }),
              paragraph(docx, performer.short, { alignment: docx.AlignmentType.LEFT }),
              ...(await signatureParagraphs(docx, form, assets)),
            ],
            4680,
          ),
          cell(
            docx,
            [
              paragraph(docx, "Заказчик:", { bold: true, alignment: docx.AlignmentType.LEFT }),
              ...(customerLines.length
                ? customerLines.map(([label, value]) =>
                    paragraph(docx, `${label}: ${value}`, {
                      alignment: docx.AlignmentType.LEFT,
                      after: 0,
                    }),
                  )
                : [
                    paragraph(docx, "Реквизиты не заполнены", {
                      alignment: docx.AlignmentType.LEFT,
                    }),
                  ]),
              paragraph(
                docx,
                `________________ /${form.customer.representative || "________________"}/`,
                { alignment: docx.AlignmentType.LEFT, before: 120 },
              ),
            ],
            4680,
          ),
        ],
      }),
    ],
  });

  return [
    paragraph(
      docx,
      `Приложение № ${form.appendixNumber || "1"} к Договору на оказание услуг № ${form.contractNumber || "________"} от ${formatDateNumeric(form.contractDate)}`,
      { bold: true, size: 24, alignment: docx.AlignmentType.CENTER, after: 180 },
    ),
    appendixMetaTable,
    paragraph(
      docx,
      buildAppendixPreamble(form.customer, form.entity, form.appendixNumber),
      { before: 150, after: 140 },
    ),
    ...imageParagraphs,
    paragraph(
      docx,
      `Место оказания услуг: ${form.servicePlace || "не указано"}.`,
      { bold: true, alignment: docx.AlignmentType.LEFT, after: 140 },
    ),
    servicesTable,
    paragraph(
      docx,
      `Итоговая стоимость всех работ ${formatMoney(total)} руб. (${amountToWords(total)}).`,
      { bold: true, alignment: docx.AlignmentType.LEFT, before: 150, after: 40 },
    ),
    paragraph(docx, entityTaxText(form.entity, vat), {
      alignment: docx.AlignmentType.LEFT,
      after: 80,
    }),
    paragraph(
      docx,
      `Заказчик производит предоплату ${formatPercent(form.prepaymentPercent)}% в размере ${formatMoney(prepayment)} руб. (${amountToWords(prepayment)}). Остаток ${formatPercent(100 - form.prepaymentPercent)}% составляет ${formatMoney(remainder)} руб. (${amountToWords(remainder)}).`,
      { after: 140 },
    ),
    signatureTable,
  ];
}

export async function buildDocx(
  kind: PreviewKind,
  form: GeneratorForm,
  images: UploadedImage[],
  assets: PerformerAssets = { signature: null, seal: null },
) {
  const docx = await import("docx");
  const children =
    kind === "contract"
      ? await contractChildren(docx, form, assets)
      : await appendixChildren(docx, form, images, assets);
  const documentFile = new docx.Document({
    creator: "Генератор договоров РекламаСтрой",
    title:
      kind === "contract"
        ? `Договор № ${form.contractNumber}`
        : `Приложение № ${form.appendixNumber}`,
    description: "Сформировано генератором договоров РекламаСтрой",
    styles: {
      default: {
        document: {
          run: { font: "Times New Roman", size: 22, color: "000000" },
          paragraph: { spacing: { after: 100, line: 276 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: {
              top: 1134,
              right: 1134,
              bottom: 1134,
              left: 1134,
              header: 567,
              footer: 567,
              gutter: 0,
            },
          },
        },
        children,
      },
    ],
  });
  const blob = await docx.Packer.toBlob(documentFile);
  const number = safeFilePart(form.contractNumber);
  const filename =
    kind === "contract"
      ? `Договор № ${number} — ${PERFORMERS[form.entity].short}.docx`
      : `Приложение № ${safeFilePart(form.appendixNumber)} к договору № ${number}.docx`;
  return { blob, filename };
}

export async function exportDocx(
  kind: PreviewKind,
  form: GeneratorForm,
  images: UploadedImage[],
  assets: PerformerAssets = { signature: null, seal: null },
) {
  const { blob, filename } = await buildDocx(kind, form, images, assets);
  downloadBlob(blob, filename);
}
