import type {
  CustomerDetails,
  EntityType,
  TemplateReport,
} from "./types";

export const PERFORMERS = {
  ooo: {
    short: "ООО «Мастера Рекламы»",
    full: "Общество с ограниченной ответственностью «Мастера Рекламы»",
    signer: "Гассиев Алан Русланович",
    signerGenitive: "Гассиева Алана Руслановича",
    signature: "Гассиев А.Р.",
    representativeTitle: "генерального директора",
    basis: "Устава",
    requisites: [
      "ООО «Мастера Рекламы»",
      "ИНН 7713394262",
      "КПП 774301001",
      "Р/с 40702810602310000342",
      "К/с 30101810200000000593",
      "БИК 044525593",
      "АО «АЛЬФА-БАНК»",
      "E-mail sales@reklamastroy.ru",
      "Телефон 8 495 008-37-95",
    ],
  },
  ip: {
    short: "ИП Гассиев Алан Русланович",
    full: "Индивидуальный предприниматель Гассиев Алан Русланович",
    signer: "Гассиев Алан Русланович",
    signature: "Гассиев А.Р.",
    representativeTitle: "",
    basis: "",
    requisites: [
      "ИП Гассиев Алан Русланович",
      "ИНН 771531062593",
      "Р/с 40802810002310003166",
      "К/с 30101810200000000593",
      "БИК 044525593",
      "АО «АЛЬФА-БАНК»",
      "E-mail sales@reklamastroy.ru",
      "Телефон 8 495 008-37-95",
    ],
  },
} as const;

export function emphasizeTextSegments(
  text: string,
  emphasizedValues: string[],
): Array<{ text: string; bold: boolean }> {
  if (!text) return [];
  const bold = Array.from({ length: text.length }, () => false);

  for (const rawValue of emphasizedValues) {
    const value = rawValue.trim();
    if (!value) continue;
    let start = text.indexOf(value);
    while (start >= 0) {
      for (let index = start; index < start + value.length; index += 1) {
        bold[index] = true;
      }
      start = text.indexOf(value, start + value.length);
    }
  }

  const segments: Array<{ text: string; bold: boolean }> = [];
  let start = 0;
  for (let index = 1; index <= text.length; index += 1) {
    if (index === text.length || bold[index] !== bold[start]) {
      segments.push({ text: text.slice(start, index), bold: bold[start] });
      start = index;
    }
  }
  return segments;
}

export function extractTemplateBody(report: TemplateReport): string {
  const paragraphs = report.paragraphs
    .map((paragraph) => paragraph.text.trim())
    .filter(Boolean);
  const start = paragraphs.findIndex((text) => /^1\.\s/.test(text));
  const end = paragraphs.findIndex((text) =>
    text.includes("ДОПОЛНИТЕЛЬНЫЕ УСЛОВИЯ"),
  );
  return paragraphs
    .slice(start >= 0 ? start : 0, end > start ? end : undefined)
    .join("\n\n");
}

export function roundMoney(value: number): number {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}

export function formatMoney(value: number): string {
  return new Intl.NumberFormat("ru-RU", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(roundMoney(value));
}

export function formatPercent(value: number): string {
  return new Intl.NumberFormat("ru-RU", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  }).format(Number(value || 0));
}

function plural(value: number, forms: [string, string, string]): string {
  const mod100 = Math.abs(value) % 100;
  const mod10 = mod100 % 10;
  if (mod100 > 10 && mod100 < 20) return forms[2];
  if (mod10 > 1 && mod10 < 5) return forms[1];
  if (mod10 === 1) return forms[0];
  return forms[2];
}

function triadToWords(value: number, feminine = false): string[] {
  const hundreds = [
    "",
    "сто",
    "двести",
    "триста",
    "четыреста",
    "пятьсот",
    "шестьсот",
    "семьсот",
    "восемьсот",
    "девятьсот",
  ];
  const tens = [
    "",
    "",
    "двадцать",
    "тридцать",
    "сорок",
    "пятьдесят",
    "шестьдесят",
    "семьдесят",
    "восемьдесят",
    "девяносто",
  ];
  const teens = [
    "десять",
    "одиннадцать",
    "двенадцать",
    "тринадцать",
    "четырнадцать",
    "пятнадцать",
    "шестнадцать",
    "семнадцать",
    "восемнадцать",
    "девятнадцать",
  ];
  const units = feminine
    ? ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
    : ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const words: string[] = [];
  const hundred = Math.floor(value / 100);
  const remainder = value % 100;
  if (hundred) words.push(hundreds[hundred]);
  if (remainder >= 10 && remainder <= 19) {
    words.push(teens[remainder - 10]);
  } else {
    const ten = Math.floor(remainder / 10);
    const unit = remainder % 10;
    if (ten) words.push(tens[ten]);
    if (unit) words.push(units[unit]);
  }
  return words;
}

function integerToWords(value: number): string {
  if (!value) return "ноль";
  const groups: Array<{
    forms: [string, string, string];
    feminine: boolean;
  }> = [
    { forms: ["", "", ""], feminine: false },
    { forms: ["тысяча", "тысячи", "тысяч"], feminine: true },
    { forms: ["миллион", "миллиона", "миллионов"], feminine: false },
    { forms: ["миллиард", "миллиарда", "миллиардов"], feminine: false },
    { forms: ["триллион", "триллиона", "триллионов"], feminine: false },
  ];
  const words: string[] = [];
  let rest = Math.floor(Math.abs(value));
  let groupIndex = 0;
  while (rest > 0 && groupIndex < groups.length) {
    const triad = rest % 1000;
    if (triad) {
      const groupWords = triadToWords(triad, groups[groupIndex].feminine);
      if (groupIndex > 0) {
        groupWords.push(plural(triad, groups[groupIndex].forms));
      }
      words.unshift(...groupWords);
    }
    rest = Math.floor(rest / 1000);
    groupIndex += 1;
  }
  return words.join(" ");
}

export function amountToWords(value: number, capitalize = true): string {
  const safeValue = Math.max(0, roundMoney(value));
  let rubles = Math.floor(safeValue);
  let kopecks = Math.round((safeValue - rubles) * 100);
  if (kopecks === 100) {
    rubles += 1;
    kopecks = 0;
  }
  let result = `${integerToWords(rubles)} ${plural(rubles, [
    "рубль",
    "рубля",
    "рублей",
  ])} ${String(kopecks).padStart(2, "0")} ${plural(kopecks, [
    "копейка",
    "копейки",
    "копеек",
  ])}`;
  if (capitalize) result = result.charAt(0).toUpperCase() + result.slice(1);
  return result;
}

export function formatDateLong(value: string): string {
  if (!value) return "дата не указана";
  const [year, month, day] = value.split("-").map(Number);
  const months = [
    "января",
    "февраля",
    "марта",
    "апреля",
    "мая",
    "июня",
    "июля",
    "августа",
    "сентября",
    "октября",
    "ноября",
    "декабря",
  ];
  if (!year || !month || !day) return value;
  return `${day} ${months[month - 1]} ${year} г.`;
}

export function formatDateQuoted(value: string): string {
  if (!value) return "«___» __________ 20___ г.";
  const [year, month, day] = value.split("-").map(Number);
  const months = [
    "января",
    "февраля",
    "марта",
    "апреля",
    "мая",
    "июня",
    "июля",
    "августа",
    "сентября",
    "октября",
    "ноября",
    "декабря",
  ];
  if (!year || !month || !day) return value;
  return `«${day}» ${months[month - 1]} ${year} г.`;
}

export function formatDateNumeric(value: string): string {
  if (!value) return "__.__.20__";
  const [year, month, day] = value.split("-").map(Number);
  if (!year || !month || !day) return value;
  return `${String(day).padStart(2, "0")}.${String(month).padStart(2, "0")}.${year}`;
}

export function buildCustomerPreamble(customer: CustomerDetails): string {
  const name = customer.name.trim() || "[наименование заказчика]";
  const representative = customer.representative.trim();
  const position = customer.position.trim();
  const basis = customer.basis.trim();
  if (customer.kind === "person") {
    return `Физическое лицо ${name} (далее — «Заказчик»)`;
  }
  if (customer.kind === "ip") {
    return `Индивидуальный предприниматель ${name}${
      basis ? `, действующий на основании ${basis}` : ""
    } (далее — «Заказчик»)`;
  }
  const signer = [position, representative].filter(Boolean).join(" ");
  return `${name}${signer ? ` в лице ${signer}` : ""}${
    basis ? `, действующего на основании ${basis}` : ""
  }, именуемое в дальнейшем «Заказчик»`;
}

export function buildPerformerPreamble(entity: EntityType): string {
  if (entity === "ooo") {
    const performer = PERFORMERS.ooo;
    return `${performer.full} в лице ${performer.representativeTitle} ${performer.signerGenitive}, действующего на основании ${performer.basis}, именуемое в дальнейшем «Исполнитель»`;
  }
  const performer = PERFORMERS.ip;
  return `${performer.full}, именуемый в дальнейшем «Исполнитель»`;
}

export function buildAppendixPreamble(
  customer: CustomerDetails,
  entity: EntityType,
  appendixNumber: string,
): string {
  return `${buildCustomerPreamble(customer)}, с одной стороны, и ${buildPerformerPreamble(entity)}, с другой стороны, при совместном упоминании по тексту Договора — «Стороны», и по отдельности — «Сторона», заключили настоящее Приложение № ${appendixNumber || "1"} о нижеследующем:`;
}

export function customerRequisites(
  customer: CustomerDetails,
): Array<[string, string]> {
  const fields: Array<[string, string]> = [
    ["Наименование", customer.name],
    ["Краткое наименование", customer.shortName],
    ["ИНН", customer.inn],
    ["КПП", customer.kpp],
    [customer.kind === "ip" ? "ОГРНИП" : "ОГРН", customer.ogrn],
    ["Юридический адрес", customer.address],
    ["Почтовый адрес", customer.postalAddress],
    ["Банк", customer.bank],
    ["БИК", customer.bik],
    ["Р/с", customer.account],
    ["К/с", customer.corrAccount],
    ["E-mail", customer.email],
    ["Телефон", customer.phone],
  ];
  return fields.filter(([, value]) => value.trim());
}

export function resolveContractBody(body: string): string {
  return body.replace(/\{\{[^}]+\}\}/g, "").trim();
}

export function entityTaxText(entity: EntityType, vat: number): string {
  return entity === "ooo"
    ? `В том числе НДС 22% — ${formatMoney(vat)} руб. (${amountToWords(vat)}).`
    : "Без НДС.";
}

export function safeFilePart(value: string): string {
  return (value || "без номера").replace(/[\\/:*?"<>|]/g, "-").trim();
}
