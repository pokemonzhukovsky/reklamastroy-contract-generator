"use client";

/* eslint-disable @next/next/no-img-element -- previews use local blob URLs and exact document assets */

import { useEffect, useMemo, useState } from "react";
import {
  deletePerformerAsset,
  loadPerformerAssets,
  savePerformerAsset,
  type PerformerAssetKind,
} from "./browser-assets";
import ipReport from "./data/ip-contract.json";
import oooReport from "./data/ooo-contract.json";
import { exportDocx } from "./docx-export";
import { exportPdf } from "./pdf-export";
import type {
  CustomerDetails,
  EntityType,
  GeneratorForm,
  PerformerAssets,
  PreviewKind,
  ServiceItem,
  TemplateReport,
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
  extractTemplateBody,
  formatDateLong,
  formatDateNumeric,
  formatDateQuoted,
  formatMoney,
  formatPercent,
  PERFORMERS,
  resolveContractBody,
  roundMoney,
} from "./utils";

const DEFAULT_BODIES = {
  ooo: extractTemplateBody(oooReport as TemplateReport),
  ip: extractTemplateBody(ipReport as TemplateReport),
};

const EMPTY_CUSTOMER: CustomerDetails = {
  kind: "ooo",
  name: "",
  shortName: "",
  representative: "",
  position: "",
  basis: "",
  inn: "",
  kpp: "",
  ogrn: "",
  address: "",
  postalAddress: "",
  bank: "",
  bik: "",
  account: "",
  corrAccount: "",
  email: "",
  phone: "",
};

function makeDefaultForm(): GeneratorForm {
  return {
    entity: "ooo",
    contractNumber: "123/1",
    contractDate: "2026-07-22",
    city: "Москва",
    servicePlace: "",
    appendixNumber: "1",
    prepaymentPercent: 70,
    customer: { ...EMPTY_CUSTOMER },
    items: [{ id: "item-1", name: "", unit: "усл.", quantity: 1, price: 0 }],
    includeSignature: false,
    includeSeal: false,
    additionalConditions: "",
    contractBodies: { ...DEFAULT_BODIES },
  };
}

async function readImageDimensions(file: File): Promise<{ width: number; height: number }> {
  if ("createImageBitmap" in window) {
    const bitmap = await createImageBitmap(file);
    const dimensions = { width: bitmap.width, height: bitmap.height };
    bitmap.close();
    return dimensions;
  }

  const url = URL.createObjectURL(file);
  try {
    return await new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
      image.onerror = reject;
      image.src = url;
    });
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function fileToUploadedImage(
  file: File,
  id: string,
): Promise<UploadedImage> {
  const dimensions = await readImageDimensions(file).catch(() => ({
    width: 1,
    height: 1,
  }));
  return {
    id,
    name: file.name,
    url: URL.createObjectURL(file),
    file,
    ...dimensions,
  };
}

function Field({
  label,
  hint,
  wide,
  children,
}: {
  label: string;
  hint?: string;
  wide?: boolean;
  children: React.ReactNode;
}) {
  return (
    <label className={`field ${wide ? "field-wide" : ""}`}>
      <span className="field-label">{label}</span>
      {children}
      {hint ? <span className="field-hint">{hint}</span> : null}
    </label>
  );
}

function AssetUploadCard({
  id,
  title,
  description,
  asset,
  disabled,
  onSelect,
  onRemove,
}: {
  id: string;
  title: string;
  description: string;
  asset: UploadedImage | null;
  disabled?: boolean;
  onSelect: (files: FileList | null) => void;
  onRemove: () => void;
}) {
  return (
    <div className={`asset-upload-card ${disabled ? "is-disabled" : ""}`}>
      <div className="asset-upload-heading">
        <div>
          <strong>{title}</strong>
          <small>{description}</small>
        </div>
        <span>{asset ? "Сохранено" : "Не загружено"}</span>
      </div>
      {asset ? (
        <div className="asset-upload-preview">
          <img src={asset.url} alt={title} />
          <div>
            <strong>{asset.name}</strong>
            <small>Хранится в этом браузере на этом компьютере</small>
          </div>
        </div>
      ) : (
        <div className="asset-empty-state">
          <span>Изображение пока не выбрано</span>
        </div>
      )}
      <div className="asset-upload-actions">
        <input
          id={id}
          type="file"
          accept="image/png,image/jpeg"
          onChange={(event) => {
            onSelect(event.target.files);
            event.target.value = "";
          }}
          disabled={disabled}
        />
        <label htmlFor={id}>{asset ? "Заменить" : "Загрузить"}</label>
        {asset ? (
          <button type="button" onClick={onRemove} disabled={disabled}>
            Удалить
          </button>
        ) : null}
      </div>
    </div>
  );
}

function Toggle({
  checked,
  onChange,
  label,
  description,
  disabled,
}: {
  checked: boolean;
  onChange: (value: boolean) => void;
  label: string;
  description: string;
  disabled?: boolean;
}) {
  return (
    <label className={`toggle-row ${disabled ? "is-disabled" : ""}`}>
      <span>
        <strong>{label}</strong>
        <small>{description}</small>
      </span>
      <input
        type="checkbox"
        checked={checked}
        onChange={(event) => onChange(event.target.checked)}
        disabled={disabled}
      />
      <span className="toggle-track" aria-hidden="true">
        <span />
      </span>
    </label>
  );
}

function SignaturePreview({
  form,
  assets,
}: {
  form: GeneratorForm;
  assets: PerformerAssets;
}) {
  const showSignature = form.includeSignature && assets.signature;
  const showSeal = form.entity === "ooo" && form.includeSeal && assets.seal;
  return (
    <div className={`signature-preview ${showSignature || showSeal ? "has-assets" : ""}`}>
      {showSeal ? (
        <img className="seal-image" src={showSeal.url} alt="Печать ООО Мастера Рекламы" />
      ) : null}
      {showSignature ? (
        <img className="signature-image" src={showSignature.url} alt="Подпись исполнителя" />
      ) : null}
      <span>________________ /Гассиев А.Р./</span>
    </div>
  );
}

function ContractPreview({
  form,
  assets,
}: {
  form: GeneratorForm;
  assets: PerformerAssets;
}) {
  const performer = PERFORMERS[form.entity];
  const preamble = `${buildCustomerPreamble(form.customer)} с одной стороны, и ${buildPerformerPreamble(form.entity)}, с другой стороны, совместно именуемые «Стороны», заключили настоящий Договор о нижеследующем:`;
  const preambleSegments = emphasizeTextSegments(preamble, [
    form.customer.name,
    form.customer.representative,
    performer.full,
    form.entity === "ooo" ? PERFORMERS.ooo.signerGenitive : PERFORMERS.ip.signer,
  ]);
  const requisites = customerRequisites(form.customer);
  const paragraphs = resolveContractBody(form.contractBodies[form.entity])
    .split(/\n\s*\n/)
    .map((text) => text.trim())
    .filter(Boolean);

  return (
    <article className="paper contract-paper" aria-label="Предварительный просмотр договора">
      <div className="paper-page-number">1</div>
      <h1>
        ДОГОВОР НА ОКАЗАНИЕ УСЛУГ № {form.contractNumber || "________"}
      </h1>
      <p className="paper-subtitle">
        (по изготовлению, монтажу и обслуживанию рекламных конструкций)
      </p>
      <div className="paper-meta">
        <span>г. {form.city || "Москва"}</span>
        <span>{formatDateLong(form.contractDate)}</span>
      </div>
      <p>
        {preambleSegments.map((segment, index) =>
          segment.bold ? (
            <strong key={`${index}-${segment.text}`}>{segment.text}</strong>
          ) : (
            <span key={`${index}-${segment.text}`}>{segment.text}</span>
          ),
        )}
      </p>

      <div className="contract-body">
        {paragraphs.map((text, index) => {
          const heading = /^\d+\.\s+[А-ЯЁ0-9\-–—, :]+$/.test(text);
          return heading ? (
            <h2 key={`${index}-${text}`}>{text}</h2>
          ) : (
            <p key={`${index}-${text.slice(0, 30)}`}>{text}</p>
          );
        })}
      </div>

      {form.additionalConditions.trim() ? (
        <section className="additional-preview">
          <h2>ДОПОЛНИТЕЛЬНЫЕ УСЛОВИЯ</h2>
          <p>{form.additionalConditions}</p>
        </section>
      ) : null}

      <h2>12. ЮРИДИЧЕСКИЕ АДРЕСА, БАНКОВСКИЕ РЕКВИЗИТЫ И ПОДПИСИ СТОРОН</h2>
      <div className="parties-grid">
        <section>
          <h3>Исполнитель:</h3>
          {performer.requisites.map((line) => (
            <p key={line}>{line}</p>
          ))}
          <SignaturePreview form={form} assets={assets} />
        </section>
        <section>
          <h3>Заказчик:</h3>
          {requisites.length ? (
            requisites.map(([label, value]) => (
              <p key={label}>
                {label}: {value}
              </p>
            ))
          ) : (
            <p className="paper-placeholder">Реквизиты не заполнены</p>
          )}
          <div className="customer-signature">
            ________________ /{form.customer.representative || "________________"}/
          </div>
        </section>
      </div>
    </article>
  );
}

function AppendixPreview({
  form,
  images,
  assets,
}: {
  form: GeneratorForm;
  images: UploadedImage[];
  assets: PerformerAssets;
}) {
  const total = roundMoney(
    form.items.reduce(
      (sum, item) => sum + Number(item.quantity || 0) * Number(item.price || 0),
      0,
    ),
  );
  const vat = form.entity === "ooo" ? roundMoney((total * 22) / 122) : 0;
  const prepayment = roundMoney((total * form.prepaymentPercent) / 100);
  const remainder = roundMoney(total - prepayment);
  const performer = PERFORMERS[form.entity];
  const requisites = customerRequisites(form.customer);

  return (
    <article className="paper appendix-paper" aria-label="Предварительный просмотр приложения">
      <div className="paper-page-number">1</div>
      <h1 className="appendix-main-heading">
        Приложение № {form.appendixNumber || "1"} к Договору на оказание услуг № {" "}
        {form.contractNumber || "________"} от {formatDateNumeric(form.contractDate)}
      </h1>
      <div className="paper-meta appendix-meta">
        <span>г. {form.city || "Москва"}</span>
        <span>{formatDateQuoted(form.contractDate)}</span>
      </div>
      <p className="appendix-preamble">
        {buildAppendixPreamble(form.customer, form.entity, form.appendixNumber)}
      </p>

      {images.length ? (
        <div className="attachment-gallery">
          {images.map((image, index) => (
            <figure key={image.id}>
              <img src={image.url} alt={`Изображение ${index + 1}: ${image.name}`} />
            </figure>
          ))}
        </div>
      ) : null}

      <p className="service-place">
        <strong>Место оказания услуг:</strong>{" "}
        {form.servicePlace || <span className="paper-placeholder">не указано</span>}
      </p>

      <div className="document-table-wrap">
        <table className="document-table">
          <thead>
            <tr>
              <th>№</th>
              <th>Наименование</th>
              <th>Кол-во</th>
              <th>{form.entity === "ooo" ? "Цена за единицу с НДС, руб." : "Цена за единицу, руб."}</th>
              <th>{form.entity === "ooo" ? "Стоимость с НДС, руб." : "Стоимость, руб."}</th>
            </tr>
          </thead>
          <tbody>
            {form.items.map((item, index) => (
              <tr key={item.id}>
                <td>{index + 1}</td>
                <td className="text-left">
                  {item.name || <span className="paper-placeholder">—</span>}
                </td>
                <td>{item.quantity || 0} {item.unit || "усл."}</td>
                <td className="number-cell">{formatMoney(item.price)}</td>
                <td className="number-cell">
                  {formatMoney(Number(item.quantity || 0) * Number(item.price || 0))}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="document-totals">
        <p>
          <strong>
            Итоговая стоимость всех работ {formatMoney(total)} руб. ({amountToWords(total)}).
          </strong>
        </p>
        <p>{entityTaxText(form.entity, vat)}</p>
      </div>

      <div className="payment-terms">
        <p>
          Заказчик производит предоплату {formatPercent(form.prepaymentPercent)}% в размере {" "}
          {formatMoney(prepayment)} руб. ({amountToWords(prepayment)}). Остаток {" "}
          {formatPercent(100 - form.prepaymentPercent)}% составляет {formatMoney(remainder)} руб. ({amountToWords(remainder)}).
        </p>
      </div>

      <div className="parties-grid appendix-signatures">
        <section>
          <h3>Исполнитель:</h3>
          {performer.requisites.map((line) => (
            <p key={line}>{line}</p>
          ))}
          <SignaturePreview form={form} assets={assets} />
        </section>
        <section>
          <h3>Заказчик:</h3>
          {requisites.length ? (
            requisites.map(([label, value]) => (
              <p key={label}>{label}: {value}</p>
            ))
          ) : (
            <p className="paper-placeholder">Реквизиты не заполнены</p>
          )}
          <div className="customer-signature">
            ________________ /{form.customer.representative || "________________"}/
          </div>
        </section>
      </div>
    </article>
  );
}

export function ContractGenerator() {
  const [form, setForm] = useState<GeneratorForm>(() => makeDefaultForm());
  const [images, setImages] = useState<UploadedImage[]>([]);
  const [assets, setAssets] = useState<PerformerAssets>({
    signature: null,
    seal: null,
  });
  const [preview, setPreview] = useState<PreviewKind>("contract");
  const [busy, setBusy] = useState<"word" | "pdf" | null>(null);
  const [toast, setToast] = useState("");
  const [draftLoaded, setDraftLoaded] = useState(false);

  useEffect(() => {
    const frame = window.requestAnimationFrame(() => {
      try {
        const raw = window.localStorage.getItem("rs-contract-generator-draft-v3");
        if (raw) {
          const saved = JSON.parse(raw) as Partial<GeneratorForm>;
          const defaults = makeDefaultForm();
          setForm({
            ...defaults,
            ...saved,
            customer: { ...defaults.customer, ...(saved.customer || {}) },
            contractBodies: {
              ...defaults.contractBodies,
              ...(saved.contractBodies || {}),
            },
            items: saved.items?.length ? saved.items : defaults.items,
          });
        }
      } catch {
        window.localStorage.removeItem("rs-contract-generator-draft-v3");
      } finally {
        setDraftLoaded(true);
      }
    });
    return () => window.cancelAnimationFrame(frame);
  }, []);

  useEffect(() => {
    let active = true;
    void loadPerformerAssets()
      .then(async (savedAssets) => {
        const [signature, seal] = await Promise.all([
          savedAssets.signature
            ? fileToUploadedImage(savedAssets.signature, "performer-signature")
            : Promise.resolve(null),
          savedAssets.seal
            ? fileToUploadedImage(savedAssets.seal, "performer-seal")
            : Promise.resolve(null),
        ]);
        if (!active) {
          if (signature) URL.revokeObjectURL(signature.url);
          if (seal) URL.revokeObjectURL(seal.url);
          return;
        }
        setAssets({ signature, seal });
        setForm((current) => ({
          ...current,
          includeSignature: Boolean(signature),
          includeSeal: current.entity === "ooo" && Boolean(seal),
        }));
      })
      .catch(() => {
        setToast("Не удалось прочитать сохранённые подпись и печать");
      });
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 2600);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const totals = useMemo(() => {
    const total = roundMoney(
      form.items.reduce(
        (sum, item) => sum + Number(item.quantity || 0) * Number(item.price || 0),
        0,
      ),
    );
    const vat = form.entity === "ooo" ? roundMoney((total * 22) / 122) : 0;
    const prepayment = roundMoney((total * form.prepaymentPercent) / 100);
    return {
      total,
      vat,
      prepayment,
      remainder: roundMoney(total - prepayment),
    };
  }, [form.entity, form.items, form.prepaymentPercent]);

  const completed = [
    Boolean(form.contractNumber && form.contractDate),
    Boolean(form.customer.name),
    Boolean(form.servicePlace),
    form.items.some((item) => item.name.trim() && Number(item.price) > 0),
  ].filter(Boolean).length;

  const updateCustomer = <K extends keyof CustomerDetails>(
    key: K,
    value: CustomerDetails[K],
  ) => {
    setForm((current) => ({
      ...current,
      customer: { ...current.customer, [key]: value },
    }));
  };

  const updateItem = <K extends keyof ServiceItem>(
    id: string,
    key: K,
    value: ServiceItem[K],
  ) => {
    setForm((current) => ({
      ...current,
      items: current.items.map((item) =>
        item.id === id ? { ...item, [key]: value } : item,
      ),
    }));
  };

  const addItem = () => {
    setForm((current) => ({
      ...current,
      items: [
        ...current.items,
        {
          id: `item-${Date.now()}`,
          name: "",
          unit: "усл.",
          quantity: 1,
          price: 0,
        },
      ],
    }));
  };

  const removeItem = (id: string) => {
    setForm((current) => ({
      ...current,
      items:
        current.items.length === 1
          ? current.items
          : current.items.filter((item) => item.id !== id),
    }));
  };

  const handleImages = async (files: FileList | null) => {
    if (!files) return;
    const accepted = Array.from(files).filter((file) =>
      ["image/png", "image/jpeg"].includes(file.type),
    );
    const slots = Math.max(0, 3 - images.length);
    const next = await Promise.all(
      accepted.slice(0, slots).map(async (file) => {
        const dimensions = await readImageDimensions(file).catch(() => ({ width: 1, height: 1 }));
        return {
          id: `${file.name}-${file.lastModified}-${Math.random()}`,
          name: file.name,
          url: URL.createObjectURL(file),
          file,
          ...dimensions,
        };
      }),
    );
    setImages((current) => [...current, ...next]);
    if (accepted.length > slots) {
      setToast("Можно добавить не более трёх изображений");
    }
  };

  const handlePerformerAsset = async (
    kind: PerformerAssetKind,
    files: FileList | null,
  ) => {
    const file = files?.[0];
    if (!file) return;
    if (!["image/png", "image/jpeg"].includes(file.type)) {
      setToast("Поддерживаются только PNG и JPG");
      return;
    }
    if (file.size > 5 * 1024 * 1024) {
      setToast("Размер файла не должен превышать 5 МБ");
      return;
    }

    const next = await fileToUploadedImage(file, `performer-${kind}`);
    let persisted = true;
    try {
      await savePerformerAsset(kind, file);
    } catch {
      persisted = false;
    }
    setAssets((current) => {
      const previous = current[kind];
      if (previous) URL.revokeObjectURL(previous.url);
      return { ...current, [kind]: next };
    });
    setForm((current) => ({
      ...current,
      includeSignature:
        kind === "signature" ? true : current.includeSignature,
      includeSeal:
        kind === "seal" && current.entity === "ooo"
          ? true
          : current.includeSeal,
    }));
    setToast(
      persisted
        ? kind === "signature"
          ? "Подпись сохранена в этом браузере"
          : "Печать сохранена в этом браузере"
        : "Файл добавлен, но браузер не смог сохранить его для следующего запуска",
    );
  };

  const removePerformerAsset = async (kind: PerformerAssetKind) => {
    try {
      await deletePerformerAsset(kind);
    } catch {
      setToast("Не удалось удалить сохранённый файл из браузера");
      return;
    }
    setAssets((current) => {
      const previous = current[kind];
      if (previous) URL.revokeObjectURL(previous.url);
      return { ...current, [kind]: null };
    });
    setForm((current) => ({
      ...current,
      includeSignature:
        kind === "signature" ? false : current.includeSignature,
      includeSeal: kind === "seal" ? false : current.includeSeal,
    }));
    setToast(kind === "signature" ? "Подпись удалена" : "Печать удалена");
  };

  const removeImage = (id: string) => {
    setImages((current) => {
      const target = current.find((image) => image.id === id);
      if (target) URL.revokeObjectURL(target.url);
      return current.filter((image) => image.id !== id);
    });
  };

  const changeEntity = (entity: EntityType) => {
    setForm((current) => ({
      ...current,
      entity,
      includeSeal: entity === "ooo" && Boolean(assets.seal),
    }));
  };

  const saveDraft = () => {
    window.localStorage.setItem(
      "rs-contract-generator-draft-v3",
      JSON.stringify(form),
    );
    setToast("Черновик сохранён на этом устройстве");
  };

  const resetTemplate = () => {
    setForm((current) => ({
      ...current,
      contractBodies: {
        ...current.contractBodies,
        [current.entity]: DEFAULT_BODIES[current.entity],
      },
    }));
    setToast("Текст шаблона восстановлен");
  };

  const handleWordExport = async () => {
    setBusy("word");
    try {
      await exportDocx(
        preview,
        { ...form, includeSeal: false },
        images,
        assets,
      );
      setToast(preview === "contract" ? "Договор Word сформирован без печати" : "Приложение Word сформировано без печати");
    } catch (error) {
      console.error(error);
      setToast("Не удалось сформировать Word — проверьте изображения и повторите");
    } finally {
      setBusy(null);
    }
  };

  const handlePdfExport = async () => {
    setBusy("pdf");
    try {
      await exportPdf(preview, form, images, assets);
      setToast(preview === "contract" ? "Договор PDF сформирован" : "Приложение PDF сформировано");
    } catch (error) {
      console.error(error);
      setToast("Не удалось сформировать PDF — проверьте изображения и повторите");
    } finally {
      setBusy(null);
    }
  };

  const pdfExportLabel =
    form.entity === "ooo"
      ? form.includeSeal && assets.seal
        ? "PDF с печатью"
        : form.includeSignature && assets.signature
          ? "PDF с подписью"
          : "PDF без печати"
      : form.includeSignature && assets.signature
        ? "PDF с подписью"
        : "PDF";

  return (
    <main className="app-shell">
      <header className="app-header">
        <div className="brand-block">
          <img className="brand-logo" src="/reklamastroy-logo.png" alt="Логотип РекламаСтрой" />
          <div>
            <span className="eyebrow">РекламаСтрой</span>
            <h1>Генератор договоров РекламаСтрой</h1>
          </div>
        </div>
        <div className="header-status" aria-label={`Заполнено ${completed} из 4 основных блоков`}>
          <span className="status-dot" />
          Черновик · {completed}/4
        </div>
        <div className="header-actions">
          <button className="button button-secondary" type="button" onClick={saveDraft}>
            Сохранить
          </button>
        </div>
      </header>

      <div className="workspace">
        <aside className="editor-panel" aria-label="Параметры договора">
          <nav className="section-nav" aria-label="Разделы формы">
            <a href="#main-data">Договор</a>
            <a href="#customer-data">Заказчик</a>
            <a href="#services-data">Услуги</a>
            <a href="#design-data">Оформление</a>
          </nav>

          <div className="editor-content">
            <section className="form-section" id="main-data">
              <div className="section-heading">
                <span>01</span>
                <div>
                  <h2>Основные данные</h2>
                  <p>Исполнитель определяет налогообложение и оформление.</p>
                </div>
              </div>

              <div className="entity-switch" role="group" aria-label="Выбор исполнителя">
                <button
                  type="button"
                  aria-pressed={form.entity === "ooo"}
                  className={form.entity === "ooo" ? "is-active" : ""}
                  onClick={() => changeEntity("ooo")}
                >
                  <strong>ООО</strong>
                  <span>НДС 22%</span>
                </button>
                <button
                  type="button"
                  aria-pressed={form.entity === "ip"}
                  className={form.entity === "ip" ? "is-active" : ""}
                  onClick={() => changeEntity("ip")}
                >
                  <strong>ИП</strong>
                  <span>Без НДС</span>
                </button>
              </div>

              <div className="field-grid">
                <Field label="Номер договора">
                  <input
                    value={form.contractNumber}
                    onChange={(event) =>
                      setForm((current) => ({ ...current, contractNumber: event.target.value }))
                    }
                    placeholder="123/1"
                  />
                </Field>
                <Field label="Дата договора">
                  <input
                    type="date"
                    value={form.contractDate}
                    onChange={(event) =>
                      setForm((current) => ({ ...current, contractDate: event.target.value }))
                    }
                  />
                </Field>
                <Field label="Город">
                  <input
                    value={form.city}
                    onChange={(event) =>
                      setForm((current) => ({ ...current, city: event.target.value }))
                    }
                    placeholder="Москва"
                  />
                </Field>
                <Field label="Номер приложения">
                  <input
                    value={form.appendixNumber}
                    onChange={(event) =>
                      setForm((current) => ({ ...current, appendixNumber: event.target.value }))
                    }
                    placeholder="1"
                  />
                </Field>
                <Field
                  label="Место оказания услуг"
                  hint="Попадёт в Приложение"
                  wide
                >
                  <input
                    value={form.servicePlace}
                    onChange={(event) =>
                      setForm((current) => ({ ...current, servicePlace: event.target.value }))
                    }
                    placeholder="Адрес объекта или несколько адресов"
                  />
                </Field>
              </div>
            </section>

            <section className="form-section" id="customer-data">
              <div className="section-heading">
                <span>02</span>
                <div>
                  <h2>Заказчик</h2>
                  <p>Все графы отдельные и необязательные.</p>
                </div>
              </div>
              <div className="field-grid">
                <Field label="Тип заказчика">
                  <select
                    value={form.customer.kind}
                    onChange={(event) =>
                      updateCustomer("kind", event.target.value as CustomerDetails["kind"])
                    }
                  >
                    <option value="ooo">Юридическое лицо</option>
                    <option value="ip">Индивидуальный предприниматель</option>
                    <option value="person">Физическое лицо</option>
                  </select>
                </Field>
                <Field label="Краткое наименование">
                  <input
                    value={form.customer.shortName}
                    onChange={(event) => updateCustomer("shortName", event.target.value)}
                    placeholder="ООО «Космо»"
                  />
                </Field>
                <Field label="Полное наименование / ФИО" wide>
                  <input
                    value={form.customer.name}
                    onChange={(event) => updateCustomer("name", event.target.value)}
                    placeholder="Общество с ограниченной ответственностью…"
                  />
                </Field>
                <Field label="Подписант">
                  <input
                    value={form.customer.representative}
                    onChange={(event) => updateCustomer("representative", event.target.value)}
                    placeholder="Иванов Иван Иванович"
                  />
                </Field>
                <Field label="Должность">
                  <input
                    value={form.customer.position}
                    onChange={(event) => updateCustomer("position", event.target.value)}
                    placeholder="Генеральный директор"
                  />
                </Field>
                <Field label="Основание полномочий" wide>
                  <input
                    value={form.customer.basis}
                    onChange={(event) => updateCustomer("basis", event.target.value)}
                    placeholder="Устава / доверенности №…"
                  />
                </Field>
              </div>

              <details className="details-block">
                <summary>
                  <span>Банковские и регистрационные реквизиты</span>
                  <small>0 обязательных полей</small>
                </summary>
                <div className="field-grid details-content">
                  <Field label="ИНН">
                    <input value={form.customer.inn} onChange={(event) => updateCustomer("inn", event.target.value)} />
                  </Field>
                  <Field label="КПП">
                    <input value={form.customer.kpp} onChange={(event) => updateCustomer("kpp", event.target.value)} />
                  </Field>
                  <Field label={form.customer.kind === "ip" ? "ОГРНИП" : "ОГРН"} wide>
                    <input value={form.customer.ogrn} onChange={(event) => updateCustomer("ogrn", event.target.value)} />
                  </Field>
                  <Field label="Юридический адрес" wide>
                    <input value={form.customer.address} onChange={(event) => updateCustomer("address", event.target.value)} />
                  </Field>
                  <Field label="Почтовый адрес" wide>
                    <input value={form.customer.postalAddress} onChange={(event) => updateCustomer("postalAddress", event.target.value)} />
                  </Field>
                  <Field label="Банк" wide>
                    <input value={form.customer.bank} onChange={(event) => updateCustomer("bank", event.target.value)} />
                  </Field>
                  <Field label="БИК">
                    <input value={form.customer.bik} onChange={(event) => updateCustomer("bik", event.target.value)} />
                  </Field>
                  <Field label="Расчётный счёт">
                    <input value={form.customer.account} onChange={(event) => updateCustomer("account", event.target.value)} />
                  </Field>
                  <Field label="Корреспондентский счёт" wide>
                    <input value={form.customer.corrAccount} onChange={(event) => updateCustomer("corrAccount", event.target.value)} />
                  </Field>
                  <Field label="E-mail">
                    <input type="email" value={form.customer.email} onChange={(event) => updateCustomer("email", event.target.value)} />
                  </Field>
                  <Field label="Телефон">
                    <input type="tel" value={form.customer.phone} onChange={(event) => updateCustomer("phone", event.target.value)} />
                  </Field>
                </div>
              </details>
            </section>

            <section className="form-section" id="services-data">
              <div className="section-heading">
                <span>03</span>
                <div>
                  <h2>Услуги и расчёты</h2>
                  <p>Количество строк не ограничено.</p>
                </div>
              </div>

              <div className="items-editor">
                {form.items.map((item, index) => (
                  <div className="item-card" key={item.id}>
                    <div className="item-number">{String(index + 1).padStart(2, "0")}</div>
                    <div className="item-fields">
                      <Field label="Наименование" wide>
                        <input
                          value={item.name}
                          onChange={(event) => updateItem(item.id, "name", event.target.value)}
                          placeholder="Изготовление и монтаж вывески"
                        />
                      </Field>
                      <div className="item-numbers">
                        <Field label="Ед.">
                          <input value={item.unit} onChange={(event) => updateItem(item.id, "unit", event.target.value)} />
                        </Field>
                        <Field label="Кол-во">
                          <input
                            type="number"
                            min="0"
                            step="0.01"
                            value={item.quantity}
                            onChange={(event) => updateItem(item.id, "quantity", Number(event.target.value))}
                          />
                        </Field>
                        <Field label="Цена, руб.">
                          <input
                            type="number"
                            min="0"
                            step="0.01"
                            value={item.price}
                            onChange={(event) => updateItem(item.id, "price", Number(event.target.value))}
                          />
                        </Field>
                      </div>
                    </div>
                    <button
                      className="icon-button"
                      type="button"
                      onClick={() => removeItem(item.id)}
                      disabled={form.items.length === 1}
                      aria-label={`Удалить позицию ${index + 1}`}
                    >
                      ×
                    </button>
                  </div>
                ))}
              </div>
              <button className="add-button" type="button" onClick={addItem}>
                <span>＋</span> Добавить позицию
              </button>

              <div className="calculation-panel">
                <div className="prepayment-control">
                  <label htmlFor="prepayment">Предоплата</label>
                  <div>
                    <input
                      id="prepayment"
                      type="range"
                      min="0"
                      max="100"
                      step="1"
                      value={form.prepaymentPercent}
                      onChange={(event) =>
                        setForm((current) => ({
                          ...current,
                          prepaymentPercent: Number(event.target.value),
                        }))
                      }
                    />
                    <input
                      className="percent-input"
                      type="number"
                      min="0"
                      max="100"
                      value={form.prepaymentPercent}
                      onChange={(event) =>
                        setForm((current) => ({
                          ...current,
                          prepaymentPercent: Math.min(100, Math.max(0, Number(event.target.value))),
                        }))
                      }
                    />
                    <span>%</span>
                  </div>
                </div>
                <dl className="calculation-list">
                  <div>
                    <dt>Стоимость</dt>
                    <dd>{formatMoney(totals.total)} ₽</dd>
                  </div>
                  <div>
                    <dt>{form.entity === "ooo" ? "НДС 22/122" : "Налогообложение"}</dt>
                    <dd>{form.entity === "ooo" ? `${formatMoney(totals.vat)} ₽` : "Без НДС"}</dd>
                  </div>
                  <div>
                    <dt>Предоплата</dt>
                    <dd>{formatMoney(totals.prepayment)} ₽</dd>
                  </div>
                  <div className="calculation-total">
                    <dt>Остаток</dt>
                    <dd>{formatMoney(totals.remainder)} ₽</dd>
                  </div>
                </dl>
                <p className="words-preview">{amountToWords(totals.total)}</p>
              </div>
            </section>

            <section className="form-section" id="design-data">
              <div className="section-heading">
                <span>04</span>
                <div>
                  <h2>Оформление</h2>
                  <p>Изображения попадут перед таблицей; подпись — по выбору в Word и PDF, печать ООО — только в PDF.</p>
                </div>
              </div>

              <div className="upload-zone">
                <input
                  id="image-upload"
                  type="file"
                  accept="image/png,image/jpeg"
                  multiple
                  onChange={(event) => {
                    void handleImages(event.target.files);
                    event.target.value = "";
                  }}
                  disabled={images.length >= 3}
                />
                <label htmlFor="image-upload">
                  <span className="upload-symbol">↥</span>
                  <strong>Добавить изображения</strong>
                  <small>PNG или JPG · до 3 файлов</small>
                </label>
              </div>
              {images.length ? (
                <div className="uploaded-list">
                  {images.map((image, index) => (
                    <div key={image.id}>
                      <img src={image.url} alt="" />
                      <span>
                        <strong>Изображение {index + 1}</strong>
                        <small>{image.name}</small>
                      </span>
                      <button type="button" onClick={() => removeImage(image.id)} aria-label={`Удалить ${image.name}`}>
                        ×
                      </button>
                    </div>
                  ))}
                </div>
              ) : null}

              <div className="performer-assets-heading">
                <strong>Подпись и печать исполнителя</strong>
                <small>PNG или JPG до 5 МБ. Файлы сохраняются только в текущем браузере.</small>
              </div>
              <div className="performer-assets-grid">
                <AssetUploadCard
                  id="signature-upload"
                  title="Подпись исполнителя"
                  description="Используется для ООО и ИП"
                  asset={assets.signature}
                  onSelect={(files) => {
                    void handlePerformerAsset("signature", files);
                  }}
                  onRemove={() => {
                    void removePerformerAsset("signature");
                  }}
                />
                <AssetUploadCard
                  id="seal-upload"
                  title="Печать ООО"
                  description={
                    form.entity === "ooo"
                      ? "Используется только в PDF"
                      : "Доступна при выборе ООО"
                  }
                  asset={assets.seal}
                  disabled={form.entity !== "ooo"}
                  onSelect={(files) => {
                    void handlePerformerAsset("seal", files);
                  }}
                  onRemove={() => {
                    void removePerformerAsset("seal");
                  }}
                />
              </div>

              <div className="toggle-group">
                <Toggle
                  checked={form.includeSignature}
                  onChange={(includeSignature) =>
                    setForm((current) => ({ ...current, includeSignature }))
                  }
                  label="Подпись исполнителя"
                  description={
                    assets.signature
                      ? "Добавить загруженную подпись в Word и PDF"
                      : "Сначала загрузите изображение подписи"
                  }
                  disabled={!assets.signature}
                />
                <Toggle
                  checked={form.includeSeal}
                  onChange={(includeSeal) =>
                    setForm((current) => ({ ...current, includeSeal }))
                  }
                  label="Печать ООО"
                  description={
                    form.entity !== "ooo"
                      ? "Для ИП печать не используется"
                      : assets.seal
                        ? "Добавить загруженную печать только в PDF"
                        : "Сначала загрузите изображение печати"
                  }
                  disabled={form.entity !== "ooo" || !assets.seal}
                />
              </div>

              <Field label="Дополнительные условия" hint="Появятся перед реквизитами" wide>
                <textarea
                  rows={5}
                  value={form.additionalConditions}
                  onChange={(event) =>
                    setForm((current) => ({ ...current, additionalConditions: event.target.value }))
                  }
                  placeholder="Дополнительные условия, правки или уточнения…"
                />
              </Field>

              <details className="details-block template-editor">
                <summary>
                  <span>Редактор полного текста договора</span>
                  <small>расширенный режим</small>
                </summary>
                <div className="details-content">
                  <p className="editor-note">
                    Здесь загружен текст выбранного действующего шаблона. Изменения сразу
                    отражаются в предпросмотре и DOCX.
                  </p>
                  <textarea
                    rows={18}
                    value={form.contractBodies[form.entity]}
                    onChange={(event) =>
                      setForm((current) => ({
                        ...current,
                        contractBodies: {
                          ...current.contractBodies,
                          [current.entity]: event.target.value,
                        },
                      }))
                    }
                  />
                  <button className="text-button" type="button" onClick={resetTemplate}>
                    Восстановить исходный текст {form.entity === "ooo" ? "ООО" : "ИП"}
                  </button>
                </div>
              </details>
            </section>
          </div>
        </aside>

        <section className="preview-panel" aria-label="Предварительный просмотр">
          <div className="preview-toolbar">
            <div className="preview-tabs" role="tablist" aria-label="Выбор документа">
              <button
                type="button"
                role="tab"
                aria-selected={preview === "contract"}
                className={preview === "contract" ? "is-active" : ""}
                onClick={() => setPreview("contract")}
              >
                Договор
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={preview === "appendix"}
                className={preview === "appendix" ? "is-active" : ""}
                onClick={() => setPreview("appendix")}
              >
                Приложение № {form.appendixNumber || "1"}
              </button>
            </div>
            <div className="preview-tools">
              <span>{PERFORMERS[form.entity].short}</span>
              <button type="button" onClick={handleWordExport} disabled={Boolean(busy)}>
                {busy === "word" ? "Формируем…" : "Word без печати"}
              </button>
              <button className="preview-download-primary" type="button" onClick={handlePdfExport} disabled={Boolean(busy)}>
                {busy === "pdf" ? "Формируем…" : pdfExportLabel}
              </button>
            </div>
          </div>
          <div className="preview-canvas">
            {preview === "contract" ? (
              <ContractPreview form={form} assets={assets} />
            ) : (
              <AppendixPreview form={form} images={images} assets={assets} />
            )}
          </div>
        </section>
      </div>

      {toast ? (
        <div className="toast" role="status">
          {toast}
        </div>
      ) : null}
      {!draftLoaded ? <span className="sr-only">Загрузка черновика</span> : null}
    </main>
  );
}
