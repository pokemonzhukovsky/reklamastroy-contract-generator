export type PerformerAssetKind = "signature" | "seal";

interface StoredAsset {
  kind: PerformerAssetKind;
  name: string;
  type: string;
  lastModified: number;
  blob: Blob;
}

const DATABASE_NAME = "reklamastroy-contract-generator";
const DATABASE_VERSION = 1;
const STORE_NAME = "performer-assets";

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = window.indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        database.createObjectStore(STORE_NAME, { keyPath: "kind" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Не удалось открыть хранилище"));
  });
}

function waitForTransaction(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () =>
      reject(transaction.error || new Error("Не удалось сохранить файл"));
    transaction.onabort = () =>
      reject(transaction.error || new Error("Сохранение файла отменено"));
  });
}

export async function savePerformerAsset(
  kind: PerformerAssetKind,
  file: File,
): Promise<void> {
  const database = await openDatabase();
  try {
    const transaction = database.transaction(STORE_NAME, "readwrite");
    const completion = waitForTransaction(transaction);
    transaction.objectStore(STORE_NAME).put({
      kind,
      name: file.name,
      type: file.type,
      lastModified: file.lastModified,
      blob: file,
    } satisfies StoredAsset);
    await completion;
  } finally {
    database.close();
  }
}

export async function loadPerformerAssets(): Promise<
  Partial<Record<PerformerAssetKind, File>>
> {
  const database = await openDatabase();
  try {
    const transaction = database.transaction(STORE_NAME, "readonly");
    const completion = waitForTransaction(transaction);
    const request = transaction.objectStore(STORE_NAME).getAll();
    const records = await new Promise<StoredAsset[]>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result as StoredAsset[]);
      request.onerror = () =>
        reject(request.error || new Error("Не удалось загрузить сохранённые файлы"));
    });
    await completion;
    return Object.fromEntries(
      records.map((record) => [
        record.kind,
        new File([record.blob], record.name, {
          type: record.type,
          lastModified: record.lastModified,
        }),
      ]),
    );
  } finally {
    database.close();
  }
}

export async function deletePerformerAsset(
  kind: PerformerAssetKind,
): Promise<void> {
  const database = await openDatabase();
  try {
    const transaction = database.transaction(STORE_NAME, "readwrite");
    const completion = waitForTransaction(transaction);
    transaction.objectStore(STORE_NAME).delete(kind);
    await completion;
  } finally {
    database.close();
  }
}
