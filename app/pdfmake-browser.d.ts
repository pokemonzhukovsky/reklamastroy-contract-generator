declare module "pdfmake/build/pdfmake" {
  interface BrowserPdf {
    download(filename?: string): void;
    getBlob(callback: (blob: Blob) => void): void;
    getBuffer(callback: (buffer: Uint8Array) => void): void;
  }

  interface PdfMakeBrowser {
    vfs?: Record<string, string>;
    addVirtualFileSystem?: (vfs: Record<string, string>) => void;
    createPdf(definition: Record<string, unknown>): BrowserPdf;
  }

  const pdfMake: PdfMakeBrowser;
  export default pdfMake;
}

declare module "pdfmake/build/vfs_fonts" {
  const vfs: Record<string, string>;
  export default vfs;
}
