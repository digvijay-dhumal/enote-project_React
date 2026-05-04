import * as pdfjsLib from 'pdfjs-dist';

class PdfViewerEngine {
  constructor() {
    this.pdfDocument = null;
    this.numPages = 0;
    this.zoom = 1.0;
    this.currentPage = 1;
    this.pageCache = new Map(); // pageNumber -> canvas data URL or image element
    this.eventListeners = new Map();
    this.plugins = [];
  }

  // Load a PDF from an ArrayBuffer or Uint8Array
  async loadDocument(data) {
    this.pdfDocument = await pdfjsLib.getDocument({ data }).promise;
    this.numPages = this.pdfDocument.numPages;
    this.emit('documentLoaded', { numPages: this.numPages });
    return this;
  }

  // Render a specific page and cache it
  async renderPage(pageNum) {
    if (this.pageCache.has(pageNum)) return this.pageCache.get(pageNum);
    const page = await this.pdfDocument.getPage(pageNum);
    const viewport = page.getViewport({ scale: this.zoom });
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport }).promise;
    const dataUrl = canvas.toDataURL();
    this.pageCache.set(pageNum, dataUrl);
    this.emit('pageRendered', { pageNum, dataUrl });
    return dataUrl;
  }

  // Render a range of pages
  async renderRange(start, end) {
    const promises = [];
    for (let i = start; i <= end; i++) {
      if (!this.pageCache.has(i)) {
        promises.push(this.renderPage(i));
      }
    }
    await Promise.all(promises);
  }

  // Zoom control
  setZoom(newZoom) {
    if (newZoom === this.zoom) return;
    this.zoom = newZoom;
    this.pageCache.clear(); // invalidate cache
    this.emit('zoomChanged', newZoom);
  }

  // Page navigation
  setCurrentPage(page) {
    if (page < 1 || page > this.numPages) return;
    this.currentPage = page;
    this.emit('pageChanged', page);
  }

  // Plugin registration
  registerPlugin(plugin) {
    plugin.install(this);
    this.plugins.push(plugin);
  }

  // Event handling
  on(event, callback) {
    if (!this.eventListeners.has(event)) this.eventListeners.set(event, []);
    this.eventListeners.get(event).push(callback);
  }

  emit(event, data) {
    const listeners = this.eventListeners.get(event);
    if (listeners) listeners.forEach(cb => cb(data));
  }

  // Utility for printing – generate all pages (could be plugin)
  async getAllPages() {
    const pages = [];
    for (let i = 1; i <= this.numPages; i++) {
      pages.push(await this.renderPage(i));
    }
    return pages;
  }

  // Utility for download – return original bytes
  getOriginalData() { /* ... */ }
}

export default PdfViewerEngine;