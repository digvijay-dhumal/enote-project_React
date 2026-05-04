// shared/pdf-viewer/hooks/usePdfViewer.js
import { useEffect, useState, useMemo, useCallback } from 'react';
import PdfViewerEngine from '../core/PdfViewerEngine';

const usePdfViewer = (pdfData, plugins = [], containerRef) => {
  const [pages, setPages] = useState([]);
  const [currentPage, setCurrentPage] = useState(1);
  const [numPages, setNumPages] = useState(0);
  const [zoom, setZoom] = useState(1.0);

  const engine = useMemo(() => {
    const eng = new PdfViewerEngine();
    plugins.forEach(p => eng.registerPlugin(p));
    return eng;
  }, [plugins]); // plugins are static, but could be dynamic if needed

  useEffect(() => {
    if (!pdfData) return;
    engine.loadDocument(pdfData).then(() => {
      setNumPages(engine.numPages);
      engine.renderRange(1, 5); // initial buffer
    });

    const onPageRendered = ({ pageNum, dataUrl }) => {
      setPages(prev => {
        const newPages = [...prev];
        newPages[pageNum - 1] = dataUrl;
        return newPages;
      });
    };
    const onPageChanged = (page) => setCurrentPage(page);
    const onZoomChanged = (newZoom) => setZoom(newZoom);

    engine.on('pageRendered', onPageRendered);
    engine.on('pageChanged', onPageChanged);
    engine.on('zoomChanged', onZoomChanged);

    return () => {
      // cleanup listeners (if engine supports removal)
    };
  }, [engine, pdfData]);

  // Navigation actions
  const goToPage = useCallback((page) => engine.setCurrentPage(page), [engine]);
  const nextPage = useCallback(() => engine.nextPage?.(), [engine]);
  const prevPage = useCallback(() => engine.previousPage?.(), [engine]);

  // Zoom actions
  const setZoomLevel = useCallback((level) => engine.setZoom(level), [engine]);

  // Scroll detection – if navigationPlugin is installed, it might handle this.
  // Otherwise we can add a manual effect:
  useEffect(() => {
    if (!containerRef.current || !engine) return;
    const handleScroll = () => { /* detect page from scrollTop and call setCurrentPage */ };
    containerRef.current.addEventListener('scroll', handleScroll);
    return () => containerRef.current?.removeEventListener('scroll', handleScroll);
  }, [containerRef, engine]);

  return {
    pages,
    currentPage,
    numPages,
    zoom,
    setZoom: setZoomLevel,
    goToPage,
    nextPage,
    prevPage,
  };
};

export default usePdfViewer;