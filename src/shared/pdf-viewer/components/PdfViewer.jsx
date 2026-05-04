import React, { useRef } from 'react';
import usePdfViewer from '../hooks/usePdfViewer';

const PdfViewer = ({ pdfData, plugins = [] }) => {
  const containerRef = useRef(null);
  const { pages, currentPage, numPages, zoom, setZoom, goToPage, nextPage, prevPage } =
    usePdfViewer(pdfData, plugins, containerRef);

  return (
    <div className="pdf-viewer-container">
      <div className="toolbar">
        <button onClick={prevPage} disabled={currentPage === 1}>Previous</button>
        <span>{currentPage} / {numPages}</span>
        <button onClick={nextPage} disabled={currentPage === numPages}>Next</button>
        <select value={zoom} onChange={(e) => setZoom(parseFloat(e.target.value))}>
          <option value={0.5}>50%</option>
          <option value={0.75}>75%</option>
          <option value={1}>100%</option>
          <option value={1.25}>125%</option>
          <option value={1.5}>150%</option>
        </select>
      </div>
      <div className="pdf-pages" ref={containerRef}>
        {pages.map((page, index) => (
          <div key={index} className="pdf-page" data-page={index + 1}>
            {page && <img src={page} alt={`Page ${index + 1}`} />}
          </div>
        ))}
      </div>
    </div>
  );
};

export default PdfViewer;