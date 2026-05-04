import React, { useState, useRef, useEffect, useCallback } from 'react';
import * as pdfjsLib from 'pdfjs-dist';

// Set PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js`;

const PDFAnnotator = ({ 
  pdfBlob, 
  fileName, 
  onAnnotationsSave, 
  onPdfSave, 
  isLocked, 
  currentUser,
  onEditModeChange,
  onLockStatusChange
}) => {
  const canvasRef = useRef(null);
  const containerRef = useRef(null);
  const [currentTool, setCurrentTool] = useState('pen');
  const [drawing, setDrawing] = useState(false);
  const [lastPosition, setLastPosition] = useState({ x: 0, y: 0 });
  const [penColor, setPenColor] = useState('#ff0000');
  const [highlightColor, setHighlightColor] = useState('#ffff00');
  const [pdfDocument, setPdfDocument] = useState(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [scale, setScale] = useState(1.0);
  const [annotations, setAnnotations] = useState([]);
  const [isEditing, setIsEditing] = useState(false);
  const [isLoading, setIsLoading] = useState(false);

  // Initialize PDF
  useEffect(() => {
    if (pdfBlob) {
      loadPDF(pdfBlob);
    }
  }, [pdfBlob]);

  const loadPDF = async (blob) => {
    try {
      setIsLoading(true);
      const arrayBuffer = await blob.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ 
        data: arrayBuffer,
        cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/',
        cMapPacked: true
      }).promise;
      setPdfDocument(pdf);
      await renderPage(pdf, currentPage, scale);
    } catch (error) {
      console.error('Error loading PDF:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const renderPage = async (pdf, pageNumber, renderScale) => {
    try {
      const page = await pdf.getPage(pageNumber);
      const canvas = canvasRef.current;
      if (!canvas) return;

      const ctx = canvas.getContext('2d');
      
      // Get the viewport with correct orientation
      const viewport = page.getViewport({ 
        scale: renderScale,
        rotation: 0 // Explicitly set rotation to 0 degrees
      });
      
      // Set canvas dimensions
      canvas.width = viewport.width;
      canvas.height = viewport.height;

      // Clear canvas with white background
      ctx.fillStyle = 'white';
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      // Render PDF page with correct transformation
      const renderContext = {
        canvasContext: ctx,
        viewport: viewport,
        transform: [1, 0, 0, 1, 0, 0] // Identity matrix - no transformation
      };

      await page.render(renderContext).promise;

      // Redraw existing annotations for this page
      redrawAnnotations();

    } catch (error) {
      console.error('Error rendering PDF page:', error);
    }
  };

  const redrawAnnotations = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    
    const ctx = canvas.getContext('2d');
    const pageAnnotations = annotations.filter(anno => anno.page === currentPage);
    
    pageAnnotations.forEach(annotation => {
      if (annotation.type === 'pen') {
        ctx.strokeStyle = annotation.color;
        ctx.lineWidth = annotation.lineWidth || 2;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        
        ctx.beginPath();
        annotation.points.forEach((point, index) => {
          if (index === 0) {
            ctx.moveTo(point.x, point.y);
          } else {
            ctx.lineTo(point.x, point.y);
          }
        });
        ctx.stroke();
      } else if (annotation.type === 'highlight') {
        ctx.fillStyle = annotation.color + '80';
        ctx.fillRect(annotation.x, annotation.y, annotation.width, annotation.height);
      }
    });
  };

  const getMousePos = (e) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    
    const rect = canvas.getBoundingClientRect();
    return {
      x: e.clientX - rect.left,
      y: e.clientY - rect.top
    };
  };

  const handleMouseDown = (e) => {
    if (isLocked || !isEditing || isLoading) return;

    const pos = getMousePos(e);
    setDrawing(true);
    setLastPosition(pos);

    if (currentTool === 'pen') {
      const newAnnotation = {
        id: Date.now().toString(),
        type: 'pen',
        page: currentPage,
        color: penColor,
        lineWidth: 2,
        points: [pos]
      };
      setAnnotations(prev => [...prev, newAnnotation]);
    } else if (currentTool === 'highlight') {
      const newAnnotation = {
        id: Date.now().toString(),
        type: 'highlight',
        page: currentPage,
        color: highlightColor + '80',
        x: pos.x,
        y: pos.y,
        width: 150,
        height: 25
      };
      setAnnotations(prev => [...prev, newAnnotation]);
      redrawAnnotations();
    }
  };

  const handleMouseMove = (e) => {
    if (!drawing || isLocked || !isEditing || isLoading) return;

    const pos = getMousePos(e);
    const canvas = canvasRef.current;
    if (!canvas) return;

    const ctx = canvas.getContext('2d');

    if (currentTool === 'pen') {
      setAnnotations(prev => {
        const updated = [...prev];
        const lastAnnotation = updated[updated.length - 1];
        if (lastAnnotation && lastAnnotation.type === 'pen') {
          lastAnnotation.points.push(pos);
          
          // Draw the line
          ctx.strokeStyle = lastAnnotation.color;
          ctx.lineWidth = lastAnnotation.lineWidth;
          ctx.lineCap = 'round';
          ctx.lineJoin = 'round';
          ctx.beginPath();
          ctx.moveTo(lastPosition.x, lastPosition.y);
          ctx.lineTo(pos.x, pos.y);
          ctx.stroke();
        }
        return updated;
      });
    }

    setLastPosition(pos);
  };

  const handleMouseUp = () => {
    if (!drawing) return;
    setDrawing(false);
    saveAnnotations();
  };

  const saveAnnotations = useCallback(() => {
    if (onAnnotationsSave && annotations.length > 0) {
      const annotationData = {
        annotations: annotations,
        timestamp: new Date().toISOString(),
        user: currentUser
      };
      onAnnotationsSave(annotationData);
    }
  }, [annotations, onAnnotationsSave, currentUser]);

  const handleClear = () => {
    if (isLocked || !isEditing) return;
    
    setAnnotations(prev => prev.filter(anno => anno.page !== currentPage));
    
    if (pdfDocument) {
      renderPage(pdfDocument, currentPage, scale);
    }
  };

  const handleDownload = async () => {
    if (!pdfDocument || !canvasRef.current) return;

    try {
      const canvas = canvasRef.current;
      canvas.toBlob((blob) => {
        if (onPdfSave) {
          onPdfSave(blob);
        }
      }, 'image/png');
    } catch (error) {
      console.error('Error downloading annotated PDF:', error);
    }
  };

  const nextPage = async () => {
    if (pdfDocument && currentPage < pdfDocument.numPages) {
      const newPage = currentPage + 1;
      setCurrentPage(newPage);
      await renderPage(pdfDocument, newPage, scale);
    }
  };

  const prevPage = async () => {
    if (pdfDocument && currentPage > 1) {
      const newPage = currentPage - 1;
      setCurrentPage(newPage);
      await renderPage(pdfDocument, newPage, scale);
    }
  };

  const zoomIn = async () => {
    const newScale = Math.min(scale + 0.2, 3.0);
    setScale(newScale);
    if (pdfDocument) {
      await renderPage(pdfDocument, currentPage, newScale);
    }
  };

  const zoomOut = async () => {
    const newScale = Math.max(0.5, scale - 0.2);
    setScale(newScale);
    if (pdfDocument) {
      await renderPage(pdfDocument, currentPage, newScale);
    }
  };

  const fitToWidth = async () => {
    if (!pdfDocument || !containerRef.current) return;
    
    const container = containerRef.current;
    const page = await pdfDocument.getPage(currentPage);
    const viewport = page.getViewport({ scale: 1 });
    const containerWidth = container.clientWidth - 40; // Padding
    const newScale = containerWidth / viewport.width;
    
    setScale(newScale);
    await renderPage(pdfDocument, currentPage, newScale);
  };

  const startEditing = () => {
    setIsEditing(true);
    if (onEditModeChange) {
      onEditModeChange(true);
    }
  };

  const stopEditing = () => {
    setIsEditing(false);
    if (onEditModeChange) {
      onEditModeChange(false);
    }
    saveAnnotations();
  };

  // Auto-fit on initial load and resize
  useEffect(() => {
    if (pdfDocument && containerRef.current) {
      const handleResize = () => {
        fitToWidth();
      };

      window.addEventListener('resize', handleResize);
      // Initial fit
      setTimeout(() => fitToWidth(), 100);

      return () => window.removeEventListener('resize', handleResize);
    }
  }, [pdfDocument, currentPage]);

  return (
    <div className="pdf-annotator" ref={containerRef}>
      {/* Header Controls */}
      <div className="pdf-controls-header">
        <div className="page-navigation">
          <button 
            onClick={prevPage} 
            disabled={!pdfDocument || currentPage <= 1 || isLoading}
            className="control-btn"
          >
            ← Prev
          </button>
          <span className="page-info">
            {isLoading ? 'Loading...' : `Page ${currentPage} of ${pdfDocument?.numPages || 0}`}
          </span>
          <button 
            onClick={nextPage} 
            disabled={!pdfDocument || currentPage >= pdfDocument?.numPages || isLoading}
            className="control-btn"
          >
            Next →
          </button>
        </div>

        <div className="zoom-controls">
          <button onClick={zoomOut} disabled={isLoading} className="control-btn">-</button>
          <span>{Math.round(scale * 100)}%</span>
          <button onClick={zoomIn} disabled={isLoading} className="control-btn">+</button>
          <button onClick={fitToWidth} disabled={isLoading} className="control-btn">
            Fit Width
          </button>
        </div>

        {!isLocked && (
          <div className="edit-controls">
            {!isEditing ? (
              <button onClick={startEditing} disabled={isLoading} className="control-btn edit-btn">
                ✏️ Start Editing
              </button>
            ) : (
              <button onClick={stopEditing} className="control-btn save-btn">
                💾 Stop Editing
              </button>
            )}
          </div>
        )}
      </div>

      {/* Toolbar */}
      {isEditing && (
        <div className="annotation-toolbar">
          <div className="tool-group">
            <button 
              onClick={() => setCurrentTool('pen')}
              className={`tool-btn ${currentTool === 'pen' ? 'active' : ''}`}
            >
              ✏️ Pen
            </button>
            <button 
              onClick={() => setCurrentTool('highlight')}
              className={`tool-btn ${currentTool === 'highlight' ? 'active' : ''}`}
            >
              🖍️ Highlight
            </button>
          </div>

          <div className="color-group">
            {currentTool === 'pen' && (
              <div className="color-picker">
                <label>Pen:</label>
                <input 
                  type="color" 
                  value={penColor}
                  onChange={(e) => setPenColor(e.target.value)}
                />
              </div>
            )}
            
            {currentTool === 'highlight' && (
              <div className="color-picker">
                <label>Highlight:</label>
                <input 
                  type="color" 
                  value={highlightColor}
                  onChange={(e) => setHighlightColor(e.target.value)}
                />
              </div>
            )}
          </div>

          <div className="action-group">
            <button onClick={handleClear} className="tool-btn clear-btn">
              🗑️ Clear
            </button>
            <button onClick={handleDownload} className="tool-btn download-btn">
              💾 Save PDF
            </button>
          </div>
        </div>
      )}

      {/* Canvas Container */}
      <div className="canvas-container">
        {isLoading && (
          <div className="pdf-loading">
            <div className="loading-spinner"></div>
            <p>Loading PDF...</p>
          </div>
        )}
        
        <canvas
          ref={canvasRef}
          className={`annotation-canvas ${isEditing ? 'editing' : ''} ${isLoading ? 'loading' : ''}`}
          onMouseDown={handleMouseDown}
          onMouseMove={handleMouseMove}
          onMouseUp={handleMouseUp}
          onMouseLeave={handleMouseUp}
          style={{ 
            cursor: isEditing ? (currentTool === 'pen' ? 'crosshair' : 'text') : 'default',
            display: isLoading ? 'none' : 'block'
          }}
        />
        
        {isLocked && (
          <div className="pdf-locked-overlay">
            <div className="locked-message">
              <span className="lock-icon">🔒</span>
              <p>PDF is locked for editing</p>
              <p>Currently being edited by: {currentUser?.name}</p>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default PDFAnnotator;