import React, { useState, useRef, useEffect } from 'react';
import * as pdfjsLib from 'pdfjs-dist';

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

const ReactPDFViewerAnnotated = ({ 
  pdfBlob, 
  fileName, 
  onAnnotationsSave, 
  onPdfSave,
  isLocked, 
  currentUser 
}) => {
  const canvasRef = useRef(null);
  const [pdfDocument, setPdfDocument] = useState(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [numPages, setNumPages] = useState(0);
  const [scale, setScale] = useState(1.5);
  const [annotations, setAnnotations] = useState([]);
  const [currentTool, setCurrentTool] = useState('select');
  const [selectedAnnotation, setSelectedAnnotation] = useState(null);

  useEffect(() => {
    if (pdfBlob) {
      loadPDF();
    }
  }, [pdfBlob]);

  const loadPDF = async () => {
    try {
      const url = URL.createObjectURL(pdfBlob);
      const loadingTask = pdfjsLib.getDocument(url);
      const pdf = await loadingTask.promise;
      
      setPdfDocument(pdf);
      setNumPages(pdf.numPages);
      await renderPage(pdf, currentPage);
    } catch (error) {
      console.error('Error loading PDF:', error);
    }
  };

  const renderPage = async (pdf, pageNumber) => {
    try {
      const page = await pdf.getPage(pageNumber);
      const viewport = page.getViewport({ scale });
      
      const canvas = canvasRef.current;
      const context = canvas.getContext('2d');
      
      canvas.height = viewport.height;
      canvas.width = viewport.width;
      
      const renderContext = {
        canvasContext: context,
        viewport: viewport
      };
      
      await page.render(renderContext).promise;
      
      // Render annotations
      renderAnnotations();
    } catch (error) {
      console.error('Error rendering page:', error);
    }
  };

  useEffect(() => {
    if (pdfDocument) {
      renderPage(pdfDocument, currentPage);
    }
  }, [currentPage, scale, pdfDocument, annotations]);

  const renderAnnotations = () => {
    const canvas = canvasRef.current;
    const context = canvas.getContext('2d');
    
    // Clear existing annotations
    context.clearRect(0, 0, canvas.width, canvas.height);
    
    // Re-render PDF page
    if (pdfDocument) {
      renderPage(pdfDocument, currentPage);
    }
    
    // Render annotations
    annotations
      .filter(ann => ann.page === currentPage)
      .forEach(annotation => {
        renderAnnotation(annotation, context);
      });
  };

  const renderAnnotation = (annotation, context) => {
    context.save();
    
    if (annotation.type === 'highlight') {
      context.fillStyle = 'rgba(255, 255, 0, 0.3)';
      context.fillRect(annotation.x, annotation.y, annotation.width, annotation.height);
    } else if (annotation.type === 'comment') {
      // Comment markers are handled in the overlay
    } else if (annotation.type === 'rectangle') {
      context.strokeStyle = 'rgba(255, 0, 0, 0.8)';
      context.lineWidth = 2;
      context.strokeRect(annotation.x, annotation.y, annotation.width, annotation.height);
    }
    
    context.restore();
  };

  const handleCanvasClick = (e) => {
    if (isLocked) return;
    
    const rect = canvasRef.current.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    
    if (currentTool === 'comment') {
      addComment(x, y);
    } else if (currentTool === 'highlight') {
      startHighlight(x, y);
    }
  };

  const addComment = (x, y) => {
    const newAnnotation = {
      id: Date.now().toString(),
      type: 'comment',
      page: currentPage,
      x: x,
      y: y,
      text: '',
      user: currentUser,
      timestamp: new Date().toISOString()
    };
    
    setAnnotations(prev => [...prev, newAnnotation]);
    setSelectedAnnotation(newAnnotation);
    saveAnnotations();
  };

  const startHighlight = (startX, startY) => {
    const mouseMoveHandler = (e) => {
      const rect = canvasRef.current.getBoundingClientRect();
      const currentX = e.clientX - rect.left;
      const currentY = e.clientY - rect.top;
      
      const width = currentX - startX;
      const height = currentY - startY;
      
      // Update temporary highlight
      const tempAnnotation = {
        id: 'temp',
        type: 'highlight',
        page: currentPage,
        x: startX,
        y: startY,
        width: width,
        height: height,
        user: currentUser,
        timestamp: new Date().toISOString()
      };
      
      setAnnotations(prev => [
        ...prev.filter(ann => ann.id !== 'temp'),
        tempAnnotation
      ]);
    };
    
    const mouseUpHandler = (e) => {
      const rect = canvasRef.current.getBoundingClientRect();
      const endX = e.clientX - rect.left;
      const endY = e.clientY - rect.top;
      
      const width = endX - startX;
      const height = endY - startY;
      
      if (Math.abs(width) > 5 && Math.abs(height) > 5) {
        const newAnnotation = {
          id: Date.now().toString(),
          type: 'highlight',
          page: currentPage,
          x: startX,
          y: startY,
          width: width,
          height: height,
          user: currentUser,
          timestamp: new Date().toISOString()
        };
        
        setAnnotations(prev => [
          ...prev.filter(ann => ann.id !== 'temp'),
          newAnnotation
        ]);
        saveAnnotations();
      } else {
        setAnnotations(prev => prev.filter(ann => ann.id !== 'temp'));
      }
      
      document.removeEventListener('mousemove', mouseMoveHandler);
      document.removeEventListener('mouseup', mouseUpHandler);
    };
    
    document.addEventListener('mousemove', mouseMoveHandler);
    document.addEventListener('mouseup', mouseUpHandler);
  };

  const updateAnnotation = (id, updates) => {
    setAnnotations(prev => 
      prev.map(ann => 
        ann.id === id ? { ...ann, ...updates } : ann
      )
    );
    saveAnnotations();
  };

  const deleteAnnotation = (id) => {
    setAnnotations(prev => prev.filter(ann => ann.id !== id));
    setSelectedAnnotation(null);
    saveAnnotations();
  };

  const saveAnnotations = () => {
    if (onAnnotationsSave) {
      onAnnotationsSave({
        annotations: annotations,
        timestamp: new Date().toISOString(),
        user: currentUser
      });
    }
  };

  const savePdfWithAnnotations = async () => {
    if (onPdfSave) {
      // Create a new canvas with annotations
      const canvas = canvasRef.current;
      const newCanvas = document.createElement('canvas');
      newCanvas.width = canvas.width;
      newCanvas.height = canvas.height;
      const context = newCanvas.getContext('2d');
      
      // Draw original PDF content
      context.drawImage(canvas, 0, 0);
      
      // Draw annotations
      annotations
        .filter(ann => ann.page === currentPage)
        .forEach(annotation => {
          renderAnnotation(annotation, context);
        });
      
      // Convert to blob
      newCanvas.toBlob((blob) => {
        onPdfSave(blob);
      }, 'application/pdf');
    }
  };

  if (!pdfBlob) {
    return <div>Loading PDF...</div>;
  }

  return (
    <div className="pdf-annotator-container">
      {/* Annotation Toolbar */}
      <div className="annotation-toolbar">
        <button 
          className={`tool-btn ${currentTool === 'select' ? 'active' : ''}`}
          onClick={() => setCurrentTool('select')}
          disabled={isLocked}
        >
          Select
        </button>
        <button 
          className={`tool-btn ${currentTool === 'comment' ? 'active' : ''}`}
          onClick={() => setCurrentTool('comment')}
          disabled={isLocked}
        >
          Add Comment
        </button>
        <button 
          className={`tool-btn ${currentTool === 'highlight' ? 'active' : ''}`}
          onClick={() => setCurrentTool('highlight')}
          disabled={isLocked}
        >
          Highlight
        </button>
        <button 
          className="tool-btn"
          onClick={() => setScale(prev => Math.min(prev + 0.25, 3))}
        >
          Zoom In
        </button>
        <button 
          className="tool-btn"
          onClick={() => setScale(prev => Math.max(prev - 0.25, 0.5))}
        >
          Zoom Out
        </button>
        <button 
          className="tool-btn"
          onClick={savePdfWithAnnotations}
          disabled={isLocked}
        >
          Save PDF
        </button>
        <div className="page-controls">
          <button 
            onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
            disabled={currentPage === 1}
          >
            Previous
          </button>
          <span>Page {currentPage} of {numPages}</span>
          <button 
            onClick={() => setCurrentPage(prev => Math.min(prev + 1, numPages))}
            disabled={currentPage === numPages}
          >
            Next
          </button>
        </div>
      </div>

      {/* PDF Viewer with Annotations */}
      <div className="pdf-viewer-area" onClick={handleCanvasClick}>
        <canvas 
          ref={canvasRef} 
          className="pdf-canvas"
        />
        
        {/* Render comment annotations as overlay */}
        {annotations
          .filter(ann => ann.page === currentPage && ann.type === 'comment')
          .map(annotation => (
            <AnnotationOverlay
              key={annotation.id}
              annotation={annotation}
              onUpdate={updateAnnotation}
              onDelete={deleteAnnotation}
              isSelected={selectedAnnotation?.id === annotation.id}
              onSelect={() => setSelectedAnnotation(annotation)}
              isLocked={isLocked}
            />
          ))
        }
      </div>

      {/* Annotations Sidebar */}
      <div className="annotations-sidebar">
        <h4>Annotations (Page {currentPage})</h4>
        {annotations
          .filter(ann => ann.page === currentPage)
          .map(annotation => (
            <AnnotationItem
              key={annotation.id}
              annotation={annotation}
              onUpdate={updateAnnotation}
              onDelete={deleteAnnotation}
              isLocked={isLocked}
            />
          ))
        }
        {annotations.filter(ann => ann.page === currentPage).length === 0 && (
          <p>No annotations on this page</p>
        )}
      </div>
    </div>
  );
};

// Annotation Overlay Component
const AnnotationOverlay = ({ annotation, onUpdate, onDelete, isSelected, onSelect, isLocked }) => {
  const [isEditing, setIsEditing] = useState(false);
  const [text, setText] = useState(annotation.text || '');

  const handleSave = () => {
    onUpdate(annotation.id, { text });
    setIsEditing(false);
  };

  return (
    <div 
      className={`annotation-comment ${isSelected ? 'selected' : ''}`}
      style={{
        left: annotation.x,
        top: annotation.y,
        position: 'absolute'
      }}
      onClick={(e) => {
        e.stopPropagation();
        onSelect();
      }}
    >
      <div className="annotation-header">
        <span className="user-badge">{annotation.user.name}</span>
        {!isLocked && (
          <div className="annotation-actions">
            <button onClick={() => setIsEditing(!isEditing)}>Edit</button>
            <button onClick={() => onDelete(annotation.id)}>Delete</button>
          </div>
        )}
      </div>
      {isEditing ? (
        <div>
          <textarea 
            value={text}
            onChange={(e) => setText(e.target.value)}
            rows="3"
            onClick={(e) => e.stopPropagation()}
          />
          <button onClick={handleSave}>Save</button>
          <button onClick={() => setIsEditing(false)}>Cancel</button>
        </div>
      ) : (
        <div className="annotation-text">
          {annotation.text || 'Click to add comment...'}
        </div>
      )}
    </div>
  );
};

// Annotation List Item Component
const AnnotationItem = ({ annotation, onUpdate, onDelete, isLocked }) => {
  const [isEditing, setIsEditing] = useState(false);
  const [text, setText] = useState(annotation.text || '');

  const handleSave = () => {
    onUpdate(annotation.id, { text });
    setIsEditing(false);
  };

  return (
    <div className="annotation-list-item">
      <div className="annotation-meta">
        <strong>{annotation.user.name}</strong>
        <span className="annotation-time">
          {new Date(annotation.timestamp).toLocaleString()}
        </span>
      </div>
      <div className="annotation-type">
        Type: {annotation.type}
      </div>
      {annotation.type === 'comment' && (
        isEditing ? (
          <div>
            <textarea 
              value={text}
              onChange={(e) => setText(e.target.value)}
              rows="3"
            />
            <button onClick={handleSave}>Save</button>
            <button onClick={() => setIsEditing(false)}>Cancel</button>
          </div>
        ) : (
          <div className="annotation-content">
            {annotation.text || 'No comment text'}
            {!isLocked && (
              <div className="annotation-actions">
                <button onClick={() => setIsEditing(true)}>Edit</button>
                <button onClick={() => onDelete(annotation.id)}>Delete</button>
              </div>
            )}
          </div>
        )
      )}
    </div>
  );
};

export default ReactPDFViewerAnnotated;