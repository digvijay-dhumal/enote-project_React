
import React, { forwardRef, useImperativeHandle } from "react";
import * as pdfjsLib from "pdfjs-dist";
import { useEffect, useRef, useState, useCallback, useMemo } from "react";
//priti shelke - 06-04-2026


import * as PDFLib from "pdf-lib";
import "../styles/pdf-text-layer.css";
import { getAccessToken, useTabContext } from "../App";
import { loginRequest, API_COMMON_HEADERS, API_BASE_URL, API_ENDPOINTS } from "../config";
import { useMsal, useAccount } from "@azure/msal-react";

// Import Kendo SVG Icons
import {
  zoomInIcon,
  zoomOutIcon,
  caretAltLeftIcon,
  caretAltRightIcon,
  brushIcon,
  commentIcon,
  downloadIcon,
  printIcon,
  trashIcon,
  filePdfIcon,
  pencilIcon, infoCircleIcon
} from "@progress/kendo-svg-icons";
import { SvgIcon } from "@progress/kendo-react-common";

// Configure PDF.js worker with proper path
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

////////////////////////////////////////////////////////////
// EXPORT PDF WITH ANNOTATIONS Digvijay Dhumal 06-04
////////////////////////////////////////////////////////////
async function exportPdfWithAnnotations(file, annotations, scale = 1.5) {
  const bytes = await file.arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(bytes);
  const pages = pdfDoc.getPages();

  for (const a of annotations) {
    const pageIndex = a.page - 1;
    if (pageIndex < 0 || pageIndex >= pages.length) continue;

    const page = pages[pageIndex];
    const pageHeight = page.getHeight();
    const pageWidth = page.getWidth();

    if (a.type === "highlight") {
      let rectsToUse = [];

      if (a.normalizedRects && a.normalizedRects.length > 0) {
        rectsToUse = a.normalizedRects.map(rect => ({
          x: (rect.x / 100) * pageWidth,
          y: (rect.y / 100) * pageHeight,
          width: (rect.width / 100) * pageWidth,
          height: (rect.height / 100) * pageHeight,
        }));
      } else if (a.rects && a.rects.length > 0) {
        rectsToUse = a.rects;
      }

      for (const r of rectsToUse) {
        const pdfX = r.x;
        const pdfY = pageHeight - r.y - r.height;

        page.drawRectangle({
          x: pdfX,
          y: pdfY,
          width: r.width,
          height: r.height,
          color: PDFLib.rgb(1, 1, 0),
          opacity: 0.4,
        });
      }
    }

    if (a.type === "comment" && a.text?.trim()) {
      const iconSize = 24;
      const rightMargin = 12;
      let commentY = 50;

      if (a.normalizedRects && a.normalizedRects.length > 0) {
        commentY = (a.normalizedRects[0].y / 100) * pageHeight;
      } else if (a.rects && a.rects.length > 0) {
        commentY = a.rects[0].y;
      }

      commentY = Math.max(iconSize + 20, Math.min(commentY, pageHeight - 100));
      const commentX = pageWidth - iconSize - rightMargin;
      const pdfY = pageHeight - commentY;

      let annots = page.node.Annots();
      if (!annots) {
        annots = pdfDoc.context.obj([]);
        page.node.set(PDFLib.PDFName.of("Annots"), annots);
      }

      const date = new Date(a.timestamp).toLocaleString();

      // FIX: Line number properly formatted from database
      let lineNumberDisplay = 'N/A';
      if (a.lineNumber && a.lineNumber !== 'N/A' && a.lineNumber !== null && a.lineNumber !== undefined) {
        lineNumberDisplay = a.lineNumber;
      }
      if (a.lineNumberRange && a.lineNumberRange !== 'N/A') {
        lineNumberDisplay = a.lineNumberRange;
      }

      const commentText = `Page: ${a.page}\nLine: ${lineNumberDisplay}\nSelected Text: ${a.selectedText || 'N/A'}\n\nComment:\n${a.text}\n\n---\nAdded by: ${a.user || 'User'} on ${date}`;
      annots.push(
        pdfDoc.context.obj({
          Type: PDFLib.PDFName.of("Annot"),
          Subtype: PDFLib.PDFName.of("Text"),
          Contents: PDFLib.PDFString.of(commentText),
          Rect: [commentX, pdfY - iconSize, commentX + iconSize, pdfY],
          Name: PDFLib.PDFName.of("Comment"),
          T: PDFLib.PDFString.of(a.user || "User"),
          Open: false,
        })
      );
    }
  }

  return await pdfDoc.save();
}



function wrapTextForPDF(text, maxCharsPerLine) {
  const words = text.split(' ');
  const lines = [];
  let currentLine = '';

  for (const word of words) {
    const testLine = currentLine ? `${currentLine} ${word}` : word;
    if (testLine.length > maxCharsPerLine && currentLine) {
      lines.push(currentLine);
      currentLine = word;
    } else {
      currentLine = testLine;
    }
  }
  if (currentLine) {
    lines.push(currentLine);
  }
  return lines;
}



////////////////////////////////////////////////////////////
// HELPER FUNCTION TO PREPARE ANNOTATIONS FOR EXPORT Digvijay Dhumal 06-04
////////////////////////////////////////////////////////////
async function prepareAnnotationsForExport(annotations, pdfDoc) {
  if (!pdfDoc || !annotations || annotations.length === 0) return annotations;

  const preparedAnnotations = [];

  for (const annotation of annotations) {
    if (annotation.type === "highlight") {
      // If annotation has normalizedRects but no rects, convert them
      if (annotation.normalizedRects && annotation.normalizedRects.length > 0 && (!annotation.rects || annotation.rects.length === 0)) {
        const page = await pdfDoc.getPage(annotation.page - 1);
        const pageWidth = page.getWidth();
        const pageHeight = page.getHeight();

        const rects = annotation.normalizedRects.map(rect => ({
          x: (rect.x / 100) * pageWidth,
          y: (rect.y / 100) * pageHeight,
          width: (rect.width / 100) * pageWidth,
          height: (rect.height / 100) * pageHeight,
        }));

        preparedAnnotations.push({ ...annotation, rects });
      } else {
        preparedAnnotations.push(annotation);
      }
    } else {
      preparedAnnotations.push(annotation);
    }
  }

  return preparedAnnotations;
}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////
// RESET CONFIRMATION DIALOG COMPONENT
////////////////////////////////////////////////////////////
function ResetConfirmationDialog({ isOpen, onClose, onConfirm, annotationCount }) {
  if (!isOpen) return null;

  return (
    <div
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: "rgba(0,0,0,0.5)",
        zIndex: 10000,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        animation: "fadeIn 0.2s ease-out"
      }}
      onClick={onClose}
    >
      <div
        style={{
          backgroundColor: "white",
          borderRadius: "8px",
          padding: "24px",
          width: "400px",
          maxWidth: "90%",
          boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
          zIndex: 10001,
          animation: "slideUp 0.3s ease-out"
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ textAlign: "center", marginBottom: "20px" }}>
          <h3 style={{
            margin: "0 0 14px 0",
            fontSize: "17px",
            fontWeight: "bold",
            textAlign: "left",
            color: "#333"
          }}>
            Delete All Annotations
          </h3>
          <p style={{
            margin: "0",
            fontSize: "14px",
            color: "#666",
            lineHeight: "1.5",
            textAlign: "left",
          }}>
            Are you sure you want to delete all annotations?
          </p>
        </div>

        <div style={{
          display: "flex",
          gap: "11px",
          justifyContent: "flex-end",
          marginTop: "8px"
        }}>
          <button
            onClick={onClose}
            style={{
              padding: "8px 16px",
              backgroundColor: "#e9ecef",
              border: "none",
              borderRadius: "6px",
              cursor: "pointer",
              fontSize: "14px",
              fontWeight: "500",
              color: "#495057",
              transition: "all 0.2s",
              fontFamily: "inherit"
            }}
            onMouseEnter={(e) => e.target.style.backgroundColor = "#dee2e6"}
            onMouseLeave={(e) => e.target.style.backgroundColor = "#e9ecef"}
          >
            Cancel
          </button>
          <button
            onClick={onConfirm}
            style={{
              padding: "8px 24px",
              backgroundColor: "#0056b3",
              border: "none",
              borderRadius: "6px",
              cursor: "pointer",
              fontSize: "14px",
              fontWeight: "500",
              color: "white",
              transition: "all 0.2s",
              fontFamily: "inherit"
            }}
            onMouseEnter={(e) => e.target.style.backgroundColor = "#0056b3"}
            onMouseLeave={(e) => e.target.style.backgroundColor = "#0056b3"}
          >
            Delete All
          </button>
        </div>
      </div>
    </div>
  );
}

const PdfPage = React.memo(({
  pdfDoc,
  pageNum,
  scale,
  annotations,
  hoveredId,
  setHoveredId,
  onCommentClick,
  onViewOnlyComment,
  onEditComment,
  onViewComment,
  onTextLinesReady,
}) => {
  const canvasRef = useRef(null);
  const textLayerRef = useRef(null);
  const renderTaskRef = useRef(null);
  const [isRendering, setIsRendering] = useState(false);
  const [textLines, setTextLines] = useState([]);
  const [renderError, setRenderError] = useState(null);
  const [pageDimensions, setPageDimensions] = useState({ width: 0, height: 0 });



  useEffect(() => {
    if (!pdfDoc) return;

    let cancelled = false;

    async function renderPage() {
      if (isRendering) return;
      setIsRendering(true);
      setRenderError(null);

      try {
        const page = await pdfDoc.getPage(pageNum);
        const viewport = page.getViewport({ scale });

        // Store page dimensions at current scale for coordinate transformation
        setPageDimensions({
          width: viewport.width,
          height: viewport.height
        });

        const canvas = canvasRef.current;
        if (!canvas) return;

        const ctx = canvas.getContext("2d");

        canvas.width = viewport.width;
        canvas.height = viewport.height;

        if (renderTaskRef.current) {
          try {
            await renderTaskRef.current.cancel();
          } catch (err) {
            console.log("Cancel error:", err);
          }
        }

        const renderTask = page.render({
          canvasContext: ctx,
          viewport,
        });

        renderTaskRef.current = renderTask;
        await renderTask.promise;

        if (cancelled) return;

        // In PdfPage component, update the text layer rendering section:

        // Render text layer for selection
        const textContent = await page.getTextContent();
        const container = textLayerRef.current;

        if (container && !cancelled) {
          container.innerHTML = "";

          // Set proper styles for text layer
          container.style.setProperty('--scale-factor', viewport.scale);
          container.style.setProperty('--page-width', viewport.width + 'px');
          container.style.setProperty('--page-height', viewport.height + 'px');

          const textLayer = await pdfjsLib.renderTextLayer({
            textContentSource: textContent,
            container: container,
            viewport: viewport,
            textDivs: [],
            enhanceTextSelection: true, // This is important for accurate selection
          });

          await textLayer;

          // Mark text layer as rendered
          container.setAttribute('data-rendered', 'true');

          // Get all text spans and process line information
          const textSpans = container.querySelectorAll('span');
          const lines = [];
          let currentLineNumber = 1;
          let currentLineY = null;
          let currentLineSpans = [];

          const spansWithRects = Array.from(textSpans).map(span => {
            const rect = span.getBoundingClientRect();
            const containerRect = container.getBoundingClientRect();
            return {
              span,
              rect,
              relativeY: Math.round(rect.top - containerRect.top),
              relativeBottom: Math.round(rect.bottom - containerRect.top),
              relativeX: Math.round(rect.left - containerRect.left)
            };
          });

          // Sort by Y position first, then by X position
          spansWithRects.sort((a, b) => {
            if (Math.abs(a.relativeY - b.relativeY) > 5) {
              return a.relativeY - b.relativeY;
            }
            return a.relativeX - b.relativeX;
          });

          const TOLERANCE = 5;

          spansWithRects.forEach(({ span, relativeY, relativeBottom, relativeX }) => {
            if (currentLineY === null || Math.abs(relativeY - currentLineY) > TOLERANCE) {
              if (currentLineSpans.length > 0) {
                const lineSpans = [...currentLineSpans];
                const minY = Math.min(...lineSpans.map(s => s.relativeY));
                const maxY = Math.max(...lineSpans.map(s => s.relativeBottom));

                lines.push({
                  lineNumber: currentLineNumber,
                  yPosition: currentLineY,
                  yMin: minY,
                  yMax: maxY,
                  spans: lineSpans.map(s => s.span)
                });
                currentLineNumber++;
              }
              currentLineY = relativeY;
              currentLineSpans = [{ span, relativeY, relativeBottom, relativeX }];
            } else {
              currentLineSpans.push({ span, relativeY, relativeBottom, relativeX });
            }
          });

          // Add the last line
          if (currentLineSpans.length > 0) {
            const lineSpans = [...currentLineSpans];
            const minY = Math.min(...lineSpans.map(s => s.relativeY));
            const maxY = Math.max(...lineSpans.map(s => s.relativeBottom));

            lines.push({
              lineNumber: currentLineNumber,
              yPosition: currentLineY,
              yMin: minY,
              yMax: maxY,
              spans: lineSpans.map(s => s.span)
            });
          }

          setTextLines(lines);

          if (onTextLinesReady && !cancelled) {
            onTextLinesReady(pageNum, lines);
          }
        }
      } catch (err) {
        if (err?.name !== "RenderingCancelledException") {
          console.error(`Error rendering page ${pageNum}:`, err);
          setRenderError(err.message);
        }
      } finally {
        if (!cancelled) {
          setIsRendering(false);
        }
      }
    }

    renderPage();

    return () => {
      cancelled = true;
      if (renderTaskRef.current) {
        try {
          renderTaskRef.current.cancel();
        } catch (err) {
          console.log("Cleanup cancel error:", err);
        }
      }
    };
  }, [pdfDoc, pageNum, scale, onTextLinesReady]);

  // In PdfPage component, replace the existing pageAnnotations useMemo with:

  const pageAnnotations = useMemo(() => {
    return annotations
      .filter((a) => a.page === pageNum)
      .map((annotation) => {
        const normalizedRects = Array.isArray(annotation.normalizedRects)
          ? annotation.normalizedRects
          : [];

        const transformedRects =
          normalizedRects.length > 0 && pageDimensions.width > 0 && pageDimensions.height > 0
            ? normalizedRects.map((rect) => ({
              x: (Number(rect.x) / 100) * pageDimensions.width,
              y: (Number(rect.y) / 100) * pageDimensions.height,
              width: (Number(rect.width) / 100) * pageDimensions.width,
              height: (Number(rect.height) / 100) * pageDimensions.height,
            }))
            : [];

        return {
          ...annotation,
          rects: transformedRects,
        };
      });
  }, [annotations, pageNum, pageDimensions.width, pageDimensions.height]);
  //this is newly added
  ////////////////////////Digvijay update this code for find correct line no on pop up
  const getLineNumberFromY = useCallback((yPosition) => {
    if (!textLines.length) return 1;

    // Find the line that contains this Y position
    let closestLine = textLines[0];
    let minDistance = Infinity;

    for (const line of textLines) {
      // Check if yPosition falls within this line's vertical range
      if (line.yMin !== undefined && line.yMax !== undefined) {
        if (yPosition >= line.yMin && yPosition <= line.yMax) {
          return line.lineNumber;
        }
      }

      // Fallback: find closest line by yPosition
      const distance = Math.abs(line.yPosition - yPosition);
      if (distance < minDistance) {
        minDistance = distance;
        closestLine = line;
      }
    }

    // If within 20px, return the closest line
    if (minDistance < 20) {
      return closestLine.lineNumber;
    }

    return 1;
  }, [textLines]);
  ///////////////////////////////////////////////////////////////

  const getPopupPositionStyle = useCallback((rect) => {
    const THRESHOLD = 300; // top detection (adjust 80–120 if needed)

    const isNearTop = rect?.y < THRESHOLD;

    return {
      position: "absolute",
      left: "0px",

      // 👇 MAIN LOGIC
      top: isNearTop ? "calc(100% + 8px)" : "-10px",
      transform: isNearTop ? "none" : "translateY(-100%)",

      background: "#fffdf5",
      padding: "12px 16px",
      border: "1px solid #f4c542",
      borderRadius: "8px",
      fontSize: "12px",
      minWidth: "260px",
      maxWidth: "360px",
      boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
      zIndex: 9999,
      wordWrap: "break-word",
      whiteSpace: "pre-wrap",
      pointerEvents: "auto"
    };
  }, []);

  const getLineNumberFromAnnotation = useCallback((annotation) => {
    // First check if we have a stored line number
    if (annotation?.lineNumber && annotation?.lineNumber !== 'N/A' && annotation?.lineNumber !== null && annotation?.lineNumber !== undefined) {
      return annotation.lineNumber;
    }

    // Check for lineNumberRange
    if (annotation?.lineNumberRange && annotation?.lineNumberRange !== 'N/A') {
      return annotation.lineNumberRange;
    }

    // Fallback to calculate from text lines
    if (!textLines.length || pageDimensions.height <= 0) {
      return "N/A";
    }

    let actualY = null;

    if (Array.isArray(annotation?.normalizedRects) && annotation.normalizedRects.length > 0) {
      actualY = (Number(annotation.normalizedRects[0].y || 0) / 100) * pageDimensions.height;
    } else if (Array.isArray(annotation?.rects) && annotation.rects.length > 0) {
      actualY = Number(annotation.rects[0].y || 0);
    } else if (annotation?.yposition !== undefined && annotation?.yposition !== null) {
      actualY = (Number(annotation.yposition || 0) / 100) * pageDimensions.height;
    }

    if (actualY === null || Number.isNaN(actualY)) {
      return "N/A";
    }

    let closestLine = null;
    let minDistance = Infinity;

    for (const line of textLines) {
      if (line.yMin !== undefined && line.yMax !== undefined) {
        if (actualY >= line.yMin && actualY <= line.yMax) {
          return line.lineNumber;
        }
      }
      const distance = Math.abs((line.yPosition ?? 0) - actualY);
      if (distance < minDistance) {
        minDistance = distance;
        closestLine = line;
      }
    }

    return closestLine && minDistance < 50 ? closestLine.lineNumber : "N/A";
  }, [textLines, pageDimensions.height]);







  if (renderError) {
    return (
      <div style={{
        padding: "20px",
        textAlign: "center",
        color: "#dc3545",
        backgroundColor: "#f8d7da",
        borderRadius: "8px",
        margin: "20px"
      }}>
        Error rendering page: {renderError}
      </div>
    );
  }

  return (
    //Digvijay 01-04 for page gap
    <div
      className="pdf-page-container"
      data-page={pageNum}
      style={{
        position: "relative",
        marginBottom: 12,
        boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
        backgroundColor: "white",
        userSelect: "text",
        width: "fit-content",
        minWidth: "fit-content"
      }}
    >
      <canvas
        ref={canvasRef}
        style={{
          display: "block",
          width: `${pageDimensions.width}px`,
          height: `${pageDimensions.height}px`,
          pointerEvents: "none",
          position: "relative",
          zIndex: 1
        }}
      />

      <div
        ref={textLayerRef}
        className="pdf-text-layer"
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          width: `${pageDimensions.width}px`,
          height: `${pageDimensions.height}px`,
          pointerEvents: "auto",
          userSelect: "text",
          WebkitUserSelect: "text",
          zIndex: 10,
          backgroundColor: "transparent",
          cursor: "text"
        }}
      />


      {/* Highlight overlay with hover popup */}
      {/* Highlight overlay with hover popup Digvijay Dhumal For Highlight pop up */}
      {/* Highlight overlay with hover popup - FIXED VERSION */}
      {pageAnnotations.map((a) =>
        a.type === "highlight" &&
        a.rects?.map((r, i) => (
          <React.Fragment key={`${a.id}-highlight-${i}`}>
            {/* Highlight overlay div (transparent/yellow background) */}
            <div
              style={{
                position: "absolute",
                left: r.x,
                top: r.y,
                width: r.width,
                height: r.height,
                background: "rgba(255, 193, 7, 0.38)",
                border: hoveredId === a.id ? "1px solid rgba(255, 153, 0, 0.9)" : "none",
                pointerEvents: "auto",
                zIndex: hoveredId === a.id ? 30 : 12,
                borderRadius: "2px",
                cursor: "pointer",
              }}
              onMouseEnter={() => {
                if (window[`hideTimeout_${a.id}`]) {
                  clearTimeout(window[`hideTimeout_${a.id}`]);
                  window[`hideTimeout_${a.id}`] = null;
                }
                setHoveredId(a.id);
              }}
              onMouseLeave={() => {
                const hideTimeout = setTimeout(() => {
                  const popupElement = document.getElementById(`highlight-popup-${a.id}`);
                  if (popupElement && !popupElement.matches(":hover")) {
                    setHoveredId(null);
                  } else if (!popupElement) {
                    setHoveredId(null);
                  }
                }, 200);
                window[`hideTimeout_${a.id}`] = hideTimeout;
              }}
            />

            {i === 0 && hoveredId === a.id && (
              <div
                id={`highlight-popup-${a.id}`}
                style={{
                  position: "absolute",
                  left: `${r.x}px`,
                  top: r.y < 200 ? `${r.y + r.height + 10}px` : `${r.y - 10}px`,
                  transform: r.y < 200 ? "none" : "translateY(-100%)",
                  background: "white",
                  border: "1px solid #f4c542",
                  borderRadius: "10px",
                  fontSize: "12px",
                  width: "340px",
                  maxWidth: "calc(100vw - 40px)",
                  boxShadow: "0 6px 18px rgba(0,0,0,0.2)",
                  zIndex: 99999,
                  pointerEvents: "auto",
                  display: "flex",
                  flexDirection: "column",
                  overflow: "hidden",
                }}
                onMouseEnter={() => {
                  if (window[`hideTimeout_${a.id}`]) {
                    clearTimeout(window[`hideTimeout_${a.id}`]);
                    window[`hideTimeout_${a.id}`] = null;
                  }
                  setHoveredId(a.id);
                }}
                onMouseLeave={() => {
                  setHoveredId(null);
                }}
              >
                <div
                  style={{
                    position: "sticky",
                    top: 0,
                    zIndex: 10,
                    background: "white",
                    borderBottom: "2px solid #f4c542",
                    padding: "12px 14px 8px 14px",
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      marginBottom: "6px",
                    }}
                  >
                    <div
                      style={{
                        fontWeight: 700,
                        fontSize: "13px",
                        color: "#c28a00",
                        letterSpacing: "0.3px",
                      }}
                    >
                      Highlight
                    </div>
                  </div>

                  {/* Page and Line Info */}
                  <div
                    style={{
                      fontSize: "11px",
                      color: "#888",
                      display: "flex",
                      gap: "16px",
                      marginTop: "2px",
                    }}
                  >
                    <span>Page {a.page}</span>
                    <span>Line {getLineNumberFromAnnotation(a)}</span>
                  </div>
                </div>

                <div
                  style={{
                    padding: "10px 14px",
                    maxHeight: "180px",
                    overflowY: "auto",
                    overflowX: "hidden",
                    backgroundColor: "#fffef7",
                  }}
                >
                  <div
                    style={{
                      background: "#fff8e1",
                      borderLeft: "3px solid #ffc107",
                      padding: "10px 12px",
                      borderRadius: "6px",
                      fontSize: "12px",
                      lineHeight: "1.55",
                      color: "#2d2d2d",
                      wordBreak: "break-word",
                      whiteSpace: "pre-wrap",
                    }}
                  >
                    {a.selectedText || "No highlighted text available"}
                  </div>
                </div>

                <div
                  style={{
                    position: "sticky",
                    bottom: 0,
                    background: "white",
                    borderTop: "1px solid #f0e6c5",
                    padding: "8px 14px 10px 14px",
                    fontSize: "10px",
                    color: "#aaa",
                    display: "flex",
                    justifyContent: "space-between",
                  }}
                >
                  <span> By: {a.user || "User"}</span>
                  <span>
                    {a.timestamp ? new Date(a.timestamp).toLocaleString() : "N/A"}
                  </span>
                </div>
              </div>
            )}
          </React.Fragment>
        ))
      )}

      {/* // In the PdfPage component, replace the comment rendering section with this updated code: */}

      {/* Comment Icons - Keep 💬 for everyone */}
      {pageAnnotations.map((a) => (
        a.type === "comment" && a.rects.length > 0 && (
          <div
            key={a.id}
            data-comment-id={a.id}
            style={{
              position: "absolute",
              right: "10px",
              top: `${a.rects[0]?.y ?? 50}px`,
              cursor: "pointer",
              fontSize: "18px",
              zIndex: 20,
              background: "white",
              borderRadius: "50%",
              width: "28px",
              height: "28px",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              boxShadow: "0 2px 6px rgba(0,0,0,0.2)",
              transition: "transform 0.2s ease",
              pointerEvents: "auto"
            }}
            onMouseEnter={(e) => {
              if (window[`hideTimeout_${a.id}`]) {
                clearTimeout(window[`hideTimeout_${a.id}`]);
                window[`hideTimeout_${a.id}`] = null;
              }
              setHoveredId(a.id);
              e.currentTarget.style.transform = "scale(1.1)";
            }}
            onMouseLeave={(e) => {
              const hideTimeout = setTimeout(() => {
                const popupElement = document.getElementById(`comment-popup-${a.id}`);
                if (popupElement && !popupElement.matches(":hover")) {
                  setHoveredId(null);
                  if (e.currentTarget) {
                    e.currentTarget.style.transform = "scale(1)";
                  }
                } else if (!popupElement) {
                  setHoveredId(null);
                  if (e.currentTarget) {
                    e.currentTarget.style.transform = "scale(1)";
                  }
                }
              }, 200);

              window[`hideTimeout_${a.id}`] = hideTimeout;
            }}
            onClick={() => {
              if (!a.canEdit && onViewOnlyComment) {
                onViewOnlyComment(a);
              }
            }}
          >
            💬

            {hoveredId === a.id && (
              <div
                id={`comment-popup-${a.id}`}
                style={{
                  position: "absolute",
                  right: "35px",
                  top: `${(a.rects?.[0]?.y ?? 50) < 120 ? "0px" : "50%"}`,
                  transform: `${(a.rects?.[0]?.y ?? 50) < 120 ? "none" : "translateY(-50%)"}`,
                  background: "white",
                  padding: "0 16px 12px 16px",
                  border: "1px solid #e0e0e0",
                  borderRadius: "8px",
                  fontSize: "12px",
                  minWidth: "360px",
                  maxWidth: "520px",
                  maxHeight: "360px",
                  overflowY: "auto",
                  overflowX: "hidden",
                  boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
                  zIndex: 99999,
                  wordWrap: "break-word",
                  whiteSpace: "normal",
                  pointerEvents: "auto",
                }}
                onMouseEnter={() => {
                  if (window[`hideTimeout_${a.id}`]) {
                    clearTimeout(window[`hideTimeout_${a.id}`]);
                    window[`hideTimeout_${a.id}`] = null;
                  }
                  setHoveredId(a.id);
                }}
                onMouseLeave={() => {
                  setHoveredId(null);
                }}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    marginBottom: "8px",
                    borderBottom: "1px solid #e0e0e0",
                    padding: "10px 0 6px 0",
                    position: "sticky",
                    top: 0,
                    background: "white",
                    zIndex: 10,
                  }}
                >
                  <div style={{ fontWeight: "bold", color: "#17a2b8" }}>
                    Comment
                  </div>

                  <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                    {/* View button - always visible next to infoCircleIcon */}
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        // Open view-only popup (same as edit but with update button hidden)
                        if (onViewOnlyComment) {
                          onViewOnlyComment(a);
                        } else if (onEditComment) {
                          // Fallback: if onViewOnlyComment doesn't exist, use edit with a view flag
                          // You'll need to modify CommentInputDialog to support view mode
                          console.log("View comment:", a);
                        }
                      }}
                      style={{
                        border: "none",
                        background: "#f8f9fa",
                        cursor: "pointer",
                        padding: "4px 8px",
                        fontSize: "12px",
                        color: "#666",
                        display: "flex",
                        alignItems: "center",
                        gap: "4px",
                        borderRadius: "4px"
                      }}
                      title="View comment"
                    >
                      👁️ View
                    </button>

                    {/* Show edit icon (pencil) ONLY for comment owner */}
                    {a.canEdit ? (
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          onEditComment?.(a);
                        }}
                        style={{
                          border: "none",
                          background: "#f8f9fa",
                          cursor: "pointer",
                          padding: "4px 8px",
                          fontSize: "12px",
                          color: "#666",
                          display: "flex",
                          alignItems: "center",
                          gap: "4px",
                          borderRadius: "4px"
                        }}
                        title="Edit comment"
                      >
                        <SvgIcon icon={pencilIcon} size="small" />
                        Edit
                      </button>
                    ) : (
                      <div
                        style={{
                          fontSize: "12px",
                          color: "#999"
                        }}
                        title="View only - You cannot edit this comment"
                      >
                        <SvgIcon icon={infoCircleIcon} size="small" />
                      </div>
                    )}
                  </div>
                </div>

                <div style={{ fontSize: "11px", color: "#666", marginBottom: "8px" }}>
                  Page: {a.page} | Line: {getLineNumberFromAnnotation(a)}
                </div>

                {/* Display selected text */}
                {a.selectedText && a.selectedText.trim() !== "" && (
                  <div
                    style={{
                      marginBottom: "12px",
                      padding: "8px",
                      backgroundColor: "#f8f9fa",
                      borderLeft: "3px solid #17a2b8",
                      fontSize: "11px",
                      fontStyle: "italic",
                      color: "#555",
                      borderRadius: "4px",
                      whiteSpace: "pre-wrap",
                      wordBreak: "break-word",
                      lineHeight: "1.5",
                    }}
                  >
                    <div style={{ fontWeight: "bold", marginBottom: "6px" }}>
                      Selected Text:
                    </div>

                    <div
                      style={{
                        background: "#ffffff",
                        border: "1px solid #e0e0e0",
                        borderRadius: "6px",
                        padding: "8px",
                        marginBottom: "10px",
                        maxHeight: "200px",
                        overflowY: "auto",
                        overflowX: "hidden",
                        whiteSpace: "pre-wrap",
                        wordBreak: "break-word",
                        lineHeight: "1.5",
                      }}
                    >
                      {a.selectedText || "No selected text available"}
                    </div>
                  </div>
                )}

                <div style={{ marginBottom: "8px", lineHeight: "1.4" }}>
                  {a.text}
                </div>

                <div
                  style={{
                    fontSize: "10px",
                    color: "#999",
                    marginTop: "8px",
                    borderTop: "1px solid #f0f0f0",
                    paddingTop: "6px",
                  }}
                >
                  By: {a.user || "User"} •{" "}
                  {a.timestamp && new Date(a.timestamp).toLocaleString()}
                </div>

                {/* Show message for non-owners */}
                {!a.canEdit && (
                  <div
                    style={{
                      marginTop: "8px",
                      fontSize: "9px",
                      color: "#ff9800",
                      fontStyle: "italic",
                      textAlign: "center",
                    }}
                  ></div>
                )}
              </div>
            )}
          </div>
        )
      ))}
    </div>
  );
}, (prevProps, nextProps) => {
  return (
    prevProps.pageNum === nextProps.pageNum &&
    prevProps.scale === nextProps.scale &&
    prevProps.annotations === nextProps.annotations &&
    prevProps.pdfDoc === nextProps.pdfDoc &&
    prevProps.hoveredId === nextProps.hoveredId
  );
});

function CommentInputDialog({ isOpen, onClose, onSubmit, selectedText, pageNo, lineNumber, editComment, viewOnly = false }) {
  const [comment, setComment] = useState("");
  const textareaRef = useRef(null);

  useEffect(() => {
    if (isOpen && textareaRef.current) {
      setTimeout(() => textareaRef.current?.focus(), 100);
      if (editComment) {
        setComment(editComment.text);
      } else {
        setComment("");
      }
    }
  }, [isOpen, editComment]);

  const handleSubmit = () => {
    if (comment.trim()) {
      onSubmit(comment.trim());
    }
  };

  const handleKeyDown = (e) => {
    if (e.key === 'Enter' && e.ctrlKey) {
      handleSubmit();
    }
  };

  if (!isOpen) return null;

  return (
    <div
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: "rgba(0,0,0,0.5)",
        zIndex: 10000,
        display: "flex",
        alignItems: "center",
        justifyContent: "center"
      }}
      onClick={onClose}
    >
      <div
        style={{
          backgroundColor: "white",
          borderRadius: "12px",
          padding: "24px",
          width: "500px",
          maxWidth: "90%",
          boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
          zIndex: 10001
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <h3 style={{ margin: "0 0 16px 0", fontSize: "18px", fontWeight: "bold", color: "#333" }}>
          {viewOnly ? "View Comment" : editComment ? "Edit Comment" : "Add Comment"}        </h3>

        <div style={{
          backgroundColor: "#f8f9fa",
          padding: "12px",
          borderRadius: "6px",
          marginBottom: "16px",
          borderLeft: "3px solid #17a2b8"
        }}>
          <div style={{ fontSize: "12px", color: "#666", marginBottom: "8px", fontWeight: "500" }}>
            Reference Information:
          </div>
          <div style={{ display: "flex", gap: "16px", fontSize: "13px" }}>
            <span><strong>Page:</strong> {pageNo}</span>
            <span><strong>Line:</strong> {lineNumber || 'N/A'}</span>
          </div>
          {selectedText && !editComment && (
            <>
              <div style={{ fontSize: "12px", color: "#666", marginTop: "12px", marginBottom: "6px", fontWeight: "500" }}>
                Selected Text:
              </div>
              <div style={{
                fontSize: "13px",
                color: "#333",
                fontStyle: "italic",
                maxHeight: "200px",
                overflowY: "auto",
                overflowX: "hidden",
                whiteSpace: "pre-wrap",
                wordBreak: "break-word",
                lineHeight: "1.5",
                backgroundColor: "white",
                padding: "8px",
                borderRadius: "4px",
                border: "1px solid #e0e0e0"
              }}>
                {selectedText}
              </div>
            </>
          )}
          {editComment && selectedText && (
            <>
              <div style={{ fontSize: "12px", color: "#666", marginTop: "12px", marginBottom: "6px", fontWeight: "500" }}>
                Selected Text:
              </div>
              <div style={{
                fontSize: "13px",
                color: "#333",
                fontStyle: "italic",
                maxHeight: "80px",
                overflow: "auto",
                backgroundColor: "white",
                padding: "8px",
                borderRadius: "4px",
                border: "1px solid #e0e0e0"
              }}>
                {selectedText}
              </div>
            </>
          )}
        </div>

        <textarea
          ref={textareaRef}
          value={comment}
          readOnly={viewOnly}

          onChange={(e) => setComment(e.target.value)}
          onKeyDown={handleKeyDown}
          placeholder="Enter your comment here... (Ctrl+Enter to submit)"
          style={{
            width: "100%",
            minHeight: "120px",
            padding: "12px",
            border: "1px solid #ddd",
            borderRadius: "6px",
            fontSize: "14px",
            fontFamily: "inherit",
            resize: "vertical",
            marginBottom: "16px",
            outline: "none"
          }}
        />

        <div style={{ display: "flex", gap: "12px", justifyContent: "flex-end" }}>
          <button
            onClick={onClose}
            style={{
              padding: "8px 20px",
              backgroundColor: "#f8f9fa",
              border: "1px solid #ddd",
              borderRadius: "6px",
              cursor: "pointer",
              fontSize: "14px",
              color: "#666",
              transition: "all 0.2s"
            }}
          >
            Cancel
          </button>
          {!viewOnly && (
            <button
              onClick={handleSubmit}
              style={{
                padding: "8px 20px",
                backgroundColor: "#17a2b8",
                border: "none",
                borderRadius: "6px",
                cursor: "pointer",
                fontSize: "14px",
                color: "white",
                fontWeight: "500",
                transition: "all 0.2s"
              }}
            >
              {editComment ? "Update Comment" : "Add Comment"}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
////////////////////////////////////////////////////////////
// ICON BUTTON COMPONENT
////////////////////////////////////////////////////////////
function IconButton({ icon, onClick, title, disabled, active, loading }) {
  const [isHovered, setIsHovered] = useState(false);

  return (
    <button
      onClick={onClick}
      title={title}
      disabled={disabled || loading}
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
      style={{
        padding: "8px",
        backgroundColor: active ? "#17a2b8" : "#f8f9fa",
        border: active ? "none" : "1px solid #dee2e6",
        borderRadius: "6px",
        cursor: (disabled || loading) ? "not-allowed" : "pointer",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        transition: "all 0.2s ease",
        opacity: (disabled || loading) ? 0.5 : 1,
        width: "36px",
        height: "36px",
        color: active ? "white" : "#495057",
        position: "relative",
        ...(isHovered && !disabled && !loading && {
          transform: "translateY(-2px)",
          boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
          backgroundColor: active ? "#138496" : "#e9ecef"
        })
      }}
    >
      {loading ? (
        <div style={{
          width: "16px",
          height: "16px",
          border: "2px solid #f3f3f3",
          borderTop: "2px solid #0056b3",
          borderRadius: "50%",
          animation: "spin 0.8s linear infinite"
        }} />
      ) : (
        <SvgIcon icon={icon} size="medium" />
      )}
    </button>
  );
}


//new
function ZoomDropdown({
  scale,
  zoomMode,
  isOpen,
  onToggle,
  onActualSize,
  onPageFit,
  onPageWidth,
  onPercentageZoom,
  getCurrentZoomLabel,
  dropdownRef,
}) {
  const percentOptions = [0.5, 0.75, 1, 1.25, 1.5, 2, 3];

  const itemStyle = {
    padding: "10px 16px",
    fontSize: "14px",
    cursor: "pointer",
    background: "#fff",
    border: "none",
    textAlign: "left",
    width: "100%",
  };

  const selectedStyle = {
    ...itemStyle,
    fontWeight: 600,
    background: "#f3f3f3",
  };

  const isSelected = (type, value = null) => {
    if (type === "percentage") {
      return zoomMode === "percentage" && Number(scale) === Number(value);
    }
    return zoomMode === type;
  };

  return (
    <div
      ref={dropdownRef}
      style={{
        position: "relative",
        display: "inline-block",
      }}
    >
      <button
        onClick={onToggle}
        title="Zoom options"
        style={{
          minWidth: "95px",
          height: "36px",
          padding: "0 12px",
          border: "1px solid #d9d9d9",
          background: "#fff",
          borderRadius: "4px",
          cursor: "pointer",
          fontSize: "14px",
        }}
      >
        {getCurrentZoomLabel()}
      </button>

      {isOpen && (
        <div
          style={{
            position: "absolute",
            top: "40px",
            left: 0,
            width: "155px",
            background: "#fff",
            border: "1px solid #d9d9d9",
            boxShadow: "0 4px 10px rgba(0,0,0,0.12)",
            zIndex: 9999,
          }}
        >
          <button
            style={isSelected("actual-size") ? selectedStyle : itemStyle}
            onClick={onActualSize}
          >
            Actual size
          </button>

          <button
            style={isSelected("page-fit") ? selectedStyle : itemStyle}
            onClick={onPageFit}
          >
            Page fit
          </button>

          <button
            style={isSelected("page-width") ? selectedStyle : itemStyle}
            onClick={onPageWidth}
          >
            Page width
          </button>

          <div style={{ borderTop: "1px solid #d9d9d9", margin: "6px 0" }} />

          {percentOptions.map((value) => (
            <button
              key={value}
              style={isSelected("percentage", value) ? selectedStyle : itemStyle}
              onClick={() => onPercentageZoom(value)}
            >
              {Math.round(value * 100)}%
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
//new
////////////////////////////////////////////////////////////
// LAZY LOADING PDF PAGE COMPONENT (for performance)
////////////////////////////////////////////////////////////
// LAZY LOADING PDF PAGE COMPONENT (for performance)
////////////////////////////////////////////////////////////
// LAZY LOADING PDF PAGE COMPONENT (for performance)
////////////////////////////////////////////////////////////
const LazyPdfPage = React.memo(({
  pageNum,
  pdfDoc,
  scale,
  annotations,
  hoveredId,
  setHoveredId,
  onCommentClick,
  onEditComment,
  onViewComment,
  onViewOnlyComment,
  onTextLinesReady,
  onPageRenderComplete,
  isReadOnly
}) => {
  const [isVisible, setIsVisible] = useState(false);
  const containerRef = useRef(null);
  const [isRendered, setIsRendered] = useState(false);

  useEffect(() => {
    const observer = new IntersectionObserver(
      (entries) => {
        if (entries[0].isIntersecting) {
          setIsVisible(true);
          observer.disconnect();
        }
      },
      { rootMargin: "200px" }
    );

    if (containerRef.current) {
      observer.observe(containerRef.current);
    }

    return () => observer.disconnect();
  }, []);

  // Notify when page is rendered
  useEffect(() => {
    if (isVisible && !isRendered) {
      setIsRendered(true);
      if (onPageRenderComplete) {
        setTimeout(() => onPageRenderComplete(pageNum), 100);
      }
    }
  }, [isVisible, pageNum, onPageRenderComplete, isRendered]);

  return (
    <div
      ref={containerRef}
      data-page-num={pageNum}
      style={{
        width: "100%",
        marginBottom: "12px"
      }}
    >
      {isVisible ? (
        <PdfPage
          pdfDoc={pdfDoc}
          pageNum={pageNum}
          scale={scale}
          annotations={annotations}
          hoveredId={hoveredId}
          setHoveredId={setHoveredId}
          onCommentClick={onCommentClick}
          onEditComment={onEditComment}
          onViewOnlyComment={onViewOnlyComment}  // Add this line
          onTextLinesReady={onTextLinesReady}
          isReadOnly={isReadOnly}
        />
      ) : (
        <div
          style={{
            height: "500px",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            backgroundColor: "#f8f9fa",
            borderRadius: "8px",
            border: "1px solid #e0e0e0"
          }}
        >
          <div style={{ textAlign: "center" }}>
            <div style={{
              width: "30px",
              height: "30px",
              border: "2px solid #f3f3f3",
              borderTop: "2px solid #0056b3",
              borderRadius: "50%",
              animation: "spin 0.8s linear infinite",
              margin: "0 auto 10px"
            }} />
            <span style={{ fontSize: "13px", color: "#666" }}>Loading page {pageNum}...</span>
          </div>
        </div>
      )}
    </div>
  );
});
////////////////////////////////////////////////////////////
// MAIN PDF ANNOTATOR COMPONENT
////////////////////////////////////////////////////////////
const PdfAnnotator = forwardRef(({
  pdfBlob,
  fileName = "document.pdf",
  onAddComment,
  onUpdateComment,
  onAddHighlight,
  onResetAnnotations,
  currentUser = "User",
  isReadOnly = false,
  generalDetails = null,
  initialZoom = 1.2,
  initialAnnotations = [],
  isFullScreen = false,
  saveCurrentUserDraftAnnotations,
  clearCurrentUserDraftAnnotations
}, ref) => {
  //priti shelke - 06-04-2026
  // this is newly added


  const [pdfDoc, setPdfDoc] = useState(null);
  const [pdfFile, setPdfFile] = useState(null);
  const [numPages, setNumPages] = useState(0);

  const applyCanEdit = useCallback(
    (list) =>
      (Array.isArray(list) ? list : []).map((item) => {
        console.log("Processing annotation from DB:", {
          id: item.id,
          type: item.type,
          hasLineNumber: !!item.lineNumber,
          hasLineNumberRange: !!item.lineNumberRange,
          rawItem: item
        });

        // Try to get line number from multiple possible field names
        let lineNumber = null;
        let lineNumberRange = null;

        if (item.type === "comment") {
          // Check different possible field names from database
          lineNumber = item.lineNumber ||
            item.LineNumber ||
            item.line_no ||
            item.lineNo ||
            item.selectedLineNumber ||
            null;

          lineNumberRange = item.lineNumberRange ||
            item.LineNumberRange ||
            item.line_range ||
            item.lineRange ||
            null;

          // If still null, try to extract from selectedText
          if (!lineNumber && item.selectedText) {
            const lineMatch = item.selectedText.match(/[Ll]ine\s*[:]?\s*(\d+)/);
            if (lineMatch && lineMatch[1]) {
              lineNumber = lineMatch[1];
              console.log(`Extracted line number from selectedText: ${lineNumber}`);
            }
          }

          console.log(`Comment ${item.id} - Extracted lineNumber: ${lineNumber}, range: ${lineNumberRange}`);
        }

        return {
          ...item,
          canEdit:
            (item?.user || "").toLowerCase() ===
            (currentUser || "").toLowerCase(),
          lineNumber: lineNumber,
          lineNumberRange: lineNumberRange,
        };
      }),
    [currentUser]
  );
  const [annotations, setAnnotations] = useState(
    applyCanEdit(initialAnnotations)
  );

  useEffect(() => {
    console.log("=== LOADING INITIAL ANNOTATIONS FROM DATABASE ===");
    console.log("Raw initialAnnotations:", initialAnnotations);

    // Log each comment's data
    initialAnnotations.forEach((ann, idx) => {
      if (ann.type === 'comment') {
        console.log(`DB Comment ${idx + 1}:`, {
          id: ann.id,
          lineNumber: ann.lineNumber,
          lineNumberRange: ann.lineNumberRange,
          selectedText: ann.selectedText?.substring(0, 50)
        });
      }
    });

    const applied = applyCanEdit(initialAnnotations);

    // Log after processing
    applied.forEach((ann, idx) => {
      if (ann.type === 'comment') {
        console.log(`After processing Comment ${idx + 1} lineNumber:`, ann.lineNumber);
      }
    });

    setAnnotations(applied);
  }, [initialAnnotations, currentUser, applyCanEdit]);

  const [hoveredId, setHoveredId] = useState(null);
  // const [scale, setScale] = useState(1.2);
  const [scale, setScale] = useState(initialZoom);
  const [currentPage, setCurrentPage] = useState(1);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [notification, setNotification] = useState(null);
  const [isCommentDialogOpen, setIsCommentDialogOpen] = useState(false);
  const [pendingSelection, setPendingSelection] = useState(null);
  const [selectedTextForComment, setSelectedTextForComment] = useState("");
  const [currentPageNo, setCurrentPageNo] = useState(1);
  const [currentLineNumber, setCurrentLineNumber] = useState(null);
  const [editComment, setEditComment] = useState(null);
  const [history, setHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  const [pageTextLines, setPageTextLines] = useState({});
  const [isAddingHighlight, setIsAddingHighlight] = useState(false);
  const [isResetting, setIsResetting] = useState(false);
  const [showResetDialog, setShowResetDialog] = useState(false);
  const { accounts, instance } = useMsal();
  const account = useAccount(accounts[0] || {});
  const [isViewOnlyComment, setIsViewOnlyComment] = useState(false);
  const [isZoomDropdownOpen, setIsZoomDropdownOpen] = useState(false);
  const [zoomMode, setZoomMode] = useState("percentage");
  const [allPagesRendered, setAllPagesRendered] = useState(false);
  const [renderedPagesCount, setRenderedPagesCount] = useState(0);
  const [isTextLinesReady, setIsTextLinesReady] = useState(false);
  //priti - Add Comment and Annotation option on text selection
  const [selectionToolbar, setSelectionToolbar] = useState({
    visible: false,
    x: 0,
    y: 0,
    selectionData: null,
  });
  //priti - Add Comment and Annotation option on text selection


  //priti shelke - 08-04-2026
  const canEditComment = useCallback((comment) => {
    const commentUser =
      comment?.user ||
      comment?.createdBy ||
      comment?.approverEmail ||
      "";

    return (
      commentUser.toLowerCase() ===
      (currentUser || "").toLowerCase()
    );
  }, [currentUser]);
  //priti shelke - 08-04-2026


  const viewerContainerRef = useRef(null);
  const zoomDropdownRef = useRef(null);

  const pageRefs = useRef({});
  const suppressSelectionToolbarRef = useRef(false);
  const ZOOM_OPTIONS = [0.5, 0.75, 1, 1.25, 1.5, 2, 3];

  const clampZoom = useCallback((value) => {
    return Math.min(3, Math.max(0.5, Number(value)));
  }, []);



  const normalizeCoordinates = useCallback((rects, renderedPageWidth, renderedPageHeight) => {
    if (!rects || rects.length === 0) return [];
    if (renderedPageWidth <= 0 || renderedPageHeight <= 0) return [];

    return rects.map((rect) => ({
      x: (rect.x / renderedPageWidth) * 100,
      y: (rect.y / renderedPageHeight) * 100,
      width: (rect.width / renderedPageWidth) * 100,
      height: (rect.height / renderedPageHeight) * 100,
    }));
  }, []);

  /* PASTE THESE 3 FUNCTIONS HERE ↓↓↓ */
  const getCurrentApproverContext = useCallback(() => {
    const currentEmail = (currentUser || "").toLowerCase();

    const currentApprover = generalDetails?.noteApproversDTO?.find(
      (item) =>
        (item?.approverEmail || "").toLowerCase() === currentEmail
    );

    const currentReferrer = generalDetails?.noteReferrerDTO?.find(
      (item) =>
        (item?.referrerEmail || "").toLowerCase() === currentEmail &&
        Number(item?.referrerStatus) === 1
    );

    return {
      noteId: generalDetails?.noteId || generalDetails?.id || 0,

      approverId:
        currentApprover?.noteApproverId ||
        currentApprover?.approverId ||
        currentReferrer?.noteApproverId ||
        currentReferrer?.approverId ||
        currentReferrer?.referrerId ||
        0,
    };
  }, [generalDetails, currentUser]);

  const persistAnnotationToServer = useCallback(async (annotation) => {
    const { noteId, approverId } = getCurrentApproverContext();

    if (!noteId || !approverId) {
      throw new Error("Unable to resolve noteId/approverId for annotation save.");
    }

    let accessToken;
    try {
      accessToken = await getAccessToken({ ...loginRequest, account }, instance);
    } catch {
      throw new Error("Failed to get access token.");
    }

    const rects = Array.isArray(annotation?.normalizedRects)
      ? annotation.normalizedRects.filter(Boolean)
      : [];

    if (!rects.length) {
      throw new Error("No normalized rects found for annotation.");
    }

    const annotationUrl = `${API_BASE_URL.replace(/\/$/, "")}/${API_ENDPOINTS.AddAnnotationPDF.replace(/^\//, "")}`;

    const highlightGroupId =
      annotation?.highlightGroupId ||
      (typeof crypto !== "undefined" && crypto.randomUUID
        ? crypto.randomUUID()
        : `${Date.now()}-${Math.random().toString(16).slice(2)}`);

    const payload = rects.map((rect) => ({
      annotationId: 0,
      noteId: noteId,
      approverId: approverId,
      filePath: fileName || "test.pdf",
      pageNo: annotation.page || 1,

      xposition: Number(rect.x || 0),
      yposition: Number(rect.y || 0),
      width: Number(rect.width || 0),
      height: Number(rect.height || 0),
      endXPosition: Number((rect.x || 0) + (rect.width || 0)),
      endYPosition: Number((rect.y || 0) + (rect.height || 0)),

      // ✅ For highlights: fillColor = "yellow"
      // ✅ For comments: fillColor = "transparent"
      fillColor:
        (annotation.type || "").toLowerCase() === "highlight"
          ? "yellow"
          : "transparent",
      opacity: Number(annotation.opacity ?? 1),
      borderWidth: Number(annotation.borderWidth ?? 1),
      isDeleted: false,

      createdBy: annotation.user || currentUser,
      modifiedBy: annotation.user || currentUser,

      highlightGroupId: highlightGroupId,

      // Save selected text for both highlights and comments
      selectedText: annotation.selectedText || "",
    }));

    const response = await fetch(annotationUrl, {
      method: "POST",
      headers: {
        ...API_COMMON_HEADERS,
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Annotation save failed: ${errorText}`);
    }

    const result = await response.json();
    const savedRows = Array.isArray(result) ? result : [];

    return {
      annotationId: savedRows[0]?.annotationId || savedRows[0]?.AnnotationId || 0,
      highlightGroupId: savedRows[0]?.highlightGroupId || savedRows[0]?.HighlightGroupId || highlightGroupId,
      allSavedAnnotations: savedRows,
    };
  }, [account, instance, currentUser, fileName, getCurrentApproverContext]);



  const hydrateAnnotationWithServerIds = useCallback((annotation, serverResponse) => {
    const annotationId =
      serverResponse?.annotationId ||
      serverResponse?.AnnotationId ||
      serverResponse?.data?.annotationId ||
      serverResponse?.data?.AnnotationId ||
      annotation.annotationId ||
      0;

    // Get line number from server response if available
    const savedAnnotation = serverResponse?.allSavedAnnotations?.[0] || {};

    return {
      ...annotation,
      annotationId,
      serverResponse,
      // Preserve or restore line number
      lineNumber: annotation.lineNumber || savedAnnotation.lineNumber || null,
      lineNumberRange: annotation.lineNumberRange || savedAnnotation.lineNumberRange || null,
    };
  }, []);

  // this is newly added



  const handleTextLinesReady = useCallback((pageNum, lines) => {
    setPageTextLines(prev => {
      const newLines = { ...prev, [pageNum]: lines };
      if (pdfDoc && Object.keys(newLines).length === numPages) {
        setIsTextLinesReady(true);
        console.log("All pages text lines ready!");
      }
      return newLines;
    });
  }, [pdfDoc, numPages]);
  //priti shelke - 06-04-2026

  const ensureTextLinesForExport = useCallback(async () => {
    console.log("Ensuring text lines are available for export...");

    // If we already have text lines for all pages with comments
    const commentPages = [...new Set(annotations.filter(a => a.type === 'comment').map(a => a.page))];
    const missingPages = commentPages.filter(page => !pageTextLines[page]);

    if (missingPages.length === 0 && Object.keys(pageTextLines).length > 0) {
      console.log("All required text lines already available");
      return true;
    }

    console.log("Missing text lines for pages:", missingPages);
    console.log("Waiting for pages to render...");

    // Wait for pages to render
    return new Promise((resolve) => {
      let attempts = 0;
      const checkInterval = setInterval(() => {
        const nowMissing = commentPages.filter(page => !pageTextLines[page]);
        attempts++;
        console.log(`Attempt ${attempts}: Still missing pages:`, nowMissing);

        if (nowMissing.length === 0 && Object.keys(pageTextLines).length > 0) {
          clearInterval(checkInterval);
          console.log("All text lines are now available!");
          resolve(true);
        } else if (attempts > 50) { // 5 seconds timeout
          clearInterval(checkInterval);
          console.warn("Timeout waiting for text lines");
          resolve(false);
        }
      }, 100);
    });
  }, [annotations, pageTextLines]);

  const handlePageRenderComplete = useCallback((pageNum) => {
    setRenderedPagesCount(prev => {
      const newCount = Math.max(prev, pageNum);
      console.log(`Page ${pageNum} rendered, total rendered: ${newCount}/${numPages}`);
      return newCount;
    });
  }, [numPages]);

  useEffect(() => {
    if (numPages > 0 && renderedPagesCount >= numPages) {
      setAllPagesRendered(true);
      console.log("All pages rendered, text lines available for all pages");
    }
  }, [renderedPagesCount, numPages]);
  // Save to history when annotations change
  useEffect(() => {
    if (annotations.length > 0 && history[historyIndex] !== annotations) {
      const newHistory = history.slice(0, historyIndex + 1);
      newHistory.push(annotations);
      setHistory(newHistory);
      setHistoryIndex(newHistory.length - 1);
    }
  }, [annotations, history, historyIndex]);

  // Clear notification after 3 seconds
  useEffect(() => {
    if (notification) {
      const timer = setTimeout(() => setNotification(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [notification]);


  useEffect(() => {
    setScale(initialZoom);
    setZoomMode("percentage");
  }, [initialZoom]);


  useEffect(() => {
    const handleClickOutside = (event) => {
      if (
        zoomDropdownRef.current &&
        !zoomDropdownRef.current.contains(event.target)
      ) {
        setIsZoomDropdownOpen(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  // Load PDF from blob prop
  useEffect(() => {
    if (pdfBlob && pdfBlob instanceof Blob && pdfBlob.size > 0) {
      loadPDFFromBlob(pdfBlob);
    } else if (pdfBlob) {
      setError("Invalid PDF data provided");
      setLoading(false);
    } else {
      setError("No PDF data provided");
      setLoading(false);
    }
  }, [pdfBlob]);
  ///Digvijay  Add for Add Scroll Listener to Update Current Page

  //digvijay

  ///Digvijay  Add for Add Scroll Listener to Update Current Page Remove Scroller

  // Update current page based on scroll position
  useEffect(() => {
    const container = viewerContainerRef.current;
    if (!container || !pdfDoc) return;

    let ticking = false;

    const updateVisiblePage = () => {
      const pages = container.querySelectorAll(".pdf-page-container");
      if (!pages.length) return;

      const containerRect = container.getBoundingClientRect();
      let visiblePage = 1;
      let maxVisibleArea = 0;

      pages.forEach((page) => {
        const rect = page.getBoundingClientRect();
        const pageNum = Number(page.dataset.page);

        const visibleTop = Math.max(rect.top, containerRect.top);
        const visibleBottom = Math.min(rect.bottom, containerRect.bottom);
        const visibleHeight = Math.max(0, visibleBottom - visibleTop);

        if (visibleHeight > maxVisibleArea) {
          maxVisibleArea = visibleHeight;
          visiblePage = pageNum;
        }
      });

      setCurrentPage((prev) => (prev !== visiblePage ? visiblePage : prev));
      ticking = false;
    };

    const handleScroll = () => {
      if (!ticking) {
        window.requestAnimationFrame(updateVisiblePage);
        ticking = true;
      }
    };

    container.addEventListener("scroll", handleScroll);
    updateVisiblePage();

    return () => {
      container.removeEventListener("scroll", handleScroll);
    };
  }, [pdfDoc]);
  //////////////////////////////////////////////////
  // Update current page based on scroll position
  useEffect(() => {
    const container = viewerContainerRef.current;
    if (!container || !pdfDoc) return;

    let ticking = false;

    const updateVisiblePage = () => {
      const pages = container.querySelectorAll(".pdf-page-container");
      if (!pages.length) return;

      const containerRect = container.getBoundingClientRect();
      let visiblePage = 1;
      let maxVisibleArea = 0;

      pages.forEach((page) => {
        const rect = page.getBoundingClientRect();
        const pageNum = Number(page.dataset.page);

        const visibleTop = Math.max(rect.top, containerRect.top);
        const visibleBottom = Math.min(rect.bottom, containerRect.bottom);
        const visibleHeight = Math.max(0, visibleBottom - visibleTop);

        if (visibleHeight > maxVisibleArea) {
          maxVisibleArea = visibleHeight;
          visiblePage = pageNum;
        }
      });

      setCurrentPage((prev) => (prev !== visiblePage ? visiblePage : prev));
      ticking = false;
    };

    const handleScroll = () => {
      if (!ticking) {
        window.requestAnimationFrame(updateVisiblePage);
        ticking = true;
      }
    };

    container.addEventListener("scroll", handleScroll);
    updateVisiblePage();

    return () => {
      container.removeEventListener("scroll", handleScroll);
    };
  }, [pdfDoc]);

  // ADD THE NEW useEffect HERE
  useEffect(() => {
    if (!pdfDoc || !viewerContainerRef.current) return;

    const timer = setTimeout(() => {
      const container = viewerContainerRef.current;
      if (!container) return;

      const pageContainers = container.querySelectorAll('[data-page-num], .pdf-page-container');
      pageContainers.forEach((el) => {
        const pageNumAttr = el.getAttribute('data-page-num') || el.getAttribute('data-page');
        if (pageNumAttr) {
          const pageNum = parseInt(pageNumAttr);
          if (!isNaN(pageNum)) {
            pageRefs.current[pageNum] = el;
          }
        }
      });
    }, 500);

    return () => clearTimeout(timer);
  }, [pdfDoc, numPages, annotations]);


  /////////////////////////////////////////////////////////////////////////////////////////////////////

  const loadPDFFromBlob = async (blob) => {
    try {
      setLoading(true);
      setError(null);

      if (blob.size === 0) {
        throw new Error("PDF file is empty");
      }

      // Convert blob to array buffer
      const arrayBuffer = await blob.arrayBuffer();

      // Validate PDF header
      const uint8Array = new Uint8Array(arrayBuffer);
      const header = String.fromCharCode.apply(null, uint8Array.slice(0, 5));

      if (!header.includes('%PDF')) {
        throw new Error("File does not appear to be a valid PDF");
      }

      // Create file object
      const file = new File([arrayBuffer], fileName, { type: 'application/pdf' });
      setPdfFile(file);

      // Load PDF document with timeout
      const loadingTask = pdfjsLib.getDocument({
        data: arrayBuffer,
        useSystemFonts: true,
        disableFontFace: false,
        verbosity: 0 // Reduce console noise
      });

      // Add timeout for loading
      const loadPromise = loadingTask.promise;
      const timeoutPromise = new Promise((_, reject) => {
        setTimeout(() => reject(new Error("PDF loading timeout after 30 seconds")), 30000);
      });

      const doc = await Promise.race([loadPromise, timeoutPromise]);
      setPdfDoc(doc);
      setNumPages(doc.numPages);
      setCurrentPage(1);

      setLoading(false);
    } catch (error) {
      console.error('Error loading PDF:', error);
      let errorMessage = "Failed to load PDF";

      if (error.message.includes("timeout")) {
        errorMessage = "PDF loading timed out. The file might be too large or corrupted.";
      } else if (error.message.includes("Invalid") || error.message.includes("not appear")) {
        errorMessage = "Invalid PDF file format. Please ensure it's a valid PDF.";
      } else if (error.message.includes("empty")) {
        errorMessage = "PDF file is empty.";
      } else {
        errorMessage = `Failed to load PDF: ${error.message}`;
      }

      setError(errorMessage);
      setLoading(false);
    }
  };

  // Save annotations to localStorage
  const saveAnnotationsToStorage = useCallback((newAnnotations) => {
    try {
      const normalized = Array.isArray(newAnnotations) ? newAnnotations : [];

      const onlyComments = normalized.filter((item) => item?.type !== "highlight");

      if (typeof saveCurrentUserDraftAnnotations === "function") {
        saveCurrentUserDraftAnnotations(onlyComments);
      }
    } catch (err) {
      console.error("Failed to save annotations:", err);
    }
  }, [saveCurrentUserDraftAnnotations]);

  // Get accurate line number from Y position
  // Get accurate line number from Y position - IMPROVED VERSION
  const getLineNumberFromY = useCallback((page, yPosition) => {
    const lines = pageTextLines[page];
    if (!lines || lines.length === 0) return 1;

    // Find the line that contains this Y position
    // We need to check if the Y position falls within the line's vertical span
    let foundLine = null;
    let bestMatch = null;
    let minDistance = Infinity;

    for (const line of lines) {
      // Get the vertical bounds of this line
      if (line.spans && line.spans.length > 0) {
        // Calculate the line's vertical range
        let minY = Infinity;
        let maxY = -Infinity;

        line.spans.forEach(span => {
          if (span.getBoundingClientRect) {
            const rect = span.getBoundingClientRect();
            minY = Math.min(minY, rect.top);
            maxY = Math.max(maxY, rect.bottom);
          }
        });

        // Check if yPosition falls within this line's vertical bounds
        const textLayerContainer = document.querySelector('.pdf-text-layer');
        if (textLayerContainer) {
          const containerRect = textLayerContainer.getBoundingClientRect();
          const relativeY = yPosition + containerRect.top; // Convert to absolute if needed

          if (relativeY >= minY && relativeY <= maxY) {
            foundLine = line;
            break;
          }

          // Also track the closest line
          const distance = Math.min(Math.abs(relativeY - minY), Math.abs(relativeY - maxY));
          if (distance < minDistance) {
            minDistance = distance;
            bestMatch = line;
          }
        }
      } else {
        // Fallback to Y position matching
        const distance = Math.abs(line.yPosition - yPosition);
        if (distance < minDistance) {
          minDistance = distance;
          bestMatch = line;
        }
      }
    }

    const selectedLine = foundLine || bestMatch;

    if (selectedLine) {
      return selectedLine.lineNumber;
    }

    return 1;
  }, [pageTextLines]);

  const getSelectionData = useCallback(() => {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || sel.toString().trim() === "") {
      setNotification({ type: 'warning', message: 'Please select some text first.' });
      return null;
    }

    const range = sel.getRangeAt(0);
    const selectedText = sel.toString().trim();

    // Find the text layer container
    let textLayerContainer = null;
    let pageContainer = null;

    const startNode = range.startContainer;
    const element = startNode.nodeType === Node.TEXT_NODE ? startNode.parentElement : startNode;

    textLayerContainer = element?.closest?.('.pdf-text-layer');
    pageContainer = element?.closest?.('.pdf-page-container');

    if (!pageContainer || !textLayerContainer) {
      setNotification({ type: 'warning', message: 'Please select text within the PDF content.' });
      return null;
    }

    const page = parseInt(pageContainer.dataset.page, 10);
    if (!page || isNaN(page)) {
      setNotification({ type: 'warning', message: 'Invalid page number' });
      return null;
    }

    // Get text layer dimensions
    const textLayerRect = textLayerContainer.getBoundingClientRect();

    // Get all client rects - these are already in correct coordinates
    const clientRects = Array.from(range.getClientRects()).filter(
      (rect) => rect.width > 0 && rect.height > 0
    );

    if (clientRects.length === 0) {
      setNotification({ type: 'warning', message: 'Could not get selection coordinates' });
      return null;
    }

    // Calculate rects relative to text layer
    const rects = clientRects.map((r) => ({
      x: Math.max(0, r.left - textLayerRect.left),
      y: Math.max(0, r.top - textLayerRect.top),
      width: r.width,
      height: r.height,
    }));

    const firstRect = clientRects[0];
    const lastRect = clientRects[clientRects.length - 1];

    const xposition = Math.round(firstRect.left - textLayerRect.left);
    const yposition = Math.round(firstRect.top - textLayerRect.top);
    const width = Math.round(lastRect.right - firstRect.left);
    const height = Math.round(lastRect.bottom - firstRect.top);
    const endXPosition = Math.round(lastRect.right - textLayerRect.left);
    const endYPosition = Math.round(lastRect.bottom - textLayerRect.top);

    // Get line numbers from stored text lines
    const pageLines = pageTextLines[page];
    let lineNumber = 1;
    let lastLineNumber = 1;

    if (pageLines && pageLines.length > 0) {
      const firstLineY = yposition;
      const lastLineY = Math.round(lastRect.top - textLayerRect.top);

      for (const line of pageLines) {
        if (line.yMin !== undefined && line.yMax !== undefined) {
          if (firstLineY >= line.yMin && firstLineY <= line.yMax) {
            lineNumber = line.lineNumber;
          }
          if (lastLineY >= line.yMin && lastLineY <= line.yMax) {
            lastLineNumber = line.lineNumber;
          }
        }
      }
    }

    const lineNumberRange = lineNumber === lastLineNumber
      ? String(lineNumber)
      : `${lineNumber}-${lastLineNumber}`;

    return {
      page,
      rects,
      selectedText,
      lineNumber,
      lineNumberRange,
      xposition,
      yposition,
      endXPosition,
      endYPosition,
      width,
      height,
      multiLine: lineNumber !== lastLineNumber,
      textLayerWidth: textLayerRect.width,
      textLayerHeight: textLayerRect.height
    };
  }, [pageTextLines, setNotification]);

  const showSelectionToolbar = useCallback(() => {
    if (isReadOnly || isCommentDialogOpen || suppressSelectionToolbarRef.current) {
      setSelectionToolbar({
        visible: false,
        x: 0,
        y: 0,
        selectionData: null,
      });
      return;
    }

    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || sel.toString().trim() === "") {
      setSelectionToolbar((prev) => ({
        ...prev,
        visible: false,
        selectionData: null,
      }));
      return;
    }

    // Small delay to ensure selection is complete
    setTimeout(() => {
      const currentSel = window.getSelection();
      if (!currentSel || currentSel.isCollapsed || currentSel.toString().trim() === "") {
        return;
      }

      const data = getSelectionData();
      if (!data) {
        setSelectionToolbar((prev) => ({
          ...prev,
          visible: false,
          selectionData: null,
        }));
        return;
      }

      const range = currentSel.getRangeAt(0);
      const rect = range.getBoundingClientRect();

      // Ensure we have valid coordinates
      if (rect.width === 0 || rect.height === 0) {
        return;
      }

      setSelectionToolbar({
        visible: true,
        x: rect.left + rect.width / 2,
        y: rect.top - 8,
        selectionData: {
          ...data,
          selectedText: data.selectedText || currentSel.toString().trim()
        },
      });
    }, 20);
  }, [getSelectionData, isReadOnly, isCommentDialogOpen]);

  // Add after your other useEffects (around line 500-600)

  // Force annotation recalculation when scale changes
  useEffect(() => {
    if (pdfDoc && annotations.length > 0) {
      // Trigger a re-render of all pages to recalculate annotation positions
      setAnnotations(prev => [...prev]);
    }
  }, [scale, pdfDoc]);


  //priti - Add Comment and Annotation option on text selection
  // Find the useEffect that handles selection (around line 700)
  useEffect(() => {
    const handleMouseUp = () => {
      // Don't show toolbar in read-only mode
      if (!isReadOnly) {
        setTimeout(() => {
          showSelectionToolbar();
        }, 80);
      }
    };

    const handleScrollOrResize = () => {
      setSelectionToolbar((prev) => ({
        ...prev,
        visible: false,
      }));
    };

    const handleMouseDownOutside = (e) => {
      const toolbar = document.getElementById("pdf-selection-toolbar");
      if (toolbar && !toolbar.contains(e.target)) {
        setTimeout(() => {
          const sel = window.getSelection();
          if (!sel || sel.isCollapsed) {
            setSelectionToolbar((prev) => ({
              ...prev,
              visible: false,
              selectionData: null,
            }));
          }
        }, 50);
      }
    };

    document.addEventListener("mouseup", handleMouseUp);
    document.addEventListener("mousedown", handleMouseDownOutside);
    window.addEventListener("scroll", handleScrollOrResize, true);
    window.addEventListener("resize", handleScrollOrResize);

    return () => {
      document.removeEventListener("mouseup", handleMouseUp);
      document.removeEventListener("mousedown", handleMouseDownOutside);
      window.removeEventListener("scroll", handleScrollOrResize, true);
      window.removeEventListener("resize", handleScrollOrResize);
    };
  }, [showSelectionToolbar, isReadOnly]); // Add isReadOnly to dependencies


  //priti - Add Comment and Annotation option on text selection
  //priti shelke - 06-04-2026
  useEffect(() => {
    if (isCommentDialogOpen) {
      setSelectionToolbar({
        visible: false,
        x: 0,
        y: 0,
        selectionData: null,
      });
    }
  }, [isCommentDialogOpen]);
  //priti shelke - 06-04-2026

  // Function to update a single annotation as deleted
  const updateAnnotationAsDeleted = useCallback(async (annotation) => {
    try {
      const loggedInUser = (account?.username || currentUser || "").toLowerCase();
      const annotationOwner = (annotation?.user || "").toLowerCase();

      if (!annotationOwner || annotationOwner !== loggedInUser) {
        setNotification({
          type: "warning",
          message: "You can delete only your own annotations/comments."
        });
        return {
          success: false,
          annotationId:
            annotation?.serverResponse?.annotationId ||
            annotation?.annotationId ||
            annotation?.id,
          error: "Unauthorized delete attempt"
        };
      }

      let accessToken;
      try {
        accessToken = await getAccessToken({ ...loginRequest, account }, instance);
      } catch (tokenError) {
        console.error("Failed to get access token:", tokenError);
        throw new Error("Authentication failed. Please log in again.");
      }

      const currentUserEmail = account?.username || currentUser;

      const currentApprover = generalDetails?.noteApproversDTO?.find(
        (approver) =>
          (approver?.approverEmail || "").toLowerCase() ===
          (currentUserEmail || "").toLowerCase()
      );

      if (!currentApprover) {
        throw new Error("Current user not found in approvers list");
      }

      // ✅ If grouped highlight exists, delete all rows of same group
      const allSavedRows = Array.isArray(annotation?.serverResponse?.allSavedAnnotations)
        ? annotation.serverResponse.allSavedAnnotations
        : [];

      let deletePayload = [];

      if (allSavedRows.length > 0) {
        deletePayload = allSavedRows.map((row) => ({
          annotationId: row.annotationId || row.AnnotationId || 0,
          noteId: row.noteId || row.NoteId || currentApprover.noteId || 0,
          approverId: row.approverId || row.ApproverId || currentApprover.noteApproverId || 0,
          filePath: row.filePath || row.FilePath || fileName,
          pageNo: row.pageNo || row.PageNo || annotation?.page || 1,
          xposition: row.xposition || row.Xposition || 0,
          yposition: row.yposition || row.Yposition || 0,
          width: row.width || row.Width || 0,
          height: row.height || row.Height || 0,
          endXPosition: row.endXPosition || row.EndXPosition || 0,
          endYPosition: row.endYPosition || row.EndYPosition || 0,
          fillColor:
            row.fillColor ||
            row.FillColor ||
            (annotation?.type === "highlight" ? "yellow" : "transparent"),
          opacity: row.opacity || row.Opacity || 1,
          borderWidth: row.borderWidth || row.BorderWidth || 1,
          isDeleted: true,
          createdBy: row.createdBy || row.CreatedBy || annotation?.user || currentUser,
          createdDate: row.createdDate || row.CreatedDate || annotation?.timestamp || new Date().toISOString(),
          highlightGroupId:
            row.highlightGroupId || row.HighlightGroupId || annotation?.highlightGroupId || null,
          modifiedBy: currentUser,
          modifiedDate: new Date().toISOString()
        }));
      } else {
        // ✅ fallback for old single-row annotations
        deletePayload = [
          {
            annotationId:
              annotation?.serverResponse?.annotationId ||
              annotation?.annotationId ||
              annotation?.id,
            noteId: currentApprover.noteId || 0,
            approverId: currentApprover.noteApproverId || 0,
            filePath: fileName,
            pageNo: annotation?.page || 1,
            xposition: annotation?.xposition || 0,
            yposition: annotation?.yposition || 0,
            width: annotation?.width || 0,
            height: annotation?.height || 0,
            endXPosition: annotation?.endXPosition || 0,
            endYPosition: annotation?.endYPosition || 0,
            fillColor: annotation?.type === "highlight" ? "yellow" : "transparent",
            opacity: 1,
            borderWidth: 1,
            isDeleted: true,
            createdBy: annotation?.user || currentUser,
            createdDate: annotation?.timestamp || new Date().toISOString(),
            highlightGroupId: annotation?.highlightGroupId || null,
            modifiedBy: currentUser,
            modifiedDate: new Date().toISOString()
          }
        ];
      }

      console.log("Updating annotation(s) to deleted:", deletePayload);

      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.updateAnnotationPDF}`,
        {
          method: "POST",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          // ✅ backend expects List<AnnotationDetails>
          body: JSON.stringify(deletePayload)
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`API call failed with status ${response.status}: ${errorText}`);
      }

      const apiResponse = await response.json();

      return {
        success: true,
        annotationId:
          annotation?.serverResponse?.annotationId ||
          annotation?.annotationId ||
          annotation?.id,
        response: apiResponse
      };
    } catch (error) {
      console.error(`Failed to delete annotation ${annotation?.id}:`, error);
      return {
        success: false,
        annotationId:
          annotation?.serverResponse?.annotationId ||
          annotation?.annotationId ||
          annotation?.id,
        error: error.message
      };
    }
  }, [account, instance, currentUser, fileName, generalDetails]);
  // ============================================================
  // FIXED handleAddHighlight
  // ============================================================
  // In handleAddHighlight function (around line 850), ensure normalizedRects are properly stored:

  // ============================================================
  // FIXED handleAddHighlight - NOW SAVES TO BACKEND
  // ============================================================
 const handleAddHighlight = useCallback(async (selectionData = null) => {
  if (isReadOnly) {
    setNotification({ type: "warning", message: "Read-only mode - cannot add highlights" });
    return;
  }

  if (isAddingHighlight) return;

  let data = selectionData;

  if (!data) {
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed && sel.toString().trim() !== "") {
      data = getSelectionData();
    }
  }

  if (!data) {
    if (selectionToolbar.selectionData) {
      data = selectionToolbar.selectionData;
    } else {
      setNotification({ type: "warning", message: "Please select some text first." });
      return;
    }
  }

  try {
    setIsAddingHighlight(true);

    const normalizedRects = normalizeCoordinates(
      data.rects,
      data.textLayerWidth,
      data.textLayerHeight
    );

    if (!normalizedRects.length) {
      setNotification({ type: "warning", message: "Unable to calculate highlight area." });
      return;
    }

    const firstRect = normalizedRects[0];

    const newHighlight = {
      id: crypto.randomUUID(),
      type: "highlight",
      user: currentUser,
      page: data.page,
      timestamp: new Date().toISOString(),
      selectedText: data.selectedText || "",
      lineNumber: data.lineNumber || null,
      lineNumberRange: data.lineNumberRange || null,
      rects: data.rects || [],
      normalizedRects,
      xposition: firstRect.x,
      yposition: firstRect.y,
      width: firstRect.width,
      height: firstRect.height,
      endXPosition: firstRect.x + firstRect.width,
      endYPosition: firstRect.y + firstRect.height,
      fillColor: "yellow",
      opacity: 1,
      borderWidth: 1,
      canEdit: true,

      // temporary only, DB save will happen on Approve/Noted flow
      annotationId: null,
      serverResponse: null,
      isTemporary: true,
    };

    setAnnotations((prev) => {
      return [...prev, newHighlight];
    });

    if (typeof onAddHighlight === "function") {
      await onAddHighlight(newHighlight);
    }

    setNotification({
      type: "success",
      message: `Highlight added on line ${data.lineNumberRange || data.lineNumber || "N/A"}`,
    });

    setSelectionToolbar({
      visible: false,
      x: 0,
      y: 0,
      selectionData: null,
    });

    const sel = window.getSelection();
    if (sel) sel.removeAllRanges();

  } catch (error) {
    console.error("Error adding highlight:", error);
    setNotification({
      type: "error",
      message: error?.message || "Failed to add highlight.",
    });
  } finally {
    setIsAddingHighlight(false);
  }
}, [
  isReadOnly,
  isAddingHighlight,
  getSelectionData,
  selectionToolbar,
  normalizeCoordinates,
  currentUser,
  onAddHighlight,
]);
  // ============================================================
  // FIXED handleAddComment
  // ============================================================
  const handleAddComment = useCallback((selectionData = null) => {
    if (isReadOnly) {
      setNotification({ type: "warning", message: "Read-only mode - cannot add comments" });
      return;
    }

    let data = selectionData;

    if (!data) {
      const sel = window.getSelection();
      if (sel && !sel.isCollapsed && sel.toString().trim() !== "") {
        data = getSelectionData();
      }
    }

    if (!data) {
      if (selectionToolbar.selectionData) {
        data = selectionToolbar.selectionData;
      } else {
        setNotification({ type: "warning", message: "Please select some text first." });
        return;
      }
    }

    // Ensure line number is calculated
    const lineNumber = data.lineNumber || 1;
    const lineNumberRange = data.lineNumberRange || String(lineNumber);

    suppressSelectionToolbarRef.current = true;
    setSelectionToolbar({ visible: false, x: 0, y: 0, selectionData: null });
    setPendingSelection({
      ...data,
      lineNumber,
      lineNumberRange,
    });
    setSelectedTextForComment(data.selectedText || "");
    setCurrentPageNo(data.page);
    setCurrentLineNumber(lineNumberRange);
    setEditComment(null);
    setIsCommentDialogOpen(true);

    setTimeout(() => {
      const sel = window.getSelection();
      if (sel) sel.removeAllRanges();
      suppressSelectionToolbarRef.current = false;
    }, 100);
  }, [getSelectionData, isReadOnly, selectionToolbar.selectionData]);
  const handleEditComment = useCallback((comment) => {
    setIsViewOnlyComment(false);
    setEditComment(comment);
    setSelectedTextForComment(comment.selectedText || "");
    setCurrentPageNo(comment.page || 1);
    setCurrentLineNumber(comment.lineNumber || comment.lineNumberRange || null);
    setIsCommentDialogOpen(true);
  }, []);


  const handleViewComment = useCallback((comment) => {
    setIsViewOnlyComment(true);
    setEditComment(comment);
    setSelectedTextForComment(comment.selectedText || "");
    setCurrentPageNo(comment.page || 1);
    setCurrentLineNumber(comment.lineNumber || comment.lineNumberRange || null);
    setIsCommentDialogOpen(true);
  }, []);
  const handleCommentSubmit = useCallback(async (commentText) => {
    if (isReadOnly) {
      setNotification({ type: "warning", message: "Read-only mode - cannot add comments" });
      return;
    }

    try {
      if (editComment) {
        // For edit mode - update both frontend AND backend
        const updatedComment = {
          ...editComment,
          text: commentText,
          modifiedBy: currentUser,
          modifiedDate: new Date().toISOString(),
          canEdit: true,
          // Preserve existing line numbers
          lineNumber: editComment.lineNumber,
          lineNumberRange: editComment.lineNumberRange,
        };

        // Update backend if it has serverResponse
        if (editComment.serverResponse?.annotationId || editComment.annotationId) {
          try {
            const { noteId, approverId } = getCurrentApproverContext();
            let accessToken = await getAccessToken({ ...loginRequest, account }, instance);

            // Prepare update payload for backend
            const updatePayload = [{
              annotationId: editComment.serverResponse?.annotationId || editComment.annotationId,
              noteId: noteId,
              approverId: approverId,
              filePath: fileName,
              pageNo: editComment.page,
              xposition: editComment.xposition || 0,
              yposition: editComment.yposition || 0,
              width: editComment.width || 0,
              height: editComment.height || 0,
              endXPosition: editComment.endXPosition || 0,
              endYPosition: editComment.endYPosition || 0,
              fillColor: "transparent",
              opacity: 1,
              borderWidth: 1,
              isDeleted: false,
              createdBy: editComment.user,
              createdDate: editComment.timestamp,
              modifiedBy: currentUser,
              modifiedDate: new Date().toISOString(),
              commentText: commentText, // Add comment text to payload
              highlightGroupId: editComment.highlightGroupId || null,
              lineNumber: editComment.lineNumber || null,
              lineNumberRange: editComment.lineNumberRange || null,
            }];

            const updateUrl = `${API_BASE_URL.replace(/\/$/, "")}/${API_ENDPOINTS.updateAnnotationPDF.replace(/^\//, "")}`;

            const response = await fetch(updateUrl, {
              method: "POST",
              headers: {
                ...API_COMMON_HEADERS,
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json",
              },
              body: JSON.stringify(updatePayload),
            });

            if (response.ok) {
              console.log("Backend comment updated successfully");
            } else {
              console.warn("Backend update failed, but keeping local update");
            }
          } catch (backendError) {
            console.error("Backend update error:", backendError);
            // Continue with local update even if backend fails
          }
        }

        // Update frontend state
        setAnnotations((prev) => {
          const updated = prev.map((item) =>
            item.id === editComment.id ? updatedComment : item
          );
          saveAnnotationsToStorage(updated);
          return updated;
        });

        if (typeof onUpdateComment === "function") {
          await onUpdateComment(updatedComment);
        }

        setNotification({
          type: "success",
          message: `Comment updated on line ${updatedComment.lineNumber || "N/A"}`
        });
      } else {
        // ADD NEW COMMENT - THIS PART WAS MISSING
        const data = pendingSelection || selectionToolbar.selectionData;

        if (!data) {
          setNotification({ type: "warning", message: "No selection data found." });
          return;
        }

        const normalizedRects = normalizeCoordinates(
          data.rects,
          data.textLayerWidth,
          data.textLayerHeight
        );

        if (!normalizedRects.length) {
          setNotification({ type: "warning", message: "Unable to calculate comment area." });
          return;
        }

        const firstRect = normalizedRects[0];

        const newComment = {
          id: crypto.randomUUID(),
          type: "comment",
          user: currentUser,
          page: data.page,
          text: commentText,
          timestamp: new Date().toISOString(),
          selectedText: data.selectedText || "",
          lineNumber: data.lineNumber || 1,
          lineNumberRange: data.lineNumberRange || String(data.lineNumber || 1),
          rects: data.rects || [],
          normalizedRects: normalizedRects,
          yposition: firstRect.y,
          xposition: firstRect.x,
          width: firstRect.width,
          height: firstRect.height,
          endXPosition: firstRect.x + firstRect.width,
          endYPosition: firstRect.y + firstRect.height,
          fillColor: "transparent",
          opacity: 1,
          borderWidth: 1,
          canEdit: true,
        };

        const serverResponse = await persistAnnotationToServer(newComment);
        const hydratedComment = hydrateAnnotationWithServerIds(newComment, serverResponse);

        setAnnotations((prev) => {
          const updated = [...prev, hydratedComment];
          saveAnnotationsToStorage(updated);
          return updated;
        });

        if (typeof onAddComment === "function") {
          await onAddComment(hydratedComment);
        }

        setNotification({
          type: "success",
          message: `Comment added on line ${data.lineNumberRange || data.lineNumber || "N/A"}`
        });
      }

      // Close dialog and reset states
      setIsCommentDialogOpen(false);
      setEditComment(null);
      setPendingSelection(null);
      setSelectedTextForComment("");
      setCurrentLineNumber(null);
      setIsViewOnlyComment(false);

      setSelectionToolbar({
        visible: false,
        x: 0,
        y: 0,
        selectionData: null,
      });

      const sel = window.getSelection();
      if (sel) sel.removeAllRanges();
    } catch (error) {
      console.error("Error saving comment:", error);
      setNotification({
        type: "error",
        message: error?.message || "Failed to save comment."
      });
    }
  }, [
    isReadOnly,
    editComment,
    pendingSelection,
    selectionToolbar,
    normalizeCoordinates,
    currentUser,
    onAddComment,
    onUpdateComment,
    saveAnnotationsToStorage,
    persistAnnotationToServer,
    hydrateAnnotationWithServerIds,
    getCurrentApproverContext,
    account,
    instance,
    fileName
  ]);

  // Open reset confirmation dialog
  const handleReset = useCallback(() => {
    const loggedInUser = (account?.username || currentUser || "").toLowerCase();

    const myAnnotations = annotations.filter(
      (a) => (a?.user || "").toLowerCase() === loggedInUser
    );

    if (myAnnotations.length === 0) {
      setNotification({
        type: "warning",
        message: "You do not have any annotations/comments to delete."
      });
      return;
    }

    setShowResetDialog(true);
  }, [annotations, account, currentUser]);



  //this is newly added
  // Handle reset confirmation
  const handleResetConfirm = useCallback(async () => {
    setShowResetDialog(false);
    setIsResetting(true);

    try {
      const loggedInUser = (account?.username || currentUser || "").toLowerCase();

      const myAnnotations = annotations.filter(
        (a) => (a?.user || "").toLowerCase() === loggedInUser
      );

      const otherUsersAnnotations = annotations.filter(
        (a) => (a?.user || "").toLowerCase() !== loggedInUser
      );

      if (myAnnotations.length === 0) {
        setNotification({
          type: "warning",
          message: "You do not have any annotations/comments to delete."
        });
        return;
      }

      // Only my backend-saved highlights/comments go to API delete
      const myBackendAnnotations = myAnnotations.filter(
        (a) => a.serverResponse?.annotationId || a.annotationId
      );

      // My local-only items
      const myLocalOnlyAnnotations = myAnnotations.filter(
        (a) => !(a.serverResponse?.annotationId || a.annotationId)
      );

      let failedUpdates = [];

      if (myBackendAnnotations.length > 0) {
        const results = await Promise.all(
          myBackendAnnotations.map((annotation) => updateAnnotationAsDeleted(annotation))
        );

        failedUpdates = results.filter((result) => !result.success);
      }

      // IDs of my backend annotations that failed
      const failedIds = failedUpdates.map(
        (r) => r.annotationId
      );

      // Keep:
      // 1. all other users' annotations
      // 2. only my backend annotations that failed to delete
      const myFailedAnnotations = myBackendAnnotations.filter((a) => {
        const id = a.serverResponse?.annotationId || a.annotationId || a.id;
        return failedIds.includes(id);
      });

      const finalRemainingAnnotations = [
        ...otherUsersAnnotations,
        ...myFailedAnnotations
      ];

      setAnnotations(finalRemainingAnnotations);
      saveAnnotationsToStorage(finalRemainingAnnotations);
      setHistory([]);
      setHistoryIndex(-1);
      setPageTextLines({});

      if (onResetAnnotations) {
        onResetAnnotations(finalRemainingAnnotations);
      }

      if (failedUpdates.length === 0) {
        setNotification({
          type: "success",
          message: `Successfully deleted your ${myAnnotations.length} annotation${myAnnotations.length !== 1 ? "s" : ""}.`
        });
      } else {
        setNotification({
          type: "warning",
          message: `Some of your annotations could not be deleted. Other users' annotations were not affected.`
        });
      }
    } catch (error) {
      console.error("Failed to reset current user's annotations:", error);

      setNotification({
        type: "error",
        message: `Failed to delete your annotations: ${error.message}`
      });
    } finally {
      setIsResetting(false);
    }
  }, [
    annotations,
    account,
    currentUser,
    updateAnnotationAsDeleted,
    saveAnnotationsToStorage,
    onResetAnnotations
  ]);
  //this is newly added
  const enrichCommentsWithLineNumbers = useCallback(() => {
    console.log("=== Enriching comments with dynamic line numbers ===");
    console.log("Available pageTextLines pages:", Object.keys(pageTextLines));

    const enrichedAnnotations = annotations.map((annotation) => {
      if (annotation.type !== "comment") return annotation;

      // Skip if already has line number
      if (annotation.lineNumber && annotation.lineNumber !== 'N/A' && annotation.lineNumber !== null) {
        return annotation;
      }

      const pageLines = pageTextLines[annotation.page];
      if (!pageLines || pageLines.length === 0) {
        console.log(`No text lines for page ${annotation.page}, comment ${annotation.id.substring(0, 8)}`);
        return annotation;
      }

      // Get Y position from rects
      let yPosition = null;

      if (annotation.normalizedRects && annotation.normalizedRects.length > 0) {
        yPosition = annotation.normalizedRects[0].y;
        console.log(`Comment ${annotation.id.substring(0, 8)} - normalizedRects Y: ${yPosition}`);
      }
      else if (annotation.rects && annotation.rects.length > 0) {
        yPosition = annotation.rects[0].y;
        console.log(`Comment ${annotation.id.substring(0, 8)} - rects Y: ${yPosition}`);
      }
      else if (annotation.yposition !== undefined) {
        yPosition = annotation.yposition;
        console.log(`Comment ${annotation.id.substring(0, 8)} - yposition: ${yPosition}`);
      }

      if (yPosition === null) {
        console.log(`No Y position for comment ${annotation.id.substring(0, 8)}`);
        return annotation;
      }

      // Find line number from Y position
      let foundLineNumber = null;

      for (const line of pageLines) {
        if (line.yMin !== undefined && line.yMax !== undefined) {
          if (yPosition >= line.yMin && yPosition <= line.yMax) {
            foundLineNumber = line.lineNumber;
            break;
          }
        }
        else if (line.yPosition !== undefined) {
          if (Math.abs(yPosition - line.yPosition) < 25) {
            foundLineNumber = line.lineNumber;
            break;
          }
        }
      }

      if (foundLineNumber === null) {
        console.log(`No line found for Y=${yPosition} on page ${annotation.page}`);
        return annotation;
      }

      console.log(`Comment ${annotation.id.substring(0, 8)}: Y=${yPosition} → Line=${foundLineNumber}`);

      return {
        ...annotation,
        lineNumber: foundLineNumber,
        lineNumberRange: String(foundLineNumber),
      };
    });

    return enrichedAnnotations;
  }, [annotations, pageTextLines]);



  const extractLineNumberFromSelectedText = useCallback((selectedText) => {
    if (!selectedText) return null;

    // Try multiple patterns
    const patterns = [
      /[Ll]ine\s*[:]?\s*(\d+)/,
      /Ln\.?\s*(\d+)/i,
      /L\.?\s*(\d+)/i,
      /line\s*(\d+)/i,
      /at line\s*(\d+)/i
    ];

    for (const pattern of patterns) {
      const match = selectedText.match(pattern);
      if (match && match[1]) {
        return parseInt(match[1], 10);
      }
    }
    return null;
  }, []);



  const handleExport = useCallback(async () => {
    if (!pdfFile) {
      setNotification({ type: "warning", message: "No PDF to export" });
      return;
    }

    try {
      console.log("=== STARTING EXPORT ===");

      await ensureTextLinesForExport();
      console.log("Available pageTextLines pages:", Object.keys(pageTextLines));

      const getLineNumberFromPageY = (pageLines, y) => {
        if (!Array.isArray(pageLines) || pageLines.length === 0) return null;
        if (y === null || y === undefined || Number.isNaN(Number(y))) return null;

        const numericY = Number(y);

        // 1. exact range match
        for (const line of pageLines) {
          const yMin = Number(line?.yMin);
          const yMax = Number(line?.yMax);

          if (!Number.isNaN(yMin) && !Number.isNaN(yMax)) {
            if (numericY >= yMin && numericY <= yMax) {
              return line.lineNumber;
            }
          }
        }

        // 2. nearest line fallback
        let closestLine = null;
        let minDistance = Infinity;

        for (const line of pageLines) {
          const yMin = Number(line?.yMin);
          const yMax = Number(line?.yMax);
          const yPosition = Number(line?.yPosition);

          let anchorY = null;

          if (!Number.isNaN(yMin) && !Number.isNaN(yMax)) {
            anchorY = (yMin + yMax) / 2;
          } else if (!Number.isNaN(yPosition)) {
            anchorY = yPosition;
          }

          if (anchorY !== null) {
            const distance = Math.abs(numericY - anchorY);
            if (distance < minDistance) {
              minDistance = distance;
              closestLine = line;
            }
          }
        }

        return closestLine ? closestLine.lineNumber : null;
      };

      const getRectYsFromAnnotation = (annotation) => {
        const ys = [];

        if (Array.isArray(annotation?.rects) && annotation.rects.length > 0) {
          annotation.rects.forEach((rect) => {
            const y = Number(rect?.y);
            const h = Number(rect?.height);

            if (!Number.isNaN(y)) ys.push(y);
            if (!Number.isNaN(y) && !Number.isNaN(h)) ys.push(y + h);
          });
          return ys;
        }

        if (Array.isArray(annotation?.normalizedRects) && annotation.normalizedRects.length > 0) {
          const pageEl = pageRefs.current?.[annotation.page];
          const textLayerEl = pageEl?.querySelector?.(".pdf-text-layer");
          const renderedHeight = textLayerEl?.getBoundingClientRect?.().height || null;

          if (renderedHeight) {
            annotation.normalizedRects.forEach((rect) => {
              const y = Number(rect?.y);
              const h = Number(rect?.height);

              if (!Number.isNaN(y)) ys.push((y / 100) * renderedHeight);
              if (!Number.isNaN(y) && !Number.isNaN(h)) {
                ys.push(((y + h) / 100) * renderedHeight);
              }
            });
            return ys;
          }
        }

        const y1 =
          annotation?.yposition ??
          annotation?.YPosition ??
          null;

        const y2 =
          annotation?.endYPosition ??
          annotation?.EndYPosition ??
          null;

        if (y1 !== null && y1 !== undefined && !Number.isNaN(Number(y1))) {
          ys.push(Number(y1));
        }
        if (y2 !== null && y2 !== undefined && !Number.isNaN(Number(y2))) {
          ys.push(Number(y2));
        }

        return ys;
      };

      const enrichedAnnotations = annotations.map((annotation) => {
        if (annotation.type !== "comment") {
          return annotation;
        }

        const pageLines = pageTextLines[annotation.page];
        if (!Array.isArray(pageLines) || pageLines.length === 0) {
          console.log(`No pageTextLines available for page ${annotation.page}`);
          return annotation;
        }

        // keep already-correct saved values
        if (
          annotation.lineNumberRange &&
          annotation.lineNumberRange !== "N/A" &&
          annotation.lineNumberRange !== ""
        ) {
          return annotation;
        }

        if (
          annotation.lineNumber &&
          annotation.lineNumber !== "N/A" &&
          annotation.lineNumber !== ""
        ) {
          return {
            ...annotation,
            lineNumberRange: String(annotation.lineNumber),
          };
        }

        const ys = getRectYsFromAnnotation(annotation);

        if (!ys.length) {
          console.log(`Could not calculate rect Y values for comment ${annotation.id}`);
          return annotation;
        }

        const mappedLines = ys
          .map((y) => getLineNumberFromPageY(pageLines, y))
          .filter((lineNo) => lineNo !== null && lineNo !== undefined);

        if (!mappedLines.length) {
          console.log(`Could not map any line for comment ${annotation.id}`);
          return annotation;
        }

        const minLine = Math.min(...mappedLines);
        const maxLine = Math.max(...mappedLines);
        const middleLine = mappedLines[Math.floor(mappedLines.length / 2)] ?? minLine;

        const lineNumber = middleLine;
        const lineNumberRange =
          minLine === maxLine ? String(minLine) : `${minLine}-${maxLine}`;

        console.log(
          `Comment ${annotation.id}: Page=${annotation.page}, mappedLines=${mappedLines.join(",")}, lineNumber=${lineNumber}, range=${lineNumberRange}`
        );

        return {
          ...annotation,
          lineNumber,
          lineNumberRange,
        };
      });

      console.log("=== FINAL ENRICHED ANNOTATIONS ===");
      enrichedAnnotations.forEach((ann, idx) => {
        if (ann.type === "comment") {
          console.log(
            `Comment ${idx + 1}: Page=${ann.page}, lineNumber=${ann.lineNumber ?? "N/A"}, lineNumberRange=${ann.lineNumberRange ?? "N/A"}`
          );
        }
      });

      const bytes = await exportPdfWithAnnotations(pdfFile, enrichedAnnotations, scale);
      const blob = new Blob([bytes], { type: "application/pdf" });

      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `annotated_${fileName}`;
      document.body.appendChild(anchor);
      anchor.click();
      document.body.removeChild(anchor);
      URL.revokeObjectURL(url);

      setNotification({
        type: "success",
        message: "PDF exported successfully with annotations!",
      });
    } catch (error) {
      console.error("Export failed:", error);
      setNotification({
        type: "error",
        message: `Failed to export PDF: ${error.message}`,
      });
    }
  }, [pdfFile, annotations, scale, fileName, pageTextLines, ensureTextLinesForExport]);




  const handleUndo = useCallback(() => {
    if (historyIndex > 0) {
      const newIndex = historyIndex - 1;
      setHistoryIndex(newIndex);
      setAnnotations(history[newIndex]);
      saveAnnotationsToStorage(history[newIndex]);
      setNotification({ type: 'success', message: 'Undo successful' });
    }
  }, [history, historyIndex, saveAnnotationsToStorage]);

  const handleRedo = useCallback(() => {
    if (historyIndex < history.length - 1) {
      const newIndex = historyIndex + 1;
      setHistoryIndex(newIndex);
      setAnnotations(history[newIndex]);
      saveAnnotationsToStorage(history[newIndex]);
      setNotification({ type: 'success', message: 'Redo successful' });
    }
  }, [history, historyIndex, saveAnnotationsToStorage]);

  // Print
  const handlePrint = () => {
    if (!pdfBlob) {
      console.error("No PDF data available for printing.");
      return;
    }

    try {
      const blob = new Blob([pdfBlob], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const printWindow = window.open(url, "_blank");

      if (!printWindow) {
        console.error("Failed to open print window.");
        return;
      }

      printWindow.onload = () => {
        printWindow.focus();
        printWindow.print();
      };
    } catch (error) {
      console.error("Error handling print:", error);
    }
  };

  // Zoom controls
  const applyPercentageZoom = useCallback((value) => {
    const nextZoom = clampZoom(value);
    setScale(nextZoom);
    setZoomMode("percentage");
    setIsZoomDropdownOpen(false);
  }, [clampZoom]);

  const applyActualSize = useCallback(() => {
    setScale(1);
    setZoomMode("actual-size");
    setIsZoomDropdownOpen(false);
  }, []);

  const applyPageFit = useCallback(() => {
    const container = viewerContainerRef.current;
    const pageEl = container?.querySelector(".pdf-page-container");
    const canvas = pageEl?.querySelector("canvas");

    if (!container || !canvas) return;

    const availableWidth = container.clientWidth - 40;
    const availableHeight = container.clientHeight - 40;

    const baseWidth = canvas.width / scale;
    const baseHeight = canvas.height / scale;

    if (!baseWidth || !baseHeight) return;

    const fitScale = clampZoom(
      Math.min(availableWidth / baseWidth, availableHeight / baseHeight)
    );

    setScale(fitScale);
    setZoomMode("page-fit");
    setIsZoomDropdownOpen(false);
  }, [scale, clampZoom]);

  const applyPageWidth = useCallback(() => {
    const container = viewerContainerRef.current;
    const pageEl = container?.querySelector(".pdf-page-container");
    const canvas = pageEl?.querySelector("canvas");

    if (!container || !canvas) return;

    const availableWidth = container.clientWidth - 40;
    const baseWidth = canvas.width / scale;

    if (!baseWidth) return;

    const widthScale = clampZoom(availableWidth / baseWidth);

    setScale(widthScale);
    setZoomMode("page-width");
    setIsZoomDropdownOpen(false);
  }, [scale, clampZoom]);

  //new
  const handleZoomIn = useCallback(() => {
    const current = Number(scale || 1);
    const next = clampZoom(current + 0.1); // 🔥 +10%
    applyPercentageZoom(next);
  }, [scale, applyPercentageZoom, clampZoom]);

  const handleZoomOut = useCallback(() => {
    const current = Number(scale || 1);
    const next = clampZoom(current - 0.1); // 🔥 -10%
    applyPercentageZoom(next);
  }, [scale, applyPercentageZoom, clampZoom]);
  //new
  // ============================================================
  // EXPOSE METHODS TO PARENT COMPONENT VIA REF
  // Add this after all your useCallback hooks (around line 1400-1500)
  // ============================================================
  useImperativeHandle(ref, () => ({
    getAnnotations: () => annotations || [],

    getHighlights: () =>
      (annotations || []).filter(
        (a) => (a.type || "").toLowerCase() === "highlight"
      ),

    getComments: () =>
      (annotations || []).filter(
        (a) => (a.type || "").toLowerCase() === "comment"
      ),
    persistAnnotationToServer,

    // ✅ After approve success, remove synced draft items from localStorage/state
    removeSyncedLocalAnnotations: (ids = []) => {
      if (!Array.isArray(ids) || ids.length === 0) return;

      setAnnotations((prev) => {
        const updated = (prev || []).filter((item) => !ids.includes(item.id));
        saveAnnotationsToStorage(updated);
        return updated;
      });
    },

    getPdfDoc: () => pdfDoc,

    getAnnotationCount: () => ({
      highlights: (annotations || []).filter(
        (a) => (a.type || "").toLowerCase() === "highlight"
      ).length,
      comments: (annotations || []).filter(
        (a) => (a.type || "").toLowerCase() === "comment"
      ).length,
      total: (annotations || []).length,
    }),

    // ✅ ONLY COMMENTS go to NoteApproverCommentsDTO
    getAllFormattedForAPI: () => {
      return (annotations || [])
        .filter(
          (annotation) =>
            (annotation.type || "").toLowerCase() === "comment" &&
            annotation.text?.trim()
        )
        .map((comment) => {
          const serverAnnotationId =
            comment.serverResponse?.annotationId ||
            comment.serverResponse?.AnnotationId ||
            comment.annotationId ||
            comment.AnnotationId ||
            null;

          const lineNo =
            comment?.lineNumberRange && comment.lineNumberRange !== "N/A"
              ? String(comment.lineNumberRange)
              : Array.isArray(comment?.lineNumber)
                ? comment.lineNumber.join("-")
                : comment?.lineNumber !== undefined && comment?.lineNumber !== null
                  ? String(comment.lineNumber)
                  : "NA";

          return {
            pageNumber: comment.page?.toString() || "NA",
            docReferrence: lineNo,
            comments: `[PDF Comment]: ${comment.text || ""}`,

            pdfAnnotationId: comment.id,
            annotationType: "comment",
            lineNumber: lineNo,

            annotationId: serverAnnotationId,
            AnnotationId: serverAnnotationId,

            rawAnnotation: {
              id: comment.id,
              type: "comment",
              user: comment.user,
              page: comment.page,
              text: comment.text,
              selectedText: comment.selectedText || "",
              lineNumber: lineNo,
              lineNumberRange: lineNo,
              annotationId: serverAnnotationId,
              AnnotationId: serverAnnotationId,
            },
          };
        });
    },

    clearAnnotations: () => {
      setAnnotations([]);
      saveAnnotationsToStorage([]);
      setHistory([]);
      setHistoryIndex(-1);
    },

    setZoom: (zoomValue) => {
      setScale(clampZoom(zoomValue));
      setZoomMode("percentage");
    },

    getZoom: () => scale,

    hasAnnotations: () => (annotations || []).length > 0,

    getAnnotationsSummary: () => {
      const highlights = (annotations || []).filter(
        (a) => (a.type || "").toLowerCase() === "highlight"
      );
      const comments = (annotations || []).filter(
        (a) => (a.type || "").toLowerCase() === "comment"
      );

      let summary = "";

      if (highlights.length > 0) {
        summary += `${highlights.length} Highlight(s):\n`;
        highlights.forEach((h, i) => {
          summary += `  ${i + 1}. Page ${h.page}, Line ${h.lineNumber || "NA"}: "${h.selectedText?.substring(0, 50) || ""}"\n`;
        });
      }

      if (comments.length > 0) {
        summary += `${comments.length} Comment(s):\n`;
        comments.forEach((c, i) => {
          summary += `  ${i + 1}. Page ${c.page}, Line ${c.lineNumber || "NA"}: "${c.text?.substring(0, 50) || ""}"\n`;
        });
      }

      return summary;
    },
  }), [
    annotations,
    scale,
    clampZoom,
    saveAnnotationsToStorage,
    persistAnnotationToServer,
  ]);

  const getCurrentZoomLabel = useCallback(() => {
    if (zoomMode === "actual-size") return "Actual size";
    if (zoomMode === "page-fit") return "Page fit";
    if (zoomMode === "page-width") return "Page width";
    return `${Math.round((scale || 1) * 100)}%`;
  }, [scale, zoomMode]);

  //new

  // // Page navigation
  // const goToPrevPage = useCallback(() => {
  //   setCurrentPage(prev => Math.max(1, prev - 1));
  // }, []);

  // const goToNextPage = useCallback(() => {
  //   setCurrentPage(prev => Math.min(numPages, prev + 1));
  // }, [numPages]);
  //////////////////Digvijay Pagination New
  // Page navigation with smooth scroll


  // Page navigation with scroll to page - IMPROVED VERSION
  const scrollToPage = useCallback((pageNumber) => {
    const container = viewerContainerRef.current;

    if (!container) return;

    // Try multiple ways to find the page element
    let pageEl = pageRefs.current[pageNumber];

    if (!pageEl) {
      // Try to find by data attribute
      pageEl = container.querySelector(`[data-page-num="${pageNumber}"]`);
    }

    if (!pageEl) {
      // Try to find by class and data attribute
      pageEl = container.querySelector(`.pdf-page-container[data-page="${pageNumber}"]`);
    }

    if (!pageEl) {
      // Last resort - find by any element containing page data
      const allPages = container.querySelectorAll('[data-page-num], [data-page]');
      for (let i = 0; i < allPages.length; i++) {
        const el = allPages[i];
        const pageAttr = el.getAttribute('data-page-num') || el.getAttribute('data-page');
        if (pageAttr && parseInt(pageAttr) === pageNumber) {
          pageEl = el;
          break;
        }
      }
    }

    if (!pageEl) {
      console.warn(`Page element ${pageNumber} not found, trying fallback`);
      // Fallback: scroll to top and show notification
      container.scrollTo({ top: 0, behavior: "smooth" });
      setNotification({ type: 'warning', message: `Page ${pageNumber} is loading, please wait...` });
      return;
    }

    // Calculate scroll position
    const containerRect = container.getBoundingClientRect();
    const pageRect = pageEl.getBoundingClientRect();

    const scrollTop = container.scrollTop + (pageRect.top - containerRect.top) - 20;

    container.scrollTo({
      top: Math.max(0, scrollTop),
      behavior: "smooth",
    });

    setCurrentPage(pageNumber);
  }, [setNotification]);

  const goToPrevPage = useCallback(() => {
    const prevPage = Math.max(1, currentPage - 1);
    scrollToPage(prevPage);
  }, [currentPage, scrollToPage]);

  const goToNextPage = useCallback(() => {
    const nextPage = Math.min(numPages, currentPage + 1);
    scrollToPage(nextPage);
  }, [currentPage, numPages, scrollToPage]);


  // Memoized page component
  const memoizedPdfPages = useMemo(() => {
    if (!pdfDoc || numPages === 0) return null;

    return Array.from({ length: numPages }, (_, index) => {
      const pageNumber = index + 1;

      return (
        <div
          key={pageNumber}
          ref={(el) => {
            if (el) pageRefs.current[pageNumber] = el;
          }}
          data-page={pageNumber}
          data-page-number={pageNumber}
          style={{
            display: "flex",
            justifyContent: "center",
            width: "100%",
          }}
        >
          <LazyPdfPage
            pageNum={pageNumber}
            pdfDoc={pdfDoc}
            scale={scale}
            annotations={annotations}
            hoveredId={hoveredId}
            setHoveredId={setHoveredId}
            onCommentClick={(comment) => {
              console.log("Comment clicked:", comment);
            }}
            onEditComment={handleEditComment}
            onViewOnlyComment={handleViewComment}  // Add this line
            onTextLinesReady={handleTextLinesReady}
            isReadOnly={isReadOnly}
          />
        </div>
      );
    });
  }, [pdfDoc, numPages, scale, annotations, hoveredId, handleEditComment, handleViewComment, handleTextLinesReady, isReadOnly]); // Add handleViewComment to dependencies
  if (loading) {
    return (
      <div style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        height: "600px",
        color: "#666",
        fontSize: "14px",
        gap: "16px"
      }}>
        <div style={{
          width: "40px",
          height: "40px",
          border: "3px solid #f3f3f3",
          borderTop: "3px solid #0056b3",
          borderRadius: "50%",
          animation: "spin 1s linear infinite"
        }} />
        <div>Loading PDF... Please wait</div>
      </div>
    );
  }

  if (error) {
    return (
      <div style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        height: "600px",
        color: "#dc3545",
        gap: "16px",
        textAlign: "center",
        padding: "20px"
      }}>
        <div style={{ fontSize: "48px" }}>⚠️</div>
        <div style={{ fontSize: "14px", maxWidth: "400px" }}>{error}</div>
        <div style={{ fontSize: "12px", color: "#999" }}>
          Please ensure you have a valid PDF file
        </div>
      </div>
    );
  }

  return (
    <div style={{
      display: "flex",
      flexDirection: "column",
      height: "100%",
      fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
      backgroundColor: "#f5f5f5",
      borderRadius: "8px",
      overflow: "hidden"
    }}>
      {/* Reset Confirmation Dialog */}
      <ResetConfirmationDialog
        isOpen={showResetDialog}
        onClose={() => setShowResetDialog(false)}
        onConfirm={handleResetConfirm}
        annotationCount={annotations.length}
        currentUserCount={
          annotations.filter(
            (a) =>
              (a?.user || "").toLowerCase() ===
              (account?.username || currentUser || "").toLowerCase()
          ).length
        }
      />

      {/* Notification Toast */}
      {notification && (
        <div style={{
          position: "fixed",
          top: "80px",
          right: "20px",
          zIndex: 1000,
          backgroundColor: notification.type === 'success' ? '#28a745' :
            notification.type === 'error' ? '#dc3545' :
              '#ffc107',
          color: "white",
          padding: "12px 20px",
          borderRadius: "8px",
          boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
          animation: "slideIn 0.3s ease-out",
          fontSize: "13px",
          fontWeight: "500",
          pointerEvents: "none"
        }}>
          {notification.message}
        </div>
      )}

      {/* // Update the selection toolbar rendering - hide entirely in read-only mode */}
      {!isReadOnly && selectionToolbar.visible && selectionToolbar.selectionData && (
        <div
          id="pdf-selection-toolbar"
          style={{
            position: "fixed",
            top: selectionToolbar.y,
            left: selectionToolbar.x,
            transform: "translate(-50%, -100%)",
            background: "#ffffff",
            border: "1px solid #e0e0e0",
            borderRadius: "8px",
            boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
            zIndex: 99999,
            display: "flex",
            alignItems: "center",
            gap: "6px",
            padding: "6px 8px",
          }}
          onMouseDown={(e) => e.preventDefault()}
        >
          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => handleAddHighlight(selectionToolbar.selectionData)}
            style={{
              width: "32px",
              height: "32px",
              border: "none",
              background: "#f5f5f5",
              borderRadius: "6px",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
            onMouseEnter={(e) => e.currentTarget.style.background = "#e6f0ff"}
            onMouseLeave={(e) => e.currentTarget.style.background = "#f5f5f5"}
            title="Highlight selected text"
          >
            ✏️
          </button>

          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => handleAddComment(selectionToolbar.selectionData)}
            style={{
              width: "32px",
              height: "32px",
              border: "none",
              background: "#f5f5f5",
              borderRadius: "6px",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
            onMouseEnter={(e) => e.currentTarget.style.background = "#e6f0ff"}
            onMouseLeave={(e) => e.currentTarget.style.background = "#f5f5f5"}
            title="Add comment on selected text"
          >
            💬
          </button>
        </div>
      )}


      <CommentInputDialog
        isOpen={isCommentDialogOpen}
        onClose={() => {
          setIsCommentDialogOpen(false);
          setEditComment(null);
          setIsViewOnlyComment(false);
        }}
        onSubmit={handleCommentSubmit}
        selectedText={selectedTextForComment}
        pageNo={currentPageNo}
        lineNumber={currentLineNumber}
        editComment={editComment}
        viewOnly={isViewOnlyComment}  // Make sure this is passed
      />
      {/* Toolbar */}
      <div style={{
        padding: "12px 20px",
        backgroundColor: "white",
        borderBottom: "1px solid #e0e0e0",
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        flexWrap: "wrap",
        gap: "8px"
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: "4px", flexWrap: "wrap" }}>
          <div style={{
            padding: "4px 9px",
            backgroundColor: "#f8f9fa",
            borderRadius: "6px",
            fontSize: "14px",
            fontWeight: "500",
            color: "#495057",
            border: "1px solid #dee2e6",
            height: "35px",
            width: "34px",
          }}>
            <SvgIcon icon={filePdfIcon} size="small" style={{ marginRight: "6px" }} />
          </div>

          <div style={{ width: "1px", height: "30px", backgroundColor: "#e0e0e0", margin: "0 4px" }} />

          <IconButton
            icon={zoomOutIcon}
            onClick={handleZoomOut}
            title="Zoom Out"
            disabled={scale <= 0.5}
          />

          <ZoomDropdown
            scale={scale}
            zoomMode={zoomMode}
            isOpen={isZoomDropdownOpen}
            onToggle={() => setIsZoomDropdownOpen((prev) => !prev)}
            onActualSize={applyActualSize}
            onPageFit={applyPageFit}
            onPageWidth={applyPageWidth}
            onPercentageZoom={applyPercentageZoom}
            getCurrentZoomLabel={getCurrentZoomLabel}
            dropdownRef={zoomDropdownRef}
          />

          <IconButton
            icon={zoomInIcon}
            onClick={handleZoomIn}
            title="Zoom In"
            disabled={scale >= 3}
          />


          <div style={{ width: "1px", height: "30px", backgroundColor: "#e0e0e0", margin: "0 4px" }} />

          <IconButton
            icon={caretAltLeftIcon}
            onClick={goToPrevPage}
            title="Previous Page"
            disabled={currentPage <= 1}
          />


          <span style={{
            fontSize: "12px",
            minWidth: "70px",
            textAlign: "center",
            fontWeight: "600",
            color: "#495057"
          }}>
            Page {currentPage} of {numPages}
          </span>
          <IconButton
            icon={caretAltRightIcon}
            onClick={goToNextPage}
            title="Next Page"
            disabled={currentPage >= numPages}
          />
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: "4px", flexWrap: "wrap" }}>
          {!isReadOnly && (
            <>


              {/* priti */}
              <IconButton
                icon={brushIcon}
                onClick={() => handleAddHighlight(selectionToolbar.selectionData)}
                title="Highlight Selected Text"
                disabled={isAddingHighlight}
                loading={isAddingHighlight}
              />

              <IconButton
                icon={commentIcon}
                onClick={() => handleAddComment(selectionToolbar.selectionData)}
                title="Add Comment to Selected Text"
              />
              {/* priti */}



              <div style={{ width: "1px", height: "30px", backgroundColor: "#e0e0e0", margin: "0 4px" }} />
            </>
          )}

          <IconButton
            icon={downloadIcon}
            onClick={handleExport}
            title="Export PDF with Annotations"
          />
          <IconButton
            icon={printIcon}
            onClick={handlePrint}
            title="Print PDF"
          />

          {!isReadOnly && (
            <>
              <div
                style={{
                  width: "1px",
                  height: "30px",
                  backgroundColor: "#e0e0e0",
                  margin: "0 4px"
                }}
              />
              <IconButton
                icon={trashIcon}
                onClick={handleReset}
                title="Delete Your Annotations"
                disabled={isResetting}
                loading={isResetting}
              />
            </>
          )}
        </div>
      </div>

      {/* Stats Bar */}
      <div style={{
        padding: "8px 20px",
        backgroundColor: "#f8f9fa",
        borderBottom: "1px solid #e0e0e0",
        fontSize: "12px",
        color: "#6c757d",
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center"
      }}>
        <span>
          📝 <strong>{annotations.filter(a => a.type === "highlight").length}</strong> Highlights &nbsp;|&nbsp;
          💬 <strong>{annotations.filter(a => a.type === "comment").length}</strong> Comments
        </span>
        <span style={{ color: "#0056b3", fontSize: "11px" }}>
          💡 Select text → Click Highlight or Comment button
        </span>
      </div>

      {/* PDF Viewer  Update Digvijay */}
      {/* PDF Viewer - Multi-page continuous scrolling with lazy loading */}
      {/* //Digvijay 01-04 for page gap */}

      {/* PDF Viewer - Multi-page continuous scrolling with lazy loading */}
      {/* PDF Viewer - Multi-page continuous scrolling with lazy loading */}
      {pdfDoc && (
        <div
          ref={viewerContainerRef}
          style={{
            flex: 1,
            overflowY: "auto",
            overflowX: "auto",
            padding: "20px",
            backgroundColor: "#e9ecef",
            display: "flex",
            justifyContent: "center"
          }}
        >
          <div style={{
            maxWidth: "100%",
            width: "fit-content",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            gap: "12px"
          }}>
            {/* Render ALL pages with lazy loading */}
            {Array.from({ length: numPages }, (_, i) => i + 1).map((pageNum) => (
              <div
                key={pageNum}
                ref={(el) => {
                  if (el) {
                    // Store the container element reference
                    if (!pageRefs.current[pageNum]) {
                      pageRefs.current[pageNum] = el.querySelector(`[data-page-num="${pageNum}"]`) || el;
                    }
                  }
                }}
              >
                <LazyPdfPage
                  pageNum={pageNum}
                  pdfDoc={pdfDoc}
                  scale={scale}
                  annotations={annotations}
                  hoveredId={hoveredId}
                  setHoveredId={setHoveredId}
                  onCommentClick={(comment) => {
                    console.log("Comment clicked:", comment);
                  }}
                  onEditComment={handleEditComment}
                  onViewOnlyComment={handleViewComment}  // Add this line
                  onTextLinesReady={handleTextLinesReady}
                  onPageRenderComplete={(renderedPageNum) => {
                    handlePageRenderComplete(renderedPageNum);
                    const container = viewerContainerRef.current;
                    if (container) {
                      const pageElement = container.querySelector(`[data-page-num="${renderedPageNum}"]`);
                      if (pageElement) {
                        pageRefs.current[renderedPageNum] = pageElement;
                      }
                    }
                  }}
                  isReadOnly={isReadOnly}
                />
              </div>
            ))}
          </div>
        </div>
      )}
      {/* //////////////////////////////////// */}

      <style jsx>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        
        @keyframes slideIn {
          from {
            transform: translateX(100%);
            opacity: 0;
          }
          to {
            transform: translateX(0);
            opacity: 1;
          }
        }
        
        @keyframes fadeIn {
          from {
            opacity: 0;
          }
          to {
            opacity: 1;
          }
        }
        
        @keyframes slideUp {
          from {
            transform: translateY(20px);
            opacity: 0;
          }
          to {
            transform: translateY(0);
            opacity: 1;
          }
        }
        
        @media print {
          .pdf-page-container {
            page-break-after: always;
          }
          .toolbar, .stats-bar {
            display: none !important;
          }
        }
      `}</style>
    </div>
  );
});



export default PdfAnnotator;


































































































































