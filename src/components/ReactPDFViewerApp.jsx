import React, { useEffect, useState, useRef } from "react";

const ReactPDFViewerApp = ({
    pdfBlob,
    fileName,
    onAnnotationsChange,
    isLocked,
    currentUser
}) => {
    const [blobUrl, setBlobUrl] = useState(null);
    const [pdfInstance, setPdfInstance] = useState(null);
    const containerRef = useRef(null);

    // -------------------- DOWNLOAD PDF --------------------
    const handleDownload = () => {
        if (!pdfBlob) return;

        const downloadUrl = URL.createObjectURL(pdfBlob);
        const a = document.createElement("a");
        a.href = downloadUrl;
        a.download = fileName || "document.pdf";
        document.body.appendChild(a);
        a.click();

        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(downloadUrl);
        }, 0);
    };

    // -------------------- PRINT PDF --------------------
    const handlePrint = () => {
        if (!pdfBlob) {
            console.error("No PDF data available for printing.");
            return;
        }

        try {
            const blob = new Blob([pdfBlob], { type: "application/pdf" });
            const url = URL.createObjectURL(blob);
            const win = window.open(url, "_blank");

            if (!win) return console.error("Popup blocked!");

            win.onload = () => {
                win.focus();
                win.print();
            };
        } catch (err) {
            console.error("Print error:", err);
        }
    };

    // -------------------- SAVE ANNOTATIONS --------------------
    const handleSaveAnnotations = async () => {
        if (!pdfInstance) return;

        try {
            const annotations = await pdfInstance.exportAnnotations();

            if (onAnnotationsChange) {
                onAnnotationsChange(annotations);
            }

            console.log("Annotations exported successfully");
        } catch (err) {
            console.error("Failed to save annotations:", err);
        }
    };

    // -------------------- INIT PSPDFKit --------------------
    const initializePSPDFKit = async () => {
        if (!blobUrl || !containerRef.current) return;

        try {
            const PSPDFKit = await import("pspdfkit");

            const instance = await PSPDFKit.load({
                container: containerRef.current,
                document: blobUrl,
                baseUrl: `${window.location.origin}/`,
                licenseKey: "your-license-key-here",

                toolbarItems: [
                    ...PSPDFKit.defaultToolbarItems,
                    { type: "content-editor" },
                    { type: "signature" },
                    { type: "export-pdf" }
                ],

                editableAnnotationTypes: isLocked
                    ? []
                    : [
                          PSPDFKit.Annotations.FreeTextAnnotation,
                          PSPDFKit.Annotations.HighlightAnnotation,
                          PSPDFKit.Annotations.UnderlineAnnotation,
                          PSPDFKit.Annotations.StrikeOutAnnotation,
                          PSPDFKit.Annotations.NoteAnnotation,
                          PSPDFKit.Annotations.InkAnnotation,
                          PSPDFKit.Annotations.ShapeAnnotation,
                          PSPDFKit.Annotations.LineAnnotation,
                          PSPDFKit.Annotations.PolygonAnnotation
                      ],

                onAnnotationsChange: (annotations) => {
                    if (onAnnotationsChange) {
                        onAnnotationsChange(annotations);
                    }
                }
            });

            setPdfInstance(instance);
        } catch (err) {
            console.error("PSPDFKit load error:", err);
        }
    };

    // -------------------- BLOB URL SETUP --------------------
    useEffect(() => {
        if (!pdfBlob || !(pdfBlob instanceof Blob)) return;

        const url = URL.createObjectURL(pdfBlob);
        setBlobUrl(url);

        return () => {
            URL.revokeObjectURL(url);
        };
    }, [pdfBlob]);

    // -------------------- LOAD VIEWER WHEN BLOB URL READY --------------------
    useEffect(() => {
        if (!blobUrl) return;

        initializePSPDFKit();

        return () => {
            if (pdfInstance) {
                pdfInstance.destroy();
            }
        };
    }, [blobUrl]);

    return (
        <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
            {/* -------------------- CUSTOM TOOLBAR -------------------- */}
            <div
                style={{
                    padding: "10px",
                    backgroundColor: "#f5f5f5",
                    borderBottom: "1px solid #ddd",
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center"
                }}
            >
                <div>
                    <strong style={{ marginRight: "20px" }}>{fileName}</strong>

                    {isLocked && (
                        <span
                            style={{
                                color: "#ff6b35",
                                backgroundColor: "#ffeaea",
                                padding: "3px 8px",
                                borderRadius: "4px",
                                fontSize: "12px"
                            }}
                        >
                            🔒 Locked for {currentUser?.name || "User"}
                        </span>
                    )}
                </div>

                <div style={{ display: "flex", gap: "10px" }}>
                    {!isLocked && (
                        <button
                            onClick={handleSaveAnnotations}
                            style={{
                                backgroundColor: "#007bff",
                                color: "#fff",
                                border: "none",
                                padding: "8px 16px",
                                borderRadius: "4px",
                                cursor: "pointer"
                            }}
                        >
                            Save Annotations
                        </button>
                    )}

                    <button
                        onClick={handlePrint}
                        style={{
                            border: "1px solid #007bff",
                            color: "#007bff",
                            background: "transparent",
                            padding: "8px 16px",
                            borderRadius: "4px"
                        }}
                    >
                        Print
                    </button>

                    <button
                        onClick={handleDownload}
                        style={{
                            border: "1px solid #28a745",
                            color: "#28a745",
                            background: "transparent",
                            padding: "8px 16px",
                            borderRadius: "4px"
                        }}
                    >
                        Download
                    </button>
                </div>
            </div>

            {/* -------------------- PSPDFKit Viewer Container -------------------- */}
            <div
                ref={containerRef}
                style={{
                    flex: 1,
                    height: "calc(100% - 60px)",
                    position: "relative"
                }}
            />
        </div>
    );
};

export default ReactPDFViewerApp;
