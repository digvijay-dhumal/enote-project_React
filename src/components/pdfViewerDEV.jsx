import React, { useEffect, useState } from 'react';
import { SpecialZoomLevel, Viewer, Worker } from '@react-pdf-viewer/core';
import { defaultLayoutPlugin } from '@react-pdf-viewer/default-layout';
import { zoomPlugin } from '@react-pdf-viewer/zoom';

import '@react-pdf-viewer/core/lib/styles/index.css';
import '@react-pdf-viewer/default-layout/lib/styles/index.css';
import '@react-pdf-viewer/zoom/lib/styles/index.css';

const ReactPDFViewerApp = ({ pdfBlob, onZoomTo, fileName  }) => {
    
    const [blobUrl, setBlobUrl] = useState(null);

    const zoomPluginInstance = zoomPlugin();
    const { zoomTo } = zoomPluginInstance;

    const zoomFill = document.getElementById("zoomFill")
    zoomFill && zoomFill.addEventListener("click",
        ()=>{
            console.log('button is clicked')
            zoomTo(SpecialZoomLevel.PageWidth)
        }
    )

    // useEffect(()=>{
    //     zoomTo(SpecialZoomLevel.PageFit)
    // })

    const zoomExit = document.getElementById("zoomExit")
    zoomExit && zoomExit.addEventListener("click",
        ()=>{
            console.log('button is clicked')
            zoomTo(SpecialZoomLevel.PageWidth)
        }
    )

    // Handle zoom commands from parent
   

    const handleDownload = () => {
        if (!pdfBlob) return;
        
        const downloadUrl = URL.createObjectURL(pdfBlob);
        const downloadFileName = fileName
        
        const a = document.createElement('a');
        a.href = downloadUrl;
        a.download = downloadFileName;
        document.body.appendChild(a);
        a.click();
        
        // Cleanup
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(downloadUrl);
        }, 0);
    };


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


    const defaultLayoutPluginInstance = defaultLayoutPlugin({
        sidebarTabs: (defaultTabs) => [],
        renderToolbar: (Toolbar) => (
            <Toolbar>
                {(slots) => {
                    const {
                        CurrentPageInput,
                        EnterFullScreen,
                        GoToNextPage,
                        GoToPreviousPage,
                        NumberOfPages,
                        ShowSearchPopover,
                        Zoom,
                        ZoomIn,
                        ZoomOut,
                    } = slots;

                    return (
                        <div
                            style={{
                                alignItems: 'center',
                                display: 'flex',
                                width: '100%',
                            }}
                        >
                            <div style={{ padding: '0 5px' }}>
                                {/* <ShowSearchPopover /> */}
                            </div>
                            <div style={{ padding: '0 2px' }}>
                                <ZoomOut />
                            </div>
                            <div style={{ padding: '0 2px' }}>
                                <Zoom />
                            </div>
                            <div style={{ padding: '0 2px' }}>
                                <ZoomIn />
                            </div>
                            <div style={{ padding: '0 2px', marginLeft: 'auto' }}>
                                <GoToPreviousPage />
                            </div>
                            <div style={{ padding: '0 2px', width: '50px' }}>
                                <CurrentPageInput />
                            </div>
                            <div style={{ padding: '0 2px' }}>
                                / <NumberOfPages />
                            </div>
                            <div style={{ padding: '0 2px' }}>
                                <GoToNextPage />
                            </div>
                            <div style={{ padding: '0 2px', marginLeft: 'auto' }}>
                                <EnterFullScreen />
                            </div>
                           <div style={{ padding: '0 2px' }}>
                                <button
                                    onClick={handlePrint}
                                    style={{
                                        backgroundColor: 'transparent',
                                        border: 'none',
                                        borderRadius: '4px',
                                        color: '#000',
                                        cursor: 'pointer',
                                        padding: '8px',
                                    }}
                                    aria-label="Print"
                                >
                                    <svg
                                        xmlns="http://www.w3.org/2000/svg"
                                        width="16"
                                        height="16"
                                        viewBox="0 0 24 24"
                                        fill="none"
                                        stroke="currentColor"
                                        strokeWidth="2"
                                        strokeLinecap="round"
                                        strokeLinejoin="round"
                                    >
                                        <polyline points="6 9 6 2 18 2 18 9"></polyline>
                                        <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path>
                                        <rect x="6" y="14" width="12" height="8"></rect>
                                    </svg>
                                </button>
                            </div>
                            <div style={{ padding: '0 2px' }}>
                                <button
                                    style={{
                                        backgroundColor: 'transparent',
                                        border: 'none',
                                        borderRadius: '4px',
                                        color: '#000',
                                        cursor: 'pointer',
                                        padding: '8px',
                                    }}
                                    onClick={handleDownload}
                                    aria-label="Download PDF"
                                >
                                    <svg
                                        width="16"
                                        height="16"
                                        viewBox="0 0 24 24"
                                        fill="none"
                                        stroke="currentColor"
                                        strokeWidth="2"
                                        strokeLinecap="round"
                                        strokeLinejoin="round"
                                    >
                                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                                        <polyline points="7 10 12 15 17 10"></polyline>
                                        <line x1="12" y1="15" x2="12" y2="3"></line>
                                    </svg>
                                </button>
                            </div>
                        </div>
                    );
                }}
            </Toolbar>
        ),
    });

    useEffect(() => {
        if (pdfBlob && pdfBlob instanceof Blob) {
            const url = URL.createObjectURL(pdfBlob);
            setBlobUrl(url);
            console.log(url, "blobUrl");

            return () => {
                URL.revokeObjectURL(url);
            };
        }
    }, [pdfBlob]);

    return (
        <Worker workerUrl='./pdf.worker.bundle.js'>
            <div style={{ height: '100%' }}>
                {blobUrl ? (
                    <Viewer
                        fileUrl={blobUrl}
                        plugins={[
                            defaultLayoutPluginInstance,
                            zoomPluginInstance
                        ]}
                    />
                ) : (
                    <p>Loading PDF...</p>
                )}
            </div>
        </Worker>
    );
};

export default ReactPDFViewerApp;