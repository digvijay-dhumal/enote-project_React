export const downloadPlugin = {
  install(engine) {
    /**
     * Download PDF from base64 or from original data
     * @param {string} [base64Pdf] Optional base64 PDF string
     * @param {string} [filename] Optional filename
     */
    engine.download = (base64Pdf, filename = 'document.pdf') => {
      let blob;
      if (base64Pdf) {
        try {
          const binaryString = atob(base64Pdf);
          const len = binaryString.length;
          const bytes = new Uint8Array(len);
          for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
          }
          blob = new Blob([bytes], { type: 'application/pdf' });
        } catch (error) {
          console.error('Error handling download:', error);
          return;
        }
      } else {
        blob = new Blob([engine.originalData], { type: 'application/pdf' });
      }
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    };
  }
};