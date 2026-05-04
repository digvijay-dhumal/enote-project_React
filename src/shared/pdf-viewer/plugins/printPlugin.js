export const printPlugin = {
  install(engine) {
    /**
     * Print PDF from base64 or from rendered pages
     * @param {string} [base64Pdf] Optional base64 PDF string
     */
    engine.print = async (base64Pdf) => {
      if (base64Pdf) {
        try {
          const binaryString = atob(base64Pdf);
          const len = binaryString.length;
          const bytes = new Uint8Array(len);
          for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
          }
          const blob = new Blob([bytes], { type: 'application/pdf' });
          const url = URL.createObjectURL(blob);
          const printWindow = window.open(url, '_blank');
          if (!printWindow) {
            console.error('Failed to open print window.');
            return;
          }
          printWindow.onload = () => {
            printWindow.focus();
            printWindow.print();
          };
        } catch (error) {
          console.error('Error handling print:', error);
        }
        return;
      }
      // Fallback to original logic (rendered pages)
      const pages = await engine.getAllPages();
      const printWindow = window.open('', '_blank');
      printWindow.document.write('<html><head><title>Print</title></head><body>');
      pages.forEach(dataUrl => {
        printWindow.document.write(`<img src="${dataUrl}" style="width:100%" />`);
      });
      printWindow.document.write('</body></html>');
      printWindow.document.close();
      printWindow.print();
    };
  }
};