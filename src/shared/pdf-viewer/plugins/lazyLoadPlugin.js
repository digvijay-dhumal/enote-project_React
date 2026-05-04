export const lazyLoadPlugin = {
  install(engine) {
    let observer = null;
    engine.startLazyLoading = (container, buffer = 2) => {
      if (!container) return;
      const options = { root: container, rootMargin: `${buffer * 100}%` };
      observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
          if (entry.isIntersecting) {
            const pageNum = parseInt(entry.target.dataset.page, 10);
            // render this page and adjacent
            engine.renderRange(pageNum - buffer, pageNum + buffer);
          }
        });
      }, options);

      // Observe all page placeholders
      // This would be set up by the UI component when pages are added
    };
  }
};