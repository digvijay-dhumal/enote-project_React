export const navigationPlugin = {
  install(engine) {
    engine.nextPage = () => engine.setCurrentPage(engine.currentPage + 1);
    engine.previousPage = () => engine.setCurrentPage(engine.currentPage - 1);
  }
};