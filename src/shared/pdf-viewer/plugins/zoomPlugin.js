export const zoomPlugin = {
  install(engine) {
    engine.zoomIn = () => engine.setZoom(engine.zoom + 0.25);
    engine.zoomOut = () => engine.setZoom(engine.zoom - 0.25);
  }
};