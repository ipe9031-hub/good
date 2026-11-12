// Test double only. Never imported by the production application.
(() => {
  let counter = 1;
  const state = (window.__officeTest = {
    selectedSlide: "slide-1",
    selected: [],
    shapes: [],
    calls: [],
    handler: null,
  });
  const collection = (items) => ({ items, load() {} });
  function shape(type, options, text = "") {
    const obj = {
      id: `shape-${counter++}`,
      name: text || type,
      type,
      ...options,
      textFrame: { textRange: { text, font: {} } },
      fill: { setSolidColor() {} },
      delete() {
        state.calls.push("delete");
        state.shapes = state.shapes.filter((s) => s.id !== obj.id);
      },
    };
    state.shapes.push(obj);
    state.calls.push(type);
    return obj;
  }
  state.shapes.push({
    id: "original",
    name: "보존할 원본 제목",
    type: "TextBox",
    left: 10,
    top: 20,
    width: 100,
    height: 60,
    delete() {
      state.calls.push("delete");
      state.shapes = state.shapes.filter((s) => s.id !== "original");
    },
  });
  const slide = () => ({
    id: state.selectedSlide,
    shapes: {
      addTextBox: (text, b) => shape("TextBox", b, text),
      addGeometricShape: (type, b) => shape(type, b),
      getItem: (id) => state.shapes.find((s) => s.id === id),
    },
  });
  window.Office = {
    HostType: { PowerPoint: "PowerPoint" },
    EventType: { DocumentSelectionChanged: "selection" },
    CoercionType: { Image: "image" },
    AsyncResultStatus: { Succeeded: "succeeded" },
    onReady: (callback) => {
      const info = { host: "PowerPoint" };
      if (callback) callback(info);
      return Promise.resolve(info);
    },
    context: {
      requirements: { isSetSupported: () => true },
      document: {
        url: "https://example.test/안전점검.pptx",
        addHandlerAsync: (_type, handler, callback) => {
          state.handler = handler;
          callback({ status: "succeeded" });
        },
        removeHandlerAsync: () => {},
        setSelectedDataAsync: (_data, options, callback) => {
          shape("Image", {
            left: options.imageLeft,
            top: options.imageTop,
            width: options.imageWidth,
            height: options.imageHeight,
          });
          callback({ status: "succeeded" });
        },
      },
    },
  };
  window.PowerPoint = {
    run: async (fn) =>
      fn({
        sync: async () => {},
        presentation: {
          slides: collection([{ id: "slide-1" }, { id: "slide-2" }]),
          getSelectedSlides: () => collection([slide()]),
          getSelectedShapes: () =>
            collection(
              state.shapes.filter((s) => state.selected.includes(s.id)),
            ),
        },
      }),
  };
})();
