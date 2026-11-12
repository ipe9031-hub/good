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
  function shape(kind, options, text = "") {
    const type = ["TextBox", "Image", "Placeholder", "Line"].includes(kind)
      ? kind
      : "GeometricShape";
    const obj = {
      id: `shape-${counter++}`,
      slideId: state.selectedSlide,
      name: text || kind,
      type,
      ...options,
      textFrame: {
        isNullObject: false,
        load() {},
        textRange: {
          text,
          load() {},
          font: {
            name: "맑은 고딕",
            size: 24,
            color: "#172B4D",
            bold: false,
            load() {},
          },
        },
      },
      fill: {
        type: "Solid",
        foregroundColor: "#264F9E",
        load() {},
        setSolidColor(color) {
          this.foregroundColor = color;
          this.type = "Solid";
        },
      },
      lineFormat: { color: "#264F9E", weight: 1, load() {} },
      getTextFrameOrNullObject() {
        return this.textFrame;
      },
      delete() {
        state.calls.push("delete");
        state.shapes = state.shapes.filter((s) => s.id !== obj.id);
      },
    };
    state.shapes.push(obj);
    state.calls.push(kind);
    return obj;
  }
  const original = shape(
    "TextBox",
    { left: 10, top: 20, width: 100, height: 60 },
    "보존할 원본 제목",
  );
  original.id = "original";
  state.calls = [];
  state.seed = shape;
  const slide = (id = state.selectedSlide) => ({
    id,
    shapes: {
      ...collection(state.shapes.filter((s) => s.slideId === id)),
      addTextBox: (text, b) => shape("TextBox", b, text),
      addGeometricShape: (type, b) => shape(type, b),
      addLine: (_type, b) => shape("Line", b),
      getItem: (shapeId) =>
        state.shapes.find((s) => s.id === shapeId && s.slideId === id),
    },
    setSelectedShapes(ids) {
      state.selected = ids;
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
        addHandlerAsync: (_type, handler, cb) => {
          state.handler = handler;
          cb({ status: "succeeded" });
        },
        removeHandlerAsync() {},
        setSelectedDataAsync: (_data, opts, cb) => {
          shape("Image", {
            left: opts.imageLeft,
            top: opts.imageTop,
            width: opts.imageWidth,
            height: opts.imageHeight,
          });
          cb({ status: "succeeded" });
        },
      },
    },
  };
  window.PowerPoint = {
    run: async (fn) =>
      fn({
        sync: async () => {},
        presentation: {
          pageSetup: { slideWidth: 960, slideHeight: 540, load() {} },
          slides: {
            ...collection([slide("slide-1"), slide("slide-2")]),
            getItem: (id) => slide(id),
          },
          getSelectedSlides: () => collection([slide()]),
          getSelectedShapes: () =>
            collection(
              state.shapes.filter(
                (s) =>
                  s.slideId === state.selectedSlide &&
                  state.selected.includes(s.id),
              ),
            ),
        },
      }),
  };
})();
