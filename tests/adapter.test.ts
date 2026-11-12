import { beforeEach, describe, expect, it, vi } from "vitest";
import { PowerPointService } from "../packages/powerpoint-adapter/src";
const bounds = { left: 10, top: 20, width: 100, height: 60 };
const target = { slideId: "s1", shapeIds: ["shape1"] };
let slideIds: string[],
  selectedIds: string[],
  shapes: any[],
  insertImage: ReturnType<typeof vi.fn>;
const addTextBox = vi.fn();
const addGeometricShape = vi.fn();
const deleteShape = vi.fn();
beforeEach(() => {
  vi.clearAllMocks();
  slideIds = ["s1"];
  selectedIds = ["shape1"];
  const shape = {
    id: "shape1",
    name: "기존 제목",
    type: "TextBox",
    ...bounds,
    delete: deleteShape,
  };
  shapes = [shape];
  insertImage = vi.fn((_data, _options, cb) => cb({ status: "succeeded" }));
  vi.stubGlobal("Office", {
    context: {
      requirements: { isSetSupported: () => true },
      document: { url: "", setSelectedDataAsync: insertImage },
    },
    CoercionType: { Image: "image" },
    AsyncResultStatus: { Succeeded: "succeeded" },
  });
  addTextBox.mockReturnValue({ textFrame: { textRange: { font: {} } } });
  addGeometricShape.mockReturnValue({ fill: { setSolidColor: vi.fn() } });
  const collection = (items: any[]) => ({ items, load: vi.fn() });
  const slide = {
    id: "s1",
    shapes: {
      addTextBox,
      addGeometricShape,
      getItem: (id: string) => shapes.find((s) => s.id === id),
    },
  };
  vi.stubGlobal("PowerPoint", {
    run: async (fn: any) =>
      fn({
        sync: async () => {},
        presentation: {
          slides: collection([{ id: "s1" }, { id: "s2" }]),
          getSelectedSlides: () =>
            collection(slideIds.map((id) => ({ ...slide, id }))),
          getSelectedShapes: () =>
            collection(shapes.filter((s) => selectedIds.includes(s.id))),
        },
      }),
  });
});
const service = new PowerPointService();
describe("Office.js 어댑터 계약 (모의 호스트)", () => {
  it("문서·슬라이드·선택 개체 정보를 읽는다", async () => {
    const result = await service.read();
    expect(result.slideCount).toBe(2);
    expect(result.shapes[0].name).toBe("기존 제목");
  });
  it("텍스트와 도형을 실제 Office.js 메서드에 전달한다", async () => {
    await service.insertText(target, "안전조회", bounds);
    await service.insertShape(target, "Rectangle", "#112233", bounds);
    expect(addTextBox).toHaveBeenCalledWith("안전조회", bounds);
    expect(addGeometricShape).toHaveBeenCalledWith("Rectangle", bounds);
  });
  it("선택한 개체만 변경한다", async () => {
    shapes.push({ id: "untouched", ...bounds });
    await service.update(target, { ...bounds, left: 99 });
    expect(shapes[0].left).toBe(99);
    expect(shapes[1].left).toBe(10);
  });
  it("선택한 개체만 삭제한다", async () => {
    await service.remove(target);
    expect(deleteShape).toHaveBeenCalledTimes(1);
  });
  it("대상이 바뀐 후 승인되면 삭제하지 않는다", async () => {
    selectedIds = [];
    await expect(service.remove(target)).rejects.toThrow("바뀌었습니다");
    expect(deleteShape).not.toHaveBeenCalled();
  });
  it("다중 슬라이드에서는 삽입하지 않는다", async () => {
    slideIds = ["s1", "s2"];
    await expect(service.insertText(target, "x", bounds)).rejects.toThrow();
    expect(addTextBox).not.toHaveBeenCalled();
  });
  it("사진을 Common API로만 전송한다", async () => {
    await service.insertImage(
      target,
      "data:image/png;base64,iVBORw0KGgoAAA",
      bounds,
    );
    expect(insertImage).toHaveBeenCalledWith(
      "iVBORw0KGgoAAA",
      expect.objectContaining({ coercionType: "image", imageLeft: 10 }),
      expect.any(Function),
    );
  });
  it("사진 호스트 실패를 전달한다", async () => {
    insertImage.mockImplementation((_d, _o, cb) =>
      cb({ status: "failed", error: { message: "Image failed" } }),
    );
    await expect(
      service.insertImage(
        target,
        "data:image/png;base64,iVBORw0KGgoAAA",
        bounds,
      ),
    ).rejects.toThrow("Image failed");
  });
  it("지원하지 않는 버전에서는 실행하지 않는다", async () => {
    Office.context.requirements.isSetSupported = () => false;
    await expect(service.read()).rejects.toThrow("1.5");
  });
  it("텍스트/도형 명령 검증을 우회할 수 없다", async () => {
    await expect(service.insertText(target, "", bounds)).rejects.toThrow();
    await expect(
      service.insertShape(target, "Unknown" as any, "#112233", bounds),
    ).rejects.toThrow();
  });
});
