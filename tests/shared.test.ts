import { describe, expect, it } from "vitest";
import {
  assertSameTarget,
  targetOf,
  validateBounds,
  validateImage,
  isImageData,
  type Selection,
} from "../packages/shared/src";
const selection: Selection = {
  documentName: "test",
  slideCount: 3,
  slideIds: ["slide1"],
  shapes: [],
};
describe("대상 보호", () => {
  it("단일 슬라이드만 허용한다", () => {
    expect(targetOf(selection)).toEqual({ slideId: "slide1", shapeIds: [] });
    expect(() => targetOf({ ...selection, slideIds: [] })).toThrow();
    expect(() => targetOf({ ...selection, slideIds: ["a", "b"] })).toThrow();
  });
  it("삭제에는 선택 개체가 필요하다", () =>
    expect(() => targetOf(selection, true)).toThrow());
  it("승인 후 슬라이드 또는 개체가 바뀌면 중단한다", () => {
    const t = { slideId: "a", shapeIds: ["one", "two"] };
    expect(() => assertSameTarget(t, { ...t, slideId: "b" }, false)).toThrow();
    expect(() =>
      assertSameTarget(t, { ...t, shapeIds: ["one"] }, true),
    ).toThrow();
    expect(() =>
      assertSameTarget(t, { ...t, shapeIds: ["two", "one"] }, true),
    ).not.toThrow();
  });
});
describe("입력 검증", () => {
  it.each([NaN, Infinity, -1, 10001])("잘못된 좌표 %s를 거부한다", (v) =>
    expect(() =>
      validateBounds({ left: v, top: 0, width: 100, height: 80 }),
    ).toThrow(),
  );
  it("0 크기를 거부하고 소수 좌표를 허용한다", () => {
    expect(() =>
      validateBounds({ left: 0, top: 0, width: 0, height: 80 }),
    ).toThrow();
    expect(
      validateBounds({ left: 0.5, top: 0, width: 100, height: 80 }).left,
    ).toBe(0.5);
  });
  it("Office 개체 메타데이터와 좌표를 분리한다", () => {
    const shape = {
      id: "shape-1",
      name: "제목",
      left: 20,
      top: 30,
      width: 100,
      height: 80,
    };
    expect(validateBounds(shape)).toEqual({
      left: 20,
      top: 30,
      width: 100,
      height: 80,
    });
  });
  it("큰 이미지와 비이미지 업로드를 거부한다", () => {
    expect(() =>
      validateImage({ type: "image/png", size: 6 * 1024 * 1024 }),
    ).toThrow();
    expect(() => validateImage({ type: "image/svg+xml", size: 100 })).toThrow();
    expect(() =>
      validateImage({ type: "image/jpeg", size: 100 }),
    ).not.toThrow();
    expect(isImageData("data:image/png;base64,not-an-image")).toBe(false);
    expect(isImageData("data:image/png;base64,iVBORw0KGgoAAA")).toBe(true);
  });
});
