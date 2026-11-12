import { describe, expect, it } from "vitest";
import {
  automaticSimilarity,
  arrange,
  grid,
  swap,
  colorDistance,
  toPoints,
  inside,
  union,
} from "../packages/layout-engine/src";
import {
  similar,
  type Snapshot,
} from "../packages/powerpoint-adapter/src/editor";
const canvas = { width: 960, height: 540 };
const items = [
  { id: "a", left: 10, top: 10, width: 100, height: 80 },
  { id: "b", left: 200, top: 150, width: 200, height: 100 },
  { id: "c", left: 700, top: 200, width: 100, height: 60 },
];
describe("레이아웃 결과", () => {
  it("명시적 기준 개체에 정렬하고 크기를 통일한다", () => {
    expect(arrange(items, "left", canvas, "b").map((x) => x.left)).toEqual([
      200, 200, 200,
    ]);
    expect(
      arrange(items, "size", canvas, "c").map((x) => [x.width, x.height]),
    ).toEqual([
      [100, 60],
      [100, 60],
      [100, 60],
    ]);
  });
  it("서로 다른 너비에서도 가로 간격을 균등하게 만든다", () => {
    const r = arrange(items, "horizontal", canvas);
    expect(r[1].left - r[0].left - r[0].width).toBe(
      r[2].left - r[1].left - r[1].width,
    );
    expect(r[0].left).toBe(10);
    expect(r[2].left + r[2].width).toBe(800);
  });
  it("부족한 대상과 음수 간격을 거부한다", () => {
    expect(() => arrange(items.slice(0, 1), "left", canvas)).toThrow();
    expect(() => arrange(items.slice(0, 2), "horizontal", canvas)).toThrow();
    expect(() =>
      arrange(
        items.map((x) => ({ ...x, left: 0 })),
        "horizontal",
        canvas,
      ),
    ).toThrow();
  });
  it("3장 사진을 비율을 유지하며 영역 안에 배치한다", () => {
    const area = { left: 20, top: 40, width: 800, height: 400 };
    const r = grid(items, area, 2, 12, true, canvas);
    r.forEach((x, i) => {
      expect(x.width / x.height).toBeCloseTo(items[i].width / items[i].height);
      expect(x.left + x.width).toBeLessThanOrEqual(820);
      expect(x.top + x.height).toBeLessThanOrEqual(440);
    });
  });
  it("영역 밖 그리드와 과도한 간격을 거부한다", () => {
    expect(() =>
      grid(
        items,
        { left: 900, top: 0, width: 400, height: 400 },
        2,
        10,
        true,
        canvas,
      ),
    ).toThrow();
    expect(() =>
      grid(
        items,
        { left: 0, top: 0, width: 10, height: 10 },
        2,
        100,
        true,
        canvas,
      ),
    ).toThrow();
  });
  it("사진 위치 교환에서 크기를 보존한다", () => {
    const r = swap(items.slice(0, 2), true, canvas);
    expect(r[0].left).toBe(200);
    expect(r[1].left).toBe(10);
    expect(r[0].width).toBe(100);
    expect(r[1].width).toBe(200);
  });
  it("슬라이드 경계와 물리 단위를 처리한다", () => {
    expect(toPoints(25.4, "mm")).toBeCloseTo(72);
    expect(toPoints(2.54, "cm")).toBeCloseTo(72);
    expect(
      inside({ left: -10, top: 530, width: 100, height: 80 }, canvas),
    ).toEqual({ left: 0, top: 460, width: 100, height: 80 });
    expect(() =>
      inside({ left: 0, top: 0, width: 1000, height: 10 }, canvas),
    ).toThrow();
    expect(union(items)).toEqual({
      left: 10,
      top: 10,
      width: 790,
      height: 250,
    });
  });
});
describe("유사 개체 조건", () => {
  const a: Snapshot = {
    ...items[0],
    slideId: "s",
    name: "a",
    type: "GeometricShape",
    fill: "#ff0000",
  };
  const b: Snapshot = { ...a, id: "b", fill: "#fe0000" };
  const c: Snapshot = { ...a, id: "c", type: "Image" };
  it("색상 차이와 복합 조건을 적용한다", () => {
    expect(colorDistance("#f00", "red")).toBe(0);
    expect(colorDistance("theme:unknown", "#ff0000")).toBe(Infinity);
    expect(
      similar([a, b, c], a, {
        type: true,
        color: true,
        size: true,
        tolerance: 1,
        sizeTolerance: 0,
      }).map((x) => x.id),
    ).toEqual(["a", "b"]);
  });
  it("기준이 불명확한 검색을 거부한다", () => {
    expect(() =>
      similar([a], a, {
        type: false,
        color: false,
        size: false,
        tolerance: 0,
        sizeTolerance: 0,
      }),
    ).toThrow();
  });
});

describe("자동 유사 오차", () => {
  it("기준 크기에 비례한 가로·세로 오차를 사용하고 지나치게 넓어지지 않는다", () => {
    const small = automaticSimilarity({ width: 20, height: 10 });
    expect(small.sizeToleranceX).toBe(1.5);
    expect(small.sizeToleranceY).toBe(1.5);
    const photo = automaticSimilarity({ width: 300, height: 200 });
    expect(photo.sizeToleranceX).toBe(15);
    expect(photo.sizeToleranceY).toBe(10);
    const large = automaticSimilarity({ width: 3000, height: 2000 });
    expect(large.sizeToleranceX).toBe(18);
    expect(large.sizeToleranceY).toBe(18);
  });
  it("자동 범위 안의 사진·도형을 찾고 명확히 다른 크기는 제외한다", () => {
    const base: Snapshot = {
      id: "a",
      slideId: "s",
      name: "a",
      type: "Image",
      left: 0,
      top: 0,
      width: 300,
      height: 200,
    };
    const close = {
      ...base,
      id: "b",
      type: "GeometricShape",
      width: 312,
      height: 207,
    };
    const different = { ...base, id: "c", width: 350, height: 220 };
    const auto = automaticSimilarity(base);
    expect(
      similar([base, close, different], base, {
        type: false,
        color: false,
        size: true,
        tolerance: auto.colorTolerance,
        sizeTolerance: 0,
        sizeToleranceX: auto.sizeToleranceX,
        sizeToleranceY: auto.sizeToleranceY,
      }).map((x) => x.id),
    ).toEqual(["a", "b"]);
  });
});
