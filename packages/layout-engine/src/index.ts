import { validateBounds, type Bounds } from "../../shared/src";
export type Canvas = { width: number; height: number };
export type Item = Bounds & { id: string };
export type LayoutAction =
  | "left"
  | "center"
  | "right"
  | "top"
  | "middle"
  | "bottom"
  | "horizontal"
  | "vertical"
  | "width"
  | "height"
  | "size"
  | "position";
export function union(items: Bounds[]): Bounds {
  if (!items.length) throw new Error("개체를 선택해 주세요.");
  const left = Math.min(...items.map((x) => x.left)),
    top = Math.min(...items.map((x) => x.top));
  return {
    left,
    top,
    width: Math.max(...items.map((x) => x.left + x.width)) - left,
    height: Math.max(...items.map((x) => x.top + x.height)) - top,
  };
}
export function inside(bounds: Bounds, canvas: Canvas): Bounds {
  if (!Number.isFinite(bounds.left) || !Number.isFinite(bounds.top))
    throw new Error("좌표를 숫자로 입력해 주세요.");
  const b = validateBounds({ ...bounds, left: 0, top: 0 });
  validateBounds({
    left: 0,
    top: 0,
    width: canvas.width,
    height: canvas.height,
  });
  if (b.width > canvas.width || b.height > canvas.height)
    throw new Error("개체가 슬라이드보다 큽니다. 크기를 먼저 줄여 주세요.");
  return {
    ...b,
    left: Math.min(Math.max(0, bounds.left), canvas.width - b.width),
    top: Math.min(Math.max(0, bounds.top), canvas.height - b.height),
  };
}
export function arrange(
  items: Item[],
  action: LayoutAction,
  canvas: Canvas,
  referenceId?: string,
): Item[] {
  if (items.length < 2) throw new Error("개체를 두 개 이상 선택해 주세요.");
  const reference = items.find((x) => x.id === referenceId) || items[0];
  if (action === "horizontal" || action === "vertical") {
    if (items.length < 3)
      throw new Error("간격 배분은 개체 세 개 이상이 필요합니다.");
    const axis = action === "horizontal" ? "left" : "top",
      size = axis === "left" ? "width" : "height";
    const sorted = [...items].sort((a, b) => a[axis] - b[axis]);
    const start = sorted[0][axis],
      end = Math.max(...sorted.map((x) => x[axis] + x[size]));
    const gap =
      (end - start - sorted.reduce((n, x) => n + x[size], 0)) /
      (items.length - 1);
    if (gap < 0)
      throw new Error("개체가 겹칩니다. 먼저 선택 영역을 넓혀 주세요.");
    let cursor = start;
    return sorted.map((x) => {
      const next = { ...x, [axis]: cursor };
      cursor += x[size] + gap;
      return { ...next, ...inside(next, canvas) };
    });
  }
  return items.map((x) => {
    const b = { ...x };
    switch (action) {
      case "left":
        b.left = reference.left;
        break;
      case "center":
        b.left = reference.left + (reference.width - x.width) / 2;
        break;
      case "right":
        b.left = reference.left + reference.width - x.width;
        break;
      case "top":
        b.top = reference.top;
        break;
      case "middle":
        b.top = reference.top + (reference.height - x.height) / 2;
        break;
      case "bottom":
        b.top = reference.top + reference.height - x.height;
        break;
      case "width":
        b.width = reference.width;
        break;
      case "height":
        b.height = reference.height;
        break;
      case "size":
        b.width = reference.width;
        b.height = reference.height;
        break;
      case "position":
        b.left = reference.left;
        b.top = reference.top;
        break;
    }
    return { ...b, ...inside(b, canvas) };
  });
}
export function grid(
  items: Item[],
  area: Bounds,
  columns: number,
  gap: number,
  keepRatio: boolean,
  canvas: Canvas,
): Item[] {
  validateBounds(area);
  if (
    !items.length ||
    !Number.isInteger(columns) ||
    columns < 1 ||
    columns > 10 ||
    !Number.isFinite(gap) ||
    gap < 0
  )
    throw new Error("열 수(1~10)와 간격(0 이상)을 확인해 주세요.");
  if (
    area.left + area.width > canvas.width ||
    area.top + area.height > canvas.height
  )
    throw new Error("배치 영역이 슬라이드 밖입니다.");
  const rows = Math.ceil(items.length / columns),
    w = (area.width - gap * (columns - 1)) / columns,
    h = (area.height - gap * (rows - 1)) / rows;
  if (w < 1 || h < 1) throw new Error("사진 간격이 배치 영역보다 큽니다.");
  return items.map((x, i) => {
    const scale = Math.min(w / x.width, h / x.height),
      width = keepRatio ? x.width * scale : w,
      height = keepRatio ? x.height * scale : h;
    return {
      ...x,
      ...validateBounds({
        left: area.left + (i % columns) * (w + gap) + (w - width) / 2,
        top: area.top + Math.floor(i / columns) * (h + gap) + (h - height) / 2,
        width,
        height,
      }),
    };
  });
}
export function swap(items: Item[], keepSize: boolean, canvas: Canvas): Item[] {
  if (items.length !== 2) throw new Error("사진을 정확히 두 개 선택해 주세요.");
  return items.map((x, i) => {
    const other = items[1 - i];
    return {
      ...x,
      ...inside(
        {
          left: other.left,
          top: other.top,
          width: keepSize ? x.width : other.width,
          height: keepSize ? x.height : other.height,
        },
        canvas,
      ),
    };
  });
}
export function normalizeColor(value: string): string {
  const named: Record<string, string> = {
    red: "#ff0000",
    white: "#ffffff",
    black: "#000000",
    blue: "#0000ff",
    navy: "#000080",
    green: "#008000",
    yellow: "#ffff00",
  };
  const s = value.toLowerCase();
  if (named[s]) return named[s];
  if (/^#?[a-f0-9]{6}$/.test(s)) return "#" + s.replace("#", "");
  if (/^#[a-f0-9]{3}$/.test(s))
    return "#" + [...s.slice(1)].map((x) => x + x).join("");
  return s;
}
export function colorDistance(a: string, b: string): number {
  a = normalizeColor(a);
  b = normalizeColor(b);
  if (a === b) return 0;
  if (!/^#[a-f0-9]{6}$/.test(a) || !/^#[a-f0-9]{6}$/.test(b)) return Infinity;
  return Math.sqrt(
    [1, 3, 5].reduce(
      (n, i) =>
        n +
        (parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16)) **
          2,
      0,
    ),
  );
}
export const toPoints = (value: number, unit: "pt" | "mm" | "cm") =>
  value * (unit === "mm" ? 72 / 25.4 : unit === "cm" ? 72 / 2.54 : 1);
export const fromPoints = (value: number, unit: "pt" | "mm" | "cm") =>
  value / (unit === "mm" ? 72 / 25.4 : unit === "cm" ? 72 / 2.54 : 1);
