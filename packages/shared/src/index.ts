export type Bounds = {
  left: number;
  top: number;
  width: number;
  height: number;
};
export type SelectedShape = Bounds & { id: string; name: string; type: string };
export type Selection = {
  documentName: string;
  slideCount: number;
  slideIds: string[];
  shapes: SelectedShape[];
};
export type Target = { slideId: string; shapeIds: string[] };
export const SHAPE_LIBRARY = [
  { id: "Rectangle", label: "사각형" },
  { id: "RoundRectangle", label: "둥근 사각형" },
  { id: "Ellipse", label: "원" },
  { id: "Triangle", label: "삼각형" },
  { id: "Diamond", label: "마름모" },
  { id: "RightArrow", label: "오른쪽 화살표" },
  { id: "LeftRightArrow", label: "양방향 화살표" },
  { id: "DownArrow", label: "아래 화살표" },
  { id: "BentArrow", label: "꺾은 화살표" },
  { id: "CurvedRightArrow", label: "곡선 화살표" },
  { id: "WedgeRRectCallout", label: "말풍선" },
  { id: "Frame", label: "강조 프레임" },
  { id: "Plus", label: "십자 표시" },
  { id: "Hexagon", label: "육각형" },
  { id: "Line", label: "직선" },
] as const;
export type ShapeKind = (typeof SHAPE_LIBRARY)[number]["id"];
export const MAX_IMAGE_BYTES = 5 * 1024 * 1024;
export function validateBounds(input: Bounds): Bounds {
  const b: Bounds = {
    left: input.left,
    top: input.top,
    width: input.width,
    height: input.height,
  };
  if (Object.values(b).some((v) => !Number.isFinite(v)))
    throw new Error("좌표와 크기는 숫자로 입력해 주세요.");
  if (
    b.left < 0 ||
    b.top < 0 ||
    b.width < 1 ||
    b.height < 1 ||
    Object.values(b).some((v) => v > 10000)
  )
    throw new Error(
      "위치는 0 이상, 크기는 1 이상이어야 하며 최대 10,000pt까지 입력할 수 있습니다.",
    );
  return b;
}
export function targetOf(s: Selection, requireShapes = false): Target {
  if (s.slideIds.length !== 1)
    throw new Error("작업할 슬라이드를 한 장만 선택해 주세요.");
  if (requireShapes && !s.shapes.length)
    throw new Error("PowerPoint에서 개체를 먼저 선택해 주세요.");
  return { slideId: s.slideIds[0], shapeIds: s.shapes.map((x) => x.id).sort() };
}
export function assertSameTarget(
  expected: Target,
  actual: Target,
  shapes: boolean,
) {
  if (
    expected.slideId !== actual.slideId ||
    (shapes &&
      JSON.stringify([...expected.shapeIds].sort()) !==
        JSON.stringify([...actual.shapeIds].sort()))
  )
    throw new Error(
      "선택 대상이 바뀌었습니다. 선택 정보를 새로고침한 후 다시 실행해 주세요.",
    );
}
export function validateImage(file: Pick<File, "size" | "type">) {
  if (!["image/png", "image/jpeg"].includes(file.type))
    throw new Error("PNG 또는 JPEG 사진만 선택해 주세요.");
  if (file.size <= 0 || file.size > MAX_IMAGE_BYTES)
    throw new Error("사진은 0바이트 초과, 5MB 이하로 선택해 주세요.");
}
export function isImageData(data: string) {
  return (
    /^data:image\/png;base64,iVBORw0KGgo/.test(data) ||
    /^data:image\/jpeg;base64,\/9j\//.test(data)
  );
}
