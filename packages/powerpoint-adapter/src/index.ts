import {
  assertSameTarget,
  isImageData,
  targetOf,
  validateBounds,
  type Bounds,
  type Selection,
  type ShapeKind,
  type Target,
} from "../../shared/src";

export class PowerPointService {
  private guard() {
    if (
      typeof Office === "undefined" ||
      !Office.context?.requirements?.isSetSupported("PowerPointApi", "1.5")
    )
      throw new Error(
        "PowerPointApi 1.5 이상을 지원하는 Windows Microsoft 365 PowerPoint가 필요합니다. Office를 업데이트해 주세요.",
      );
  }
  async read(): Promise<Selection> {
    this.guard();
    return PowerPoint.run(async (context) => {
      const all = context.presentation.slides;
      const slides = context.presentation.getSelectedSlides();
      const shapes = context.presentation.getSelectedShapes();
      all.load("items/id");
      slides.load("items/id");
      shapes.load(
        "items/id,items/name,items/type,items/left,items/top,items/width,items/height",
      );
      await context.sync();
      const url = Office.context.document.url;
      const documentName = url
        ? decodeURIComponent(
            url.split(/[\\/]/).pop()?.split("?")[0] || "현재 프레젠테이션",
          )
        : "현재 프레젠테이션 (저장 전 또는 이름 미제공)";
      return {
        documentName,
        slideCount: all.items.length,
        slideIds: slides.items.map((s) => s.id),
        shapes: shapes.items.map((s) => ({
          id: s.id,
          name: s.name,
          type: s.type,
          left: s.left,
          top: s.top,
          width: s.width,
          height: s.height,
        })),
      };
    });
  }
  private async mutate(
    target: Target,
    requireShapes: boolean,
    action: (
      slide: PowerPoint.Slide,
      context: PowerPoint.RequestContext,
    ) => void,
  ) {
    this.guard();
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      const shapes = context.presentation.getSelectedShapes();
      slides.load("items/id");
      shapes.load("items/id");
      await context.sync();
      const current = {
        slideId: slides.items.length === 1 ? slides.items[0].id : "",
        shapeIds: shapes.items.map((s) => s.id),
      };
      assertSameTarget(target, current, requireShapes);
      action(slides.items[0], context);
      await context.sync();
    });
  }
  async insertText(target: Target, text: string, bounds: Bounds) {
    validateBounds(bounds);
    if (!text.trim() || text.length > 5000)
      throw new Error("텍스트는 1~5,000자로 입력해 주세요.");
    await this.mutate(target, false, (slide) => {
      const shape = slide.shapes.addTextBox(text, bounds);
      shape.textFrame.textRange.font.size = 24;
      shape.textFrame.textRange.font.color = "#172B4D";
    });
  }
  async insertShape(
    target: Target,
    kind: ShapeKind,
    color: string,
    bounds: Bounds,
  ) {
    validateBounds(bounds);
    if (
      !["Rectangle", "Ellipse", "Triangle", "RightArrow"].includes(kind) ||
      !/^#[0-9a-f]{6}$/i.test(color)
    )
      throw new Error("지원하는 도형과 색상을 선택해 주세요.");
    await this.mutate(target, false, (slide) => {
      const shape = slide.shapes.addGeometricShape(kind, bounds);
      shape.fill.setSolidColor(color);
    });
  }
  async update(target: Target, bounds: Bounds) {
    validateBounds(bounds);
    if (target.shapeIds.length !== 1)
      throw new Error("위치·크기는 개체 한 개를 선택해서 변경해 주세요.");
    await this.mutate(target, true, (slide) => {
      const shape = slide.shapes.getItem(target.shapeIds[0]);
      shape.left = bounds.left;
      shape.top = bounds.top;
      shape.width = bounds.width;
      shape.height = bounds.height;
    });
  }
  async remove(target: Target) {
    if (!target.shapeIds.length)
      throw new Error("삭제할 개체를 선택해 주세요.");
    await this.mutate(target, true, (slide) =>
      target.shapeIds.forEach((id) => slide.shapes.getItem(id).delete()),
    );
  }
  async insertImage(target: Target, data: string, bounds: Bounds) {
    this.guard();
    validateBounds(bounds);
    if (!isImageData(data) || data.length > 7 * 1024 * 1024)
      throw new Error("유효한 5MB 이하 PNG/JPEG 이미지가 필요합니다.");
    if (!Office.context.requirements.isSetSupported("ImageCoercion", "1.1"))
      throw new Error("이 PowerPoint 버전은 사진 삽입을 지원하지 않습니다.");
    assertSameTarget(target, targetOf(await this.read()), false);
    // Common API inserts a new picture on the active slide; no replacement or upload.
    await new Promise<void>((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        data.split(",")[1],
        {
          coercionType: Office.CoercionType.Image,
          imageLeft: bounds.left,
          imageTop: bounds.top,
          imageWidth: bounds.width,
          imageHeight: bounds.height,
        },
        (result) =>
          result.status === Office.AsyncResultStatus.Succeeded
            ? resolve()
            : reject(new Error(result.error.message)),
      );
    });
  }
}
export function friendlyError(error: unknown) {
  const code = (error as { code?: string })?.code;
  if (code === "InvalidSelection" || code === "ItemNotFound")
    return "선택한 개체를 찾을 수 없습니다. 슬라이드에서 다시 선택해 주세요.";
  if (code === "AccessDenied")
    return "편집 권한이 없습니다. 보호된 보기와 읽기 전용 설정을 확인해 주세요.";
  if (code === "GeneralException")
    return "PowerPoint가 작업을 완료하지 못했습니다. 편집 모드에서 다시 시도해 주세요.";
  return error instanceof Error
    ? error.message
    : "작업 중 오류가 발생했습니다. 선택 상태를 확인하고 다시 시도해 주세요.";
}
