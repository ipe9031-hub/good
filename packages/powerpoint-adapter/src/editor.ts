import {
  assertSameTarget,
  validateBounds,
  type Bounds,
  type Target,
} from "../../shared/src";
import { colorDistance, type Canvas } from "../../layout-engine/src";
export type Scope = "selected" | "slide" | "all";
export type FontStyle = {
  name?: string;
  size?: number;
  color?: string;
  bold?: boolean;
};
export type Snapshot = Bounds & {
  id: string;
  slideId: string;
  name: string;
  type: string;
  text?: string;
  fill?: string;
  lineColor?: string;
  lineWeight?: number;
  font?: FontStyle;
};
export type ReadFields = {
  text: boolean;
  fill: boolean;
  line: boolean;
  canvas: boolean;
};
const defaultFields: ReadFields = {
  text: true,
  fill: true,
  line: true,
  canvas: true,
};
export type Scan = {
  fields?: ReadFields;
  scope: Scope;
  target: Target;
  items: Snapshot[];
  canvas: Canvas | null;
  excludedText: number;
};
export type Patch = {
  before: Snapshot;
  bounds?: Bounds;
  fill?: string;
  lineColor?: string;
  lineWeight?: number;
  font?: FontStyle;
  text?: string;
  remove?: boolean;
};
export type Plan = { title: string; scan: Scan; patches: Patch[] };
const supports = (version: string) =>
  Office.context.requirements.isSetSupported("PowerPointApi", version);
const textTypes = ["TextBox", "GeometricShape"];
const fillTypes = ["TextBox", "GeometricShape", "Callout"];
export const canSetFill = (type: string) => fillTypes.includes(type);
function readFont(font: PowerPoint.ShapeFont): FontStyle {
  return {
    ...(font.name ? { name: font.name } : {}),
    ...(font.size ? { size: font.size } : {}),
    ...(font.color ? { color: font.color } : {}),
    ...(font.bold !== null ? { bold: font.bold } : {}),
  };
}
async function inspect(
  context: PowerPoint.RequestContext,
  slides: PowerPoint.Slide[],
  scope: Scope,
  ids: string[],
  fields: ReadFields = defaultFields,
): Promise<{ items: Snapshot[]; excludedText: number }> {
  slides.forEach((s) =>
    s.shapes.load(
      "items/id,items/name,items/type,items/left,items/top,items/width,items/height",
    ),
  );
  await context.sync();
  const refs = slides.flatMap((s) =>
    s.shapes.items
      .filter((x) => scope !== "selected" || ids.includes(x.id))
      .map((x) => ({ shape: x, slideId: s.id })),
  );
  if (refs.length > 2000)
    throw new Error("한 번에 2,000개까지 처리합니다. 선택 범위를 줄여 주세요.");
  const frames = new Map<PowerPoint.Shape, PowerPoint.TextFrame>();
  for (const { shape } of refs) {
    if (fields.text && textTypes.includes(shape.type))
      frames.set(shape, shape.textFrame);
    else if (fields.text && shape.type === "Placeholder" && supports("1.10")) {
      const frame = shape.getTextFrameOrNullObject();

      frames.set(shape, frame);
    }
    if (fields.fill && fillTypes.includes(shape.type)) shape.fill.load("type");
    if (
      fields.line &&
      ["GeometricShape", "Line", "TextBox"].includes(shape.type)
    )
      shape.lineFormat.load("color,weight");
  }
  await context.sync();
  for (const { shape } of refs) {
    if (
      fields.fill &&
      fillTypes.includes(shape.type) &&
      shape.fill.type === "Solid"
    )
      shape.fill.load("foregroundColor");
  }
  for (const [shape, frame] of frames) {
    if (shape.type === "Placeholder" && frame.isNullObject) {
      frames.delete(shape);
      continue;
    }
    frame.textRange.load("text");
    frame.textRange.font.load("name,size,color,bold");
  }
  await context.sync();
  return {
    excludedText: refs.filter(
      ({ shape }) => shape.type === "Placeholder" && !frames.has(shape),
    ).length,
    items: refs.map(({ shape: s, slideId }) => ({
      id: s.id,
      slideId,
      name: s.name,
      type: s.type,
      left: s.left,
      top: s.top,
      width: s.width,
      height: s.height,
      ...(frames.has(s)
        ? {
            text: frames.get(s)!.textRange.text,
            font: readFont(frames.get(s)!.textRange.font),
          }
        : {}),
      ...(fields.fill && fillTypes.includes(s.type) && s.fill.type === "Solid"
        ? { fill: s.fill.foregroundColor }
        : {}),
      ...(fields.line && ["GeometricShape", "Line", "TextBox"].includes(s.type)
        ? { lineColor: s.lineFormat.color, lineWeight: s.lineFormat.weight }
        : {}),
    })),
  };
}
export class EditorService {
  async scan(
    scope: Scope,
    requested: Partial<ReadFields> = {},
    requireReference = false,
  ): Promise<Scan> {
    if (!supports("1.5"))
      throw new Error("PowerPointApi 1.5 이상이 필요합니다.");
    const fields = { ...defaultFields, ...requested };
    return PowerPoint.run(async (c) => {
      const selected = c.presentation.getSelectedSlides(),
        shapes = c.presentation.getSelectedShapes();
      selected.load("items/id");
      shapes.load("items/id");
      await c.sync();
      if (scope !== "all" && selected.items.length !== 1)
        throw new Error("슬라이드를 한 장 선택해 주세요.");
      const target = {
        slideId: selected.items.length === 1 ? selected.items[0].id : "",
        shapeIds: shapes.items.map((s) => s.id).sort(),
      };
      if (scope === "selected" && !target.shapeIds.length)
        throw new Error("개체를 먼저 선택해 주세요.");
      if (requireReference && (!target.slideId || target.shapeIds.length !== 1))
        throw new Error("슬라이드 한 장에서 기준 개체 한 개만 선택해 주세요.");
      let slides = selected.items;
      if (scope === "all") {
        const all = c.presentation.slides;
        all.load("items/id");
        await c.sync();
        slides = all.items;
      }
      let canvas: Canvas | null = null;
      if (fields.canvas && supports("1.10")) {
        const page = c.presentation.pageSetup;
        page.load("slideWidth,slideHeight");
        await c.sync();
        canvas = { width: page.slideWidth, height: page.slideHeight };
      }
      return {
        scope,
        fields,
        target,
        canvas,
        ...(await inspect(c, slides, scope, target.shapeIds, fields)),
      };
    });
  }
  async apply(plan: Plan) {
    if (!plan.patches.length) throw new Error("변경할 개체가 없습니다.");
    if (plan.patches.length > 2000) throw new Error("처리 범위를 줄여 주세요.");
    // Validate the complete patch set before sending any mutation to Office.
    for (const p of plan.patches) {
      if (p.bounds) validateBounds(p.bounds);
      if ((p.font || p.text !== undefined) && p.before.text === undefined)
        throw new Error(
          "이 대상에서 텍스트를 읽지 않았으므로 텍스트 변경을 적용할 수 없습니다.",
        );
      if (p.fill !== undefined && !fillTypes.includes(p.before.type))
        throw new Error("이 개체는 채우기 색상을 지원하지 않습니다.");
      if (
        p.font?.size !== undefined &&
        (!Number.isFinite(p.font.size) || p.font.size < 1 || p.font.size > 400)
      )
        throw new Error("글자 크기는 1~400pt로 입력해 주세요.");
      for (const color of [p.fill, p.lineColor, p.font?.color])
        if (color !== undefined && !/^#[a-f0-9]{6}$/i.test(color))
          throw new Error("색상은 #RRGGBB 형식으로 입력해 주세요.");
      if (
        p.font?.name !== undefined &&
        (!p.font.name.trim() || p.font.name.length > 100)
      )
        throw new Error("글꼴 이름을 확인해 주세요.");
    }
    await PowerPoint.run(async (c) => {
      const selected = c.presentation.getSelectedSlides(),
        shapes = c.presentation.getSelectedShapes();
      selected.load("items/id");
      shapes.load("items/id");
      await c.sync();
      if (plan.scan.scope !== "all")
        assertSameTarget(
          plan.scan.target,
          {
            slideId: selected.items.length === 1 ? selected.items[0].id : "",
            shapeIds: shapes.items.map((s) => s.id),
          },
          plan.scan.scope === "selected",
        );
      const slideIds = [...new Set(plan.patches.map((p) => p.before.slideId))];
      const slides = slideIds.map((id) => c.presentation.slides.getItem(id));
      const current = await inspect(
        c,
        slides,
        "slide",
        [],
        plan.scan.fields || defaultFields,
      );
      const key = (s: Snapshot) => s.slideId + "\0" + s.id;
      const actual = new Map(current.items.map((x) => [key(x), x]));
      for (const p of plan.patches) {
        const found = actual.get(key(p.before));
        if (!found || JSON.stringify(found) !== JSON.stringify(p.before))
          throw new Error(
            "대상 개체가 변경되었습니다. 다시 대상을 조회한 뒤 적용해 주세요.",
          );
      }
      for (const p of plan.patches) {
        const slide = slides.find((s) => s.id === p.before.slideId)!;
        const shape = slide.shapes.getItem(p.before.id);
        if (p.remove) {
          shape.delete();
          continue;
        }
        if (p.bounds) {
          shape.left = p.bounds.left;
          shape.top = p.bounds.top;
          shape.width = p.bounds.width;
          shape.height = p.bounds.height;
        }
        if (p.fill !== undefined) {
          if (!fillTypes.includes(p.before.type))
            throw new Error("이 개체는 채우기 색상을 지원하지 않습니다.");
          shape.fill.setSolidColor(p.fill);
        }
        if (p.lineColor !== undefined) shape.lineFormat.color = p.lineColor;
        if (p.lineWeight !== undefined) shape.lineFormat.weight = p.lineWeight;
        if (p.text !== undefined) shape.textFrame.textRange.text = p.text;
        if (p.font) {
          const f = shape.textFrame.textRange.font;
          if (p.font.name !== undefined) f.name = p.font.name;
          if (p.font.size !== undefined) f.size = p.font.size;
          if (p.font.color !== undefined) f.color = p.font.color;
          if (p.font.bold !== undefined) f.bold = p.font.bold;
        }
      }
      await c.sync();
    });
  }
  async select(scan: Scan, items: Snapshot[]) {
    if (scan.scope === "all")
      throw new Error("선택은 현재 슬라이드에서만 지원합니다.");
    await PowerPoint.run(async (c) => {
      const selected = c.presentation.getSelectedSlides();
      selected.load("items/id");
      await c.sync();
      if (
        selected.items.length !== 1 ||
        selected.items[0].id !== scan.target.slideId
      )
        throw new Error("선택 슬라이드가 바뀌었습니다.");
      selected.items[0].setSelectedShapes(items.map((x) => x.id));
      await c.sync();
    });
  }
}
export function similar(
  items: Snapshot[],
  reference: Snapshot,
  criteria: {
    type: boolean;
    color: boolean;
    size: boolean;
    tolerance: number;
    sizeTolerance: number;
    sizeToleranceX?: number;
    sizeToleranceY?: number;
  },
): Snapshot[] {
  if (!criteria.type && !criteria.color && !criteria.size)
    throw new Error("검색 조건을 하나 이상 선택해 주세요.");
  if (criteria.color && !reference.fill)
    throw new Error("기준 개체가 단색 채우기가 아닙니다.");
  return items.filter(
    (x) =>
      (!criteria.type || x.type === reference.type) &&
      (!criteria.color ||
        (!!x.fill &&
          colorDistance(x.fill, reference.fill!) <= criteria.tolerance)) &&
      (!criteria.size ||
        (Math.abs(x.width - reference.width) <=
          (criteria.sizeToleranceX ?? criteria.sizeTolerance) &&
          Math.abs(x.height - reference.height) <=
            (criteria.sizeToleranceY ?? criteria.sizeTolerance))),
  );
}
