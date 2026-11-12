import { useEffect, useRef, useState } from "react";
import { Check, LoaderCircle, RefreshCw, Trash2, X } from "lucide-react";
import {
  EditorService,
  similar,
  type Scan,
  type Snapshot,
  type Patch,
  type Plan,
  type Scope,
  type FontStyle,
} from "../../../packages/powerpoint-adapter/src/editor";
import { friendlyError } from "../../../packages/powerpoint-adapter/src";
import {
  arrange,
  grid,
  swap,
  union,
  inside,
  toPoints,
  normalizeColor,
  colorDistance,
  type Canvas,
  type LayoutAction,
} from "../../../packages/layout-engine/src";
import { type Selection } from "../../../packages/shared/src";
const editor = new EditorService();
const alignActions: { id: LayoutAction; label: string }[] = [
  { id: "left", label: "왼쪽" },
  { id: "center", label: "가로 중앙" },
  { id: "right", label: "오른쪽" },
  { id: "top", label: "위쪽" },
  { id: "middle", label: "세로 중앙" },
  { id: "bottom", label: "아래쪽" },
  { id: "horizontal", label: "가로 간격" },
  { id: "vertical", label: "세로 간격" },
  { id: "width", label: "너비 통일" },
  { id: "height", label: "높이 통일" },
  { id: "size", label: "크기 복제" },
  { id: "position", label: "위치 복제" },
];
export function EditorTools({
  available,
  selection,
  onBusy,
  onRefresh,
}: {
  available: boolean;
  selection: Selection | null;
  onBusy: (busy: boolean) => void;
  onRefresh: () => Promise<void>;
}) {
  const [mode, setMode] = useState<"layout" | "photos" | "similar" | "text">(
    "layout",
  );
  const [notice, setNotice] = useState<{ error: boolean; text: string } | null>(
    null,
  );
  const [busy, setBusy] = useState(false);
  const lock = useRef(false);
  const [reference, setReference] = useState("");
  const [canvas, setCanvas] = useState<Canvas>({ width: 960, height: 540 });
  const [unit, setUnit] = useState<"pt" | "mm" | "cm">("pt");
  const [dx, setDx] = useState(0),
    [dy, setDy] = useState(0);
  const [columns, setColumns] = useState(2),
    [gap, setGap] = useState(12),
    [keepRatio, setKeepRatio] = useState(true);
  const [deleteId, setDeleteId] = useState("");
  const [reflow, setReflow] = useState<"grid" | "pull" | "keep">("grid");
  const [scope, setScope] = useState<Scope>("selected");
  const [scan, setScan] = useState<Scan | null>(null);
  const [matches, setMatches] = useState<Snapshot[]>([]);
  const [matchType, setMatchType] = useState(true),
    [matchColor, setMatchColor] = useState(false),
    [matchSize, setMatchSize] = useState(false);
  const [tolerance, setTolerance] = useState(0),
    [sizeTolerance, setSizeTolerance] = useState(1);
  const [fill, setFill] = useState("#264F9E");
  const [oldColor, setOldColor] = useState("#ff0000");
  const [find, setFind] = useState(""),
    [replace, setReplace] = useState("");
  const [fontName, setFontName] = useState("맑은 고딕"),
    [fontSize, setFontSize] = useState(24),
    [fontColor, setFontColor] = useState("#172B4D");
  const [useFont, setUseFont] = useState(true),
    [useSize, setUseSize] = useState(true),
    [useColor, setUseColor] = useState(false);
  const [copied, setCopied] = useState<Snapshot | null>(null);
  const [confirm, setConfirm] = useState<Plan | null>(null);
  useEffect(() => {
    setScan(null);
    setMatches([]);
  }, [
    scope,
    matchType,
    matchColor,
    matchSize,
    tolerance,
    sizeTolerance,
    reference,
  ]);
  const disabled = !available || busy;
  async function run(action: () => Promise<string>) {
    if (lock.current) return;
    lock.current = true;
    setBusy(true);
    onBusy(true);
    setNotice(null);
    try {
      const text = await action();
      setNotice({ error: false, text });
      await onRefresh();
    } catch (e) {
      setNotice({ error: true, text: friendlyError(e) });
    } finally {
      lock.current = false;
      setBusy(false);
      onBusy(false);
    }
  }
  async function execute(plan: Plan) {
    await editor.apply(plan);
    setScan(null);
    setMatches([]);
    return `${plan.title} 완료 · ${plan.patches.length}개`;
  }
  async function selected() {
    const s = await editor.scan("selected");
    if (!s.items.length) throw new Error("개체를 선택해 주세요.");
    if (s.canvas) setCanvas(s.canvas);
    return s;
  }
  const canvasFor = (s: Scan) => s.canvas || canvas;
  async function layout(action: LayoutAction) {
    const s = await selected();
    const result = arrange(s.items, action, canvasFor(s), reference);
    return execute({
      title: alignActions.find((a) => a.id === action)!.label,
      scan: s,
      patches: result.map((x) => ({
        before: s.items.find((b) => b.id === x.id)!,
        bounds: x,
      })),
    });
  }
  async function photos() {
    const s = await selected();
    if (s.items.some((x) => x.type !== "Image"))
      throw new Error(
        "사진 개체만 선택해 주세요. 도형 채우기 이미지는 제외됩니다.",
      );
    return s;
  }
  async function arrangePhotos() {
    const s = await photos();
    const result = grid(
      s.items,
      union(s.items),
      columns,
      gap,
      keepRatio,
      canvasFor(s),
    );
    return execute({
      title: "사진 배치",
      scan: s,
      patches: result.map((x) => ({
        before: s.items.find((b) => b.id === x.id)!,
        bounds: x,
      })),
    });
  }
  async function stagePhotoDelete() {
    const s = await photos();
    if (s.items.length < 2)
      throw new Error("삭제할 사진과 재배치할 사진을 함께 선택해 주세요.");
    const removed = s.items.find((x) => x.id === deleteId) || s.items[0];
    const rest = s.items.filter((x) => x.id !== removed.id);
    let patches: Patch[] = rest.map((x) => ({ before: x }));
    if (reflow === "grid") {
      const result = grid(
        rest,
        union(s.items),
        Math.min(columns, rest.length),
        gap,
        keepRatio,
        canvasFor(s),
      );
      patches = result.map((x) => ({
        before: rest.find((b) => b.id === x.id)!,
        bounds: x,
      }));
    } else if (reflow === "pull") {
      const sorted = [...s.items].sort(
        (a, b) => a.top - b.top || a.left - b.left,
      );
      const remaining = sorted.filter((x) => x.id !== removed.id);
      patches = remaining.map((x, i) => ({
        before: x,
        bounds: inside(
          { ...x, left: sorted[i].left, top: sorted[i].top },
          canvasFor(s),
        ),
      }));
    }
    patches.push({ before: removed, remove: true });
    setConfirm({ title: "사진 삭제 및 재배치", scan: s, patches });
    return "삭제할 사진과 남은 사진의 배치를 확인해 주세요.";
  }
  async function search() {
    const s = await editor.scan(scope === "selected" ? "slide" : scope);
    const referenceShape = s.items.find(
      (x) =>
        x.id === (reference || selection?.shapes[0]?.id) &&
        x.slideId === s.target.slideId,
    );
    if (!referenceShape)
      throw new Error(
        "기준 개체 하나를 선택하고 현재 슬라이드 범위로 검색해 주세요.",
      );
    const found = similar(s.items, referenceShape, {
      type: matchType,
      color: matchColor,
      size: matchSize,
      tolerance,
      sizeTolerance,
    });
    setScan(s);
    setMatches(found);
    return `${found.length}개 찾았습니다. 검색된 개체만 변경됩니다.`;
  }
  async function batchText(replaceText: boolean) {
    const s = await editor.scan(scope);
    const font: FontStyle = {
      ...(useFont ? { name: fontName } : {}),
      ...(useSize ? { size: fontSize } : {}),
      ...(useColor ? { color: fontColor } : {}),
    };
    if (replaceText && !find) throw new Error("찾을 문구를 입력해 주세요.");
    if (!replaceText && !Object.keys(font).length)
      throw new Error("변경할 서식을 선택해 주세요.");
    const items = s.items.filter(
      (x) =>
        x.text !== undefined &&
        x.text.length > 0 &&
        (!replaceText || x.text.includes(find)),
    );
    const plan: Plan = {
      title: replaceText ? "텍스트 바꾸기" : "텍스트 서식 변경",
      scan: s,
      patches: items.map((x) => ({
        before: x,
        ...(replaceText
          ? { text: x.text!.split(find).join(replace) }
          : { font }),
      })),
    };
    if (!items.length) throw new Error("변경할 텍스트가 없습니다.");
    if (scope === "all") {
      setConfirm(plan);
      return `전체 슬라이드 ${items.length}개 텍스트 상자 변경을 확인해 주세요.`;
    }
    return execute(plan);
  }
  async function colorReplace() {
    const s = await editor.scan(scope);
    const patches: Patch[] = [];
    for (const item of s.items) {
      const p: Patch = { before: item };
      if (item.fill && colorDistance(item.fill, oldColor) <= tolerance)
        p.fill = fill;
      if (
        item.font?.color &&
        colorDistance(item.font.color, oldColor) <= tolerance
      )
        p.font = { color: fill };
      if (
        item.lineColor &&
        colorDistance(item.lineColor, oldColor) <= tolerance
      )
        p.lineColor = fill;
      if (p.fill || p.font || p.lineColor) patches.push(p);
    }
    if (!patches.length) throw new Error("일치하는 색상이 없습니다.");
    const plan = { title: "색상 일괄 교체", scan: s, patches };
    if (scope === "all") {
      setConfirm(plan);
      return `${patches.length}개 개체의 색상 교체를 확인해 주세요.`;
    }
    return execute(plan);
  }
  async function copyStyle() {
    const s = await selected();
    if (s.items.length !== 1)
      throw new Error("복사할 기준 개체 하나만 선택해 주세요.");
    setCopied(s.items[0]);
    return "기본 서식을 복사했습니다. 적용할 개체를 선택해 주세요.";
  }
  async function pasteStyle() {
    if (!copied) throw new Error("먼저 서식을 복사해 주세요.");
    const s = await selected();
    const patches = s.items
      .map((x) => ({
        before: x,
        ...(copied.fill &&
        ["GeometricShape", "TextBox", "Callout"].includes(x.type)
          ? { fill: normalizeColor(copied.fill) }
          : {}),
        ...(copied.lineColor &&
        ["GeometricShape", "TextBox", "Line"].includes(x.type)
          ? {
              lineColor: normalizeColor(copied.lineColor),
              ...(Number.isFinite(copied.lineWeight)
                ? { lineWeight: copied.lineWeight }
                : {}),
            }
          : {}),
        ...(copied.font && x.text !== undefined ? { font: copied.font } : {}),
      }))
      .filter((x) => Object.keys(x).length > 1);
    return execute({ title: "기본 서식 붙여넣기", scan: s, patches });
  }
  return (
    <section className="editor-card advanced-tools">
      <nav className="tool-modes" aria-label="기본 편집 도구">
        {(
          [
            { id: "layout", label: "정렬·크기" },
            { id: "photos", label: "사진 배치" },
            { id: "similar", label: "유사 검색" },
            { id: "text", label: "텍스트·색상" },
          ] as const
        ).map((t) => (
          <button
            key={t.id}
            aria-pressed={mode === t.id}
            className={mode === t.id ? "active" : ""}
            onClick={() => {
              setMode(t.id);
              setNotice(null);
            }}
          >
            {t.label}
          </button>
        ))}
      </nav>
      {notice && (
        <div
          role={notice.error ? "alert" : "status"}
          className={`notice ${notice.error ? "error" : "success"}`}
        >
          <p>{notice.text}</p>
          <button aria-label="편집 알림 닫기" onClick={() => setNotice(null)}>
            <X size={14} />
          </button>
        </div>
      )}
      {busy && (
        <p className="processing" role="status">
          <LoaderCircle className="spin" size={15} /> 처리 중…
        </p>
      )}
      {(mode === "layout" || mode === "photos") && (
        <>
          <div className="section-heading">
            <div>
              <h3>{mode === "layout" ? "정렬 및 크기 통일" : "사진 배치"}</h3>
              <p>현재 슬라이드의 선택 개체 {selection?.shapes.length || 0}개</p>
            </div>
          </div>
          <details className="canvas-settings">
            <summary>슬라이드 크기 설정</summary>
            <p>
              API가 지원하면 실제 크기를 자동으로 읽습니다. 지원하지 않는 버전은
              디자인 → 슬라이드 크기의 값을 pt로 입력하세요.
            </p>
            <div className="form-row">
              <label>
                너비 (pt)
                <input
                  type="number"
                  min={1}
                  value={canvas.width}
                  onChange={(e) =>
                    setCanvas({ ...canvas, width: Number(e.target.value) })
                  }
                />
              </label>
              <label>
                높이 (pt)
                <input
                  type="number"
                  min={1}
                  value={canvas.height}
                  onChange={(e) =>
                    setCanvas({ ...canvas, height: Number(e.target.value) })
                  }
                />
              </label>
            </div>
          </details>
        </>
      )}
      {mode === "layout" && (
        <>
          <label className="field">
            기준 개체
            <select
              value={reference}
              onChange={(e) => setReference(e.target.value)}
            >
              <option value="">첫 번째 선택 개체</option>
              {selection?.shapes.map((x) => (
                <option key={x.id} value={x.id}>
                  {x.name}
                </option>
              ))}
            </select>
          </label>
          <div className="action-grid">
            {alignActions.map((a) => (
              <button
                key={a.id}
                disabled={disabled}
                onClick={() => void run(() => layout(a.id))}
              >
                {a.label}
              </button>
            ))}
          </div>
          <h4>미세 이동</h4>
          <div className="form-row">
            <label>
              단위
              <select
                value={unit}
                onChange={(e) => setUnit(e.target.value as typeof unit)}
              >
                <option>pt</option>
                <option>mm</option>
                <option>cm</option>
              </select>
            </label>
            <label>
              가로 이동
              <input
                type="number"
                value={dx}
                onChange={(e) => setDx(Number(e.target.value))}
              />
            </label>
            <label>
              세로 이동
              <input
                type="number"
                value={dy}
                onChange={(e) => setDy(Number(e.target.value))}
              />
            </label>
          </div>
          <button
            className="secondary full"
            disabled={disabled}
            onClick={() =>
              void run(async () => {
                const s = await selected();
                return execute({
                  title: "미세 이동",
                  scan: s,
                  patches: s.items.map((x) => ({
                    before: x,
                    bounds: inside(
                      {
                        ...x,
                        left: x.left + toPoints(dx, unit),
                        top: x.top + toPoints(dy, unit),
                      },
                      canvasFor(s),
                    ),
                  })),
                });
              })
            }
          >
            선택 개체 이동
          </button>
          <h4>기본 서식 복제</h4>
          <p className="field-hint">
            단색 채우기·테두리 색/두께·균일한 글꼴 서식을 복사합니다. 내용과
            위치는 유지됩니다.
          </p>
          <div className="form-row">
            <button
              className="secondary"
              disabled={disabled}
              onClick={() => void run(copyStyle)}
            >
              서식 복사
            </button>
            <button
              className="secondary"
              disabled={disabled || !copied}
              onClick={() => void run(pasteStyle)}
            >
              서식 붙여넣기
            </button>
          </div>
          {copied && <p className="field-hint">복사한 개체: {copied.name}</p>}
        </>
      )}
      {mode === "photos" && (
        <>
          <div className="form-row">
            <label>
              열 수
              <input
                type="number"
                min={1}
                max={10}
                value={columns}
                onChange={(e) => setColumns(Number(e.target.value))}
              />
            </label>
            <label>
              간격 (pt)
              <input
                type="number"
                min={0}
                value={gap}
                onChange={(e) => setGap(Number(e.target.value))}
              />
            </label>
          </div>
          <label className="check-field">
            <input
              type="checkbox"
              checked={keepRatio}
              onChange={(e) => setKeepRatio(e.target.checked)}
            />
            현재 사진 비율 유지
          </label>
          <button
            className="primary full"
            disabled={disabled}
            onClick={() => void run(arrangePhotos)}
          >
            선택 사진 그리드 배치
          </button>
          <button
            className="secondary full spaced"
            disabled={disabled}
            onClick={() =>
              void run(async () => {
                const s = await photos();
                return execute({
                  title: "사진 위치 맞교환",
                  scan: s,
                  patches: swap(s.items, true, canvasFor(s)).map((x) => ({
                    before: s.items.find((b) => b.id === x.id)!,
                    bounds: x,
                  })),
                });
              })
            }
          >
            사진 2개 위치 맞교환
          </button>
          <p className="field-hint">
            사진 개체의 위치만 바꿉니다. 크기·사진 내용·주변 설명문은
            유지됩니다.
          </p>
          <h4>삭제 후 재배치</h4>
          <label className="field">
            삭제할 사진
            <select
              value={deleteId}
              onChange={(e) => setDeleteId(e.target.value)}
            >
              <option value="">첫 번째 선택 사진</option>
              {selection?.shapes
                .filter((x) => x.type === "Image")
                .map((x) => (
                  <option key={x.id} value={x.id}>
                    {x.name}
                  </option>
                ))}
            </select>
          </label>
          <label className="field spaced">
            남은 사진 배치
            <select
              value={reflow}
              onChange={(e) => setReflow(e.target.value as typeof reflow)}
            >
              <option value="grid">빈 공간 채우기 (그리드)</option>
              <option value="pull">기존 크기로 앞자리 당기기</option>
              <option value="keep">기존 위치 유지</option>
            </select>
          </label>
          <button
            className="danger full"
            disabled={disabled}
            onClick={() => void run(stagePhotoDelete)}
          >
            <Trash2 size={15} /> 삭제 대상 확인
          </button>
          <p className="field-hint">
            삭제할 사진과 남길 사진을 함께 선택하세요. 선택한 사진의 전체 영역
            안에서만 재배치합니다.
          </p>
        </>
      )}
      {(mode === "similar" || mode === "text") && (
        <label className="field">
          작업 범위
          <select
            value={scope}
            onChange={(e) => {
              setScope(e.target.value as Scope);
              setScan(null);
              setMatches([]);
            }}
          >
            <option value="selected">선택 개체</option>
            <option value="slide">현재 슬라이드</option>
            <option value="all">전체 프레젠테이션</option>
          </select>
        </label>
      )}
      {mode === "similar" && (
        <>
          <h4>선택한 첫 개체와 유사한 개체 찾기</h4>
          <p className="field-hint">
            ‘선택 개체’ 범위도 검색은 현재 슬라이드 전체에서 수행합니다. 개체
            유형은 사진·텍스트·도형 같은 큰 분류입니다. 사각형과 화살표의 세부
            형태는 구분하지 않습니다.
          </p>
          <label className="check-field">
            <input
              type="checkbox"
              checked={matchType}
              onChange={(e) => setMatchType(e.target.checked)}
            />
            같은 개체 유형
          </label>
          <label className="check-field">
            <input
              type="checkbox"
              checked={matchColor}
              onChange={(e) => setMatchColor(e.target.checked)}
            />
            같은/유사한 단색 채우기
          </label>
          <label className="check-field">
            <input
              type="checkbox"
              checked={matchSize}
              onChange={(e) => setMatchSize(e.target.checked)}
            />
            같은/유사한 크기
          </label>
          <div className="form-row">
            <label>
              색상 허용 차이 (0~442)
              <input
                type="number"
                min={0}
                max={442}
                value={tolerance}
                onChange={(e) => setTolerance(Number(e.target.value))}
              />
            </label>
            <label>
              크기 허용 차이 (pt)
              <input
                type="number"
                min={0}
                value={sizeTolerance}
                onChange={(e) => setSizeTolerance(Number(e.target.value))}
              />
            </label>
          </div>
          <button
            className="secondary full"
            disabled={disabled}
            onClick={() => void run(search)}
          >
            <RefreshCw size={14} />
            유사 개체 찾기
          </button>
          {scan && (
            <>
              <ResultList items={matches} />
              <div className="form-row">
                <button
                  className="secondary"
                  disabled={disabled || scope === "all" || !matches.length}
                  onClick={() =>
                    void run(async () => {
                      await editor.select(scan, matches);
                      return `${matches.length}개 선택 완료`;
                    })
                  }
                >
                  검색 결과 선택
                </button>
                <button
                  className="danger"
                  disabled={disabled || !matches.length}
                  onClick={() =>
                    setConfirm({
                      title: "검색 개체 삭제",
                      scan,
                      patches: matches.map((x) => ({
                        before: x,
                        remove: true,
                      })),
                    })
                  }
                >
                  검색 결과 삭제
                </button>
              </div>
              <label className="field spaced">
                변경할 채우기 색상
                <input
                  aria-label="변경할 채우기 색상"
                  type="color"
                  value={fill}
                  onChange={(e) => setFill(e.target.value)}
                />
              </label>
              <button
                className="secondary full spaced"
                disabled={disabled || !matches.some((x) => x.fill)}
                onClick={() => {
                  const plan = {
                    title: "검색 개체 색상 변경",
                    scan,
                    patches: matches
                      .filter((x) => x.fill)
                      .map((x) => ({ before: x, fill })),
                  };
                  if (scope === "all") setConfirm(plan);
                  else void run(() => execute(plan));
                }}
              >
                검색 결과 색상 변경
              </button>
            </>
          )}
        </>
      )}
      {mode === "text" && (
        <>
          <h4>글꼴 및 텍스트 서식</h4>
          <div className="form-row">
            <label>
              <span>
                <input
                  type="checkbox"
                  checked={useFont}
                  onChange={(e) => setUseFont(e.target.checked)}
                />
                글꼴
              </span>
              <input
                aria-label="글꼴 이름"
                value={fontName}
                onChange={(e) => setFontName(e.target.value)}
              />
            </label>
            <label>
              <span>
                <input
                  type="checkbox"
                  checked={useSize}
                  onChange={(e) => setUseSize(e.target.checked)}
                />
                크기 (pt)
              </span>
              <input
                aria-label="글자 크기"
                type="number"
                min={1}
                max={400}
                value={fontSize}
                onChange={(e) => setFontSize(Number(e.target.value))}
              />
            </label>
          </div>
          <label className="check-field">
            <input
              type="checkbox"
              checked={useColor}
              onChange={(e) => setUseColor(e.target.checked)}
            />
            글자 색상 변경
            <input
              aria-label="글자 색상"
              type="color"
              value={fontColor}
              onChange={(e) => setFontColor(e.target.value)}
            />
          </label>
          <button
            className="primary full"
            disabled={disabled}
            onClick={() => void run(() => batchText(false))}
          >
            텍스트 서식 적용
          </button>
          <h4>문구 찾기 및 바꾸기</h4>
          <label className="field">
            찾을 문구
            <input value={find} onChange={(e) => setFind(e.target.value)} />
          </label>
          <label className="field spaced">
            바꿀 문구
            <input
              value={replace}
              onChange={(e) => setReplace(e.target.value)}
            />
          </label>
          <button
            className="secondary full spaced"
            disabled={disabled}
            onClick={() => void run(() => batchText(true))}
          >
            문구 바꾸기
          </button>
          <p className="field-hint">
            대소문자를 구분합니다. 텍스트 상자 전체 문자열을 갱신하므로 문장
            중간의 서로 다른 서식은 유지되지 않을 수 있습니다.
          </p>
          <h4>색상 일괄 교체</h4>
          <div className="form-row">
            <label>
              기존 색
              <input
                type="color"
                aria-label="기존 색"
                value={oldColor}
                onChange={(e) => setOldColor(e.target.value)}
              />
            </label>
            <label>
              새 색
              <input
                type="color"
                aria-label="새 색"
                value={fill}
                onChange={(e) => setFill(e.target.value)}
              />
            </label>
            <label>
              허용 차이
              <input
                type="number"
                min={0}
                max={442}
                value={tolerance}
                onChange={(e) => setTolerance(Number(e.target.value))}
              />
            </label>
          </div>
          <button
            className="secondary full"
            disabled={disabled}
            onClick={() => void run(colorReplace)}
          >
            일치하는 색상 교체
          </button>
          <p className="field-hint">
            단색 채우기·테두리·균일한 글자 색상이 대상입니다. 그룹 내부·표
            셀·이미지 속 문자는 제외됩니다. 전체 프레젠테이션 변경에는 대상
            확인이 표시됩니다.
          </p>
        </>
      )}
      {confirm && (
        <div className="modal-backdrop">
          <section
            role="dialog"
            aria-modal="true"
            aria-labelledby="batch-confirm-title"
            className="modal"
            onKeyDown={(e) => {
              if (e.key === "Escape" && !busy) setConfirm(null);
              if (e.key === "Tab") {
                const buttons = [
                  ...e.currentTarget.querySelectorAll<HTMLButtonElement>(
                    "button:not(:disabled)",
                  ),
                ];
                if (e.shiftKey && document.activeElement === buttons[0]) {
                  e.preventDefault();
                  buttons.at(-1)?.focus();
                } else if (
                  !e.shiftKey &&
                  document.activeElement === buttons.at(-1)
                ) {
                  e.preventDefault();
                  buttons[0]?.focus();
                }
              }
            }}
          >
            <h3 id="batch-confirm-title">{confirm.title}</h3>
            <p>
              슬라이드{" "}
              {new Set(confirm.patches.map((x) => x.before.slideId)).size}장 ·
              개체 {confirm.patches.length}개
              {confirm.patches.some((x) => x.remove)
                ? ` · 삭제 ${confirm.patches.filter((x) => x.remove).length}개`
                : ""}
            </p>
            <ResultList
              items={confirm.patches.map((x) => ({
                ...x.before,
                name:
                  (x.remove ? "[삭제] " : x.bounds ? "[재배치] " : "[변경] ") +
                  x.before.name,
              }))}
            />
            <p className="field-hint">
              표시된 개체만 변경합니다. 원본을 저장한 후 실행하세요.
            </p>
            <div className="modal-actions">
              <button
                className="secondary"
                autoFocus
                disabled={busy}
                onClick={() => setConfirm(null)}
              >
                취소
              </button>
              <button
                className={
                  confirm.patches.some((x) => x.remove)
                    ? "danger-solid"
                    : "primary"
                }
                disabled={busy}
                onClick={() => {
                  const plan = confirm;
                  setConfirm(null);
                  void run(() => execute(plan));
                }}
              >
                확인 후 적용
              </button>
            </div>
          </section>
        </div>
      )}
    </section>
  );
}
function ResultList({ items }: { items: Snapshot[] }) {
  return (
    <div className="search-results">
      <strong>
        대상 {items.length}개 · {new Set(items.map((x) => x.slideId)).size}장
      </strong>
      <ul>
        {items.map((x) => (
          <li key={x.slideId + "-" + x.id}>
            <span>{x.name}</span>
            <small>
              {x.type} · 슬라이드 {x.slideId}
            </small>
          </li>
        ))}
      </ul>
    </div>
  );
}
