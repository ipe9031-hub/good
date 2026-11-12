import { EditorTools } from "./EditorTools";
import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type ReactNode,
} from "react";
import {
  ArrowRight,
  Check,
  ChevronRight,
  CircleHelp,
  ImagePlus,
  Layers3,
  LoaderCircle,
  MousePointer2,
  PanelRight,
  Plus,
  RefreshCw,
  ShieldCheck,
  Square,
  Trash2,
  Type,
  X,
} from "lucide-react";
import {
  PowerPointService,
  friendlyError,
} from "../../../packages/powerpoint-adapter/src";
import {
  SHAPE_LIBRARY,
  targetOf,
  validateBounds,
  validateImage,
  type Bounds,
  type Selection,
  type ShapeKind,
  type Target,
} from "../../../packages/shared/src";
const service = new PowerPointService();
const initialBounds: Bounds = { left: 60, top: 80, width: 320, height: 100 };
type Pending = {
  title: string;
  detail: string;
  target: Target;
  destructive?: boolean;
  execute: () => Promise<void>;
  preview: ReactNode;
};
const kinds = SHAPE_LIBRARY;
export function App() {
  const [host, setHost] = useState<
    "loading" | "ready" | "browser" | "unsupported"
  >("loading");
  const [selection, setSelection] = useState<Selection | null>(null);
  const [tab, setTab] = useState<"insert" | "selection" | "tools" | "guide">(
    "insert",
  );
  const [tool, setTool] = useState<"text" | "shape" | "image">("text");
  const [text, setText] = useState("안전한 현장, 함께 만드는 내일");
  const [kind, setKind] = useState<ShapeKind>("Rectangle");
  const [color, setColor] = useState("#264F9E");
  const [bounds, setBounds] = useState<Bounds>(initialBounds);
  const [editBounds, setEditBounds] = useState<Bounds>(initialBounds);
  const [image, setImage] = useState<{ name: string; data: string } | null>(
    null,
  );
  const [pending, setPending] = useState<Pending | null>(null);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<{ error: boolean; text: string } | null>(
    null,
  );
  const mounted = useRef(true);
  const locked = useRef(false);
  const imageRead = useRef(0);
  const refresh = useCallback(async () => {
    try {
      const value = await service.read();
      if (mounted.current) {
        setSelection(value);
        if (value.shapes.length === 1)
          setEditBounds({
            left: value.shapes[0].left,
            top: value.shapes[0].top,
            width: value.shapes[0].width,
            height: value.shapes[0].height,
          });
      }
    } catch (e) {
      if (mounted.current) {
        setSelection(null);
        setNotice({ error: true, text: friendlyError(e) });
      }
    }
  }, []);
  useEffect(() => {
    mounted.current = true;
    let timer: ReturnType<typeof setTimeout>;
    let registered = false;
    const handler = () => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        void refresh();
      }, 200);
    };
    const timeout = setTimeout(() => {
      if (mounted.current) setHost((h) => (h === "loading" ? "browser" : h));
    }, 7000);
    if (typeof Office === "undefined") setHost("browser");
    else
      Office.onReady((info) => {
        if (!mounted.current) return;
        clearTimeout(timeout);
        if (info.host !== Office.HostType.PowerPoint) {
          setHost("browser");
          return;
        }
        if (
          !Office.context.requirements.isSetSupported("PowerPointApi", "1.5")
        ) {
          setHost("unsupported");
          return;
        }
        setHost("ready");
        void refresh();
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          handler,
          (result) => {
            registered = result.status === Office.AsyncResultStatus.Succeeded;
            if (!mounted.current && registered)
              Office.context.document.removeHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                { handler },
              );
          },
        );
      }).catch(() => {
        if (mounted.current) setHost("browser");
      });
    return () => {
      mounted.current = false;
      clearTimeout(timeout);
      clearTimeout(timer);
      if (registered)
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          { handler },
        );
    };
  }, [refresh]);
  const enabled = host === "ready" && !busy;
  const stage = (make: (s: Selection) => Pending) => {
    try {
      if (!selection)
        throw new Error(
          "선택 정보가 없습니다. PowerPoint에서 새로고침해 주세요.",
        );
      if (locked.current) return;
      const operation = make(selection);
      setNotice(null);
      if (operation.destructive) setPending(operation);
      else void executeOperation(operation);
    } catch (e) {
      setNotice({ error: true, text: friendlyError(e) });
    }
  };
  const stageInsert = () =>
    stage((s) => {
      const target = targetOf(s);
      const b = validateBounds({ ...bounds });
      const title =
        tool === "text"
          ? "텍스트 상자 삽입"
          : tool === "shape"
            ? `${kinds.find((k) => k.id === kind)?.label} 삽입`
            : "사진 삽입";
      if (tool === "text" && (!text.trim() || text.length > 5000))
        throw new Error("텍스트는 1~5,000자로 입력해 주세요.");
      if (tool === "image" && !image)
        throw new Error("사진을 먼저 선택해 주세요.");
      return {
        title,
        target,
        detail: `현재 선택 슬라이드에 새 개체 1개를 추가합니다. 기존 개체는 유지됩니다.`,
        execute: () =>
          tool === "text"
            ? service.insertText(target, text, b)
            : tool === "shape"
              ? service.insertShape(target, kind, color, b)
              : service.insertImage(target, image!.data, b),
        preview: (
          <div className="slide-preview">
            <span className="preview-label">삽입 내용</span>
            {tool === "text" ? (
              <p>{text}</p>
            ) : tool === "image" ? (
              <img src={image!.data} alt="삽입할 사진" />
            ) : (
              <div
                className={`shape-preview ${kind}`}
                style={{ backgroundColor: color }}
              />
            )}
            <small>
              X {b.left} · Y {b.top} · {b.width} × {b.height} pt
            </small>
          </div>
        ),
      };
    });
  const executeOperation = async (operation: Pending) => {
    if (locked.current) return;
    locked.current = true;
    setBusy(true);
    try {
      await operation.execute();
      setPending(null);
      setNotice({
        error: false,
        text: `${operation.title} 완료.`,
      });
      await refresh();
    } catch (e) {
      setPending(null);
      setNotice({ error: true, text: friendlyError(e) });
    } finally {
      locked.current = false;
      setBusy(false);
    }
  };
  async function loadImage(file?: File) {
    const request = ++imageRead.current;
    setImage(null);
    try {
      if (!file) return;
      validateImage(file);
      const data = await new Promise<string>((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(String(r.result));
        r.onerror = () => reject(new Error("사진을 읽을 수 없습니다."));
        r.readAsDataURL(file);
      });
      await new Promise<void>((resolve, reject) => {
        const i = new Image();
        i.onload = () => resolve();
        i.onerror = () =>
          reject(new Error("사진이 손상되었거나 지원하지 않는 형식입니다."));
        i.src = data;
      });
      if (request === imageRead.current) setImage({ name: file.name, data });
    } catch (e) {
      if (request === imageRead.current)
        setNotice({ error: true, text: friendlyError(e) });
    }
  }
  return (
    <div className="app-shell">
      <div className="workspace">
        <header>
          <div className="brand">
            <div>
              <h1>PPT AI Studio</h1>
            </div>
          </div>
          <span className={`connection ${host === "ready" ? "online" : ""}`}>
            <i />
            {host === "ready"
              ? "PowerPoint 연결됨"
              : host === "loading"
                ? "연결 확인 중"
                : "PowerPoint 연결 필요"}
          </span>
        </header>
        <main>
          {host !== "ready" && (
            <div className="host-message">
              <PanelRight size={19} />
              <div>
                <strong>
                  {host === "loading"
                    ? "PowerPoint 연결을 확인하고 있습니다"
                    : host === "unsupported"
                      ? "PowerPoint 업데이트가 필요합니다"
                      : "PowerPoint 안에서 열어 주세요"}
                </strong>
                <p>
                  {host === "unsupported"
                    ? "PowerPointApi 1.5 이상이 필요합니다."
                    : "Windows PowerPoint에 manifest.xml을 등록하면 편집 도구를 사용할 수 있습니다."}
                </p>
                <button className="text-button" onClick={() => setTab("guide")}>
                  설치 안내 보기 <ArrowRight size={13} />
                </button>
              </div>
            </div>
          )}
          <section className="document-card">
            <div className="document-icon">
              <Layers3 size={19} />
            </div>
            <div className="document-info">
              <span>현재 작업 문서</span>
              <strong>
                {selection?.documentName || "연결된 프레젠테이션 없음"}
              </strong>
              <p>
                {selection
                  ? `전체 ${selection.slideCount}장 · 선택 슬라이드 ${selection.slideIds.length}장 · 선택 개체 ${selection.shapes.length}개`
                  : "문서를 열고 슬라이드를 선택해 주세요"}
              </p>
            </div>
            <button
              className="icon-button"
              title="선택 정보 새로고침"
              aria-label="선택 정보 새로고침"
              disabled={!enabled}
              onClick={() => void refresh()}
            >
              <RefreshCw size={16} />
            </button>
          </section>
          <nav className="tabs" aria-label="편집 메뉴">
            {(
              [
                { id: "insert", label: "빠른 삽입", icon: Plus },
                { id: "selection", label: "선택 개체", icon: MousePointer2 },
                { id: "tools", label: "편집 도구", icon: Layers3 },
                { id: "guide", label: "사용 안내", icon: CircleHelp },
              ] as const
            ).map((t) => (
              <button
                key={t.id}
                className={tab === t.id ? "active" : ""}
                aria-current={tab === t.id ? "page" : undefined}
                onClick={() => setTab(t.id)}
              >
                <t.icon size={16} />
                {t.label}
              </button>
            ))}
          </nav>
          {notice && (
            <div
              role={notice.error ? "alert" : "status"}
              className={`notice ${notice.error ? "error" : "success"}`}
            >
              <span>{notice.error ? "!" : <Check size={17} />}</span>
              <p>{notice.text}</p>
              <button aria-label="알림 닫기" onClick={() => setNotice(null)}>
                <X size={15} />
              </button>
            </div>
          )}
          {tab === "insert" && (
            <section className="editor-card">
              <div className="section-heading">
                <div>
                  <h3>슬라이드에 추가</h3>
                  <p>추가할 요소를 선택하세요</p>
                </div>
                <span className="step">01 — INSERT</span>
              </div>
              <div className="tool-grid">
                {(
                  [
                    {
                      id: "text",
                      label: "텍스트",
                      subtitle: "메시지를 선명하게",
                      icon: Type,
                    },
                    {
                      id: "shape",
                      label: "기본 도형",
                      subtitle: "핵심을 시각적으로",
                      icon: Square,
                    },
                    {
                      id: "image",
                      label: "사진",
                      subtitle: "현장을 생생하게",
                      icon: ImagePlus,
                    },
                  ] as const
                ).map((t) => (
                  <button
                    key={t.id}
                    aria-pressed={tool === t.id}
                    className={tool === t.id ? "selected" : ""}
                    onClick={() => setTool(t.id)}
                  >
                    <t.icon size={23} />
                    <strong>{t.label}</strong>
                    <span>{t.subtitle}</span>
                    {tool === t.id && (
                      <i>
                        <Check size={10} />
                      </i>
                    )}
                  </button>
                ))}
              </div>
              {tool === "text" && (
                <label className="field">
                  텍스트 내용
                  <textarea
                    maxLength={5000}
                    value={text}
                    onChange={(e) => setText(e.target.value)}
                    placeholder="슬라이드에 넣을 내용을 입력하세요"
                  />
                  <span className="field-hint">
                    PowerPoint에서 직접 수정할 수 있는 텍스트로 삽입됩니다.
                    <span>{text.length} / 5,000</span>
                  </span>
                </label>
              )}
              {tool === "shape" && (
                <div className="shape-fields">
                  <label className="field">
                    도형 종류
                    <select
                      value={kind}
                      onChange={(e) => setKind(e.target.value as ShapeKind)}
                    >
                      {kinds.map((k) => (
                        <option key={k.id} value={k.id}>
                          {k.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="field">
                    채우기 색상
                    <div className="color-field">
                      <input
                        aria-label="채우기 색상"
                        type="color"
                        value={color}
                        onChange={(e) => setColor(e.target.value)}
                      />
                      <code>{color.toUpperCase()}</code>
                    </div>
                  </label>
                </div>
              )}
              {tool === "image" && (
                <label className="upload">
                  <ImagePlus size={26} />
                  <strong>{image?.name || "사진 파일 선택"}</strong>
                  <span>PNG, JPEG · 최대 5MB · 외부 전송 없음</span>
                  <input
                    aria-label="사진 파일 선택"
                    type="file"
                    accept="image/png,image/jpeg"
                    onChange={(e) => void loadImage(e.target.files?.[0])}
                  />
                  {image && <img src={image.data} alt="선택한 사진" />}
                </label>
              )}
              <BoundsFields value={bounds} onChange={setBounds} />
              <button
                className="primary full"
                disabled={
                  !enabled || !selection || (tool === "image" && !image)
                }
                onClick={stageInsert}
              >
                {busy ? "삽입 중…" : "삽입"} <ArrowRight size={17} />
              </button>
              <p className="safe-note">
                <ShieldCheck size={13} /> 현재 선택한 슬라이드에만 추가됩니다
              </p>
            </section>
          )}
          {tab === "selection" && (
            <section className="editor-card">
              <div className="section-heading">
                <div>
                  <h3>선택 개체 편집</h3>
                  <p>PowerPoint에서 개체를 선택한 후 새로고침하세요</p>
                </div>
                <span className="count-badge">
                  {selection?.shapes.length || 0}개
                </span>
              </div>
              {!selection?.shapes.length ? (
                <div className="empty">
                  <MousePointer2 size={30} />
                  <strong>아직 선택한 개체가 없어요</strong>
                  <p>슬라이드에서 텍스트, 도형 또는 사진을 선택해 주세요.</p>
                </div>
              ) : (
                <>
                  <ul className="object-list">
                    {selection.shapes.map((s) => (
                      <li key={s.id}>
                        <Square size={16} />
                        <div>
                          <strong>{s.name}</strong>
                          <small>
                            {s.type} · {Math.round(s.width)} ×{" "}
                            {Math.round(s.height)} pt
                          </small>
                        </div>
                      </li>
                    ))}
                  </ul>
                  {selection.shapes.length === 1 ? (
                    <>
                      <BoundsFields
                        value={editBounds}
                        onChange={setEditBounds}
                      />
                      <button
                        className="primary full"
                        disabled={!enabled}
                        onClick={() =>
                          stage((s) => {
                            const target = targetOf(s, true);
                            const b = validateBounds({ ...editBounds });
                            return {
                              title: "위치·크기 변경",
                              target,
                              detail:
                                "선택 개체 1개의 위치와 크기만 변경합니다.",
                              execute: () => service.update(target, b),
                              preview: (
                                <div className="change-preview">
                                  {Object.entries(b).map(([key, val]) => (
                                    <p key={key}>
                                      <span>
                                        {
                                          {
                                            left: "X",
                                            top: "Y",
                                            width: "너비",
                                            height: "높이",
                                          }[key]
                                        }
                                      </span>
                                      {Math.round(
                                        s.shapes[0][key as keyof Bounds],
                                      )}{" "}
                                      → <strong>{val} pt</strong>
                                    </p>
                                  ))}
                                </div>
                              ),
                            };
                          })
                        }
                      >
                        변경 적용 <ChevronRight size={16} />
                      </button>
                    </>
                  ) : (
                    <p className="field-hint">
                      위치·크기 변경은 개체 한 개씩 지원합니다.
                    </p>
                  )}
                  <button
                    className="danger full"
                    disabled={!enabled}
                    onClick={() =>
                      stage((s) => {
                        const target = targetOf(s, true);
                        return {
                          title: "선택 개체 삭제",
                          target,
                          destructive: true,
                          detail: `선택 슬라이드 1장에서 개체 ${target.shapeIds.length}개를 삭제합니다.`,
                          execute: () => service.remove(target),
                          preview: (
                            <ul className="delete-list">
                              {s.shapes.map((shape) => (
                                <li key={shape.id}>
                                  {shape.name} · {shape.type}
                                </li>
                              ))}
                            </ul>
                          ),
                        };
                      })
                    }
                  >
                    <Trash2 size={16} /> 선택 개체 {selection.shapes.length}개
                    삭제
                  </button>
                </>
              )}
            </section>
          )}
          {tab === "tools" && (
            <EditorTools
              available={enabled}
              selection={selection}
              onRefresh={refresh}
              onBusy={(value) => {
                locked.current = value;
                setBusy(value);
              }}
            />
          )}
          {tab === "guide" && (
            <section className="editor-card guide">
              <div className="section-heading">
                <div>
                  <h3>PowerPoint에서 시작하기</h3>
                  <p>Windows Microsoft 365 · Phase 1</p>
                </div>
                <PanelRight size={22} />
              </div>
              <ol>
                <li>
                  <strong>실행 방식 선택</strong>
                  <p>
                    배포형은 https://good-red.vercel.app에 연결하므로 Node.js나
                    개발 인증서가 필요 없습니다. 로컬 개발형은
                    Windows-Setup.cmd로 설치 후 Windows-Start.cmd를 실행합니다.
                  </p>
                </li>
                <li>
                  <strong>애드인 등록</strong>
                  <p>
                    배포형 설치 ZIP의 manifest.xml을 공유 폴더에 넣고
                    PowerPoint의 보안 센터 → 신뢰할 수 있는 추가 기능 카탈로그에
                    등록합니다. 자세한 절차는 docs/WINDOWS.md를 확인하세요.
                  </p>
                </li>
                <li>
                  <strong>리본에서 실행</strong>
                  <p>
                    추가 기능을 삽입한 뒤 PPT AI Studio 리본 버튼을 누릅니다.
                    슬라이드를 한 장 선택하고 편집을 시작하세요.
                  </p>
                </li>
              </ol>
              <div className="guide-note">
                <ShieldCheck size={18} />
                <p>
                  삽입과 위치 변경은 바로 적용됩니다. 삭제 전에만 대상을
                  확인합니다. 기본 실행 취소는 PowerPoint의 Ctrl+Z를 사용하세요.
                  애드인 자체 선택 복원은 후속 단계입니다.
                </p>
              </div>
              <h4>다음으로 준비 중</h4>
              <p>
                AI 생성과 번역은 후속 단계입니다. 정렬·사진 배치·일괄 편집은
                편집 도구에서 사용할 수 있습니다.
              </p>
              <p className="field-hint">
                Windows PowerPoint 실기 검증은 배포 문서의 체크리스트를 따라
                진행해 주세요.
              </p>
            </section>
          )}
          <footer>
            <span>
              <ShieldCheck size={13} /> 원본 파일을 먼저 저장해 주세요
            </span>
            <span>v0.2.1</span>
          </footer>
        </main>
      </div>
      {pending && (
        <div className="modal-backdrop">
          <section
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="confirm-title"
            onKeyDown={(e) => {
              if (e.key === "Escape" && !busy) setPending(null);
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
            <div className="modal-top">
              <span className="eyebrow">REVIEW & APPLY</span>
              <button
                aria-label="미리보기 닫기"
                disabled={busy}
                onClick={() => setPending(null)}
              >
                <X size={19} />
              </button>
            </div>
            <h3 id="confirm-title">{pending.title}</h3>
            <p>{pending.detail}</p>
            <div className="target-summary">
              대상 슬라이드 ID <code>{pending.target.slideId}</code>
            </div>
            {pending.preview}
            {pending.destructive && (
              <p className="delete-warning">
                삭제 전 파일을 저장해 주세요. 애드인 자체 복원은 지원하지
                않습니다.
              </p>
            )}
            <div className="modal-actions">
              <button
                autoFocus
                disabled={busy}
                className="secondary"
                onClick={() => setPending(null)}
              >
                취소
              </button>
              <button
                className={pending.destructive ? "danger-solid" : "primary"}
                disabled={busy}
                onClick={() => pending && void executeOperation(pending)}
              >
                {busy ? (
                  <>
                    <LoaderCircle className="spin" size={16} /> 적용 중
                  </>
                ) : pending.destructive ? (
                  "확인 후 삭제"
                ) : (
                  "확인 후 적용"
                )}
              </button>
            </div>
          </section>
        </div>
      )}
    </div>
  );
}
function BoundsFields({
  value,
  onChange,
}: {
  value: Bounds;
  onChange: (b: Bounds) => void;
}) {
  return (
    <fieldset className="bounds">
      <legend>
        위치 및 크기 <span>단위: pt</span>
      </legend>
      <div>
        {(
          [
            { key: "left", label: "X 위치" },
            { key: "top", label: "Y 위치" },
            { key: "width", label: "너비" },
            { key: "height", label: "높이" },
          ] as const
        ).map((f) => (
          <label key={f.key}>
            {f.label}
            <input
              aria-label={f.label}
              type="number"
              min={f.key === "width" || f.key === "height" ? 1 : 0}
              max={10000}
              step={1}
              value={Number.isNaN(value[f.key]) ? "" : value[f.key]}
              onChange={(e) =>
                onChange({
                  ...value,
                  [f.key]: e.target.value === "" ? NaN : Number(e.target.value),
                })
              }
            />
          </label>
        ))}
      </div>
    </fieldset>
  );
}
