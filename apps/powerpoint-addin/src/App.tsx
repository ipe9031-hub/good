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
  Layers3,
  LoaderCircle,
  MousePointer2,
  PanelRight,
  RefreshCw,
  ShieldCheck,
  Trash2,
  X,
} from "lucide-react";
import {
  PowerPointService,
  friendlyError,
} from "../../../packages/powerpoint-adapter/src";
import {
  targetOf,
  validateBounds,
  type Bounds,
  type Selection,
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
export function App() {
  const [host, setHost] = useState<
    "loading" | "ready" | "browser" | "unsupported"
  >("loading");
  const [selection, setSelection] = useState<Selection | null>(null);
  const [tab, setTab] = useState<"selection" | "tools" | "guide">("tools");
  const [editBounds, setEditBounds] = useState<Bounds>(initialBounds);
  const [pending, setPending] = useState<Pending | null>(null);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<{ error: boolean; text: string } | null>(
    null,
  );
  const mounted = useRef(true);
  const locked = useRef(false);
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
          <div className="selection-bar">
            <span>
              {selection
                ? `${selection.shapes.length}개 선택됨`
                : "선택 정보 없음"}
            </span>
            <button
              className="icon-button"
              title="선택 정보 새로고침"
              aria-label="선택 정보 새로고침"
              disabled={!enabled}
              onClick={() => void refresh()}
            >
              <RefreshCw size={15} />
            </button>
          </div>
          <nav className="tabs" aria-label="편집 메뉴">
            {(
              [
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
                  {selection.shapes.length === 1 ? (
                    <details className="option-box">
                      <summary>위치·크기 변경</summary>
                      <div className="option-content">
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
                      </div>
                    </details>
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
                          preview: null,
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
                  개체 추가는 PowerPoint 기본 메뉴를 사용하세요. 위치 변경은
                  바로 적용됩니다. 삭제 전에만 대상을 확인합니다. 기본 실행
                  취소는 PowerPoint의 Ctrl+Z를 사용하세요. 애드인 자체 선택
                  복원은 후속 단계입니다.
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
            <span>v0.2.3</span>
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
