import { beforeEach, afterEach, it, expect, vi } from "vitest";
import {
  SHORTCUTS,
  decodeShortcut,
  dispatchShortcut,
  subscribeShortcuts,
  registerOfficeShortcuts,
  globalShortcutsRegistered,
} from "../apps/powerpoint-addin/src/shortcuts";
let unsubscribe: () => void;
let receiver: ReturnType<typeof vi.fn>;
let actions: Record<string, () => Promise<void>>;
let show: ReturnType<typeof vi.fn>;
beforeEach(() => {
  receiver = vi.fn();
  unsubscribe = subscribeShortcuts(receiver);
  actions = {};
  show = vi.fn(async () => {});
  vi.stubGlobal("Office", {
    HostType: { PowerPoint: "PowerPoint" },
    onReady: async () => ({ host: "PowerPoint" }),
    context: { requirements: { isSetSupported: () => true } },
    actions: {
      associate: (id: string, fn: () => Promise<void>) => {
        actions[id] = fn;
      },
    },
    addin: { showAsTaskpane: show },
  });
});
afterEach(() => {
  unsubscribe();
  vi.unstubAllGlobals();
});
it("안전한 수정키 조합과 물리 키를 해석한다", () => {
  const event = {
    ctrlKey: true,
    altKey: true,
    shiftKey: true,
    metaKey: false,
    code: "KeyF",
    repeat: false,
    isComposing: false,
  };
  expect(decodeShortcut(event)).toBe("find");
  expect(decodeShortcut({ ...event, code: "Digit7" })).toBe("horizontal");
  expect(decodeShortcut({ ...event, shiftKey: false })).toBeNull();
  expect(decodeShortcut({ ...event, repeat: true })).toBeNull();
  expect(decodeShortcut({ ...event, isComposing: true })).toBeNull();
});
it("화면 준비 전 요청을 한 번 전달한다", () => {
  unsubscribe();
  dispatchShortcut("find");
  const next = vi.fn();
  unsubscribe = subscribeShortcuts(next);
  expect(next).toHaveBeenCalledExactlyOnceWith("find");
});
it("공식 Office 동작 ID에 모든 키를 연결하고 작업창을 연 후 요청한다", async () => {
  expect(await registerOfficeShortcuts(true)).toBe(true);
  expect(Object.keys(actions)).toEqual(SHORTCUTS.map((x) => x.id));
  await actions["PPTAI.Find"]();
  expect(show).toHaveBeenCalledOnce();
  expect(receiver).toHaveBeenCalledExactlyOnceWith("find");
  expect(globalShortcutsRegistered()).toBe(true);
});
it("기본 등록에서는 전역 키를 연결하지 않는다", async () => {
  expect(await registerOfficeShortcuts(false)).toBe(false);
  expect(Object.keys(actions)).toEqual([]);
});
it("지원하지 않는 호스트에서 전역 지원으로 표시하지 않는다", async () => {
  Office.context.requirements.isSetSupported = () => false;
  expect(await registerOfficeShortcuts(true)).toBe(false);
  expect(globalShortcutsRegistered()).toBe(false);
  expect(Object.keys(actions)).toEqual([]);
});
