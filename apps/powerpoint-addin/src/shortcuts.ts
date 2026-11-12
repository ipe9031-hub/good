export const SHORTCUTS = [
  { command: "find", id: "PPTAI.Find", code: "KeyF", key: "F", label: "찾기" },
  {
    command: "size",
    id: "PPTAI.Size",
    code: "KeyC",
    key: "C",
    label: "크기 복제",
  },
  {
    command: "left",
    id: "PPTAI.Left",
    code: "Digit1",
    key: "1",
    label: "왼쪽 정렬",
  },
  {
    command: "center",
    id: "PPTAI.Center",
    code: "Digit2",
    key: "2",
    label: "가로 중앙",
  },
  {
    command: "right",
    id: "PPTAI.Right",
    code: "Digit3",
    key: "3",
    label: "오른쪽 정렬",
  },
  {
    command: "top",
    id: "PPTAI.Top",
    code: "Digit4",
    key: "4",
    label: "위쪽 정렬",
  },
  {
    command: "middle",
    id: "PPTAI.Middle",
    code: "Digit5",
    key: "5",
    label: "세로 중앙",
  },
  {
    command: "bottom",
    id: "PPTAI.Bottom",
    code: "Digit6",
    key: "6",
    label: "아래쪽 정렬",
  },
  {
    command: "horizontal",
    id: "PPTAI.Horizontal",
    code: "Digit7",
    key: "7",
    label: "가로 간격 배분",
  },
  {
    command: "vertical",
    id: "PPTAI.Vertical",
    code: "Digit8",
    key: "8",
    label: "세로 간격 배분",
  },
] as const;
export type ShortcutCommand = (typeof SHORTCUTS)[number]["command"];
export type ShortcutRequest = { id: number; command: ShortcutCommand };
let receiver: ((command: ShortcutCommand) => void) | null = null;
let queued: ShortcutCommand | null = null;
let globalRegistered = false;
export function subscribeShortcuts(
  callback: (command: ShortcutCommand) => void,
) {
  receiver = callback;
  if (queued) {
    const command = queued;
    queued = null;
    callback(command);
  }
  return () => {
    if (receiver === callback) receiver = null;
  };
}
export function dispatchShortcut(command: ShortcutCommand) {
  if (receiver) receiver(command);
  else queued = command;
}
export function decodeShortcut(
  event: Pick<
    KeyboardEvent,
    | "ctrlKey"
    | "altKey"
    | "shiftKey"
    | "metaKey"
    | "code"
    | "repeat"
    | "isComposing"
  >,
): ShortcutCommand | null {
  if (
    !event.ctrlKey ||
    !event.altKey ||
    !event.shiftKey ||
    event.metaKey ||
    event.repeat ||
    event.isComposing
  )
    return null;
  return SHORTCUTS.find((x) => x.code === event.code)?.command || null;
}
export function isTypingTarget(target: EventTarget | null) {
  return (
    target instanceof Element &&
    !!target.closest(
      'input,textarea,select,[contenteditable="true"],[contenteditable=""],[role="textbox"]',
    )
  );
}
export function installPaneShortcuts() {
  const handler = (event: KeyboardEvent) => {
    if (event.defaultPrevented || isTypingTarget(event.target)) return;
    const command = decodeShortcut(event);
    if (command) {
      event.preventDefault();
      event.stopPropagation();
      dispatchShortcut(command);
    }
  };
  window.addEventListener("keydown", handler);
  return () => window.removeEventListener("keydown", handler);
}
export const globalShortcutsRegistered = () => globalRegistered;
export async function registerOfficeShortcuts(
  allowGlobal = new URLSearchParams(location.search).get("keyboard") === "1",
) {
  globalRegistered = false;
  if (!allowGlobal || typeof Office === "undefined") return false;
  const info = await Office.onReady();
  if (
    info.host !== Office.HostType.PowerPoint ||
    !Office.context.requirements.isSetSupported("SharedRuntime", "1.1") ||
    !Office.actions?.associate ||
    !Office.addin?.showAsTaskpane
  )
    return false;
  for (const shortcut of SHORTCUTS) {
    Office.actions.associate(shortcut.id, async () => {
      if (
        typeof document !== "undefined" &&
        document.hasFocus() &&
        isTypingTarget(document.activeElement)
      )
        return;
      await Office.addin.showAsTaskpane();
      dispatchShortcut(shortcut.command);
    });
  }
  globalRegistered = true;
  return true;
}
