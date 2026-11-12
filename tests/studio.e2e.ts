import { test, expect } from "@playwright/test";
import { readFileSync } from "node:fs";
const mock = readFileSync("tests/fixtures/office-host.js", "utf8");
async function connect(page: import("@playwright/test").Page) {
  await page.route("https://appsforoffice.microsoft.com/**", (route) =>
    route.fulfill({ contentType: "application/javascript", body: mock }),
  );
  await page.goto("/");
  await expect(page.getByText("PowerPoint 연결됨")).toBeVisible();
}
test("PowerPoint 밖에서는 편집을 막고 설치 안내를 제공한다", async ({
  page,
}) => {
  await page.route("https://appsforoffice.microsoft.com/**", (route) =>
    route.fulfill({ contentType: "application/javascript", body: "" }),
  );
  await page.setViewportSize({ width: 360, height: 900 });
  await page.goto("/");
  await expect(
    page.getByRole("button", { name: "삽입 미리보기" }),
  ).toBeDisabled();
  await page.getByRole("button", { name: /설치 안내 보기/ }).click();
  await expect(
    page.getByRole("heading", { name: "PowerPoint에서 시작하기" }),
  ).toBeVisible();
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= innerWidth,
    ),
  ).toBe(true);
});
test("텍스트와 기본 도형은 미리보기 승인 후에만 삽입한다", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await connect(page);
  await page.getByRole("button", { name: "삽입 미리보기" }).click();
  expect(await page.evaluate(() => (window as any).__officeTest.calls)).toEqual(
    [],
  );
  await expect(page.getByRole("dialog")).toContainText("안전한 현장");
  await page.getByRole("button", { name: "확인 후 적용" }).click();
  await expect(page.getByRole("status")).toContainText("완료");
  await page.getByRole("button", { name: /기본 도형 핵심/ }).click();
  await page.getByLabel("도형 종류").selectOption("Ellipse");
  await page.getByRole("button", { name: "삽입 미리보기" }).click();
  await page.getByRole("button", { name: "확인 후 적용" }).click();
  await expect(page.getByRole("status")).toContainText("원 삽입 완료");
  expect(await page.evaluate(() => (window as any).__officeTest.calls)).toEqual(
    ["TextBox", "Ellipse"],
  );
  expect(errors).toEqual([]);
});
test("승인 전 선택이 바뀌면 삭제를 중단하고 기존 개체를 보존한다", async ({
  page,
}) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    s.selected = ["original"];
  });
  await page.getByRole("button", { name: "선택 정보 새로고침" }).click();
  await page.getByRole("button", { name: "선택 개체", exact: true }).click();
  await page.getByRole("button", { name: "선택 개체 1개 삭제" }).click();
  await page.evaluate(() => {
    (window as any).__officeTest.selected = [];
  });
  await page.getByRole("button", { name: "확인 후 삭제" }).click();
  await expect(page.getByRole("alert")).toContainText(
    "선택 대상이 바뀌었습니다",
  );
  expect(
    await page.evaluate(() => (window as any).__officeTest.shapes.length),
  ).toBe(1);
});
test("선택 개체 위치 변경과 승인 삭제", async ({ page }) => {
  await connect(page);
  await page.evaluate(() => {
    (window as any).__officeTest.selected = ["original"];
  });
  await page.getByRole("button", { name: "선택 정보 새로고침" }).click();
  await page.getByRole("button", { name: "선택 개체", exact: true }).click();
  await page.getByRole("spinbutton", { name: "X 위치" }).fill("120");
  await page.getByRole("button", { name: /변경 내용 미리보기/ }).click();
  await page.getByRole("button", { name: "확인 후 적용" }).click();
  await expect(page.getByRole("status")).toContainText("변경 완료");
  expect(
    await page.evaluate(() => (window as any).__officeTest.shapes[0].left),
  ).toBe(120);
  await page.getByRole("button", { name: "선택 개체 1개 삭제" }).click();
  await page.getByRole("button", { name: "확인 후 삭제" }).click();
  await expect(page.getByRole("status")).toContainText("삭제 완료");
  expect(
    await page.evaluate(() => (window as any).__officeTest.shapes.length),
  ).toBe(0);
});
test("사진 파일 검증, 미리보기, 삽입", async ({ page }) => {
  await connect(page);
  await page.getByRole("button", { name: /사진 현장을/ }).click();
  await page
    .getByLabel("사진 파일 선택")
    .setInputFiles({
      name: "bad.svg",
      mimeType: "image/svg+xml",
      buffer: Buffer.from("<svg/>"),
    });
  await expect(page.getByRole("alert")).toContainText("PNG 또는 JPEG");
  await page
    .getByLabel("사진 파일 선택")
    .setInputFiles({
      name: "test.png",
      mimeType: "image/png",
      buffer: readFileSync("apps/powerpoint-addin/public/assets/icon-80.png"),
    });
  await expect(page.getByAltText("선택한 사진")).toBeVisible();
  await page.getByRole("button", { name: "삽입 미리보기" }).click();
  await page.getByRole("button", { name: "확인 후 적용" }).click();
  await expect(page.getByRole("status")).toContainText("사진 삽입 완료");
});
test("좁은 작업창과 데스크톱 UI 스크린샷", async ({ page }) => {
  await connect(page);
  await page.setViewportSize({ width: 360, height: 1000 });
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= innerWidth,
    ),
  ).toBe(true);
  await page.screenshot({ path: "artifacts/taskpane-360.png", fullPage: true });
  await page.setViewportSize({ width: 1280, height: 1100 });
  await page.screenshot({
    path: "artifacts/studio-desktop.png",
    fullPage: true,
  });
});
