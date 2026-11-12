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
    page.getByRole("button", { name: "삽입", exact: true }),
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
test("텍스트와 기본 도형은 확인창 없이 한 번에 삽입한다", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await connect(page);
  await page.getByRole("button", { name: "삽입", exact: true }).click();
  await expect(page.getByRole("dialog")).toHaveCount(0);
  await expect(page.getByRole("status")).toContainText("완료");
  await page.getByRole("button", { name: /기본 도형 핵심/ }).click();
  await page.getByLabel("도형 종류").selectOption("Ellipse");
  await page.getByRole("button", { name: "삽입", exact: true }).click();
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
  await page.getByRole("button", { name: /변경 적용/ }).click();
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
  await page.getByLabel("사진 파일 선택").setInputFiles({
    name: "bad.svg",
    mimeType: "image/svg+xml",
    buffer: Buffer.from("<svg/>"),
  });
  await expect(page.getByRole("alert")).toContainText("PNG 또는 JPEG");
  await page.getByLabel("사진 파일 선택").setInputFiles({
    name: "test.png",
    mimeType: "image/png",
    buffer: readFileSync("apps/powerpoint-addin/public/assets/icon-80.png"),
  });
  await expect(page.getByAltText("선택한 사진")).toBeVisible();
  await page.getByRole("button", { name: "삽입", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("사진 삽입 완료");
});
test("좁은 작업창과 데스크톱 UI 스크린샷", async ({ page }) => {
  await connect(page);
  await page.setViewportSize({ width: 360, height: 1000 });
  await expect(page.getByText("좋은 장표의 시작,")).toHaveCount(0);
  await expect(page.getByText("아이디어를 슬라이드 위로")).toHaveCount(0);
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
  await page.setViewportSize({width:360,height:1000});
  await page.getByRole('button',{name:'편집 도구',exact:true}).click();
  await page.getByRole('button',{name:'텍스트·색상',exact:true}).click();
  expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBe(true);
  await page.screenshot({path:'artifacts/phase2-tools-360.png',fullPage:true});
});
test("편집 도구에서 선택 개체를 바로 정렬하고 크기를 통일한다", async ({
  page,
}) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    const b = s.seed("Rectangle", {
      left: 300,
      top: 80,
      width: 200,
      height: 100,
    });
    s.selected = ["original", b.id];
  });
  await page.getByRole("button", { name: "선택 정보 새로고침" }).click();
  await page.getByRole("button", { name: "편집 도구", exact: true }).click();
  await page.getByRole("button", { name: "왼쪽", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("왼쪽 완료");
  await expect(page.getByRole("dialog")).toHaveCount(0);
  expect(
    await page.evaluate(() => (window as any).__officeTest.shapes[1].left),
  ).toBe(10);
  await page.getByRole("button", { name: "크기 복제", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("크기 복제 완료");
  expect(
    await page.evaluate(() => (window as any).__officeTest.shapes[1].width),
  ).toBe(100);
});
test("사진 그리드와 위치 교환, 승인 삭제 후 재배치에서 제목을 보존한다", async ({
  page,
}) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    const a = s.seed("Image", { left: 20, top: 100, width: 200, height: 100 }),
      b = s.seed("Image", { left: 400, top: 100, width: 200, height: 100 }),
      c = s.seed("Image", { left: 20, top: 300, width: 200, height: 100 });
    s.selected = [a.id, b.id, c.id];
  });
  await page.getByRole("button", { name: "선택 정보 새로고침" }).click();
  await page.getByRole("button", { name: "편집 도구", exact: true }).click();
  await page.getByRole("button", { name: "사진 배치", exact: true }).click();
  await page.getByRole("button", { name: "선택 사진 그리드 배치" }).click();
  await expect(page.getByRole("status")).toContainText("사진 배치 완료");
  expect(
    await page.evaluate(
      () => (window as any).__officeTest.shapes[0].textFrame.textRange.text,
    ),
  ).toBe("보존할 원본 제목");
  await page.getByRole("button", { name: "삭제 대상 확인" }).click();
  await expect(page.getByRole("dialog")).toContainText("삭제 1개");
  expect(
    await page.evaluate(() => (window as any).__officeTest.shapes.length),
  ).toBe(4);
  await page.getByRole("button", { name: "확인 후 적용" }).click();
  await expect(page.getByRole("status")).toContainText(
    "사진 삭제 및 재배치 완료",
  );
  expect(
    await page.evaluate(
      () =>
        (window as any).__officeTest.shapes.filter(
          (s: any) => s.type === "Image",
        ).length,
    ),
  ).toBe(2);
});
test("전체 텍스트 서식 변경만 확인하고 적용한다", async ({ page }) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    s.selectedSlide = "slide-2";
    s.seed(
      "TextBox",
      { left: 100, top: 100, width: 300, height: 80 },
      "안전 조회",
    );
    s.selectedSlide = "slide-1";
  });
  await page.getByRole("button", { name: "편집 도구", exact: true }).click();
  await page.getByRole("button", { name: "텍스트·색상", exact: true }).click();
  await page.getByLabel("작업 범위").selectOption("all");
  await page.getByLabel("글자 크기", { exact: true }).fill("30");
  await page.getByRole("button", { name: "텍스트 서식 적용" }).click();
  await expect(page.getByRole("dialog")).toContainText("슬라이드 2장");
  expect(
    await page.evaluate(
      () =>
        (window as any).__officeTest.shapes[0].textFrame.textRange.font.size,
    ),
  ).toBe(24);
  await page.getByRole("button", { name: "확인 후 적용" }).click();
  await expect(page.getByRole("status")).toContainText("텍스트 서식 변경 완료");
  expect(
    await page.evaluate(() =>
      (window as any).__officeTest.shapes.map(
        (s: any) => s.textFrame.textRange.font.size,
      ),
    ),
  ).toEqual([30, 30]);
});
test("유사 개체 검색 결과만 선택하고 삭제 취소 시 보존한다", async ({
  page,
}) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    s.selected = ["original"];
    s.seed(
      "TextBox",
      { left: 200, top: 200, width: 100, height: 60 },
      "두 번째 제목",
    );
    s.seed("Image", { left: 400, top: 200, width: 100, height: 60 });
  });
  await page.getByRole("button", { name: "선택 정보 새로고침" }).click();
  await page.getByRole("button", { name: "편집 도구", exact: true }).click();
  await page.getByRole("button", { name: "유사 검색", exact: true }).click();
  await page.getByRole("button", { name: "유사 개체 찾기" }).click();
  await expect(page.getByRole("status")).toContainText("2개 찾았습니다");
  await page.getByRole("button", { name: "검색 결과 선택" }).click();
  await expect(page.getByRole("status")).toContainText("2개 선택 완료");
  expect(
    await page.evaluate(() => (window as any).__officeTest.selected.length),
  ).toBe(2);
  await page.getByRole("button", { name: "검색 결과 삭제" }).click();
  await page.getByRole("button", { name: "취소", exact: true }).click();
  expect(
    await page.evaluate(() => (window as any).__officeTest.shapes.length),
  ).toBe(3);
});
test("확장 도형 라이브러리의 직선과 둥근 사각형 삽입", async ({ page }) => {
  await connect(page);
  await page.getByRole("button", { name: /기본 도형 핵심/ }).click();
  await page.getByLabel("도형 종류").selectOption("RoundRectangle");
  await page.getByRole("button", { name: "삽입", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("둥근 사각형 삽입 완료");
  await page.getByLabel("도형 종류").selectOption("Line");
  await page.getByRole("button", { name: "삽입", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("직선 삽입 완료");
});
test("사진 두 개는 크기와 제목을 유지하며 위치만 교환한다", async ({
  page,
}) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    const a = s.seed("Image", { left: 40, top: 100, width: 100, height: 80 }),
      b = s.seed("Image", { left: 400, top: 200, width: 200, height: 100 });
    s.selected = [a.id, b.id];
  });
  await page.getByRole("button", { name: "선택 정보 새로고침" }).click();
  await page.getByRole("button", { name: "편집 도구", exact: true }).click();
  await page.getByRole("button", { name: "사진 배치", exact: true }).click();
  await page.getByRole("button", { name: "사진 2개 위치 맞교환" }).click();
  await expect(page.getByRole("status")).toContainText("사진 위치 맞교환 완료");
  expect(
    await page.evaluate(() =>
      (window as any).__officeTest.shapes
        .slice(1)
        .map((x: any) => [x.left, x.top, x.width, x.height]),
    ),
  ).toEqual([
    [400, 200, 100, 80],
    [40, 100, 200, 100],
  ]);
});
test("선택 텍스트 문구 교체와 색상 교체는 즉시 적용한다", async ({ page }) => {
  await connect(page);
  await page.evaluate(() => {
    (window as any).__officeTest.selected = ["original"];
  });
  await page.getByRole("button", { name: "선택 정보 새로고침" }).click();
  await page.getByRole("button", { name: "편집 도구", exact: true }).click();
  await page.getByRole("button", { name: "텍스트·색상", exact: true }).click();
  await page.getByLabel("찾을 문구", { exact: true }).fill("원본");
  await page.getByLabel("바꿀 문구", { exact: true }).fill("새로운");
  await page.getByRole("button", { name: "문구 바꾸기", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("텍스트 바꾸기 완료");
  await expect(page.getByRole("dialog")).toHaveCount(0);
  expect(
    await page.evaluate(
      () => (window as any).__officeTest.shapes[0].textFrame.textRange.text,
    ),
  ).toBe("보존할 새로운 제목");
  await page.getByLabel("기존 색", { exact: true }).fill("#264f9e");
  await page.getByLabel("새 색", { exact: true }).fill("#ff0000");
  await page.getByRole("button", { name: "일치하는 색상 교체" }).click();
  await expect(page.getByRole("status")).toContainText("색상 일괄 교체 완료");
  expect(
    await page.evaluate(
      () => (window as any).__officeTest.shapes[0].fill.foregroundColor,
    ),
  ).toBe("#ff0000");
});
test("기본 서식은 내용과 위치를 유지하며 복제한다", async ({ page }) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    s.shapes[0].fill.foregroundColor = "#ff0000";
    s.shapes[0].textFrame.textRange.font.size = 36;
    s.selected = ["original"];
  });
  await page.getByRole("button", { name: "선택 정보 새로고침" }).click();
  await page.getByRole("button", { name: "편집 도구", exact: true }).click();
  await page.getByRole("button", { name: "서식 복사", exact: true }).click();
  await expect(page.getByRole("status")).toContainText(
    "기본 서식을 복사했습니다",
  );
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    const b = s.seed(
      "TextBox",
      { left: 400, top: 200, width: 200, height: 100 },
      "유지할 내용",
    );
    s.selected = [b.id];
  });
  await page.getByRole("button", { name: "선택 정보 새로고침" }).click();
  await page
    .getByRole("button", { name: "서식 붙여넣기", exact: true })
    .click();
  await expect(page.getByRole("status")).toContainText(
    "기본 서식 붙여넣기 완료",
  );
  expect(
    await page.evaluate(() => {
      const b = (window as any).__officeTest.shapes[1];
      return [
        b.textFrame.textRange.text,
        b.left,
        b.fill.foregroundColor,
        b.textFrame.textRange.font.size,
      ];
    }),
  ).toEqual(["유지할 내용", 400, "#ff0000", 36]);
});
