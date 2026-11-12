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
    page.getByRole("button", { name: "왼쪽", exact: true }),
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
  await page.setViewportSize({ width: 360, height: 1000 });
  await page.getByRole("button", { name: "편집 도구", exact: true }).click();
  await page.getByRole("button", { name: "텍스트·색상", exact: true }).click();
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= innerWidth,
    ),
  ).toBe(true);
  await page.screenshot({
    path: "artifacts/phase2-tools-360.png",
    fullPage: true,
  });
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
  await page.getByRole("button", { name: "개체 검색", exact: true }).click();
  await page.getByLabel("같은 개체 유형", { exact: true }).check();
  await page.getByRole("button", { name: "개체 찾기" }).click();
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
test("온라인 manifest의 버전이 있는 주소에서도 초기화된다", async ({
  page,
}) => {
  await page.route("https://appsforoffice.microsoft.com/**", (route) =>
    route.fulfill({ contentType: "application/javascript", body: mock }),
  );
  await page.goto("/index.html?addin=hosted&v=0.2.1");
  await expect(page.getByText("PowerPoint 연결됨")).toBeVisible();
  await expect(
    page.getByRole("button", { name: "편집 도구", exact: true }),
  ).toBeVisible();
  await expect(
    page.getByRole("button", { name: "빠른 삽입", exact: true }),
  ).toHaveCount(0);
});
test("선택 새로고침 없이 바꾼 개체를 유사 검색 기준으로 사용한다", async ({
  page,
}) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    const a = s.seed("Rectangle", {
        left: 200,
        top: 100,
        width: 100,
        height: 80,
      }),
      b = s.seed("Rectangle", { left: 400, top: 100, width: 100, height: 80 });
    s.selected = [a.id];
  });
  // No selection event and no refresh: the visible parent cache still has zero selections.
  await page.getByRole("button", { name: "개체 검색", exact: true }).click();
  await page.getByLabel("같은 개체 유형", { exact: true }).check();
  await page.getByRole("button", { name: "개체 찾기" }).click();
  await expect(page.getByRole("status")).toContainText("2개 찾았습니다");
  await expect(page.getByText("기준 개체: Rectangle")).toBeVisible();
});
test("텍스트를 지원하지 않는 네이티브 도형도 유형 검색된다", async ({
  page,
}) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    const a = s.seed("Rectangle", {
      left: 200,
      top: 100,
      width: 100,
      height: 80,
    });
    s.seed("Rectangle", { left: 400, top: 100, width: 100, height: 80 });
    Object.defineProperty(a, "textFrame", {
      get() {
        throw new Error("TextFrame unsupported");
      },
    });
    Object.defineProperty(a, "fill", {
      get() {
        throw new Error("Fill unsupported");
      },
    });
    s.selected = [a.id];
  });
  await page.getByRole("button", { name: "개체 검색", exact: true }).click();
  await page.getByLabel("같은 개체 유형", { exact: true }).check();
  await page.getByRole("button", { name: "개체 찾기" }).click();
  await expect(page.getByRole("status")).toContainText("2개 찾았습니다");
});
test("선택 없이 도형·사진·텍스트를 함께 찾고 사진만 필터링한다", async ({
  page,
}) => {
  await connect(page);
  await page.evaluate(() => {
    const s = (window as any).__officeTest;
    s.seed("Rectangle", { left: 200, top: 100, width: 100, height: 80 });
    s.seed("Image", { left: 400, top: 100, width: 100, height: 80 });
    s.selected = [];
  });
  await page.getByRole("button", { name: "개체 검색", exact: true }).click();
  await page.getByRole("button", { name: "개체 찾기", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("3개 찾았습니다");
  await expect(
    page.getByText("도형 1개 · 사진 1개 · 텍스트 상자 1개"),
  ).toBeVisible();
  await page.getByLabel("개체 종류", { exact: true }).selectOption("Image");
  await page.getByRole("button", { name: "개체 찾기", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("1개 찾았습니다");
  await page.getByRole("button", { name: "검색 결과 선택" }).click();
  await expect(page.getByRole("status")).toContainText("1개 선택 완료");
  expect(
    await page.evaluate(() => {
      const s = (window as any).__officeTest;
      return s.shapes.find((x: any) => x.id === s.selected[0]).type;
    }),
  ).toBe("Image");
  await page.getByLabel("개체 종류", { exact: true }).selectOption("all");
  await page.getByLabel("같은/유사한 크기", { exact: true }).check();
  await page.getByRole("button", { name: "개체 찾기", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("2개 찾았습니다");
  await expect(
    page.getByText("도형 1개 · 사진 1개 · 텍스트 상자 0개"),
  ).toBeVisible();
});
