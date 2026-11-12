import { beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import { EditorService } from "../packages/powerpoint-adapter/src/editor";
let state: any;
const service = new EditorService();
beforeEach(() => {
  const win: any = {};
  vi.stubGlobal("window", win);
  new Function(readFileSync("tests/fixtures/office-host.js", "utf8"))();
  vi.stubGlobal("Office", win.Office);
  vi.stubGlobal("PowerPoint", win.PowerPoint);
  state = win.__officeTest;
  state.selected = ["original"];
});
describe("Phase 2 Office.js 명령 보호", () => {
  it("지원하는 텍스트와 실제 슬라이드 크기를 읽는다", async () => {
    const s = await service.scan("selected");
    expect(s.canvas).toEqual({ width: 960, height: 540 });
    expect(s.items[0].text).toBe("보존할 원본 제목");
  });
  it("선택 개체 서식만 바꾸고 본문을 보존한다", async () => {
    const s = await service.scan("selected");
    await service.apply({
      title: "서식",
      scan: s,
      patches: [{ before: s.items[0], font: { size: 32 } }],
    });
    expect(state.shapes[0].textFrame.textRange.font.size).toBe(32);
    expect(state.shapes[0].textFrame.textRange.text).toBe("보존할 원본 제목");
  });
  it("전체 프레젠테이션의 명시한 개체만 바꾼다", async () => {
    state.selectedSlide = "slide-2";
    state.seed(
      "TextBox",
      { left: 0, top: 0, width: 100, height: 50 },
      "안전 조회",
    );
    const s = await service.scan("all");
    expect(s.items.length).toBe(2);
    await service.apply({
      title: "교체",
      scan: s,
      patches: [{ before: s.items[1], text: "안전 교육" }],
    });
    expect(state.shapes[0].textFrame.textRange.text).toBe("보존할 원본 제목");
    expect(state.shapes[1].textFrame.textRange.text).toBe("안전 교육");
  });
  it("조회 후 다른 개체를 선택하면 삭제를 중단한다", async () => {
    const s = await service.scan("selected");
    state.selected = [];
    await expect(
      service.apply({
        title: "삭제",
        scan: s,
        patches: [{ before: s.items[0], remove: true }],
      }),
    ).rejects.toThrow("바뀌었습니다");
    expect(state.shapes.length).toBe(1);
  });
  it("조회 후 내용이 바뀌면 덮어쓰지 않는다", async () => {
    const s = await service.scan("selected");
    state.shapes[0].textFrame.textRange.text = "사용자가 수정함";
    await expect(
      service.apply({
        title: "교체",
        scan: s,
        patches: [{ before: s.items[0], text: "덮어쓰기" }],
      }),
    ).rejects.toThrow("변경되었습니다");
    expect(state.shapes[0].textFrame.textRange.text).toBe("사용자가 수정함");
  });
  it("잘못된 일괄 명령은 일부도 적용하지 않는다", async () => {
    const s = await service.scan("selected");
    await expect(
      service.apply({
        title: "불량 명령",
        scan: s,
        patches: [
          { before: s.items[0], font: { size: 20 } },
          { before: s.items[0], font: { size: NaN } },
        ],
      }),
    ).rejects.toThrow();
    expect(state.shapes[0].textFrame.textRange.font.size).toBe(24);
  });
  it("API 1.10 미지원 시 크기를 추정해서 반환하지 않는다", async () => {
    Office.context.requirements.isSetSupported = (_name, version) =>
      version === "1.5";
    const s = await service.scan("selected");
    expect(s.canvas).toBeNull();
  });
});
it('텍스트가 없는 placeholder를 읽거나 덮어쓰지 않는다',async()=>{
 const picture=state.seed('Placeholder',{left:0,top:0,width:200,height:100});picture.textFrame.isNullObject=true;
 const s=await service.scan('slide');const item=s.items.find(x=>x.id===picture.id)!;expect(item.text).toBeUndefined();expect(s.excludedText).toBe(1);
});
it('구버전 placeholder는 명시적으로 제외한다',async()=>{
 state.seed('Placeholder',{left:0,top:0,width:200,height:100},'제목');Office.context.requirements.isSetSupported=(_name,version)=>version==='1.5';
 const s=await service.scan('slide');expect(s.excludedText).toBe(1);expect(s.items.find(x=>x.type==='Placeholder')?.text).toBeUndefined();
});
