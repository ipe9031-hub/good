# PPT AI Studio 단축키

## 기본 키

| 키 | 동작 |
|---|---|
| Ctrl+Alt+Shift+F | 찾기 → 현재 슬라이드 결과 선택 → 배치 방법 표시 |
| Ctrl+Alt+Shift+C | 선택 개체 크기 복제 |
| Ctrl+Alt+Shift+1 / 2 / 3 | 왼쪽 / 가로 중앙 / 오른쪽 정렬 |
| Ctrl+Alt+Shift+4 / 5 / 6 | 위쪽 / 세로 중앙 / 아래쪽 정렬 |
| Ctrl+Alt+Shift+7 / 8 | 가로 / 세로 간격 균등 배분 |

크기 복제·정렬은 개체 2개 이상, 간격 배분은 3개 이상이 필요합니다. 입력 상자에서 글을 입력하거나 삭제 확인창이 열렸을 때는 편집 단축키를 실행하지 않습니다. PowerPoint 기본 Ctrl+Z와 Delete는 그대로 사용합니다.

## 기존 온라인 등록에서 사용

기존 ‘PPT AI Studio 온라인’ 등록을 유지할 수 있습니다. 작업창을 클릭하거나 F6로 작업창에 포커스를 이동한 후 위 키를 사용합니다. 슬라이드에 포커스가 있는 동안에는 작업창의 JavaScript가 키를 가로챌 수 없습니다.

## 슬라이드에서 직접 사용하는 선택형 등록

Microsoft 공식 문서에서 PowerPoint Windows 전역 애드인 단축키는 버전 2601 (빌드 19628.20150) 이상, 월 단위 엔터프라이즈 채널은 2604 (빌드 19929.20172) 이상에서 지원합니다. SharedRuntime 1.1 지원이 필요합니다. 구버전에서는 기존 온라인 등록과 작업창 단축키를 사용하세요.

1. `PPT-AI-Studio-Keyboard-Windows-버전.zip`을 받습니다. 파일 → 계정 → PowerPoint 정보에서 위 버전 이상인지 확인합니다.
2. ZIP의 manifest.xml을 기존 테스트 공유 카탈로그에 **manifest.keyboard.xml**이라는 이름으로 복사합니다. 일반 온라인용 파일은 보존합니다.
3. PowerPoint를 다시 열고 추가 기능 → 공유 폴더에서 **PPT AI Studio 단축키**를 추가합니다.
4. ‘PPT AI Studio 단축키’ 리본에서 스튜디오를 한 번 열고, 슬라이드에서 개체를 선택한 뒤 위 키를 사용합니다. 기존 온라인 작업창은 닫아 두면 등록을 구분하기 쉽습니다.

새 등록이 필요한 이유는 Office가 공유 런타임과 키 매핑을 웹페이지가 아닌 manifest에서 읽기 때문입니다. 기본 온라인 등록의 ID와 호환성은 변경하지 않았습니다. 단축키 등록은 별도 ID로 제공하며 구버전에서 지원되는 것처럼 표시하지 않습니다.

다른 추가 기능과 충돌하면 Office가 어떤 동작을 사용할지 물을 수 있습니다. 기본 PowerPoint 기능의 키는 덮어쓰지 않도록 Ctrl+Alt+Shift 조합을 사용했습니다. 키 조합은 애드인 내부 ‘단축키’ 박스에서도 확인할 수 있습니다.

## 검증 범위

작업창 키 이벤트, 입력 중 차단, 중복 실행 방지 및 Office.actions 매핑을 자동 테스트합니다. 이 Linux 환경에는 Windows PowerPoint가 없으므로 실제 Windows 공유 런타임 전역 키 작동은 미검증입니다. 배포된 최신 shortcuts.json 접근도 Windows에서 확인해야 합니다.

공식 근거 (2026-11-12 확인):

- https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/design/keyboard-shortcuts.md
- https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/develop/configure-your-add-in-to-use-a-shared-runtime.md
