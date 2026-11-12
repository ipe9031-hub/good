# PPT AI Studio

현재 개발 방향은 **Windows 설치형의 개체 근처 도구막대 가능성 검증**입니다. [네이티브 검증본 안내](native/README.md) · [Windows 실행 ZIP](https://github.com/ipe9031-hub/good/raw/refs/heads/main/downloads/PPT-AI-Studio-Native-Probe-Windows-0.1.0.zip). Linux 크로스 빌드·핵심 검사 13개 통과, 실제 Windows/PowerPoint 동작은 아직 확인 전입니다. 아래 웹앱은 기존 구현 기록으로 보존합니다.

Windows Microsoft PowerPoint 내부에서 실행하는 한국어 Office Add-in입니다. 일반 웹 PPT 생성기가 아닙니다. **Phase 1은 사용자 Windows PowerPoint에서 실행 확인됐고, 현재 v0.2.4에 Phase 2 기본 편집 도구를 추가했습니다. 새 Phase 2 기능은 Linux 자동 검증을 마쳤으며 Windows 실기 확인은 아직입니다.**

## Windows 다운로드

[Phase 2 전체 소스 Windows ZIP](https://github.com/ipe9031-hub/good/raw/refs/heads/main/downloads/PPT-AI-Studio-Phase2-Windows.zip) · [SHA-256](downloads/PPT-AI-Studio-Phase2-Windows.zip.sha256)

## 인증서 설치 없는 배포형 (권장)

서버 주소: https://good-red.vercel.app/

[배포형 설치 ZIP](https://github.com/ipe9031-hub/good/raw/refs/heads/main/downloads/PPT-AI-Studio-Hosted-Windows.zip)을 풀고, 안의 `manifest.xml`을 PowerPoint 테스트 카탈로그에 등록합니다. **Node.js, npm, Windows-Setup/Start, 로컬 개발 인증서 설치가 필요 없습니다.** 인터넷과 Office 추가 기능 사용 권한은 필요합니다. [배포형 설치 안내](docs/INSTALL-HOSTED.md)를 참고하세요.

GitHub 저장소 루트의 `manifest.production.xml`은 배포형, `manifest.xml`은 localhost 개발형입니다. **다운로드 ZIP에서는 기본 `manifest.xml`이 온라인용이며 `manifest.development.xml`을 로컬 개발용으로 따로 제공합니다.** 파일을 혼용하지 마세요. 배포형 ZIP 안에서는 등록하기 쉽도록 production manifest를 `manifest.xml`이라는 이름으로 제공합니다.

## 로컬 개발형 빠른 시작 (Windows)

1. Node.js **22 LTS 또는 24 LTS**, Microsoft 365 PowerPoint, WebView2를 준비합니다. PowerPointApi 1.5 이상 지원이 필요합니다. 영구 라이선스 구버전은 지원을 보장하지 않습니다.
2. Windows ZIP을 일반 폴더에 완전히 압축 해제합니다. ZIP 안에서 직접 실행하지 마세요.
3. `Windows-Setup.cmd`를 실행합니다. 패키지 설치 후 localhost용 개발 인증서 신뢰 확인이 표시될 수 있습니다. 개인 개발 PC에서만 신뢰하세요.
4. `Windows-Start.cmd`를 실행하고 창을 열어 둡니다.
5. 브라우저에서 `https://localhost:3000`이 인증서 경고 없이 열리는지 확인합니다.
6. [Windows 설치·사이드로드 안내](docs/WINDOWS.md)를 따라 `manifest.xml`을 PowerPoint에 등록합니다.
7. 리본의 **PPT AI Studio → 스튜디오 열기**로 작업창을 엽니다. 슬라이드를 한 장 선택합니다.

ZIP은 **소스·빌드 결과·설치 도구가 포함된 개발용 배포본**입니다. 독립 실행 EXE나 자동 설치되는 상용 애드인이 아닙니다. 설치 시 npm 다운로드와 Office.js CDN 접근이 필요합니다. PowerPoint 밖에서는 편집이 차단됩니다.

## 구현 범위

- PowerPoint 리본 실행 버튼, 오른쪽 Task Pane, 반응형 한국어 UI
- 프레젠테이션 슬라이드 수, 선택 슬라이드 ID, 선택 개체 ID/이름/유형/좌표/크기 읽기
- 선택 개체 하나의 위치·크기 변경, 선택 개체 여러 개 삭제
- 빠른 삽입은 사용자 요청으로 제거. 개체 추가는 PowerPoint 기본 메뉴 사용
- 일반 변경은 한 번 클릭으로 적용, 삭제·전체 프레젠테이션 변경만 확인
- 상시 크기 복제 버튼, 찾기 후 배치 창의 6방향 정렬·가로/세로 간격·크기 복제
- pt/mm/cm 단위 미세 이동, 슬라이드 경계 제한
- 사진 그리드(열 수·간격·비율 설정), 위치 맞교환, 삭제 후 그리드/당기기/위치 유지
- 선택 없이 도형·사진·텍스트 전체 조회 및 종류 필터, 유형·색상·크기의 선택적 유사 조건, 결과 선택·색상 변경·삭제
- 선택/현재 슬라이드/전체 범위의 글꼴·크기·색상 변경, 문구 바꾸기, 색상 교체
- 단색 채우기·테두리·균일 텍스트 서식 복제
- 큰 소개 배너·슬로건·장식 로고 제거, 편집 도구 우선 UI

전체 요구사항과 남은 Phase 2 세부 기능 및 Phase 3~5 기능은 [전체 요구사항](docs/PRODUCT-REQUIREMENTS.ko.txt)과 [개발 현황](docs/STATUS.md)에 보존했습니다. 미구현 AI 버튼이나 가짜 응답은 없습니다.

v0.2.1부터 공유 폴더 목록의 온라인 등록 이름은 **PPT AI Studio 온라인**입니다. 기존 등록에서 로드 오류가 나면 [등록 복구 절차](docs/INSTALL-HOSTED.md#이전-등록-오류-복구-v021)를 따라 새 온라인 등록을 추가하세요.

v0.2.2부터 빠른 삽입 메뉴를 제거하고 편집 도구를 기본 화면으로 사용합니다. 개체 검색은 기본으로 모든 종류를 찾으며, 유사 조건을 사용할 때는 화면의 이전 선택 대신 검색 시점의 실제 Office 선택을 사용합니다. 유형 검색에서는 텍스트·글꼴·채우기·슬라이드 크기를 조회하지 않습니다.

v0.2.3에서는 **찾아서 선택** 한 번으로 검색과 실제 선택을 함께 실행합니다. 검색·배치·서식 설정은 기본으로 접힌 박스에서 열고 닫습니다. 개체 이름/ID 및 상세 결과 목록을 표시하지 않습니다. 유사 크기의 허용 오차는 기준 너비·높이에 따라 자동 계산합니다.

v0.2.4에서는 ‘개체 검색’을 **찾기**로 바꿨습니다. 여러 개체를 찾고 선택하면 **어떻게 배치할까요?** 창에서 정렬·간격·크기 복제를 고릅니다. 기본 화면에는 크기 복제만 바로 표시하고 별도의 너비/높이/위치 복제 버튼은 제거했습니다.

단축키: **Ctrl+Alt+Shift+F** 찾기, **Ctrl+Alt+Shift+C** 크기 복제, **Ctrl+Alt+Shift+1~8** 정렬/간격. 기본 등록에서는 작업창 포커스가 필요합니다. 최신 Office용 [슬라이드 단축키 등록 ZIP](https://github.com/ipe9031-hub/good/raw/refs/heads/main/downloads/PPT-AI-Studio-Keyboard-Windows-0.2.4.zip)과 [설치·지원 범위](docs/SHORTCUTS.md)를 제공합니다. 기존 온라인 등록은 유지됩니다.

## Phase 2 사용

작업창의 **편집 도구**에서 크기 복제 / 사진 배치 / 찾기 / 텍스트·색상을 선택합니다. 선택한 개체에 적용하는 일반 편집은 즉시 실행합니다. 삭제와 전체 프레젠테이션 변경은 대상 목록·개수·범위를 확인합니다.

[Phase 2 사용 범위와 제한](docs/PHASE2.md)을 참고하세요. 사진 내용만 교환, 원본 비율 재구성, 세부 도형 종류별 검색, 스마트 크롭, 그림자 복제 등은 완료로 표시하지 않습니다.

## 개발 명령

```sh
npm ci
npm run certs                 # Windows localhost 인증서 설치
npm run dev                   # HTTPS 3000
npm run typecheck
npm test                      # 단위·Office.js 모의 호스트 계약 테스트
npm run build                 # dist 생성
npm run validate:manifest:local # XML/리소스 구조 검사, 전체 XSD 검사 아님
npm run validate:manifest     # Microsoft 온라인 검증, 네트워크 필요
npm run test:browser          # Chromium 브라우저 + 모의 Office 호스트
npm run package:windows      # 커밋된 소스 및 dist를 ZIP으로 패키징
```

브라우저 테스트는 Linux에서 `/usr/bin/chromium`을 사용합니다. 다른 환경은 `PPT_CHROMIUM_PATH`를 Chromium 실행 파일 경로로 설정하세요. 테스트는 **실제 PowerPoint의 동작을 보장하지 않으며** Windows 수동 검증을 대체하지 않습니다.

클라우드에서는 npm 캐시 쓰기 권한이 없을 경우 `npm ci --cache /workspace/.cache/npm`을 사용합니다. HTTPS 인증서와 키를 별도로 제공하려면 `PPT_DEV_CERT`, `PPT_DEV_KEY` 파일 경로를 지정할 수 있습니다. 인증서 검증을 끄지 마세요. 개인키는 Git/ZIP에 포함하지 않습니다.

## 구조

```text
apps/powerpoint-addin  React + TypeScript + Office.js + Tailwind
apps/api              Phase 3 API 설계 경계 (아직 서버 기능 없음)
packages/powerpoint-adapter  PowerPoint API 호출, 선택 대상 보호
packages/shared       명령 대상·좌표·파일 검증
packages/ai-engine     Phase 3 예정
packages/layout-engine 정렬·그리드·크기·색상·단위 계산
supabase/             후속 Auth/RLS/Storage 도입 지침
scripts/              Windows 실행, manifest 변환/검증, ZIP
```

## Vercel 배포

1. GitHub 저장소를 Vercel 프로젝트로 가져옵니다. Root Directory는 저장소 루트, 빌드 `npm run build`, 결과 `dist`입니다. `vercel.json`에 설정되어 있습니다.
2. HTTPS 배포 URL을 얻은 후 `npm run manifest:production -- https://YOUR-PROJECT.vercel.app`을 실행합니다.
3. `artifacts/manifest.production.xml`을 `node scripts/validate-manifest.mjs artifacts/manifest.production.xml --online`으로 검증하고 해당 manifest를 테스트 카탈로그에 배치합니다.
4. 개발용 localhost manifest와 배포용 manifest를 혼용하지 마세요. manifest URL을 바꿔도 Vercel 배포가 자동 수행되지는 않습니다.
5. 상용/조직 배포는 Microsoft 365 관리 센터의 통합 앱 또는 Marketplace 검토 절차를 별도로 거칩니다. 공유 폴더 카탈로그는 테스트용입니다.

현재 세션에서는 Vercel 배포를 수행하지 않았습니다. 웹페이지가 Office iframe 내에서 열려야 하므로 전역 `X-Frame-Options: DENY` 등을 설정하지 마세요.

## OpenAI / Supabase

Phase 1에는 API 키, 로그인, DB가 필요하지 않습니다. `.env.example`의 후속 환경변수 예시는 사용하지 않는 주석입니다. Phase 3부터 OpenAI 공식 API를 서버에서 호출하고 모델을 서버 설정으로 관리합니다. ChatGPT 구독은 API 사용권과 별개입니다. Supabase 도입 시 Auth와 RLS를 함께 구현하고 검사합니다. 현재 실행할 DB 스키마나 마이그레이션은 없습니다. 비밀키를 `VITE_` 변수나 브라우저 코드에 넣지 마세요.

## 검증 및 문제 해결

[검증 결과](docs/VALIDATION.md), [API 지원 범위와 제약](docs/API-SUPPORT.md), [Windows 체크리스트](docs/WINDOWS.md)를 참고하세요. SSL 경고는 `npm run certs`로 신뢰를 복구하고, 3000 포트 충돌은 해당 프로세스를 종료한 뒤 실행합니다. 리본이 안 보이면 카탈로그 등록 및 PowerPoint 재시작을 확인하세요. 회사 보안 정책으로 추가 기능이 차단되면 관리자에게 테스트 카탈로그 사용 권한을 요청하세요. 자동 복구나 보안 설정 전체 해제는 하지 마세요.
