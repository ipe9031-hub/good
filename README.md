# PPT AI Studio

Windows Microsoft PowerPoint 내부에서 실행하는 한국어 Office Add-in입니다. 일반 웹 PPT 생성기가 아닙니다. **Phase 1 코드 구현 및 Linux 자동 검증 완료, Windows PowerPoint 실기 검증 대기** 상태입니다.

## Windows 다운로드

[Phase 1 Windows ZIP 다운로드](https://github.com/ipe9031-hub/good/raw/refs/heads/main/downloads/PPT-AI-Studio-Phase1-Windows.zip) · [SHA-256](downloads/PPT-AI-Studio-Phase1-Windows.zip.sha256)

## 인증서 설치 없는 배포형 (권장)

서버 주소: https://good-red.vercel.app/

[배포형 설치 ZIP](https://github.com/ipe9031-hub/good/raw/refs/heads/main/downloads/PPT-AI-Studio-Hosted-Windows.zip)을 풀고, 안의 `manifest.xml`을 PowerPoint 테스트 카탈로그에 등록합니다. **Node.js, npm, Windows-Setup/Start, 로컬 개발 인증서 설치가 필요 없습니다.** 인터넷과 Office 추가 기능 사용 권한은 필요합니다. [배포형 설치 안내](docs/INSTALL-HOSTED.md)를 참고하세요.

저장소 루트의 `manifest.production.xml`은 배포형, `manifest.xml`은 localhost 개발형입니다. 파일을 혼용하지 마세요. 배포형 ZIP 안에서는 등록하기 쉽도록 production manifest를 `manifest.xml`이라는 이름으로 제공합니다.

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
- 텍스트 상자, 사각형·원·삼각형·오른쪽 화살표 삽입
- 로컬 PNG/JPEG 사진 삽입 (5MB 이하, 파일 검증, 외부 업로드 없음)
- 선택 개체 하나의 위치·크기 변경, 선택 개체 여러 개 삭제
- 변경 미리보기·사용자 승인·선택 대상 재확인·오류와 진행 상태 표시

Phase 2~5 기능은 [전체 요구사항](docs/PRODUCT-REQUIREMENTS.ko.txt)과 [개발 현황](docs/STATUS.md)에 보존했습니다. 미구현 AI 버튼이나 가짜 응답은 없습니다.

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
packages/layout-engine Phase 2 예정
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
