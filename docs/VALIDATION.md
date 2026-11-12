# v0.2.1 등록 분리 및 로드 오류 대응

사용자는 Edge에서 good-red.vercel.app 페이지가 정상이라고 보고했습니다. PowerPoint의 기존 등록 URL·캐시 및 실제 원인은 원격으로 확인하지 못했습니다. 새 온라인 등록에 별도 고정 ID와 표시명 ‘PPT AI Studio 온라인’, 버전이 포함된 HTML URL을 적용하고 HTML 캐시 방지 헤더를 추가했습니다. 소스 ZIP의 기본 manifest도 온라인용으로 바꾸며 원래 개발형은 manifest.development.xml로 보존합니다. 새 Windows 등록의 성공은 아직 검증되지 않았습니다.

클린 npm ci, 타입 검사·빌드, 단위/계약 테스트 41개, Chromium 통합 테스트 15개가 통과했습니다. 개발/배포 ID·이름·URL 분리와 ZIP 템플릿 재사용 및 새 URL의 브라우저 초기화도 검사했습니다.

# 현재 검증 — v0.2.0 / Phase 2 기본 편집 도구

2026-11-12 (KST), Linux, Node.js 24.22.1, Chromium.

- 타입 검사 및 프로덕션 빌드 통과
- 단위·계약 테스트 38개 통과 (입력/대상 보호, 정렬/간격/그리드/단위/색상, 일괄 명령 사전 검증, 조회 후 변경 시 덮어쓰기 차단)
- 실제 Chromium 통합 테스트 14개 통과. PowerPoint 호스트는 모의 객체입니다.
- 확인창 없는 삽입·위치 변경, 삭제 보호, 배너 제거, 정렬·크기, 사진 그리드·위치 교환·삭제 재배치, 유사 검색·선택·취소, 전체 텍스트 서식 승인, 문구/색상 교체, 기본 서식 복제를 테스트했습니다.
- Phase 1은 사용자가 실제 Windows PowerPoint에서 실행·삽입·변경·삭제 동작을 확인했다고 보고했습니다. 아래 과거 기록의 ‘Windows 미실행’은 당시 클라우드 검증 범위입니다.
- 새 Phase 2 기능은 Windows PowerPoint 실기 미검증입니다. Vercel의 최신 커밋 배포 상태도 클라우드 네트워크 제한으로 직접 확인하지 못했습니다.
- API 지원 근거는 공식 `@types/office-js`의 API set을 확인했습니다. 기본 기능은 1.5이며 실제 slide 크기 및 안전한 placeholder 텍스트 확인만 1.10 지원 시 사용합니다.

---

# 검증 기록 — Phase 1

검증일: 2026-11-12 (KST). Linux 클라우드, Node.js 24.22.1, npm 11.9.0, Chromium.

| 검사 | 결과 | 실제 검증 범위 |
|---|---|---|
| npm ci (lockfile) | 통과 | 고정된 의존성 설치 |
| npm run typecheck | 통과 | React/TypeScript 및 공식 Office.js API 타입 |
| npm run build | 통과 | dist 생성 |
| npm test | 20개 통과 | 입력 검증, 선택 보호, Office.js 어댑터 모의 호스트 계약 |
| npm run test:browser | 6개 통과 | 실제 Chromium에서 UI/미리보기/승인/취소/삭제/사진/반응형, Office 호스트는 모의 객체 |
| npm run validate:manifest:local | 통과 | XML 파싱, 리소스 참조, 호스트, 권한, HTTPS, 요구 API 및 리본 액션 검사 |
| 배포 manifest 생성·로컬 검사 | 통과 | 테스트용 HTTPS origin 치환, 원본 개발 manifest 보존; 배포 미실행 |
| HTTPS 개발 서버 | 통과 | 임시 localhost 인증서로 실행, curl --cacert로 TLS 검증 후 HTML·아이콘 HTTP 200 |
| Microsoft 온라인 manifest 검증 | 차단 | validationgateway.omex.office.net으로의 연결이 프록시 HTTP 403, 전체 XSD/Office 서비스 검증 완료로 표시하지 않음 |
| Windows 설치 CMD 및 개발 인증서 신뢰 설치 | 미실행 | Linux에서 실행할 수 없음. Windows 체크리스트 제공 |
| Windows PowerPoint 리본/사이드로드/실제 편집 | 미실행 | Windows PowerPoint가 없는 환경. 가장 중요한 잔여 실기 검증 |
| Vercel / Supabase / OpenAI 실연동 | 미실행 | Phase 1에 외부 서비스 기능 없음. Vercel 설정 및 후속 단계 안내만 제공 |

## 테스트 내용

- PowerPoint 외부에서는 편집 버튼 비활성화, 설치 안내 표시
- 삽입 승인 전에는 PowerPoint 호출 없음
- 텍스트와 도형 삽입 후 성공 상태
- 사진 MIME/크기/디코딩 검증 및 Common API 삽입
- 기존 선택 개체만 위치 변경/삭제
- 승인 전 선택이 바뀐 경우 삭제 차단 및 기존 자료 보존
- 360px에서 수평 넘침 없음, 데스크톱/작업창 스크린샷 확인

초기 테스트에서 발견한 좌표 필드에 개체 메타데이터가 혼합되는 오류를 수정했고, 관련 회귀 테스트와 브라우저 위치 변경 테스트가 통과했습니다.

## 재실행

```sh
npm ci
npm run build
npm test
npm run validate:manifest:local
npm run test:browser
```

브라우저 실행 파일은 `PPT_CHROMIUM_PATH`로 지정할 수 있습니다. `npm run validate:manifest`는 Microsoft 공식 서비스 응답이 accepted인 경우에만 통과합니다. 네트워크가 허용된 환경에서 다시 실행하세요. 로컬 검사는 전체 Microsoft XSD 검사나 실제 호스트 검사를 대체하지 않습니다.

## Windows에서 남은 작업

`WINDOWS.md`의 모든 실기 항목을 실행해 결과를 기록해야 Phase 1 호스트 검증이 완료됩니다. Office 버전/빌드/업데이트 채널도 함께 기록하세요. 자동 테스트 결과를 실제 Office에서 확인한 결과로 해석하지 마세요.

## good-red.vercel.app 배포형 추가

- `manifest.production.xml`과 배포 public manifest를 생성했으며 localhost 참조가 없음을 확인했습니다.
- manifest의 페이지·아이콘 URL이 로컬 빌드의 실제 파일과 일치함을 검사했습니다.
- 타입 검사·빌드와 브라우저 통합 테스트 6개가 통과했습니다.
- 배포형 설치 ZIP은 manifest와 설치 안내만 포함합니다. Node.js 또는 로컬 개발 인증서 설치가 필요 없습니다.
- 클라우드 프록시가 good-red.vercel.app 연결을 HTTP 403으로 차단하여 실제 Vercel 배포 응답은 확인하지 못했습니다. 이는 배포 서버에서 관측한 403이 아닙니다. 사용자 브라우저에서 페이지와 PowerPoint 로드를 확인해야 합니다.
