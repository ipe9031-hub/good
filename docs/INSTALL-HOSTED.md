# PPT AI Studio — 배포형 Windows 설치

연결 주소: https://good-red.vercel.app/

이 패키지는 위 HTTPS 서버를 PowerPoint 작업창에서 여는 Office Add-in 등록 파일입니다. **Node.js 설치, npm 실행, Windows-Setup.cmd, Windows-Start.cmd, 개발 인증서 등록이 필요 없습니다.**

## 1. 배포 주소 확인

Edge에서 https://good-red.vercel.app/을 엽니다. PPT AI Studio가 보여야 합니다. 일반 브라우저에서는 ‘PowerPoint 안에서 열어 주세요’가 표시되고 편집이 차단되는 것이 정상입니다.

404, 배포 오류, Vercel 로그인 화면이 나오면 먼저 Vercel 배포 상태/Deployment Protection을 확인하세요. PowerPoint에서도 페이지를 열 수 있는 공개 HTTPS 배포가 필요합니다. 사내 정책에 따른 접근 제어는 관리자와 조정하세요.

## 2. manifest 등록 (Windows 테스트용)

1. 이 ZIP을 완전히 압축 해제합니다. ZIP 안의 `manifest.xml`은 **good-red.vercel.app을 사용하는 배포형**입니다.
2. `C:\PPTAIStudioCatalog` 폴더를 만들고 이 `manifest.xml`만 복사합니다. 기존 localhost용 manifest가 있다면 이 테스트 카탈로그에서 교체합니다. 소스 ZIP의 localhost용 manifest를 쓰지 마세요.
3. 폴더 우클릭 → 속성 → 공유 → 공유. 본인 계정에 필요한 권한만 부여하고 표시되는 공유 경로(예: `\\YOUR-PC\PPTAIStudioCatalog`)를 기록합니다.
4. PowerPoint → 파일 → 옵션 → 보안 센터 → 보안 센터 설정 → 신뢰할 수 있는 추가 기능 카탈로그.
5. 공유 경로를 카탈로그 URL에 입력 → 카탈로그 추가 → 메뉴에 표시 체크.
6. PowerPoint를 모두 닫았다가 다시 엽니다.
7. 홈 → 추가 기능 → 고급 (버전에 따라 삽입 → 내 추가 기능) → 공유 폴더 → PPT AI Studio → 추가.
8. 리본의 **PPT AI Studio → 스튜디오 열기**를 누릅니다.

공유 폴더 카탈로그는 개발·테스트용입니다. 조직 배포는 Microsoft 365 관리 센터의 통합 앱 등 별도 절차가 필요합니다. 회사 정책이 추가 기능이나 공유 폴더 카탈로그도 막는 경우 관리자 지원이 필요합니다. 개발 인증서 차단을 해결하는 방식과 추가 기능 사용 권한은 별개입니다.

## 3. 간단한 동작 확인

원본 대신 복제한 PPT에서 테스트하세요.

1. 빈 슬라이드를 한 장 선택합니다.
2. 작업창에서 ‘PowerPoint 연결됨’과 선택 슬라이드 1장을 확인합니다.
3. 텍스트 입력 → 삽입. 일반 삽입은 한 번에 적용됩니다. 삭제와 전체 범위 일괄 변경에만 확인창이 표시됩니다.
4. PowerPoint에서 새 텍스트 상자가 직접 편집되는지 확인합니다.
5. 기본 도형·사진 삽입, 선택 개체 위치 변경, 승인 삭제를 차례로 확인합니다.

실제 PowerPoint 검증은 사용자 Windows에서 진행해야 합니다. Linux 자동 테스트가 실기 검증을 대체하지 않습니다.

## 오류 시

- 브라우저에만 사이트가 열림: PowerPoint에 manifest를 등록해야 개체 편집이 가능합니다.
- 기존 localhost 오류가 계속 표시됨: 기존에 추가한 개발형 애드인을 닫고 배포형 manifest로 다시 등록했는지 확인합니다. 리본이 갱신되지 않으면 PowerPoint를 재시작합니다.
- 로드 실패: Vercel 배포와 해당 도메인/Office.js CDN에 대한 PC 네트워크 접근을 확인합니다.
- 카탈로그 메뉴가 없거나 추가 기능이 차단됨: 조직 관리자에게 Office Add-in 사용 및 테스트 카탈로그 정책을 확인합니다.
