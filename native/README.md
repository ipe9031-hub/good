# PPT AI Studio 설치형 가능성 검증 · 0.1.0

오른쪽 작업창 대신 **선택한 개체 근처에 작은 도구막대**를 띄우는 Windows 검증본입니다. 기존 웹앱은 보존하며, 이 검증 결과로 설치형 개발 방향을 결정합니다.

## Windows에서 첫 시험

Windows 10/11, 데스크톱 PowerPoint, .NET Framework 4.8이 필요합니다. 웹 PowerPoint·Mac은 지원하지 않습니다. Node.js, 개발 SDK, 웹 서버, 루트 인증서 설치는 필요 없습니다.

1. ZIP을 일반 폴더에 완전히 풀고 `PPTAIStudio-Probe.exe`를 실행합니다.
2. PowerPoint에서 **복제한 시험용 문서**를 열어 둡니다. 두 프로그램 모두 일반 사용자로 실행합니다.
3. 관리 창의 **① 등록 없이 PowerPoint 연결 시험**을 누른 뒤 PowerPoint로 돌아갑니다. 관리 창을 닫으면 연결 시험도 끝납니다.
4. 개체를 선택하면 옆에 도구막대가 뜨는지 확인합니다. 버튼에 잠깐 마우스를 올리면 설명이 나옵니다.

| 선택 | 시험할 기능 |
|---|---|
| 도형 또는 사진 하나 | 돋보기: 같은 종류·비슷한 크기의 개체를 현재 슬라이드에서 자동 선택 |
| 사진 두 장 | ⇄: 두 사진의 위치를 교환. 각 사진의 크기·자르기는 유지 |
| 기준 개체 하나 | 겹친 사각형: 크기 기억. 다른 개체를 선택한 뒤 체크 사각형: 크기 적용 |
| 여러 개체 | 정렬 아이콘: 6방향 정렬, 3개 이상이면 가로·세로 간격 |

마지막 ×는 현재 선택의 도구막대를 숨깁니다. 다른 개체를 선택하면 다시 표시됩니다. 일반 편집은 확인 창 없이 적용하며 PowerPoint 실행 취소를 사용할 수 있습니다. 사진 맞바꾸기는 **위치 교환**이며 사진 내용 교체는 아닙니다.

## 실제 추가 기능 등록 시험

위 연결 시험이 되면 **연결 시험 중지 → PowerPoint 완전 종료 → ② PowerPoint 시작 시 로드 등록** 순서로 진행합니다. PowerPoint를 새로 열고 개체를 선택했을 때 관리 프로그램 없이 도구막대가 나타나는지 확인합니다.

등록은 현재 사용자(HKCU)만 사용하고 파일을 `%LOCALAPPDATA%\PPTAIStudio\NativeProbe\app`에 복사합니다. 관리자 권한·인증서 설치·VSTO 설치를 요구하지 않습니다. Office 정책에 따라 현재 사용자 COM 추가 기능도 차단될 수 있습니다. PowerPoint의 **파일 → 옵션 → 추가 기능 → COM 추가 기능**에서 `PPT AI Studio 설치형 검증` 로드 상태를 볼 수 있습니다. 다른 추가 기능이나 Office 보안 설정은 변경하지 않습니다.

해제: PowerPoint를 완전히 닫고 관리 창의 **추가 기능 등록 해제**를 누릅니다. Windows 앱 목록에서도 같은 검증본을 제거할 수 있습니다. 등록 해제는 파일·로그를 남깁니다. 이후 해당 LocalAppData 폴더와 ZIP 압축 해제 폴더를 수동 삭제할 수 있습니다.

## 현재 검증 수준과 제한

- Linux에서 Windows 바이너리 크로스 빌드와 핵심 검사 13개를 통과했습니다. **실제 Windows·PowerPoint 로딩, 화면 위치, 클릭 동작은 아직 검증하지 않았습니다.**
- 이번 단계는 **선택할 때 표시**입니다. 마우스만 올려서 찾는 기능은 포함하지 않았습니다. 선택은 200ms 간격으로 확인합니다.
- 현재 슬라이드의 최상위 개체 최대 500개를 대상으로 합니다. 그룹 내부 개체는 개별 검색하지 않습니다. 찾기는 같은 도형 종류, 자동 크기 오차(5%, 1.5~18pt), 단색 채우기 색상 오차를 사용합니다. 사진은 연결·삽입 사진과 사진 자리 표시자를 대상으로 하며 이미지 내용의 AI 유사도 검색은 아닙니다.
- 다른 앱이 앞에 있거나 드래그 중이면 도구막대를 숨깁니다. 확대율·다중 모니터·화면 배율·회전 개체·다중 PowerPoint 프로세스는 Windows 실기 확인이 필요합니다.
- 변경 전에 대상·경계를 확인하고 실행 취소 구간을 시작합니다. 일부 쓰기가 실패하면 이전 위치·크기로 복구를 시도하지만 호스트나 정책이 복구도 막을 수 있습니다.
- 개발 검증본에는 게시자 Authenticode 서명이 없습니다. DLL의 strong-name은 CLR 식별용이며 게시자 신뢰를 보장하지 않습니다. SmartScreen 또는 조직 정책이 차단하면 보안 설정을 끄지 말고 표시 문구를 기준으로 진단합니다. 정식 배포에는 코드 서명과 설치 검증이 필요합니다.

연결 시험은 되지만 시작 시 로드가 안 되면 기능 구현과 추가 기능 등록 문제를 구분할 수 있습니다. 둘 다 되면 이 구조로 설치형 개발을 이어갈 근거가 됩니다. 로그 버튼으로 `%LOCALAPPDATA%\PPTAIStudio\NativeProbe\logs`를 열 수 있습니다. 로그에는 작업명·개수·예외 종류/HRESULT만 저장하고 슬라이드 내용·이미지·문서 경로는 기록하지 않습니다.

## 개발 빌드

.NET SDK 8 필요. Windows: `powershell -File native/scripts/build.ps1` (실행 정책 변경 없음). Linux: Python의 `cryptography`와 .NET SDK를 준비하고 `bash native/scripts/build.sh`를 실행합니다. `DOTNET` 환경변수로 SDK 실행 파일 경로를 지정할 수 있습니다. 개인 개발용 strong-name 키는 저장소 밖에 생성하며 배포 ZIP에 넣지 않습니다. 의존성은 `packages.lock.json`의 잠금 모드로 복원합니다.

코어 검사는 사진 위치·크기 보존, 경계 이탈 거부, 크기 복사, 도형 종류 구분, 정렬·간격, 선택 변경 거부, 부분 실패 복구, 도구막대 위치 계산 및 AnyCPU/strong-name 바이너리 메타데이터를 검사합니다. COM 등록·UI 시험을 대체하지 않습니다.

## 구조와 API 근거

- .NET Framework 4.8 WinForms + `IDTExtensibility2` COM 추가 기능. 별도 관리 프로그램의 COM 연결 모드도 같은 편집 코드를 사용합니다. AnyCPU로 빌드하되 32/64비트 Office 모두 실기 시험이 필요합니다.
- 화면 좌표: [PowerPoint DocumentWindow.PointsToScreenPixelsX](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindow.pointstoscreenpixelsx), Y 대응 API.
- 실행 취소: [Application.StartNewUndoEntry](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.startnewundoentry).
- 등록: [.NET Framework COM 등록](https://learn.microsoft.com/en-us/dotnet/framework/interop/registering-assemblies-with-com), [HKCR의 사용자·컴퓨터 등록 병합](https://learn.microsoft.com/en-us/windows/win32/sysinfo/merged-view-of-hkey-classes-root).
- ABI 참고: [NetOffice IDTExtensibility2 선언](https://github.com/NetOfficeFw/NetOffice/blob/main/Source/NetOffice/Tools/Native/IDTExtensibility2.cs). NetOffice 패키지는 사용하지 않습니다.
