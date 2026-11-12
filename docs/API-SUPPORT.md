# API 근거와 기술적 한계

확인일: 2026-11-12. Microsoft Learn 직접 접근은 클라우드 정책에서 차단되어 Microsoft OfficeDev 공식 GitHub 문서와 설치된 `@types/office-js` 정의를 함께 확인했습니다. 타입 검사 통과는 Windows 런타임 검증과 다릅니다.

| 요구 기능 | 공식 API / 지원 | 구현 방식 | 한계 및 상태 |
|---|---|---|---|
| Windows 리본/Task Pane | Add-in only XML manifest, Presentation, ShowTaskpane | CustomTab + TaskpaneApp | 구현 완료, 실제 Windows 미검증 |
| 프레젠테이션 인식 | presentation.slides 1.2, document.url | 슬라이드 수·파일명 | URL 미제공 시 문서명 대체 표시. 전체 경로 전송 없음 |
| 선택 슬라이드/개체 | getSelectedSlides/getSelectedShapes 1.5 | ID/이름/유형/좌표/크기 로드 | 구현 완료, 지원 검사 및 manifest 최소 1.5 |
| 텍스트/기본 도형 | shapes.addTextBox/addGeometricShape 1.4 | 새 개체 추가 | 구현 완료, 실기 대기 |
| 사진 삽입 | Common API setSelectedDataAsync + ImageCoercion 1.1 | PNG/JPEG base64, 위치·크기 옵션 | 구현 완료, 직접 addImage라는 미지원 PowerPoint 메서드를 사용하지 않음 |
| 위치·크기 변경 | Shape left/top/width/height 1.4 | 선택 개체 1개 | 구현 완료, 슬라이드 경계 자동 제한은 미구현. pt 값 확인 필요 |
| 선택 삭제 | Shape.delete 1.4 | 승인한 ID 목록만 삭제 | 구현 완료, 부분 실패 시 원자적 롤백 보장 없음 |
| 선택 자동 갱신 | DocumentSelectionChanged Common API | 이벤트 + 수동 새로고침 | 이벤트 미지원/실패 시 수동 새로고침 사용 |
| 미리보기 | 애드인 렌더링 | 입력 내용·좌표·대상·삭제 목록 | 내용 미리보기이며 실제 슬라이드의 정확한 픽셀 렌더링은 아님 |
| 실행 취소 | PowerPoint 기본 Ctrl+Z | 사용자가 실행 | 실제 동작은 Windows 체크리스트로 확인. 선택적 복원은 Phase 5 |

사진 두 크기를 모두 지정하므로 지정 비율로 삽입합니다. 원본 비율 자동 유지/스마트 크롭은 Phase 2 범위입니다. 그룹·선·잠긴 개체 등 호스트가 거부하는 작업은 오류로 표시합니다. 편집 중 다른 슬라이드로 전환하면 재확인하여 중단하지만 호스트 호출 도중 선택을 원자적으로 고정하는 기능은 없습니다.

## 공식 근거

- [XML manifest 개요](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview) / [공식 GitHub 원문](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/develop/xml-manifest-overview.md)
- [PowerPoint 도형 API 사용](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/powerpoint/shapes.md)
- [Windows 공유 폴더 사이드로드](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [PowerPoint requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets)
- [Common 이미지 API](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=common-js#office-office-document-setselecteddataasync-member(1))

XML 방식은 Windows 테스트 카탈로그와 호환되는 명시적 선택입니다. unified manifest 지원 범위를 가정하지 않습니다. 상용 배포는 별도 검토가 필요합니다.
