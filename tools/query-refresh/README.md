# PS Query Refresh (단독 실행 파일)

`PSQueryRefresh.exe` 하나로 Excel 통합 문서의 파워 쿼리 표를 새로 고치고 수식을 다시 계산해 그 파일에 저장한다. 설치하지 않는다. PS 오피스를 깔기 어려운 느린 PC나, Excel이 낡아 파워 쿼리를 새로 고칠 수 없는 PC용이다.

- 쓰는 법: 파일을 exe 아이콘 위에 끌어다 놓는다. exe를 그냥 두 번 누르면 자기를 Explorer의 [보내기] 메뉴에 "PS Query Refresh"로 넣고 쓰는 법을 보여 준다(사용자 폴더, 관리자 권한 필요 없음). 콘솔 창에 PC 언어로 진행과 결과가 나오고, Enter를 누르면 닫힌다.
- 계산은 PS 오피스의 새로 고침 전용 실행(`--refresh-queries`)과 같은 코드다(`apps/sheets/src/main/refresh-file.ts`). 문구는 `apps/shell/src/main/query-refresh-run.ts`의 22개 언어를 같이 쓴다. 수식을 모두 다시 계산했으면 "열 때 다시 계산" 표시를 파일 원래 값으로 둔다(`calc-on-load.ts`, 낡은 Excel의 `#NAME?` 방지).
- 끝 코드: 모두 됨 0, 파일 없음 1, 하나라도 안 됨 3. `--no-wait`는 마지막 Enter 기다리기를 빼고, `--json`은 파일별 결과를 JSON 한 줄로 낸다(예약 작업·스크립트용).
- 만들기: GitHub Actions의 "PS Query Refresh exe"(`.github/workflows/query-refresh-exe.yml`). Node의 단일 실행 파일(SEA)에 프로그램(`refresh-cli.ts`), 계산 작업 스레드(`query-refresh-worker.ts`, 스택 64MB)와 xlsx 사이드카를 넣고, 아이콘과 이름을 PS 오피스 것으로 바꾼다(`build.mjs`). 사이드카는 처음 실행할 때 `%LOCALAPPDATA%\PSQueryRefresh\<해시>\`에 풀린다. 빌드는 작은 쿼리 통합 문서로 실제 새로 고침을 한 번 해 보고 올린다. 서명하지 않으므로 Windows가 처음 실행 때 경고할 수 있다.
- 배포할 때는 `LICENSE`, `NOTICE`를 함께 준다(아티팩트에 들어 있다).
