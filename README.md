<p align="center"><img src="apps/shell/build/icons/128x128.png" width="96" alt="PS 오피스"></p>

# PS 오피스

**한국어** · [English](README.en.md)

PS 오피스는 PS Textiles가 사내에서 쓰려고 만든 무료 데스크톱 오피스 프로그램입니다. 워드, 엑셀,
파워포인트, PDF, 마크다운, HTML, 한글(HWP) 문서와 마인드맵, 다이어그램을 한 창에서 열고 고칩니다.
AI가 옆에서 답만 하는 것이 아니라 문서 자체를 고칩니다.

오픈소스 프로그램 [이음오피스(IeumOffice)](https://github.com/pblsketch/ieumoffice)를 고쳐
만들었습니다. 이음오피스와 다른 점은 이름·아이콘·로고, 업데이트를 받는 곳, 의견·문의를 받는 곳뿐이고
기능은 같습니다. 자세한 출처는 아래 [라이선스와 출처](#라이선스와-출처)에 있습니다.

## 무엇이 드나요

- **프로그램은 무료입니다.** 설치·업데이트에 돈이 들지 않습니다.
- **AI는 이미 쓰고 있는 Claude 구독으로 돌아갑니다.** API 키가 필요 없고, PS 오피스 때문에 따로 내는
  돈이 없습니다. 앱이 스스로 요금이 나가는 연결로 바꾸지 않습니다. ChatGPT(Codex) 구독이나 무료
  AI(OpenRouter, Google Gemini)도 연결할 수 있습니다.

## 화면 언어

화면 언어는 22개이고 **스페인어(Español)와 한국어**가 모두 들어 있습니다. Windows가 스페인어면 처음부터
스페인어로 열리고, **설정 → 일반 → 언어**에서 언제든 바꿀 수 있습니다.

## 설치

1. [릴리스 페이지](https://github.com/pstextiles/ps-office/releases)에서
   `PSOffice-Setup-<버전>.exe`를 내려받아 실행합니다.
2. 서명하지 않은 설치 파일이라 Windows가 "알 수 없는 게시자" 경고를 띄웁니다. **추가 정보 → 실행**을
   누르면 됩니다.
3. 새 버전이 나오면 앱이 알려 주고, 누르면 알아서 업데이트합니다.

## AI(Claude) 연결

**설정 → AI**의 Claude 카드에서 **설치 → 로그인** 순서로 누르면 됩니다. 단계별 안내는
[CONNECT.md](CONNECT.md)(스페인어·한국어·영어)에 있습니다. 로그인은 Anthropic의 페이지와 Claude Code
프로그램에서 이루어지고, PS 오피스는 비밀번호를 보지 않습니다.

## 엑셀에서 알아 둘 것

- 파워쿼리가 들어 있는 파일은 열리고, **데이터 → 쿼리 및 연결**에서 쿼리 목록과 들어가는 위치를 볼 수
  있습니다. 저장해도 파워쿼리 부분은 그대로 남아 MS 엑셀에서 계속 쓸 수 있습니다.
- 다만 파워쿼리 **새로고침(실행)과 만들기·고치기는 안 됩니다.** 그런 작업은 MS 엑셀에서 하세요.
- 매크로(VBA)도 파일 안에 그대로 남지만 실행하거나 고치지는 않습니다.

## 설치 파일 만들기 (관리자용)

설치 파일은 GitHub Actions가 만듭니다. 저장소의 **Actions → Windows build → Run workflow**를 누르면
약 30~60분 뒤 결과 화면의 **Artifacts**에 `PSOffice-Setup-<버전>.exe`가 생깁니다. 버전을 릴리스로
내려면 먼저 릴리스(`v0.13.0` 같은 태그)를 만들고, Run workflow의 `release_tag` 칸에 그 태그를 적습니다.
직접 빌드하는 방법은 [docs/operations.md](docs/operations.md)에 있습니다.

## 개인정보

PS 오피스는 서버, 회원 계정, 사용 통계가 없습니다. 무엇이 어디로 전송되는지는 [PRIVACY.md](PRIVACY.md)에
모두 적혀 있습니다.

## 라이선스와 출처

PS 오피스는 [Apache License 2.0](LICENSE)으로 배포합니다.

**PS 오피스는 [이음오피스(IeumOffice)](https://github.com/pblsketch/ieumoffice)를 고쳐 만들었고,
이음오피스는 [GenOffice](https://github.com/genspark-ai/genoffice)를 바탕으로 만들었습니다.**
GenOffice는 Genspark(Mainfunc, Inc.)가 만든 오픈소스 오피스 프로그램이고, 이음오피스와 GenOffice 모두
Apache-2.0 라이선스입니다. PS Textiles, S.A.는 제품 이름, 아이콘과 로고, 업데이트를 받는 저장소, 의견과
문의를 받는 주소를 바꿨습니다. 원래의 저작권 고지는 그대로 두었고, 출처와 "고쳐 만든 버전"이라는 사실은
[NOTICE](NOTICE)에 적어 두었습니다. 무엇을 바꿨는지는 이 저장소의 기록에 남아 있습니다.

슬라이드 디자인 템플릿은 [frontend-slides](https://github.com/zarazhangrui/frontend-slides)(MIT
라이선스)의 디자인을 바탕으로 다시 만들었고, 템플릿 글꼴은 모두 SIL Open Font License 1.1입니다.
한글(HWP·HWPX) 편집기는 [rhwp](https://github.com/edwardkim/rhwp)(MIT 라이선스)를 바탕으로 합니다.
한글과컴퓨터가 HWP 형식 문서를 공개하며 붙인 조건에 따라 다음 문구를 밝힙니다.

> 본 제품은 한글과컴퓨터의 한글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

다이어그램 편집기는 [draw.io](https://github.com/jgraph/drawio)(Apache-2.0 라이선스)의 웹 편집기를
바탕으로 합니다. PS 오피스의 아이콘과 로고는 PS Textiles, S.A.의 PS TEXTILES 로고이며, 이 로고와
이름은 Apache License 2.0으로 배포하는 대상이 아닙니다. 그 밖의 고지와 글꼴 목록은 [NOTICE](NOTICE)와 설치 폴더의 `THIRD-PARTY-NOTICES.txt`에
있습니다.

PS 오피스는 이음오피스 기여자, Genspark / Mainfunc, Inc., Anthropic, OpenAI, Microsoft,
한글과컴퓨터, Adobe, XMind, Google과 제휴하거나 후원·보증을 받은 관계가 아닙니다. "이음오피스"와
"IeumOffice", "GenOffice"와 "Genspark"라는 이름과 로고는 각 소유자의 것이고, Apache License 2.0은
그 이름과 로고를 쓸 권리를 주지 않습니다(4조, 6조). 그래서 PS 오피스는 그 이름과 로고를 제품 이름이나
아이콘으로 쓰지 않고, 무엇을 바탕으로 만들었는지 밝히는 데만 적습니다. "Claude"는 Anthropic의,
"Codex"와 "ChatGPT"는 OpenAI의 상표입니다. "Microsoft", "Windows", "Word", "Excel", "PowerPoint"는
Microsoft Corporation의 상표이고, "한글", "한컴", "HWP", "HWPX"는 주식회사 한글과컴퓨터의 등록
상표이며, "XMind"는 XMind Ltd.의, "Google 드라이브"는 Google LLC의 상표입니다. 이 이름들은 무엇과 함께
동작하고 어떤 형식의 파일을 여는지 알리기 위해서만 적었습니다.

PS 오피스는 Apache License 2.0의 7조, 8조에 따라 있는 그대로 제공하며, 어떤 보증도 하지 않습니다. AI가
만든 내용은 틀릴 수 있으니 확인한 뒤 쓰세요.

상표나 저작권 침해 신고, 라이선스에 관한 문의는
[GitHub 이슈](https://github.com/pstextiles/ps-office/issues)로 남겨 주세요. 앱의 설정 → 정보에도 같은
주소가 있습니다.
