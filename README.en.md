<p align="center"><img src="apps/shell/build/icons/128x128.png" width="96" alt="PS Office"></p>

# PS Office

[한국어](README.md) · **English**

PS Office is a free desktop office suite PS Textiles made for its own use. It opens and edits Word,
Excel, PowerPoint, PDF, Markdown, HTML and Hangul (HWP) documents, mind maps and diagrams in one
window, and its AI edits the document itself.

It is a modified version of the open-source [IeumOffice](https://github.com/pblsketch/ieumoffice).
Only the name, icons and logo, where updates come from, and where feedback and contact go differ from
IeumOffice; the features are the same. See [License and acknowledgement](#license-and-acknowledgement).

## What it costs

- **The program is free.** Installing and updating cost nothing.
- **AI runs on the Claude subscription you already pay for.** No API key is needed and PS Office
  charges nothing on top. The app never switches to a paid connection on its own. A ChatGPT (Codex)
  subscription or a free AI (OpenRouter, Google Gemini) can be connected too.

## Languages

The app has 22 UI languages, **Spanish and Korean** among them. On a Spanish Windows it starts in
Spanish; change it any time under **Settings → General → Language**.

## Installing

1. Download `PSOffice-Setup-<version>.exe` from the
   [releases page](https://github.com/pstextiles/ps-office/releases) and run it.
2. The installer is not code-signed, so Windows shows an "unknown publisher" warning. Press
   **More info → Run anyway**.
3. When a new version is out, the app offers it and updates itself.

## Connecting AI (Claude)

Under **Settings → AI**, press **Install**, then **Sign in** on the Claude card. Step-by-step
instructions are in [CONNECT.md](CONNECT.md) (Spanish, Korean, English). Signing in happens on
Anthropic's page and in the Claude Code program; PS Office never sees your password.

## Spreadsheets: good to know

- Workbooks with Power Query open, and **Data → Queries & Connections** lists the queries and where
  their rows land. Saving keeps the Power Query parts, so the file still works in Microsoft Excel.
- Refreshing, creating or editing Power Query queries is **not** supported; do that in Excel.
- Macros (VBA) are kept in the file but not run or edited.

## Building the installer (maintainers)

GitHub Actions builds the installer: **Actions → Windows build → Run workflow**. After about 30 to
60 minutes `PSOffice-Setup-<version>.exe` is under **Artifacts** on the run page. To publish a
release, create the release (a tag such as `v0.13.0`) first and enter that tag in `release_tag` when
running the workflow. Building by hand is described in [docs/operations.md](docs/operations.md).

## Privacy

PS Office has no server, no accounts and no usage statistics. Everything that leaves the PC, and
where it goes, is listed in [PRIVACY.md](PRIVACY.md).

## License and acknowledgement

PS Office is distributed under the [Apache License 2.0](LICENSE).

**PS Office is a modified version of [IeumOffice](https://github.com/pblsketch/ieumoffice), which is
based on [GenOffice](https://github.com/genspark-ai/genoffice).** GenOffice is an open-source office
suite by Genspark (Mainfunc, Inc.); IeumOffice and GenOffice are both Apache-2.0 licensed. PS
Textiles, S.A. changed the product name, the icons and logo, the repository the app updates from, and
the feedback and contact addresses. The original copyright notices are kept, and the attribution and
the fact that this is a modified version are recorded in [NOTICE](NOTICE). The repository history
records each change.

The slide design templates are rebuilt from the designs of
[frontend-slides](https://github.com/zarazhangrui/frontend-slides) (MIT License), and their fonts are
all under the SIL Open Font License 1.1. The HWP/HWPX editor is built on
[rhwp](https://github.com/edwardkim/rhwp) (MIT License). As required by Hancom Inc. for software
developed with its published HWP format document: this product was developed with reference to
Hancom's published document on the Hangul document file format (.hwp). The diagram editor is built on
the web editor of [draw.io](https://github.com/jgraph/drawio) (Apache-2.0). The PS Office icon and
logo are the PS TEXTILES logo of PS Textiles, S.A.; that name and logo are not licensed under the
Apache License 2.0. Other
notices and the font list are in [NOTICE](NOTICE) and in `THIRD-PARTY-NOTICES.txt` in the install
folder.

PS Office is not affiliated with, sponsored by or endorsed by the IeumOffice contributors, Genspark /
Mainfunc, Inc., Anthropic, OpenAI, Microsoft, Hancom, Adobe, XMind or Google. The names and logos
"IeumOffice", "이음오피스", "GenOffice" and "Genspark" belong to their owners, and the Apache License
2.0 grants no right to use them (sections 4 and 6), so PS Office does not use them as its product name
or icon and names them only to say what it is based on. "Claude" is a trademark of Anthropic; "Codex"
and "ChatGPT" are trademarks of OpenAI. "Microsoft", "Windows", "Word", "Excel" and "PowerPoint" are
trademarks of Microsoft Corporation; "한글" (Hangul), "한컴" (Hancom), "HWP" and "HWPX" are
registered trademarks of Hancom Inc.; "XMind" is a trademark of XMind Ltd.; "Google Drive" is a
trademark of Google LLC. These names are used only to say what the app works with and which files it
opens.

PS Office is provided as is, without warranties of any kind, under sections 7 and 8 of the Apache
License 2.0. AI-generated content can be wrong: check it before you use it.

Trademark and copyright reports and licensing questions go to
[GitHub issues](https://github.com/pstextiles/ps-office/issues). The same address is in the app under
Settings → About.
