# word-engine-mcp — 한국어 안내

[English → README.md](README.md)

Word.Application 데스크톱을 COM으로 자동화해, **라이브러리(python-docx)에 조판·계산·렌더
엔진이 없어 불가능한 작업**을 핵심 도구로 노출하는 로컬 MCP 서버.

> 설계 사상: 텍스트·표·스타일 편집은 python-docx가 빠르고, **필드/목차 계산·페이지 조판·
> PDF 렌더·포맷 변환·문서 비교**는 엔진만 가능하다. 이 MCP는 후자(엔진 전용)만 담당한다.
> (python-docx로 만든 문서는 목차 페이지번호가 `?`로 남고, '페이지' 개념 자체가 없다.)

## 요구사항
- Windows 10+ / **로그인된 인터랙티브 데스크톱 세션**
- Microsoft Office(Word) 설치 — 검증 환경: **Office 2016+ (Word 16.0)**
- Python 3.10+ — 검증: **3.12**
- [Claude Code](https://claude.com/claude-code) 또는 임의의 MCP 클라이언트

## 설치

```powershell
git clone https://github.com/Feynman520/d01-p04-word-engine-mcp.git
cd d01-p04-word-engine-mcp
py -3.12 -m venv .venv          # 또는: python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Claude Code 등록

클론한 폴더 안에서 아래 한 줄 실행 (절대경로가 박히므로 이후 어디서든 동작):

```powershell
claude mcp add word-automation --scope user -- "$PWD\.venv\Scripts\python.exe" "$PWD\server.py"
```

`--scope user`는 모든 프로젝트에서 사용 가능. 한 프로젝트에서만 쓰려면 `--scope project`.

### 동작 검증

```powershell
$py = ".\.venv\Scripts\python.exe"; $env:PYTHONUTF8 = "1"
& $py tests\smoke_com.py      # COM 필드갱신·페이지통계·PDF·변환·비교·메일머지·좀비정리 검증
& $py tests\server_tools.py   # MCP 8개 도구 등록 검증(Word 미기동)
```

## 도구 (핵심 6 + 진단 2)

| # | 도구 | 입력 → 출력 | 엔진 전용 이유 |
|---|---|---|---|
| ① | `word_update_fields` | `src_path, out_path` → `{out_path,fields_updated}` | 목차·상호참조·PAGE·SEQ·색인의 **표시값 계산**(python-docx는 코드만, 값은 `?`) |
| ② | `word_read_layout` | `path, update_fields?, include_readability?` → `{pages,words,lines,characters,...}` | **실제 페이지 수**·조판 통계(라이브러리에 '페이지' 개념 없음) |
| ③ | `word_export_pdf` | `src_path, out_path, update_fields?, pdf_a?, from_page?, to_page?, create_bookmarks?` → `{out_path}` | Word 렌더 엔진 PDF(PDF/A·제목 책갈피·페이지범위) |
| ④ | `word_convert` | `src_path, out_path, repair?` → `{out_path,format,repaired}` | `.doc`/`.rtf`/`.odt`/`.html`/`.txt` ↔ `.docx` 변환·**손상 복구**(python-docx는 .docx 전용) |
| ⑤ | `word_compare` | `original_path, revised_path, out_path, author?, granularity?` → `{out_path,revisions}` | `CompareDocuments` redline(라이브러리에 동등 구현 없음) |
| ⑥ | `word_mail_merge` | `template_path, data_path, out_path` → `{out_path,records}` | 네이티브 메일 머지(레코드별 문서) |
| — | `word_health` | → `{alive, word_version}` | 세션 점검(첫 호출 시 Word 기동) |
| — | `word_restart` | → `{alive}` | COM 오류 복구 |

전형적 흐름: python-docx로 만든 `.docx` → `word_update_fields`로 목차/번호 확정 →
`word_read_layout`으로 실제 페이지 수 확인 → `word_export_pdf`로 배포용 PDF.
원본은 절대 수정하지 않고 결과는 항상 `out_path`로 저장한다.

### 메일 머지 데이터소스
가장 안정적인 소스는 **머리글 1행 + 데이터행 표를 가진 `.docx`**(드라이버 불필요·무인 동작).
`.csv`/`.xlsx`도 받지만, 환경에 따라 Word의 SQL 실행 확인 프롬프트가 뜰 수 있다(HKCU 설정 의존).

## 아키텍처 핵심
- **단일 STA 워커 스레드**(`engine/session.py`): 모든 Word 호출을 전용 스레드 1개에 직렬화.
- **세션 지연 초기화**: 첫 도구 호출 시 Word 기동, 서버 종료 시 닫음. 1개 인스턴스 재사용.
- **`DispatchEx` + `gencache.EnsureDispatch`(조기바인딩)**: 전용 인스턴스를 새로 띄우고,
  ExportAsFixedFormat/CompareDocuments 등 선택적 인자가 많은 메서드를 타입라이브러리로 정확히
  마샬링한다. 첫 호출의 makepy 생성 출력은 fd 가드로 stderr로 보내 JSON-RPC(stdout) 오염을 막는다.
- **`Visible=False`**: Word는 앱 창 숨김이 정상 동작. 매크로는 `AutomationSecurity=ForceDisable`로
  열 때 차단(보안 기본값).
- **원본 보존**: 입력은 읽기전용으로 열고 결과는 항상 새 경로로만 쓴 뒤 `Close(SaveChanges=False)`.
- **좀비 방지**: Word.Application은 안정적 Hwnd를 안 주므로, DispatchEx 전후 `WINWORD.EXE`
  프로세스 목록을 비교(diff)해 전용 인스턴스 PID를 식별하고, 종료 시 남으면 강제 종료한다.
- **RPC 거부 재시도**: 기동 직후 `RPC_E_CALL_REJECTED`를 잠깐 쉬고 재시도.

## 한계
- 무인/서비스 세션 부적합(데스크톱 세션 필요).
- 텍스트·표·스타일 편집은 python-docx가 빠름 — 의도된 역할 분담이다.

## 라이선스

[MIT](LICENSE)
