"""
engine/document.py — Word 엔진 단일 문서 작업: ① 필드/목차 갱신  ② 레이아웃/통계 산출
③ PDF 렌더링  ④ 포맷 변환·복구.

모든 함수는 `app`(Word.Application COM)을 첫 인자로 받고 전용 워커 스레드에서 호출된다
(session.run). 원본 보존: 입력은 읽기전용으로 열고 결과는 항상 새 경로(out_path)로만 쓴 뒤
Close(SaveChanges=False)로 정리해 다음 호출 오염을 막는다.

라이브러리(python-docx)가 못 하는 것만 노출한다 — python-docx는 XML 트리 편집기라 페이지·
줄 같은 **조판 산출물**을 모르고, 필드(목차·상호참조·PAGE)의 **표시값을 계산**하지 못하며,
.docx 외 포맷을 렌더/변환하지 못한다. 여기서는 진짜 Word 조판·계산·렌더 엔진만 쓴다.
"""
from __future__ import annotations

import os
from typing import Any, Optional

# --- WdStatistic (ComputeStatistics) ----------------------------------------
WD_STAT_WORDS = 0
WD_STAT_LINES = 1
WD_STAT_PAGES = 2
WD_STAT_CHARS = 3
WD_STAT_PARAS = 4
WD_STAT_CHARS_WITH_SPACES = 5

# --- WdExportFormat / 옵션 (ExportAsFixedFormat) ----------------------------
WD_EXPORT_PDF = 17
WD_EXPORT_OPTIMIZE_PRINT = 0
WD_EXPORT_ALL = 0
WD_EXPORT_FROM_TO = 3
WD_EXPORT_CONTENT = 0
WD_EXPORT_NO_BOOKMARKS = 0
WD_EXPORT_HEADING_BOOKMARKS = 1

# --- WdSaveFormat (SaveAs2 FileFormat) — 확장자로 결정 ----------------------
_FMT_BY_EXT = {
    ".docx": 16,  # wdFormatDocumentDefault
    ".docm": 13,  # wdFormatXMLDocumentMacroEnabled
    ".doc": 0,    # wdFormatDocument (Word 97-2003 바이너리)
    ".rtf": 6,    # wdFormatRTF
    ".txt": 7,    # wdFormatUnicodeText (한글 보존 위해 UTF-16)
    ".odt": 23,   # wdFormatOpenDocumentText
    ".html": 8,   # wdFormatHTML
    ".htm": 8,
    ".mht": 9,    # wdFormatWebArchive
    ".mhtml": 9,
    ".xml": 11,   # wdFormatXMLDocument
    ".pdf": 17,   # wdFormatPDF
    ".xps": 18,   # wdFormatXPS
}


def _abspath(path: str) -> str:
    return os.path.abspath(os.path.expanduser(path))


def _ensure_parent(path: str) -> None:
    os.makedirs(os.path.dirname(_abspath(path)), exist_ok=True)


def _open(app, path: str, *, read_only: bool = True, repair: bool = False):
    """문서를 연다(원본 보존: 기본 읽기전용·창없음·변환확인 끔)."""
    src = _abspath(path)
    if not os.path.exists(src):
        raise FileNotFoundError(f"원본을 찾을 수 없습니다: {src}")
    return app.Documents.Open(
        FileName=src,
        ConfirmConversions=False,
        ReadOnly=read_only,
        AddToRecentFiles=False,
        Revert=False,
        Visible=False,
        OpenAndRepair=bool(repair),
    )


def _saveas(doc, out_path: str) -> str:
    """확장자로 형식을 결정해 새 경로로 저장한다(SaveAs2)."""
    out = _abspath(out_path)
    ext = os.path.splitext(out)[1].lower()
    fmt = _FMT_BY_EXT.get(ext)
    if fmt is None:
        raise ValueError(
            f"지원하지 않는 출력 확장자: {ext} (가능: {', '.join(sorted(_FMT_BY_EXT))})"
        )
    _ensure_parent(out)
    doc.SaveAs2(FileName=out, FileFormat=fmt)
    if not os.path.exists(out):
        raise RuntimeError(f"저장 실패: {out}")
    return out


def _update_all_fields(doc) -> int:
    """모든 스토리(본문·머리말/꼬리말·각주·텍스트상자 등)의 필드와 목차/그림목차/색인을
    갱신한다. 갱신 직전의 필드 총개수를 반환한다.

    python-docx는 필드 '코드'만 쓰고 표시값을 계산하지 못한다(목차 페이지번호가 '?'로 남음).
    Word 엔진은 조판 후 PAGE/상호참조/SEQ/목차 페이지번호까지 실제 값으로 확정한다.
    """
    n = 0
    try:
        stories = list(doc.StoryRanges)
    except Exception:
        stories = []
    for story in stories:
        rng = story
        guard = 0
        while rng is not None and guard < 500:  # 안전 가드(연결 스토리 순회)
            guard += 1
            try:
                fields = rng.Fields
                n += int(fields.Count)
                fields.Update()
            except Exception:
                pass
            try:
                rng = rng.NextStoryRange
            except Exception:
                rng = None
    # 목차류는 페이지번호 산출을 위해 별도 Update (조판 의존)
    for getter in ("TablesOfContents", "TablesOfFigures", "TablesOfAuthorities", "Indexes"):
        try:
            coll = getattr(doc, getter)
            for i in range(1, int(coll.Count) + 1):
                try:
                    coll(i).Update()
                except Exception:
                    pass
        except Exception:
            pass
    return n


def _repaginate(app, doc) -> None:
    """백그라운드 페이지네이션을 켜고 강제 재조판한다(페이지 통계/PAGE 필드 정합성)."""
    try:
        app.Options.Pagination = True
    except Exception:
        pass
    try:
        doc.Repaginate()
    except Exception:
        pass


def _readability(doc) -> Optional[dict]:
    """가독성 통계(Flesch 등)를 {이름: 값}으로 반환. 접근 시 맞춤법/문법 검사가 돌며,
    영어가 아니거나 검사 불가하면 예외 → None을 반환한다(best-effort)."""
    try:
        stats = {}
        for s in doc.Content.ReadabilityStatistics:
            try:
                stats[str(s.Name)] = float(s.Value)
            except Exception:
                stats[str(s.Name)] = None
        return stats or None
    except Exception:
        return None


# ============================================================================
# ① 필드/목차 갱신 — 조판 후 모든 필드의 실제 표시값을 확정하고 저장
# ============================================================================
def update_fields(app, src_path: str, out_path: str) -> dict:
    """모든 필드(목차·상호참조·PAGE/NUMPAGES·SEQ·색인 등)를 갱신해 표시값을 확정하고 저장한다.

    python-docx로 만든 문서는 목차 페이지번호가 '?'로 남는다 — 이 도구가 진짜 조판 엔진으로
    값을 채운다. 원본은 보존되고 결과만 out_path로 저장된다.
    """
    doc = _open(app, src_path, read_only=True)
    try:
        _repaginate(app, doc)
        n = _update_all_fields(doc)
        _repaginate(app, doc)  # 갱신으로 길이가 바뀌었을 수 있어 재조판
        out = _saveas(doc, out_path)
    finally:
        doc.Close(SaveChanges=False)
    return {"out_path": out, "fields_updated": n}


# ============================================================================
# ② 레이아웃/통계 — 실제 페이지 수와 조판 통계(라이브러리가 모르는 값)
# ============================================================================
def read_layout(
    app,
    path: str,
    update_fields: bool = True,
    include_readability: bool = False,
) -> dict:
    """문서를 조판해 **실제 페이지 수**와 통계(단어·줄·문자·문단)를 반환한다.

    python-docx에는 '페이지' 개념이 없어 페이지 수·줄 수를 알 수 없다. 여기서는 Word가
    실제로 조판한 결과를 ComputeStatistics로 읽는다. include_readability=True면 가독성
    통계(Flesch-Kincaid 등, 영어 문서에서 의미)도 best-effort로 포함한다.
    """
    doc = _open(app, path, read_only=True)
    try:
        if update_fields:
            _update_all_fields(doc)
        _repaginate(app, doc)

        def stat(kind: int) -> Optional[int]:
            try:
                return int(doc.ComputeStatistics(kind))
            except Exception:
                return None

        result = {
            "path": _abspath(path),
            "pages": stat(WD_STAT_PAGES),
            "words": stat(WD_STAT_WORDS),
            "lines": stat(WD_STAT_LINES),
            "characters": stat(WD_STAT_CHARS),
            "characters_with_spaces": stat(WD_STAT_CHARS_WITH_SPACES),
            "paragraphs": stat(WD_STAT_PARAS),
            "readability": _readability(doc) if include_readability else None,
        }
        return result
    finally:
        doc.Close(SaveChanges=False)


# ============================================================================
# ③ PDF 렌더링 — Word 렌더 엔진으로 정밀 PDF 출력(필드 갱신·PDF/A·책갈피·페이지범위)
# ============================================================================
def export_pdf(
    app,
    src_path: str,
    out_path: str,
    update_fields: bool = True,
    pdf_a: bool = False,
    from_page: Optional[int] = None,
    to_page: Optional[int] = None,
    create_bookmarks: bool = True,
) -> str:
    """문서를 Word 렌더 엔진으로 PDF 출력한다(폰트·레이아웃·필드 그대로).

    python-docx에는 렌더 엔진이 없어 PDF로 구울 수 없다. 출력 전 필드를 갱신해 목차·페이지
    번호를 맞추고, 제목 스타일 기반 책갈피와 (옵션) PDF/A 보존 포맷을 적용한다.
    """
    if os.path.splitext(out_path)[1].lower() != ".pdf":
        raise ValueError("out_path 확장자는 .pdf 여야 합니다.")
    out = _abspath(out_path)
    _ensure_parent(out)
    doc = _open(app, src_path, read_only=True)
    try:
        if update_fields:
            _update_all_fields(doc)
        _repaginate(app, doc)
        use_range = bool(from_page) and bool(to_page)
        doc.ExportAsFixedFormat(
            OutputFileName=out,
            ExportFormat=WD_EXPORT_PDF,
            OpenAfterExport=False,
            OptimizeFor=WD_EXPORT_OPTIMIZE_PRINT,
            Range=(WD_EXPORT_FROM_TO if use_range else WD_EXPORT_ALL),
            From=int(from_page or 1),
            To=int(to_page or 1),
            Item=WD_EXPORT_CONTENT,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=(WD_EXPORT_HEADING_BOOKMARKS if create_bookmarks else WD_EXPORT_NO_BOOKMARKS),
            DocStructureTags=True,       # 태그드 PDF(접근성)
            BitmapMissingFonts=True,
            UseISO19005_1=bool(pdf_a),   # PDF/A-1a 보존 포맷
        )
    finally:
        doc.Close(SaveChanges=False)
    if not os.path.exists(out):
        raise RuntimeError(f"PDF 저장 실패: {out}")
    return out


# ============================================================================
# ④ 포맷 변환·복구 — .doc/.rtf/.odt/.html/.txt/.pdf ↔ .docx (+ 손상 문서 복구)
# ============================================================================
def convert(app, src_path: str, out_path: str, repair: bool = False) -> dict:
    """입력을 열어(필요 시 복구) 출력 확장자 형식으로 저장한다.

    python-docx는 .docx 전용이라 레거시 .doc(바이너리)는 읽지도 못하고, .rtf/.odt/.html
    상호변환도 못 한다. repair=True면 OpenAndRepair로 손상 문서를 복구해 연다.

    지원 확장자: .docx .docm .doc .rtf .txt .odt .html .htm .mht .xml .pdf .xps
    """
    ext = os.path.splitext(out_path)[1].lower()
    if ext not in _FMT_BY_EXT:
        raise ValueError(
            f"지원하지 않는 출력 확장자: {ext} (가능: {', '.join(sorted(_FMT_BY_EXT))})"
        )
    doc = _open(app, src_path, read_only=True, repair=repair)
    try:
        out = _saveas(doc, out_path)
    finally:
        doc.Close(SaveChanges=False)
    return {"out_path": out, "format": ext.lstrip("."), "repaired": bool(repair)}
