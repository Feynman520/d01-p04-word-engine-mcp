"""
server.py — Word 엔진 자동화 MCP 서버 (stdio).

Word.Application 데스크톱을 COM으로 자동화해, 라이브러리(python-docx 등)에 **조판/계산/렌더
엔진이 없어** 불가능한 작업을 핵심 도구로 노출한다:
  ① word_update_fields  필드·목차·상호참조·PAGE 갱신 → 표시값 확정 저장
  ② word_read_layout    실제 페이지 수·통계(조판 산출물) 읽기
  ③ word_export_pdf     Word 렌더 엔진으로 PDF 출력(PDF/A·책갈피·페이지범위)
  ④ word_convert        .doc/.rtf/.odt/.html/.txt ↔ .docx 변환(+손상 복구)
  ⑤ word_compare        두 문서 비교 → 변경내용(redline) 문서 생성
  ⑥ word_mail_merge     템플릿+데이터소스 메일 머지 → 문서/PDF

모든 Word 호출은 단일 STA 워커 스레드(engine.session)에서 직렬화되며, 도구는 async로 그
블로킹 작업을 워커 스레드에 위임해 이벤트 루프를 막지 않는다. 세션은 **지연 초기화**:
첫 도구 호출 시 Word를 띄우고, 서버 종료 시 닫는다.

실행:  python server.py        (Claude Code가 stdio로 기동)
"""
from __future__ import annotations

import threading
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from typing import Any, Optional

import anyio
from mcp.server.fastmcp import FastMCP

from engine import compare as compare_mod
from engine import document
from engine import mailmerge as mailmerge_mod
from engine.session import WordSession

# ---- 단일 세션 관리(지연 초기화) -------------------------------------------
_session: Optional[WordSession] = None
_lock = threading.Lock()


def _get_session() -> WordSession:
    global _session
    with _lock:
        if _session is None or not _session.alive:
            s = WordSession(visible=False)
            s.start()
            _session = s
        return _session


def _stop_session() -> None:
    global _session
    with _lock:
        if _session is not None:
            _session.stop()
            _session = None


@asynccontextmanager
async def _lifespan(_server: FastMCP) -> AsyncIterator[dict]:
    try:
        yield {}
    finally:
        await anyio.to_thread.run_sync(_stop_session)


mcp = FastMCP(
    "word-automation",
    lifespan=_lifespan,
    instructions=(
        "Word.Application COM 엔진 자동화. 라이브러리(python-docx)가 못 하는 조판·계산·렌더·"
        "변환·비교 작업 전용. 일반 흐름: (python-docx로 내용을 쓴 .docx를) word_update_fields로 "
        "목차/필드 갱신 → word_read_layout으로 실제 페이지 수 확인 → word_export_pdf로 PDF 출력. "
        "그 외 word_convert(포맷변환·복구), word_compare(redline 비교), word_mail_merge(메일머지). "
        "경로는 절대경로 권장. 원본은 보존되며 결과는 out_path로 저장된다."
    ),
)


def _run(fn) -> Any:
    return _get_session().run(fn)


# ============================================================================
# 핵심 6개 도구 (+ 진단 2)
# ============================================================================
@mcp.tool()
async def word_update_fields(src_path: str, out_path: str) -> dict:
    """① 모든 필드(목차·상호참조·PAGE/NUMPAGES·SEQ·색인)를 갱신해 표시값을 확정·저장한다.

    python-docx는 필드 코드만 쓰고 값을 못 채워 목차 페이지번호가 '?'로 남는다. 이 도구는
    진짜 Word 조판 엔진으로 목차/번호를 실제 값으로 확정한다.

    Args:
        src_path: 입력 문서 절대경로(.docx/.doc/.rtf 등)
        out_path: 결과 저장 경로(확장자로 형식 결정)
    Returns: {"out_path","fields_updated"}
    """
    def work():
        return _run(lambda app: document.update_fields(app, src_path, out_path))
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def word_read_layout(
    path: str,
    update_fields: bool = True,
    include_readability: bool = False,
) -> dict:
    """② 문서를 조판해 **실제 페이지 수**와 통계(단어·줄·문자·문단)를 읽는다.

    python-docx에는 '페이지' 개념이 없어 페이지·줄 수를 알 수 없다. 이 도구는 Word가 실제
    조판한 결과를 읽는다.

    Args:
        path: 문서 절대경로
        update_fields: 읽기 전 필드 갱신 여부(기본 True — 목차/번호가 길이에 영향)
        include_readability: 가독성 통계(Flesch-Kincaid 등, 영어 문서에서 의미) 포함 여부
    Returns: {"path","pages","words","lines","characters","characters_with_spaces","paragraphs","readability"}
    """
    def work():
        return _run(lambda app: document.read_layout(app, path, update_fields, include_readability))
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def word_export_pdf(
    src_path: str,
    out_path: str,
    update_fields: bool = True,
    pdf_a: bool = False,
    from_page: Optional[int] = None,
    to_page: Optional[int] = None,
    create_bookmarks: bool = True,
) -> dict:
    """③ 문서를 Word 렌더 엔진으로 PDF 출력한다(폰트·레이아웃·필드 그대로).

    python-docx에는 렌더 엔진이 없다. 출력 전 필드를 갱신하고, 제목 스타일 기반 책갈피와
    (옵션) PDF/A 보존 포맷·페이지 범위를 적용한다.

    Args:
        src_path: 입력 문서 절대경로
        out_path: 출력 .pdf 절대경로
        update_fields: 출력 전 필드 갱신(기본 True)
        pdf_a: PDF/A-1a 보존 포맷으로 출력(기본 False)
        from_page, to_page: 특정 페이지 범위만 출력(둘 다 줘야 적용)
        create_bookmarks: 제목 스타일을 PDF 책갈피로 생성(기본 True)
    Returns: {"out_path"}
    """
    def work():
        return {"out_path": _run(lambda app: document.export_pdf(
            app, src_path, out_path, update_fields, pdf_a, from_page, to_page, create_bookmarks))}
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def word_convert(src_path: str, out_path: str, repair: bool = False) -> dict:
    """④ 문서를 다른 포맷으로 변환한다(.doc/.rtf/.odt/.html/.txt/.pdf ↔ .docx).

    python-docx는 .docx 전용이라 레거시 .doc(바이너리)는 읽지도 못한다. 출력 확장자가 형식을
    결정한다. repair=True면 손상 문서를 OpenAndRepair로 복구해 연다.

    Args:
        src_path: 입력 문서 절대경로
        out_path: 출력 경로(.docx .doc .rtf .txt .odt .html .mht .xml .pdf .xps)
        repair: 손상 문서 복구 모드로 열기(기본 False)
    Returns: {"out_path","format","repaired"}
    """
    def work():
        return _run(lambda app: document.convert(app, src_path, out_path, repair))
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def word_compare(
    original_path: str,
    revised_path: str,
    out_path: str,
    author: str = "검토자",
    granularity: str = "word",
) -> dict:
    """⑤ 두 문서를 비교해 변경내용(redline)이 표시된 문서를 만든다(법률/계약 검토용).

    `CompareDocuments`는 라이브러리에 동등 구현이 없는 엔진 전용 기능. 삽입/삭제/이동/서식까지
    의미 단위로 비교한다.

    Args:
        original_path: 원본 문서 절대경로
        revised_path: 수정본 문서 절대경로
        out_path: redline 결과 저장 경로
        author: 변경내용에 표시할 검토자 이름
        granularity: "word"(단어 단위, 기본) | "char"(문자 단위)
    Returns: {"out_path","revisions"}
    """
    def work():
        return _run(lambda app: compare_mod.compare(
            app, original_path, revised_path, out_path, author, granularity))
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def word_mail_merge(template_path: str, data_path: str, out_path: str) -> dict:
    """⑥ 템플릿(MERGEFIELD)에 데이터소스를 연결해 메일 머지를 실행한다.

    네이티브 메일 머지 엔진으로 레코드를 합친다. 데이터소스는 **머리글 1행 + 데이터행 표를 가진
    .docx**가 가장 안정적이며(.csv/.xlsx도 가능). out_path가 .pdf면 PDF, 그 외면 해당 포맷 저장.

    Args:
        template_path: 메일 머지 템플릿(.docx) 절대경로
        data_path: 데이터소스 절대경로(.docx 표 / .csv / .xlsx 등)
        out_path: 병합 결과 저장 경로(.docx 또는 .pdf 등)
    Returns: {"out_path","records"}
    """
    def work():
        return _run(lambda app: mailmerge_mod.mail_merge(app, template_path, data_path, out_path))
    return await anyio.to_thread.run_sync(work)


# ---- 진단/복구 -------------------------------------------------------------
@mcp.tool()
async def word_health() -> dict:
    """엔진 상태를 점검한다(세션 기동/Word 버전). 첫 호출 시 Word를 띄운다."""
    def work():
        s = _get_session()
        try:
            ver = s.run(lambda app: str(app.Version))
        except Exception as e:  # noqa: BLE001
            ver = f"버전 조회 실패: {e}"
        return {"alive": s.alive, "word_version": ver}
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def word_restart() -> dict:
    """Word 세션을 종료 후 재기동한다(COM 오류 복구용)."""
    def work():
        _stop_session()
        s = _get_session()
        return {"alive": s.alive}
    return await anyio.to_thread.run_sync(work)


if __name__ == "__main__":
    mcp.run(transport="stdio")
