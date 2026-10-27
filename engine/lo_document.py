"""
engine/lo_document.py — macOS 엔진: 워드 MCP 의 여섯 작업을 LibreOffice Writer 로(engine/document.py ·
compare.py · mailmerge.py 와 같은 함수 이름·같은 반환).

첫 인자 `app` 은 COM 의 Word.Application 대신 engine/lo_bridge.LoBridge 다. 조판·렌더링은 LibreOffice 가 한다 —
Word 와 줄바꿈·쪽 나눔이 조금 다를 수 있고, 가독성 통계(readability)·MHT/XPS 출력은 없다.
"""
from __future__ import annotations

import os
from typing import Optional


def _abspath(path: str) -> str:
    return os.path.abspath(os.path.expanduser(path))


def _require(path: str) -> str:
    src = _abspath(path)
    if not os.path.exists(src):
        raise FileNotFoundError(f"원본을 찾을 수 없습니다: {src}")
    return src


def update_fields(app, src_path: str, out_path: str) -> dict:
    """필드(목차·상호참조·쪽 번호 등)와 색인을 다시 계산해 저장한다. → {"out_path","fields_updated"}"""
    return app.run("writer_update_fields", src=_require(src_path), out=_abspath(out_path))


def read_layout(app, path: str, update_fields: bool = True, include_readability: bool = False) -> dict:
    """조판한 결과의 쪽·단어·줄·문자·문단 수. readability 는 macOS 에서 늘 None(Word 전용 통계)."""
    return app.run("writer_layout", path=_require(path), update_fields=bool(update_fields))


def export_pdf(app, src_path: str, out_path: str, update_fields: bool = True, pdf_a: bool = False,
               from_page: Optional[int] = None, to_page: Optional[int] = None, create_bookmarks: bool = True) -> str:
    """PDF 출력(필드 갱신 · PDF/A · 쪽 범위(둘 다 줘야) · 제목 책갈피)."""
    if os.path.splitext(out_path)[1].lower() != ".pdf":
        raise ValueError("out_path 확장자는 .pdf 여야 합니다.")
    both = bool(from_page and to_page)
    return app.run("writer_pdf", src=_require(src_path), out=_abspath(out_path), update_fields=bool(update_fields),
                   pdf_a=bool(pdf_a), from_page=int(from_page) if both else None, to_page=int(to_page) if both else None,
                   create_bookmarks=bool(create_bookmarks))["out_path"]


def convert(app, src_path: str, out_path: str, repair: bool = False) -> dict:
    """다른 형식으로 변환(.docx .docm .doc .rtf .txt .odt .html .htm .xml .pdf). → {"out_path","format","repaired"}"""
    return app.run("writer_convert", src=_require(src_path), out=_abspath(out_path), repair=bool(repair))


def compare(app, original_path: str, revised_path: str, out_path: str, author: str = "검토자", granularity: str = "word") -> dict:
    """두 문서를 비교해 변경 추적(redline) 문서를 만든다. → {"out_path","revisions"}"""
    return app.run("writer_compare", original=_require(original_path), revised=_require(revised_path),
                   out=_abspath(out_path), author=author or "검토자", granularity=granularity or "word")


def mail_merge(app, template_path: str, data_path: str, out_path: str) -> dict:
    """템플릿의 MERGEFIELD 를 데이터(.csv · .xlsx/.ods 첫 시트 · .docx 첫 표)로 채워 한 문서로. → {"out_path","records"}"""
    return app.run("writer_mail_merge", template=_require(template_path), data=_require(data_path), out=_abspath(out_path))
