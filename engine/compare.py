"""
engine/compare.py — 두 문서를 비교해 변경내용(redline)이 표시된 문서를 생성한다.

`Application.CompareDocuments`는 어떤 오픈소스 라이브러리에도 동등 구현이 없는 엔진 전용
기능이다. 서식·이동(move)·표·머리글/꼬리말·각주·필드·주석까지 의미 단위로 비교해 법률/계약/
공문 검토용 redline 문서를 만든다.
"""
from __future__ import annotations

from typing import Optional

from engine.document import _open, _saveas

# WdCompareDestination / WdGranularity
WD_COMPARE_DEST_NEW = 2       # wdCompareDestinationNew
WD_GRANULARITY_CHAR = 0       # wdGranularityCharLevel
WD_GRANULARITY_WORD = 1       # wdGranularityWordLevel


def compare(
    app,
    original_path: str,
    revised_path: str,
    out_path: str,
    author: str = "검토자",
    granularity: str = "word",
) -> dict:
    """원본과 수정본을 비교해 변경내용(삽입/삭제/이동/서식)이 표시된 redline 문서를 저장한다.

    Args:
        granularity: "word"(기본, 단어 단위) | "char"(문자 단위, 더 촘촘)
    Returns: {"out_path", "revisions"} — revisions는 표시된 변경 개수(가능할 때).
    """
    gran = WD_GRANULARITY_CHAR if str(granularity).lower().startswith("char") else WD_GRANULARITY_WORD
    orig = _open(app, original_path, read_only=True)
    rev = _open(app, revised_path, read_only=True)
    cmp_doc = None
    try:
        cmp_doc = app.CompareDocuments(
            OriginalDocument=orig,
            RevisedDocument=rev,
            Destination=WD_COMPARE_DEST_NEW,
            Granularity=gran,
            CompareFormatting=True,
            CompareCaseChanges=True,
            CompareWhitespace=True,
            CompareTables=True,
            CompareHeaders=True,
            CompareFootnotes=True,
            CompareTextboxes=True,
            CompareFields=True,
            CompareComments=True,
            CompareMoves=True,
            RevisedAuthor=author,
            IgnoreAllComparisonWarnings=True,
        )
        revisions: Optional[int]
        try:
            revisions = int(cmp_doc.Revisions.Count)
        except Exception:
            revisions = None
        out = _saveas(cmp_doc, out_path)
        return {"out_path": out, "revisions": revisions}
    finally:
        for d in (cmp_doc, rev, orig):
            try:
                if d is not None:
                    d.Close(SaveChanges=False)
            except Exception:
                pass
