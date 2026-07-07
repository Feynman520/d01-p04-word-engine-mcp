"""
engine/mailmerge.py — 네이티브 메일 머지: 템플릿 + 데이터소스 → 병합 문서/PDF.

Word의 `MailMerge`는 데이터소스(Word표/CSV/Excel/Access)를 연결해 IF/NEXT 등 머지 규칙과
함께 레코드를 합치는 엔진 전용 기능이다. 라이브러리의 단순 치환(docxtpl)과 달리 진짜
메일 머지 엔진을 쓴다. 데이터소스로는 **머리글 1행 + 데이터행을 가진 .docx 표**가 가장
안정적이며(드라이버 불필요·무인 동작), .csv/.xlsx도 받는다(환경에 따라 SQL 확인 설정 필요).
"""
from __future__ import annotations

from typing import Optional

from engine.document import (
    WD_EXPORT_ALL,
    WD_EXPORT_CONTENT,
    WD_EXPORT_NO_BOOKMARKS,
    WD_EXPORT_OPTIMIZE_PRINT,
    WD_EXPORT_PDF,
    _abspath,
    _ensure_parent,
    _open,
    _saveas,
)

WD_SEND_TO_NEW_DOCUMENT = 0  # wdSendToNewDocument


def _record_count(mm) -> Optional[int]:
    try:
        n = int(mm.DataSource.RecordCount)
        return n if n >= 0 else None
    except Exception:
        return None


def mail_merge(app, template_path: str, data_path: str, out_path: str) -> dict:
    """템플릿(MERGEFIELD 포함)에 데이터소스를 연결해 메일 머지를 실행하고 결과를 저장한다.

    out_path 확장자가 .pdf면 병합 결과를 PDF로, 그 외(.docx 등)면 해당 포맷으로 저장한다.
    원본 템플릿/데이터는 보존된다.

    Returns: {"out_path", "records"} — records는 병합된 레코드 수(가능할 때).
    """
    template = _open(app, template_path, read_only=False)  # 데이터소스 연결 위해 RW
    merged = None
    try:
        mm = template.MailMerge
        mm.OpenDataSource(
            Name=_abspath(data_path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            LinkToSource=False,
        )
        mm.Destination = WD_SEND_TO_NEW_DOCUMENT
        try:
            mm.SuppressBlankLines = True
        except Exception:
            pass
        records = _record_count(mm)
        mm.Execute(Pause=False)
        merged = app.ActiveDocument  # 병합 결과(새 문서)

        out = _abspath(out_path)
        ext = out.lower().rsplit(".", 1)[-1] if "." in out else ""
        if ext == "pdf":
            _ensure_parent(out)
            merged.ExportAsFixedFormat(
                OutputFileName=out,
                ExportFormat=WD_EXPORT_PDF,
                OpenAfterExport=False,
                OptimizeFor=WD_EXPORT_OPTIMIZE_PRINT,
                Range=WD_EXPORT_ALL,
                Item=WD_EXPORT_CONTENT,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=WD_EXPORT_NO_BOOKMARKS,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
        else:
            out = _saveas(merged, out_path)
        return {"out_path": out, "records": records}
    finally:
        for d in (merged, template):
            try:
                if d is not None:
                    d.Close(SaveChanges=False)
            except Exception:
                pass
