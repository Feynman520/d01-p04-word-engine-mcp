"""
스모크 테스트 — 본 머신에서 Word COM 엔진 자동화가 실제로 동작하는지 검증한다.
MCP 계층 없이 engine(session + document/compare/mailmerge)을 직접 호출한다.

실행:  .venv\\Scripts\\python.exe tests\\smoke_com.py

통과 기준(핵심 5): 예외/팝업 없이 ① 목차 페이지번호 갱신 ② 실제 페이지 수>=2 산출
③ PDF 생성 ④ .docx→.rtf/.txt/.pdf 변환 ⑤ redline 비교 문서 생성. 메일 머지(⑥)는 데이터소스
드라이버/프롬프트 환경 의존이라 best-effort로 시도하고 결과만 보고한다.
"""
import os
import sys
import traceback

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from engine import compare as compare_mod  # noqa: E402
from engine import document  # noqa: E402
from engine import mailmerge as mailmerge_mod  # noqa: E402
from engine.session import WordSession  # noqa: E402

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out")
os.makedirs(OUT_DIR, exist_ok=True)

# 로케일 독립 내장 스타일 인덱스(WdBuiltinStyle)
WD_STYLE_NORMAL = -1
WD_STYLE_HEADING1 = -2
WD_STYLE_TITLE = -63
WD_PAGE_BREAK = 7  # wdPageBreak
WD_FORMAT_DOCX = 16


def log(msg):
    print(f"[smoke] {msg}", flush=True)


def _make_doc_with_toc(app, path):
    """제목 + 목차 + (페이지 넘기는) 여러 장(章) 본문을 가진 .docx를 COM으로 생성한다.
    python-docx로 만든 '목차 페이지번호 ?' 입력의 대역."""
    doc = app.Documents.Add()
    sel = app.Selection
    sel.Style = doc.Styles(WD_STYLE_TITLE)
    sel.TypeText("엔진 렌더 테스트 문서")
    sel.TypeParagraph()
    # 목차 자리(빈 문단) — 위치 기억 후 마지막에 TOC 삽입
    sel.Style = doc.Styles(WD_STYLE_NORMAL)
    toc_range = sel.Range.Duplicate
    sel.TypeParagraph()
    sel.InsertBreak(WD_PAGE_BREAK)
    for i in range(1, 4):
        sel.Style = doc.Styles(WD_STYLE_HEADING1)
        sel.TypeText(f"{i}장. 엔진 전용 기능")
        sel.TypeParagraph()
        sel.Style = doc.Styles(WD_STYLE_NORMAL)
        for j in range(18):
            sel.TypeText(f"{i}장 본문 문단 {j + 1}. 페이지 조판을 위해 채우는 텍스트입니다. " * 3)
            sel.TypeParagraph()
        sel.InsertBreak(WD_PAGE_BREAK)
    doc.TablesOfContents.Add(
        Range=toc_range, UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=3
    )
    if os.path.exists(path):
        os.remove(path)
    doc.SaveAs2(path, WD_FORMAT_DOCX)
    doc.Close(False)
    return path


def _make_simple_doc(app, path, lines):
    doc = app.Documents.Add()
    sel = app.Selection
    for ln in lines:
        sel.Style = doc.Styles(WD_STYLE_NORMAL)
        sel.TypeText(ln)
        sel.TypeParagraph()
    if os.path.exists(path):
        os.remove(path)
    doc.SaveAs2(path, WD_FORMAT_DOCX)
    doc.Close(False)
    return path


def _make_merge_data(app, path):
    """머리글 1행 + 데이터 2행 표를 가진 .docx 데이터소스를 생성한다(가장 안정적인 머지 소스)."""
    doc = app.Documents.Add()
    rng = doc.Range(0, 0)
    table = doc.Tables.Add(rng, 3, 2)
    rows = [("이름", "점수"), ("김철수", "95"), ("이영희", "88")]
    for r in range(3):
        for c in range(2):
            table.Cell(r + 1, c + 1).Range.Text = rows[r][c]
    if os.path.exists(path):
        os.remove(path)
    doc.SaveAs2(path, WD_FORMAT_DOCX)
    doc.Close(False)
    return path


def _make_merge_template(app, path):
    doc = app.Documents.Add()
    sel = app.Selection
    sel.TypeText("성적 통지 — 이름: ")
    doc.MailMerge.Fields.Add(sel.Range, "이름")
    sel.TypeText("  점수: ")
    doc.MailMerge.Fields.Add(sel.Range, "점수")
    sel.TypeText("점")
    if os.path.exists(path):
        os.remove(path)
    doc.SaveAs2(path, WD_FORMAT_DOCX)
    doc.Close(False)
    return path


def main():
    base = os.path.join(OUT_DIR, "base.docx")
    updated = os.path.join(OUT_DIR, "updated.docx")
    pdf = os.path.join(OUT_DIR, "out.pdf")
    rtf = os.path.join(OUT_DIR, "conv.rtf")
    txt = os.path.join(OUT_DIR, "conv.txt")
    conv_pdf = os.path.join(OUT_DIR, "conv.pdf")
    orig = os.path.join(OUT_DIR, "orig.docx")
    rev = os.path.join(OUT_DIR, "rev.docx")
    redline = os.path.join(OUT_DIR, "redline.docx")
    data = os.path.join(OUT_DIR, "data.docx")
    template = os.path.join(OUT_DIR, "template.docx")
    merged = os.path.join(OUT_DIR, "merged.docx")

    log("WordSession 시작 ...")
    s = WordSession()
    s.start()
    try:
        log(f"Word Version = {s.run(lambda app: str(app.Version))}")

        log("입력 문서(목차 포함) 생성 ...")
        s.run(lambda app: _make_doc_with_toc(app, base))
        assert os.path.exists(base), "base.docx 미생성"

        log("① word_update_fields ...")
        r1 = s.run(lambda app: document.update_fields(app, base, updated))
        log(f"   필드 {r1['fields_updated']}개 갱신 → {os.path.basename(r1['out_path'])}")
        assert os.path.exists(updated), "updated.docx 미생성"

        log("② word_read_layout ...")
        r2 = s.run(lambda app: document.read_layout(app, updated, update_fields=True))
        log(f"   pages={r2['pages']} words={r2['words']} lines={r2['lines']} paras={r2['paragraphs']}")
        assert r2["pages"] and r2["pages"] >= 2, f"페이지 수 비정상: {r2['pages']}"
        assert r2["words"] and r2["words"] > 0, "단어 수 비정상"

        log("③ word_export_pdf (책갈피 포함) ...")
        s.run(lambda app: document.export_pdf(app, updated, pdf, update_fields=False))
        assert os.path.exists(pdf) and os.path.getsize(pdf) > 0, "PDF 미생성"

        log("④ word_convert (.docx→.rtf/.txt/.pdf) ...")
        s.run(lambda app: document.convert(app, updated, rtf))
        s.run(lambda app: document.convert(app, updated, txt))
        s.run(lambda app: document.convert(app, updated, conv_pdf))
        for f in (rtf, txt, conv_pdf):
            assert os.path.exists(f) and os.path.getsize(f) > 0, f"변환 실패: {f}"

        log("⑤ word_compare (redline) ...")
        s.run(lambda app: _make_simple_doc(app, orig, [
            "첫째 문단은 동일합니다.", "둘째 문단 원본 내용.", "셋째 문단은 동일합니다."]))
        s.run(lambda app: _make_simple_doc(app, rev, [
            "첫째 문단은 동일합니다.", "둘째 문단 수정된 내용입니다.", "셋째 문단은 동일합니다.",
            "넷째 문단이 추가되었습니다."]))
        r5 = s.run(lambda app: compare_mod.compare(app, orig, rev, redline, author="검토자"))
        log(f"   변경 {r5['revisions']}건 → {os.path.basename(r5['out_path'])}")
        assert os.path.exists(redline), "redline.docx 미생성"
        assert r5["revisions"] is None or r5["revisions"] >= 1, "변경내용이 0건(비정상)"

        log("=== 핵심 5개 SMOKE PASS: 필드갱신·페이지통계·PDF·변환·비교 OK ===")

        # ⑥ 메일 머지 — 데이터소스 드라이버/프롬프트 환경 의존 → best-effort
        log("⑥ word_mail_merge (best-effort) ...")
        try:
            s.run(lambda app: _make_merge_data(app, data))
            s.run(lambda app: _make_merge_template(app, template))
            r6 = s.run(lambda app: mailmerge_mod.mail_merge(app, template, data, merged))
            assert os.path.exists(merged), "merged.docx 미생성"
            log(f"   레코드 {r6['records']}건 병합 → {os.path.basename(r6['out_path'])} : MAIL MERGE OK")
        except Exception as e:  # noqa: BLE001
            log(f"   MAIL MERGE SKIPPED(환경 의존): {e!r}")

        return 0
    finally:
        log("WordSession 종료 ...")
        s.stop()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
