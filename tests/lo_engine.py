"""
macOS 엔진 검증(LibreOffice) — engine/lo_bridge.py + engine/lo_document.py 의 여섯 작업을 진짜 soffice 로 돌린다.

시험 문서는 표준 라이브러리로 직접 쓴 docx(필드·쪽 나눔·병합 필드)와 csv 다.
실행: python tests/lo_engine.py      (LibreOffice 가 없으면 SKIP)
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from engine.lo_bridge import LoSession, find_soffice  # noqa: E402
from engine import lo_document as doc  # noqa: E402

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')


def p(text):
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def field(instr, shown):
    return (f'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{shown}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')


PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def docx(path, body):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>')


def pdf_pages(path):
    with open(path, "rb") as f:
        data = f.read()
    return len(re.findall(rb"/Type\s*/Page(?![s\w])", data)), data


def main():
    if not find_soffice():
        print("SKIP: LibreOffice 없음")
        return 0
    tmp = tempfile.mkdtemp(prefix="iris-lo-test-")
    fails = 0

    def check(cond, name):
        nonlocal fails
        print(("PASS" if cond else "FAIL"), name, flush=True)
        if not cond:
            fails += 1

    def text_of(s, path):
        out = os.path.join(tmp, "txt", os.path.basename(path) + ".txt")
        s.run(lambda a: doc.convert(a, path, out))
        with open(out, "r", encoding="utf-8-sig") as f:
            return f.read()

    try:
        s = LoSession(profile_dir=os.path.join(tmp, "profile"))
        s.start()
        # ① 필드 갱신: 표시값이 틀린 쪽 수 필드(NUMPAGES=?)가 실제 쪽 수(2)로
        two = os.path.join(tmp, "two.docx")
        docx(two, f'<w:p><w:r><w:t xml:space="preserve">전체 쪽: </w:t></w:r>{field("NUMPAGES", "?")}</w:p>' + p("첫 쪽 한글 문장입니다") + PAGE_BREAK + p("둘째 쪽"))
        upd = os.path.join(tmp, "out", "two.updated.docx")
        r = s.run(lambda a: doc.update_fields(a, two, upd))
        check(r["fields_updated"] >= 1 and os.path.exists(r["out_path"]), f"word_update_fields → 필드 {r['fields_updated']}개")
        check("전체 쪽: 2" in text_of(s, upd), "갱신된 문서에서 NUMPAGES 가 2 로 확정")
        # ② 조판 통계
        lay = s.run(lambda a: doc.read_layout(a, two))
        check(lay["pages"] == 2 and lay["paragraphs"] >= 3 and lay["words"] >= 5 and lay["lines"] and lay["lines"] >= 3
              and lay["characters_with_spaces"] > lay["characters"] > 0 and lay["readability"] is None,
              f"word_read_layout → {dict((k, lay[k]) for k in ('pages', 'words', 'lines', 'characters', 'characters_with_spaces', 'paragraphs'))}")
        # ③ PDF: 쪽 범위·한글 글꼴
        whole = s.run(lambda a: doc.export_pdf(a, two, os.path.join(tmp, "out", "two.pdf")))
        n_all, data = pdf_pages(whole)
        part = s.run(lambda a: doc.export_pdf(a, two, os.path.join(tmp, "out", "p2.pdf"), from_page=2, to_page=2))
        n_part, _ = pdf_pages(part)
        check(n_all == 2 and n_part == 1, f"word_export_pdf → 전체 {n_all}쪽, 범위(2-2) {n_part}쪽")
        check(any(k in data for k in (b"AppleSDGothic", b"Nanum", b"AppleMyungjo", b"AppleGothic")), "PDF 에 한글 글꼴이 박혔다")
        pdfa = s.run(lambda a: doc.export_pdf(a, two, os.path.join(tmp, "out", "a.pdf"), pdf_a=True))
        with open(pdfa, "rb") as f:
            check(b"pdfaid" in f.read(), "pdf_a=True → PDF/A 식별 정보")
        # ④ 변환
        for ext in (".odt", ".doc", ".rtf", ".html"):
            r = s.run(lambda a: doc.convert(a, two, os.path.join(tmp, "out", "conv" + ext)))
            check(os.path.exists(r["out_path"]) and r["format"] == ext[1:] and r["repaired"] is False, f"word_convert → {ext}")
        try:
            s.run(lambda a: doc.convert(a, two, os.path.join(tmp, "out", "x.xps")))
            check(False, "XPS 는 거부(LibreOffice 에 없음)")
        except Exception as e:  # noqa: BLE001
            check("지원하지 않는 출력 확장자" in str(e), "XPS 는 거부(LibreOffice 에 없음)")
        # ⑤ 비교
        orig, rev = os.path.join(tmp, "orig.docx"), os.path.join(tmp, "rev.docx")
        docx(orig, p("계약 금액은 백만 원이다.") + p("기간은 1년이다."))
        docx(rev, p("계약 금액은 이백만 원이다.") + p("기간은 1년이다.") + p("추가 조항."))
        r = s.run(lambda a: doc.compare(a, orig, rev, os.path.join(tmp, "out", "redline.docx"), author="검토자A"))
        check(r["revisions"] >= 2 and os.path.exists(r["out_path"]), f"word_compare → 변경 {r['revisions']}건")
        with zipfile.ZipFile(r["out_path"]) as z:
            xml = z.read("word/document.xml").decode("utf-8")
        check("<w:ins " in xml and 'w:author="검토자A"' in xml, "redline docx 에 변경 추적(w:ins)과 검토자 이름")
        # ⑥ 메일 머지
        tpl = os.path.join(tmp, "tpl.docx")
        docx(tpl, f'<w:p><w:r><w:t xml:space="preserve">이름: </w:t></w:r>{field("MERGEFIELD 이름", "«이름»")}</w:p>'
                  f'<w:p><w:r><w:t xml:space="preserve">점수: </w:t></w:r><w:fldSimple w:instr=" MERGEFIELD 점수 "><w:r><w:t>«점수»</w:t></w:r></w:fldSimple></w:p>')
        data_csv = os.path.join(tmp, "data.csv")
        with open(data_csv, "w", encoding="utf-8") as f:
            f.write("이름,점수\n홍길동,90\n김철수,85\n")
        r = s.run(lambda a: doc.mail_merge(a, tpl, data_csv, os.path.join(tmp, "out", "merged.docx")))
        merged = text_of(s, r["out_path"])
        check(r["records"] == 2 and "이름: 홍길동" in merged and "점수: 85" in merged and "«" not in merged,
              f"word_mail_merge(csv) → {r['records']}건, 값이 채워졌다")
        lay_m = s.run(lambda a: doc.read_layout(a, r["out_path"], update_fields=False))
        check(lay_m["pages"] == 2, f"레코드마다 새 쪽({lay_m['pages']}쪽)")
        table_src = os.path.join(tmp, "data.docx")
        docx(table_src, '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>이름</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>점수</w:t></w:r></w:p></w:tc></w:tr>'
                        '<w:tr><w:tc><w:p><w:r><w:t>이영희</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>77</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
        r = s.run(lambda a: doc.mail_merge(a, tpl, table_src, os.path.join(tmp, "out", "merged2.pdf")))
        check(r["records"] == 1 and pdf_pages(r["out_path"])[0] == 1, "word_mail_merge(docx 표 → PDF)")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print(f"---- FAILS: {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
