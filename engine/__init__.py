"""engine — Word.Application COM 엔진 전용 작업 묶음.

session(전용 STA 워커) 위에서 document/compare/mailmerge 모듈이 라이브러리(python-docx)가
못 하는 **계산(필드·페이지)·렌더(PDF)·변환·비교·메일머지** 작업만 수행한다.
"""
