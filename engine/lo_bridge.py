"""
engine/lo_bridge.py — macOS 엔진: LibreOffice(무료, MPL-2.0)를 화면 없이 띄워 작업을 맡긴다.

Windows 판은 Office/한컴을 COM 으로 자동화한다(engine/session.py). macOS 에는 COM 이 없으므로 같은 도구를
LibreOffice 로 처리한다. 실제 작업은 soffice 안에서 도는 UNO 매크로(engine/lo_macro/iris_uno.py)가 하고,
이 파일은 작업 파일(JSON)을 넘기고 결과 파일을 받아 온다 — LibreOffice 26 의 내장 Python 은 soffice 만
띄울 수 있어(실행 제약) 밖에서 직접 부를 수 없다.

  · 프로필: IRIS_LO_PROFILE_DIR 또는 이 저장소의 `.lo-profile/`(사용자가 쓰는 LibreOffice 설정과 섞이지 않게).
    매크로는 그 프로필의 user/Scripts/python/ 에 복사해 둔다(내용이 같으면 손대지 않음).
  · 글꼴: 화면 없는 LibreOffice 는 macOS 글꼴 시스템이 아니라 fontconfig 로 글꼴을 찾아 시스템 글꼴(한글 포함)을
    하나도 못 본다 — 한글이 □ 로 나온다(2026-10-27 실측). 그래서 시스템 글꼴 폴더를 적은 fonts.conf 를 프로필에
    만들어 FONTCONFIG_FILE 로 넘긴다(사용자가 FONTCONFIG_FILE 을 이미 정했으면 그것을 쓴다).
  · 한 번에 하나: 같은 프로필의 soffice 가 떠 있으면 새 실행이 그쪽에 요청만 넘기고 끝나 버리므로 직렬화한다.
  · 시간 초과면 그 soffice 프로세스 그룹만 끝낸다(이 파일이 띄운 것만).

`LoSession` 은 engine/session.py 의 세션과 같은 모양(start/stop/alive/run(fn))이라 server.py 는 엔진만 고르면 된다.
`run(fn)` 의 fn 은 `LoBridge` 를 받는다(COM 의 app 자리).

이 파일은 엑셀·PPT·워드 MCP 세 저장소에 글자까지 같은 사본으로 들어 있다.
"""
from __future__ import annotations

import json
import os
import pathlib
import shutil
import signal
import subprocess
import tempfile
import threading
from typing import Any, Callable, Optional

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
MACRO_SRC = os.path.join(HERE, "lo_macro", "iris_uno.py")
SCRIPT_URL = "vnd.sun.star.script:iris_uno.py$run?language=Python&location=user"
DEFAULT_TIMEOUT = float(os.environ.get("IRIS_LO_TIMEOUT", "240"))

FONT_DIRS = (
    "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental",
    "/System/Library/AssetsV2/com_apple_MobileAsset_Font7",  # 내려받기 글꼴(나눔 등)이 여기에 온다
    "/Library/Fonts",
    os.path.expanduser("~/Library/Fonts"),
)

_CANDIDATES = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
)


def find_soffice() -> Optional[str]:
    """LibreOffice 실행 파일: IRIS_SOFFICE → /Applications → ~/Applications → PATH 의 soffice."""
    for c in (os.environ.get("IRIS_SOFFICE"), *_CANDIDATES):
        if c and os.path.isfile(c) and os.access(c, os.X_OK):
            return c
    return shutil.which("soffice")


class LibreOfficeMissing(RuntimeError):
    pass


class LoBridge:
    """soffice 한 번 실행 = 작업 하나. 결과는 dict, 실패는 RuntimeError(LibreOffice 의 오류 문구 포함)."""

    _lock = threading.Lock()

    def __init__(self, profile_dir: Optional[str] = None, soffice: Optional[str] = None):
        self.soffice = soffice or find_soffice()
        self.profile = os.path.abspath(profile_dir or os.environ.get("IRIS_LO_PROFILE_DIR") or os.path.join(REPO, ".lo-profile"))

    # COM 세션의 app.Version 자리(진단 도구가 쓴다)
    @property
    def Version(self) -> str:  # noqa: N802 — COM 쪽 이름과 맞춘다
        if not self.soffice:
            raise LibreOfficeMissing(missing_message())
        r = subprocess.run([self.soffice, "--version"], capture_output=True, text=True, timeout=60)
        return (r.stdout or r.stderr).strip().splitlines()[-1] if (r.stdout or r.stderr).strip() else "LibreOffice"

    def _install_macro(self) -> None:
        dst_dir = os.path.join(self.profile, "user", "Scripts", "python")
        dst = os.path.join(dst_dir, "iris_uno.py")
        with open(MACRO_SRC, "rb") as f:
            want = f.read()
        try:
            with open(dst, "rb") as f:
                if f.read() == want:
                    return
        except OSError:
            pass
        os.makedirs(dst_dir, exist_ok=True)
        tmp = dst + ".tmp"
        with open(tmp, "wb") as f:
            f.write(want)
        os.replace(tmp, dst)

    def _fontconfig(self) -> str:
        conf = os.path.join(self.profile, "fonts.conf")
        from xml.sax.saxutils import escape
        dirs = "".join(f"  <dir>{escape(d)}</dir>\n" for d in FONT_DIRS if os.path.isdir(d))
        text = ('<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd">\n<fontconfig>\n'
                f"{dirs}  <cachedir>{escape(os.path.join(self.profile, 'fontconfig-cache'))}</cachedir>\n</fontconfig>\n")
        try:
            with open(conf, "r", encoding="utf-8") as f:
                if f.read() == text:
                    return conf
        except OSError:
            pass
        os.makedirs(self.profile, exist_ok=True)
        with open(conf, "w", encoding="utf-8") as f:
            f.write(text)
        return conf

    def run(self, op: str, timeout: float = DEFAULT_TIMEOUT, **job: Any) -> dict:
        if not self.soffice:
            raise LibreOfficeMissing(missing_message())
        with self._lock:
            self._install_macro()
            work = tempfile.mkdtemp(prefix="iris-lo-")
            try:
                job_path = os.path.join(work, "job.json")
                result_path = os.path.join(work, "result.json")
                with open(job_path, "w", encoding="utf-8") as f:
                    json.dump({"op": op, **job}, f, ensure_ascii=False)
                env = {**os.environ, "IRIS_LO_JOB": job_path, "IRIS_LO_RESULT": result_path}
                if not env.get("FONTCONFIG_FILE"):
                    env["FONTCONFIG_FILE"] = self._fontconfig()
                profile_url = pathlib.Path(self.profile).as_uri()  # 공백·한글도 올바르게 퍼센트 인코딩
                cmd = [self.soffice, f"-env:UserInstallation={profile_url}", "--headless", "--invisible",
                       "--norestore", "--nologo", "--nodefault", "--nolockcheck", SCRIPT_URL]
                proc = subprocess.Popen(cmd, env=env, cwd=work, stdin=subprocess.DEVNULL,
                                        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, start_new_session=True)
                try:
                    out, _ = proc.communicate(timeout=timeout)
                except subprocess.TimeoutExpired:
                    try:
                        os.killpg(proc.pid, signal.SIGKILL)  # 이 파일이 띄운 soffice 그룹만
                    except OSError:
                        pass
                    proc.communicate()
                    raise RuntimeError(f"LibreOffice 작업 시간 초과({int(timeout)}초): {op}")
                if not os.path.exists(result_path):
                    tail = (out or b"").decode("utf-8", "replace").strip().splitlines()[-5:]
                    raise RuntimeError(f"LibreOffice 가 결과를 남기지 않았습니다({op}, 종료 코드 {proc.returncode}): " + " | ".join(tail))
                with open(result_path, "r", encoding="utf-8") as f:
                    res = json.load(f)
                if not res.get("ok"):
                    raise RuntimeError(res.get("error") or f"LibreOffice 작업 실패: {op}")
                return res.get("result") or {}
            finally:
                shutil.rmtree(work, ignore_errors=True)


def missing_message() -> str:
    return ("LibreOffice 가 없습니다. macOS 에서는 무료 LibreOffice 로 문서를 계산·변환합니다 — "
            "https://www.libreoffice.org/download 에서 설치하거나(또는 `brew install --cask libreoffice`) "
            "IRIS_SOFFICE 에 soffice 경로를 지정하세요.")


class LoSession:
    """engine/session.py 의 세션과 같은 모양. LibreOffice 는 작업마다 새로 뜨므로 상주 프로세스가 없다."""

    def __init__(self, visible: bool = False, profile_dir: Optional[str] = None):
        self.bridge = LoBridge(profile_dir=profile_dir)
        self._started = False

    @property
    def alive(self) -> bool:
        return self._started

    def start(self) -> None:
        if not self.bridge.soffice:
            raise LibreOfficeMissing(missing_message())
        self._started = True

    def stop(self) -> None:
        self._started = False

    def run(self, fn: Callable[[LoBridge], Any]) -> Any:
        if not self._started:
            self.start()
        return fn(self.bridge)
