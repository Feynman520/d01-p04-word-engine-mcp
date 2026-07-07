"""
engine/session.py — Word.Application COM 세션을 전용 STA 워커 스레드 1개에 고정한다.

win32com COM 객체는 STA(단일 스레드 아파트)에 묶여 생성한 그 스레드에서만 안전하게
호출된다. FastMCP는 asyncio 이벤트 루프에서 도구를 실행하며 스레드를 옮겨다닐 수 있으므로,
모든 Word 호출을 여기 정의한 전용 스레드의 작업 큐로 직렬화한다.

Word 특이사항:
  - `Application.Visible = False`는 정상 동작한다(Excel과 동일, PowerPoint와 다름) →
    문서를 창 없이 화면에 띄우지 않고 작업한다.
  - Word.Application은 Excel처럼 안정적인 `Hwnd`를 노출하지 않으므로(창 핸들은
    ActiveWindow에만 있음), 전용 인스턴스 PID는 DispatchEx 전후 WINWORD.EXE 프로세스
    목록을 비교(diff)해 식별한다(종료 시 좀비 방지).
  - 조기바인딩(`gencache.EnsureDispatch`)을 쓴다 — ExportAsFixedFormat/CompareDocuments
    등 선택적 인자가 많은 메서드를 타입라이브러리로 정확히 마샬링하기 위함.

사용:
    session = WordSession()
    session.start()
    n = session.run(lambda app: app.Documents.Count)
    session.stop()
"""
from __future__ import annotations

import contextlib
import os
import queue
import sys
import threading
import time
import traceback
from concurrent.futures import Future
from typing import Any, Callable, Optional

# COM "서버 사용 중" HRESULT — 시작 직후 Word가 초기화 중일 때 흔히 발생, 재시도로 해소.
_RPC_RETRY_HRESULTS = (
    -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED ("피호출자가 호출을 거부했습니다")
    -2147417846,  # 0x8001010A RPC_E_SERVERCALL_RETRYLATER
)
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

# Word 무인 자동화용 상수
WD_ALERTS_NONE = 0                       # wdAlertsNone
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3  # msoAutomationSecurityForceDisable (열 때 매크로 차단)


@contextlib.contextmanager
def _fd1_to_stderr():
    """fd1(stdout)을 잠시 stderr로 우회 — EnsureDispatch의 1회성 makepy 출력이 JSON-RPC를
    오염시키지 않게 한다(stdout은 MCP 채널)."""
    try:
        sys.stdout.flush()
    except Exception:
        pass
    saved = os.dup(1)
    try:
        os.dup2(2, 1)
        yield
    finally:
        try:
            sys.stdout.flush()
        except Exception:
            pass
        os.dup2(saved, 1)
        os.close(saved)


def _com_retry(fn: Callable[[], Any], *, tries: int = 24, delay: float = 0.25) -> Any:
    """COM 호출이 '서버 사용 중'으로 거부되면 잠깐 쉬고 재시도한다(그 외 오류는 즉시 전파)."""
    import pythoncom  # noqa: F401  (com_error 식별용)

    last: Optional[BaseException] = None
    for _ in range(tries):
        try:
            return fn()
        except pythoncom.com_error as e:  # type: ignore[attr-defined]
            hr = e.args[0] if e.args else None
            if hr in _RPC_RETRY_HRESULTS:
                last = e
                time.sleep(delay)
                continue
            raise
    if last is not None:
        raise last


def _dbg(msg: str) -> None:
    """진단 로그를 stderr와(설정 시) 파일에 쓴다. stdout은 MCP JSON-RPC 채널이라 절대 쓰지 않는다."""
    line = f"[word-session {time.strftime('%H:%M:%S')}] {msg}"
    try:
        print(line, file=sys.stderr, flush=True)
    except Exception:
        pass
    path = os.environ.get("WORD_MCP_DEBUG")
    if path:
        try:
            with open(path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass


def _pids_by_name(name: str) -> set[int]:
    """실행 중인 특정 exe(예: 'WINWORD.EXE')의 PID 집합을 반환한다."""
    out: set[int] = set()
    try:
        import win32api
        import win32process
    except Exception:
        return out
    for pid in win32process.EnumProcesses():
        try:
            h = win32api.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            try:
                exe = win32process.GetModuleFileNameEx(h, 0)
            finally:
                win32api.CloseHandle(h)
            if os.path.basename(exe).lower() == name.lower():
                out.add(int(pid))
        except Exception:
            continue
    return out


def _force_kill(pid: Optional[int]) -> None:
    """Quit 후에도 살아있는 전용 Word 프로세스를 PID로 강제 종료한다(좀비 방지 안전망)."""
    if not pid:
        return
    try:
        import win32api
        import win32con

        h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(h, 0)
        win32api.CloseHandle(h)
        _dbg(f"잔존 Word PID {pid} 강제종료")
    except Exception:  # 이미 종료됐거나 접근 불가 → 무시
        pass


class WordSession:
    """전용 워커 스레드가 소유하는 단일 Word.Application 인스턴스."""

    def __init__(self, *, visible: bool = False):
        self._visible = visible
        self._tasks: "queue.Queue[Optional[tuple[Callable, Future]]]" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._ready = threading.Event()
        self._init_error: Optional[BaseException] = None
        self._app: Any = None
        self._pid: Optional[int] = None

    # ---- lifecycle ---------------------------------------------------------
    def start(self, timeout: float = 90.0) -> None:
        if self._thread is not None:
            return
        self._thread = threading.Thread(
            target=self._run, name="word-com-worker", daemon=True
        )
        self._thread.start()
        if not self._ready.wait(timeout):
            raise TimeoutError("Word COM 세션 초기화가 시간 초과되었습니다.")
        if self._init_error is not None:
            raise self._init_error

    def stop(self, timeout: float = 30.0) -> None:
        if self._thread is None:
            return
        self._tasks.put(None)  # sentinel → 워커 종료
        self._thread.join(timeout)
        self._thread = None

    @property
    def alive(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    # ---- 작업 제출 ---------------------------------------------------------
    def run(self, fn: Callable[[Any], Any], timeout: Optional[float] = 300.0) -> Any:
        """`fn(app)`을 워커 스레드에서 실행하고 결과를 반환(블로킹)."""
        if not self.alive:
            raise RuntimeError("Word 세션이 살아있지 않습니다. start()를 먼저 호출하세요.")
        fut: Future = Future()
        self._tasks.put((fn, fut))
        return fut.result(timeout)

    # ---- 워커 본체 ---------------------------------------------------------
    def _run(self) -> None:
        import pythoncom

        _dbg("worker thread 시작")
        try:
            pythoncom.CoInitialize()
            _dbg("CoInitialize 완료")
        except Exception as e:  # 이미 초기화된 경우 등은 무시
            _dbg(f"CoInitialize 예외(무시): {e}")

        try:
            import win32com.client

            before = _pids_by_name("WINWORD.EXE")
            _dbg("Word.Application DispatchEx 시작 ...")
            t0 = time.time()
            # DispatchEx → 사용자가 띄운 Word/좀비에 붙지 않고 전용 인스턴스를 새로 띄운다.
            raw = _com_retry(lambda: win32com.client.DispatchEx("Word.Application"))
            # 조기바인딩으로 래핑 — ExportAsFixedFormat/CompareDocuments 등 선택적 인자를
            # 타입라이브러리로 정확히 마샬링한다. 첫 호출의 makepy 출력은 fd 가드로 stderr로 보낸다.
            with _fd1_to_stderr():
                app = win32com.client.gencache.EnsureDispatch(raw)
            _com_retry(lambda: app.Version)  # 워밍업
            # 무인 자동화용 환경설정 — 모달 팝업/화면갱신/매크로/링크갱신 차단
            for prop, val in (
                ("Visible", self._visible),
                ("DisplayAlerts", WD_ALERTS_NONE),
                ("ScreenUpdating", False),
                ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
            ):
                try:
                    _com_retry(lambda p=prop, v=val: setattr(app, p, v))
                except Exception as e:  # noqa: BLE001
                    _dbg(f"{prop} 설정 실패(무시): {e}")
            # Options(존재할 때만) — 변환 확인/링크갱신 프롬프트 차단, 백그라운드 페이지네이션 켜기
            for opt, val in (
                ("ConfirmConversions", False),
                ("UpdateLinksAtOpen", False),
                ("Pagination", True),
                ("CheckSpellingAsYouType", False),
                ("CheckGrammarAsYouType", False),
            ):
                try:
                    _com_retry(lambda o=opt, v=val: setattr(app.Options, o, v))
                except Exception as e:  # noqa: BLE001
                    _dbg(f"Options.{opt} 설정 실패(무시): {e}")
            self._app = app
            # 새로 뜬 WINWORD.EXE PID 식별
            new = _pids_by_name("WINWORD.EXE") - before
            self._pid = next(iter(new), None)
            _dbg(f"Word 생성 완료 ({time.time() - t0:.1f}s, pid={self._pid})")
        except BaseException as e:  # noqa: BLE001 — 초기화 실패를 메인 스레드로 전달
            _dbg(f"Word 생성 실패: {e!r}\n{traceback.format_exc()}")
            self._init_error = e
            self._ready.set()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            return

        self._ready.set()

        try:
            while True:
                item = self._tasks.get()
                if item is None:
                    break
                fn, fut = item
                if not fut.set_running_or_notify_cancel():
                    continue
                try:
                    fut.set_result(fn(self._app))
                except BaseException as e:  # noqa: BLE001
                    fut.set_exception(e)
        finally:
            try:
                if self._app is not None:
                    # 남은 문서를 저장 없이 닫고 Word 종료
                    try:
                        for doc in list(self._app.Documents):
                            try:
                                doc.Close(SaveChanges=False)
                            except Exception:
                                pass
                    except Exception:
                        pass
                    try:
                        _com_retry(lambda: self._app.Quit(SaveChanges=False), tries=12)
                    except Exception as e:
                        _dbg(f"Quit 실패(무시): {e}")
            finally:
                # COM 참조를 모두 끊고 gc → Quit이 프로세스를 끝낼 수 있게 한다.
                self._app = None
                app = None  # type: ignore[assignment]  # noqa: F841
                raw = None  # type: ignore[assignment]  # noqa: F841
                import gc

                gc.collect()
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                time.sleep(0.4)  # Quit의 정상 종료에 잠깐 여유
                _force_kill(self._pid)  # 그래도 살아있으면 PID로 강제 종료(전용 인스턴스)
                self._pid = None
            _dbg("worker thread 종료")
