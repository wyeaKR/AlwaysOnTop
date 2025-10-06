import tkinter as tk
from tkinter import ttk
import pygetwindow as gw
import win3egu
import win32con
import keyboard
import threading
import time
import ctypes
import sys
import webbrowser
import requests # 업데이트 확인용

# --- 프로그램 설정 ---
CURRENT_VERSION = "1.0.2"
GITHUB_API_URL = "https://api.github.com/repos/wyeaKR/AlwaysOnTop/releases/latest"

# --- 상태 관리를 위한 전역 변수 ---
running = False
hwnd = None
pin_thread = None

def check_for_updates():
    """
    GitHub에서 태그 이름을 비교하여 최신 버전을 확인합니다.
    반환값:
        str: 'latest', 'update_needed', 'error' 중 하나
    """
    try:
        response = requests.get(GITHUB_API_URL, timeout=5)
        response.raise_for_status() # HTTP 에러(4xx 또는 5xx)가 있을 경우 예외 발생

        latest_version_tag = response.json()['tag_name']
        latest_version = latest_version_tag.lstrip('v') # 태그에서 'v' 접두사 제거

        # --- 수정됨: 정확한 버전 일치 확인 ---
        # 정확한 비교를 위해 버전 문자열을 정수 리스트로 변환
        current_parts = list(map(int, CURRENT_VERSION.split('.')))
        latest_parts = list(map(int, latest_version.split('.')))

        # 비교 오류를 막기 위해 버전 길이를 맞춤 (예: 1.0 -> 1.0.0)
        while len(current_parts) < len(latest_parts):
            current_parts.append(0)
        while len(latest_parts) < len(current_parts):
            latest_parts.append(0)

        # 버전이 정확히 일치할 때만 프로그램 실행을 허용
        if latest_parts == current_parts:
            return 'latest'
        else:
            return 'update_needed'

    except requests.exceptions.RequestException as e:
        # 네트워크 오류나 잘못된 API 주소 처리
        print(f"업데이트 확인 실패: {e}")
        return 'error'

def show_force_update_popup():
    """사용자에게 업데이트를 강제하는 모달 팝업을 표시합니다."""
    message = "최신 버전이 아니라서 실행이 불가능합니다.\n홈페이지로 이동하여 최신 버전을 다운로드해주세요."
    ctypes.windll.user32.MessageBoxW(
        0, message, "업데이트 필요",
        win32con.MB_OK | win32con.MB_ICONEXCLAMATION | win32con.MB_TOPMOST
    )
    webbrowser.open_new_tab("https://www.wyea.info")

def show_connection_error_popup():
    """인터넷 연결 오류 팝업을 표시합니다."""
    message = "최신 버전인지 확인이 불가능합니다.\n인터넷 연결 후 사용 바랍니다."
    ctypes.windll.user32.MessageBoxW(
        0, message, "인터넷 연결 확인",
        win32con.MB_OK | win32con.MB_ICONWARNING | win32con.MB_TOPMOST
    )

def is_admin():
    """스크립트가 관리자 권한으로 실행되고 있는지 확인합니다."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def refresh_window_list():
    """열려있는 모든 창의 제목을 가져와 콤보 박스를 업데이트합니다."""
    windows = [title for title in gw.getAllTitles() if title.strip()]
    combo['values'] = windows
    if windows:
        combo.current(0)
    status_label.config(text=f"{len(windows)}개의 창 감지됨")

def pin_window():
    """
    선택한 창을 찾아 최상위에 유지하는 스레드를 시작합니다.
    """
    global running, hwnd, pin_thread

    # 이미 고정된 창이 있다면, 먼저 고정을 해제합니다.
    if running:
        unpin_window()
        time.sleep(0.1)

    target_title = combo.get()
    if not target_title:
        status_label.config(text="선택된 창이 없습니다.")
        return

    windows = gw.getWindowsWithTitle(target_title)
    if not windows:
        status_label.config(text=f"창을 찾을 수 없음: {target_title}")
        return

    win = windows[0]
    hwnd = win._hWnd # 창 핸들(HWND) 가져오기

    # 창이 최소화되어 있다면, 먼저 복원합니다.
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.1)

    running = True
    status_label.config(text=f"'{win.title}' 최상위 유지 중 (Alt+0 또는 Alt+Esc로 해제)")

    def keep_on_top():
        """별도의 스레드에서 실행되며 창을 계속 최상위로 유지하는 핵심 로직."""
        while running:
            try:
                # 창 핸들이 여전히 유효한지 확인
                if not win32gui.IsWindow(hwnd):
                    break
                # 창의 크기나 위치를 바꾸지 않고 최상위 속성만 설정
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            except win32gui.error:
                # 창이 예기치 않게 닫혔을 경우 발생할 수 있음
                break
            time.sleep(0.5) # 0.5초마다 다시 적용

    pin_thread = threading.Thread(target=keep_on_top, daemon=True)
    pin_thread.start()

def unpin_window():
    """고정 스레드를 중지하고 창을 일반 상태로 되돌립니다."""
    global running, hwnd

    if not running:
        status_label.config(text="고정된 창이 없습니다.")
        return

    running = False # 이 변수를 바꾸면 `keep_on_top` 반복문이 멈춤

    if hwnd:
        try:
            if win32gui.IsWindow(hwnd):
                # 창을 최상위가 아닌 상태로 되돌림
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            status_label.config(text="최상위 고정 해제됨.")
        except win32gui.error:
            status_label.config(text="해제 실패 (창이 닫혔거나 권한 문제).")
        finally:
            hwnd = None # 핸들 초기화
    else:
        status_label.config(text="고정된 창이 없습니다.")

def register_hotkeys():
    """언제든지 창 고정을 해제할 수 있도록 전역 단축키를 등록합니다."""
    keyboard.add_hotkey("alt+esc", unpin_window)
    keyboard.add_hotkey("alt+0", unpin_window)

def open_link(event):
    """개발자 웹사이트를 새 브라우저 탭에서 엽니다."""
    webbrowser.open_new_tab("https://www.wyea.info")

# --- 메인 실행 로직 ---
if __name__ == "__main__":
    # 프로그램이 관리자 권한으로 실행되도록 보장. 아닐 경우, 스스로 다시 실행.
    if is_admin():
        # 1단계: 메인 어플리케이션 시작 전, 업데이트 확인
        update_status = check_for_updates()

        if update_status == 'update_needed':
            show_force_update_popup()
            sys.exit()
        elif update_status == 'error':
            show_connection_error_popup()
            sys.exit()

        # 2단계: 버전이 최신일 경우, Tkinter 어플리케이션 생성 및 실행
        root = tk.Tk()
        root.withdraw() # 메인 창을 초기에 숨김

        # 시작 팝업 표시
        ctypes.windll.user32.MessageBoxW(
            0, "Created by WYEA", "AlwaysOnTop",
            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST
        )

        root.deiconify() # 팝업을 닫으면 메인 창 표시
        root.title("AlwaysOnTop v1.0.2")

        # 시작 시 창을 잠시 맨 앞으로 가져옴
        root.attributes('-topmost', True)
        root.after(100, lambda: root.attributes('-topmost', False))

        # --- 창을 화면 정가운데에 위치시키는 로직 ---
        window_width = 400
        window_height = 280
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        position_x = int(screen_width / 2 - window_width / 2)
        position_y = int(screen_height / 2 - window_height / 2)
        root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

        root.resizable(False, False)

        # --- 위젯 생성 및 배치 ---
        frame = tk.Frame(root)
        frame.pack(pady=10, padx=10)

        combo_label = tk.Label(frame, text="현재 열린 창 목록:")
        combo_label.pack()

        combo = ttk.Combobox(frame, width=50)
        combo.pack(pady=5)

        refresh_btn = tk.Button(frame, text="새로고침", command=refresh_window_list)
        refresh_btn.pack(pady=5)

        start_btn = tk.Button(root, text="실행 (지속 고정)", command=pin_window,
                              width=20, height=2, bg="#C8E6C9")
        start_btn.pack(pady=5)

        stop_btn = tk.Button(root, text="정지 (해제)", command=unpin_window,
                             width=20, height=2, bg="#FFCDD2")
        stop_btn.pack(pady=5)

        status_label = tk.Label(root, text="", fg="blue")
        status_label.pack(pady=(10, 0))

        credit_label = tk.Label(root, text="Created by WYEA (https://wyea.info)", fg="gray", cursor="hand2")
        credit_label.pack(pady=(5, 10))
        credit_label.bind("<Button-1>", open_link)

        # UI 요소 초기화
        refresh_window_list()
        register_hotkeys()
        status_label.config(text="Alt+0 또는 Alt+Esc로 강제 정지 가능")

        def prevent_minimize():
            """주기적으로 메인 창이 최소화되었는지 확인하고 복원합니다."""
            try:
                main_hwnd = root.winfo_id()
                if win32gui.IsIconic(main_hwnd):
                    win32gui.ShowWindow(main_hwnd, win32con.SW_RESTORE)
            except tk.TclError:
                # 창이 파괴되는 과정에서 오류가 발생할 수 있음
                return
            root.after(100, prevent_minimize)

        def on_closing():
            """창 닫기(X) 버튼 이벤트를 처리합니다."""
            root.withdraw() # 먼저 메인 창을 숨김
            result = ctypes.windll.user32.MessageBoxW(
                0, "개발자의 최신 소식 확인 및 프로그램 피드백 전달해주실래요?", "프로그램 종료",
                win32con.MB_YESNO | win32con.MB_ICONQUESTION | win32con.MB_TOPMOST
            )
            if result == win32con.IDYES: # 사용자가 '예'를 클릭했을 경우
                webbrowser.open_new_tab("https://www.wyea.info")
            root.destroy()

        # 이벤트 핸들러 설정
        root.after(100, prevent_minimize)
        root.protocol("WM_DELETE_WINDOW", on_closing) # 기본 닫기 버튼의 동작을 재정의

        root.mainloop()
    else:
        # 관리자 권한이 아닐 경우, 관리자 권한으로 스크립트를 다시 실행
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

