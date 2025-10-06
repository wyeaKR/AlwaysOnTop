# AlwaysOnTop v1.0.2
# Created by WYEA
#
# 이 프로그램은 사용자가 선택한 창을 항상 다른 모든 창들 위에 있도록 고정시켜주는 유틸리티입니다.
# 멀티태스킹 시 동영상, 메모, 계산기 등을 항상 보면서 작업할 수 있어 생산성을 높여줍니다.
#
# [필요 라이브러리 설치]
# pip install pygetwindow keyboard requests pywin32 pyarmor

import tkinter as tk
from tkinter import ttk, messagebox
import pygetwindow as gw
import win32gui
import win32con
import keyboard
import threading
import time
import ctypes
import sys
import webbrowser
import requests # 업데이트 확인을 위해 추가

# --- 프로그램 정보 및 설정 ---
CURRENT_VERSION = "1.0.2"
GITHUB_API_URL = "https://api.github.com/repos/wyeaKR/AlwaysOnTop/releases/latest"

# --- 상태 관리를 위한 전역 변수 ---
running = False       # 고정 기능이 현재 실행 중인지 여부
hwnd = None           # 고정할 창의 핸들(고유 식별자)
pin_thread = None     # 창을 주기적으로 최상위로 만드는 작업을 수행할 스레드

def check_for_updates():
    """
    GitHub 저장소의 최신 릴리즈 버전을 확인합니다.
    - 현재 버전과 정확히 일치하면 'latest' 반환
    - 일치하지 않으면 'update_needed' 반환
    - 확인 중 오류 발생 시 'error' 반환
    """
    try:
        response = requests.get(GITHUB_API_URL, timeout=5)
        response.raise_for_status() # HTTP 에러가 있을 경우 예외 발생

        latest_version_tag = response.json()['tag_name']
        latest_version = latest_version_tag.lstrip('v') # 'v1.0.2' -> '1.0.2'

        # 버전이 정확히 일치하는지 확인
        if latest_version == CURRENT_VERSION:
            return 'latest'
        else:
            return 'update_needed'

    except requests.exceptions.RequestException as e:
        # 인터넷 연결이 없거나 API 주소가 잘못된 경우 등
        print(f"업데이트 확인 실패: {e}")
        return 'error'

def show_force_update_popup():
    """사용자에게 최신 버전 다운로드를 안내하는 팝업을 띄웁니다."""
    message = "최신 버전이 아니라서 실행이 불가능합니다.\n홈페이지로 이동하여 최신 버전을 다운로드해주세요."
    ctypes.windll.user32.MessageBoxW(
        0, message, "업데이트 필요",
        win32con.MB_OK | win32con.MB_ICONEXCLAMATION | win32con.MB_TOPMOST
    )
    webbrowser.open_new_tab("https://www.wyea.info")

def show_connection_error_popup():
    """업데이트 확인 실패 시 인터넷 연결을 요청하는 팝업을 띄웁니다."""
    message = "최신 버전인지 확인이 불가능합니다.\n인터넷 연결 후 사용 바랍니다."
    ctypes.windll.user32.MessageBoxW(
        0, message, "인터넷 연결 확인",
        win32con.MB_OK | win32con.MB_ICONWARNING | win32con.MB_TOPMOST
    )

def is_admin():
    """프로그램이 관리자 권한으로 실행되었는지 확인합니다."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def refresh_window_list():
    """현재 화면에 열려있는 모든 창의 목록을 새로고침하여 드롭다운 메뉴에 표시합니다."""
    windows = [title for title in gw.getAllTitles() if title.strip()]
    combo['values'] = windows
    if windows:
        combo.current(0)
    status_label.config(text=f"{len(windows)}개의 창 감지됨")

def pin_window():
    """'실행' 버튼을 눌렀을 때, 선택된 창을 최상위에 고정시키는 스레드를 시작합니다."""
    global running, hwnd, pin_thread

    # 이미 실행 중인 고정 작업이 있다면 먼저 중지
    if running:
        unpin_window()
        time.sleep(0.1)

    # 드롭다운 메뉴에서 선택된 창 제목 가져오기
    target_title = combo.get()
    if not target_title:
        status_label.config(text="선택된 창이 없습니다.")
        return

    # 창 제목으로 창 객체 찾기
    windows = gw.getWindowsWithTitle(target_title)
    if not windows:
        status_label.config(text=f"창을 찾을 수 없음: {target_title}")
        return

    win = windows[0]
    hwnd = win._hWnd

    # 창이 최소화되어 있다면, 먼저 복원
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.1)

    running = True
    status_label.config(text=f"'{win.title}' 최상위 유지 중 (Alt+0 또는 Alt+Esc로 해제)")

    # 0.5초마다 창을 최상위로 설정하는 작업을 수행할 함수
    def keep_on_top():
        while running:
            try:
                # 창이 여전히 존재하는지 확인
                if not win32gui.IsWindow(hwnd):
                    break
                # 창을 최상위로 설정
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST,
                                      0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            except win32gui.error:
                break
            time.sleep(0.5)

    # 위 함수를 별도의 스레드에서 실행 (UI가 멈추지 않도록)
    pin_thread = threading.Thread(target=keep_on_top, daemon=True)
    pin_thread.start()

def unpin_window():
    """'정지' 버튼 또는 단축키로 창의 최상위 고정을 해제합니다."""
    global running, hwnd

    if not running:
        status_label.config(text="고정된 창이 없음")
        return

    running = False # 고정 루프 중단

    if hwnd:
        try:
            if win32gui.IsWindow(hwnd):
                # 창을 최상위가 아닌 일반 상태로 변경
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                      0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            status_label.config(text="최상위 고정 해제됨")
        except win32gui.error:
            status_label.config(text="해제 실패 (창이 닫혔거나 권한 문제)")
        finally:
            hwnd = None
    else:
        status_label.config(text="고정된 창이 없음")

def register_hotkeys():
    """프로그램 전체에서 사용할 수 있는 단축키를 등록합니다."""
    keyboard.add_hotkey("alt+esc", unpin_window)
    keyboard.add_hotkey("alt+0", unpin_window)

def open_link(event):
    """'Created by WYEA' 라벨을 클릭했을 때 웹사이트를 엽니다."""
    webbrowser.open_new_tab("https://www.wyea.info")

# --- 메인 실행 로직 ---
if is_admin():
    # 1. 프로그램 시작 전, 강제로 업데이트 확인
    update_status = check_for_updates()

    if update_status == 'update_needed':
        show_force_update_popup()
        sys.exit() # 업데이트가 필요하면 프로그램 종료
    elif update_status == 'error':
        show_connection_error_popup()
        sys.exit() # 확인 실패 시 프로그램 종료

    # 2. 업데이트가 필요 없을 경우에만 프로그램 UI 실행
    root = tk.Tk()
    root.withdraw() # 팝업을 먼저 띄우기 위해 메인 창을 잠시 숨김

    # 시작 알림 팝업
    ctypes.windll.user32.MessageBoxW(
        0, "Created by WYEA", "AlwaysOnTop",
        win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST
    )

    root.deiconify() # 팝업 확인 후 메인 창 표시
    root.title("AlwaysOnTop v1.0.2")

    # 잠깐 최상위로 올렸다가 원래대로 돌려서 사용자에게 창을 확실히 보여줌
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    # 창 크기 및 위치 설정 (화면 정가운데)
    window_width = 400
    window_height = 280
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_x = int(screen_width / 2 - window_width / 2)
    position_y = int(screen_height / 2 - window_height / 2)
    root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    root.resizable(False, False) # 창 크기 변경 불가

    # --- UI 위젯 배치 ---
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

    # --- 초기 설정 및 루프 시작 ---
    refresh_window_list()
    register_hotkeys()
    status_label.config(text="Alt+0 또는 Alt+Esc로 강제 정지 가능")

    def prevent_minimize():
        """주기적으로 창 상태를 확인하여 최소화되어 있으면 복원합니다."""
        try:
            main_hwnd = root.winfo_id()
            if win32gui.IsIconic(main_hwnd):
                win32gui.ShowWindow(main_hwnd, win32con.SW_RESTORE)
        except tk.TclError:
            return
        root.after(100, prevent_minimize)

    def on_closing():
        """창의 'X' 버튼을 눌렀을 때의 동작을 정의합니다."""
        root.withdraw() # 메인 창을 먼저 숨김
        result = ctypes.windll.user32.MessageBoxW(
            0, "개발자의 최신 소식 확인 및 프로그램 피드백 전달해주실래요?", "프로그램 종료",
            win32con.MB_YESNO | win32con.MB_ICONQUESTION | win32con.MB_TOPMOST
        )
        if result == win32con.IDYES: # '예'를 눌렀을 때
            webbrowser.open_new_tab("https://www.wyea.info")
        root.destroy()

    root.after(100, prevent_minimize)
    root.protocol("WM_DELETE_WINDOW", on_closing) # 'X' 버튼 동작을 on_closing 함수로 연결

    root.mainloop() # GUI 프로그램 실행
else:
    # 관리자 권한이 없으면, 권한을 요청하며 스스로를 다시 실행
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
