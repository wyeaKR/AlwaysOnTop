# -*- coding: utf-8 -*-
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

# pywin32 라이브러리가 설치되어 있어야 합니다. (pip install pywin32)
# requests 라이브러리가 설치되어 있어야 합니다. (pip install requests)
# keyboard 라이브러리가 설치되어 있어야 합니다. (pip install keyboard)
# pygetwindow 라이브러리가 설치되어 있어야 합니다. (pip install PyGetWindow)


# --- 프로그램 정보 ---
# 현재 프로그램의 버전. GitHub 릴리즈의 태그와 정확히 일치해야 합니다.
CURRENT_VERSION = "1.0.2"
# GitHub API 주소. 'wyeaKR/AlwaysOnTop' 부분을 본인의 깃허브 ID와 저장소 이름으로 변경해야 합니다.
GITHUB_API_URL = "https://api.github.com/repos/wyeaKR/AlwaysOnTop/releases/latest"

# --- 상태 관리를 위한 전역 변수 ---
running = False  # 창 고정 기능이 활성화되어 있는지 여부
hwnd = None  # 고정할 창의 핸들(고유 ID)
pin_thread = None  # 창을 주기적으로 최상위로 만드는 작업을 수행할 스레드


def check_for_updates():
    """
    GitHub 저장소의 최신 릴리즈 정보를 가져와 현재 버전과 비교합니다.
    - 정확히 일치하면 'latest'를 반환합니다.
    - 일치하지 않으면 'update_needed'를 반환합니다.
    - 오류 발생 시 'error'를 반환합니다.
    """
    try:
        # GitHub API에 최신 릴리즈 정보 요청 (타임아웃 5초)
        response = requests.get(GITHUB_API_URL, timeout=5)
        response.raise_for_status()  # HTTP 에러가 발생하면 예외를 일으킴

        # 응답(JSON)에서 'tag_name' 필드를 가져옴 (예: "v1.0.2")
        latest_version_tag = response.json()['tag_name']
        latest_version = latest_version_tag.lstrip('v')  # 맨 앞의 'v' 문자 제거

        # 버전이 정확히 일치하는지 비교
        if latest_version == CURRENT_VERSION:
            return 'latest'
        else:
            return 'update_needed'

    except requests.exceptions.RequestException as e:
        # 인터넷 연결 문제, API 주소 오류 등 요청 관련 예외 처리
        print(f"업데이트 확인 실패: {e}")
        return 'error'

def show_force_update_popup():
    """사용자에게 업데이트가 필요함을 알리는 팝업창을 띄웁니다."""
    message = "최신 버전이 아니라서 실행이 불가능합니다.\n홈페이지로 이동하여 최신 버전을 다운로드해주세요."
    ctypes.windll.user32.MessageBoxW(
        0, message, "업데이트 필요",
        win32con.MB_OK | win32con.MB_ICONEXCLAMATION | win32con.MB_TOPMOST
    )
    webbrowser.open_new_tab("https://www.wyea.info")

def show_connection_error_popup():
    """버전 확인 실패(인터넷 연결 등)를 알리는 팝업창을 띄웁니다."""
    message = "최신 버전인지 확인이 불가능합니다.\n인터넷 연결 후 사용 바랍니다."
    ctypes.windll.user32.MessageBoxW(
        0, message, "인터넷 연결 확인",
        win32con.MB_OK | win32con.MB_ICONWARNING | win32con.MB_TOPMOST
    )

def is_admin():
    """스크립트가 관리자 권한으로 실행되었는지 확인합니다."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def refresh_window_list():
    """현재 열려있는 모든 창의 목록을 다시 불러와 드롭다운 메뉴를 업데이트합니다."""
    windows = [title for title in gw.getAllTitles() if title.strip()]
    combo['values'] = windows
    if windows:
        combo.current(0)
    status_label.config(text=f"{len(windows)}개의 창 감지됨")

def pin_window():
    """드롭다운 메뉴에서 선택된 창을 최상위에 고정하는 기능을 시작합니다."""
    global running, hwnd, pin_thread

    # 이미 다른 창을 고정 중이라면, 먼저 해제합니다.
    if running:
        unpin_window()
        time.sleep(0.1)

    target_title = combo.get()
    if not target_title:
        status_label.config(text="선택된 창이 없습니다.")
        return

    # 선택한 제목의 창을 찾습니다.
    windows = gw.getWindowsWithTitle(target_title)
    if not windows:
        status_label.config(text=f"창을 찾을 수 없음: {target_title}")
        return

    win = windows[0]
    hwnd = win._hWnd

    # 창이 최소화되어 있다면, 먼저 복원합니다.
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.1)

    running = True
    status_label.config(text=f"'{win.title}' 최상위 유지 중 (Alt+0 또는 Alt+Esc로 해제)")

    # 별도의 스레드에서 창을 주기적으로 최상위로 만드는 작업을 수행합니다.
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
                # 창이 닫히는 등 에러 발생 시 루프 종료
                break
            time.sleep(0.5)

    pin_thread = threading.Thread(target=keep_on_top, daemon=True)
    pin_thread.start()

def unpin_window():
    """창의 최상위 고정을 해제합니다."""
    global running, hwnd

    if not running:
        status_label.config(text="고정된 창이 없음")
        return

    running = False

    if hwnd:
        try:
            if win32gui.IsWindow(hwnd):
                # 창을 최상위가 아닌 상태(보통)로 되돌림
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                      0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            status_label.config(text="최상위 고정 해제됨")
        except win32gui.error:
            status_label.config(text="해제 실패 (창이 닫혔거나 권한 문제)")
        finally:
            hwnd = None  # 핸들 초기화
    else:
        status_label.config(text="고정된 창이 없음")

def register_hotkeys():
    """고정 해제를 위한 전역 단축키를 등록합니다."""
    keyboard.add_hotkey("alt+esc", unpin_window)
    keyboard.add_hotkey("alt+0", unpin_window)

def open_link(event):
    """제작자 정보 링크를 기본 웹 브라우저에서 엽니다."""
    webbrowser.open_new_tab("https://www.wyea.info")


# --- 메인 실행 로직 ---
if __name__ == "__main__":
    # 스크립트가 관리자 권한으로 실행되었는지 확인합니다.
    if is_admin():
        # 실행 전, 강제로 업데이트 확인
        update_status = check_for_updates()

        if update_status == 'update_needed':
            show_force_update_popup()
            sys.exit()  # 업데이트가 필요하면 프로그램 종료
        elif update_status == 'error':
            show_connection_error_popup()
            sys.exit()  # 확인 실패 시 프로그램 종료

        # === GUI 생성 및 설정 ===
        root = tk.Tk()
        root.withdraw()  # 메인 창을 먼저 숨김

        # 시작 팝업
        ctypes.windll.user32.MessageBoxW(
            0, "Created by WYEA", "AlwaysOnTop",
            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST
        )

        root.deiconify()  # 팝업 확인 후 메인 창 표시
        root.title("AlwaysOnTop v1.0.2")

        # 창이 맨 앞에 잠깐 나타났다가 다른 창을 클릭해도 뒤로 가도록 설정
        root.attributes('-topmost', True)
        root.after(100, lambda: root.attributes('-topmost', False))

        # 창을 화면 정가운데에 위치시키는 로직
        window_width = 400
        window_height = 280
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        position_x = int(screen_width / 2 - window_width / 2)
        position_y = int(screen_height / 2 - window_height / 2)
        root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

        root.resizable(False, False) # 창 크기 변경 불가

        # === 위젯 생성 및 배치 ===
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

        # === 초기 설정 및 이벤트 핸들러 등록 ===
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
                # 창이 닫힌 후 Tcl 에러가 발생하는 것을 방지
                return
            root.after(100, prevent_minimize)

        def on_closing():
            """프로그램 종료(X 버튼) 시 메인 창을 먼저 숨기고 사용자에게 확인 팝업을 띄웁니다."""
            root.withdraw()
            result = ctypes.windll.user32.MessageBoxW(
                0, "개발자의 최신 소식 확인 및 프로그램 피드백 전달해주실래요?", "프로그램 종료",
                win32con.MB_YESNO | win32con.MB_ICONQUESTION | win32con.MB_TOPMOST
            )
            # '예'를 눌렀을 때의 반환 값은 6 (IDYES)
            if result == win32con.IDYES:
                webbrowser.open_new_tab("https://www.wyea.info")
            root.destroy()

        # 0.1초 후에 최소화 방지 루틴 시작
        root.after(100, prevent_minimize)
        # 창 닫기 버튼(X)의 기본 동작을 on_closing 함수로 변경
        root.protocol("WM_DELETE_WINDOW", on_closing)

        # GUI 메인 루프 시작
        root.mainloop()
    else:
        # 관리자 권한이 없으면, 권한 상승을 요청하여 스크립트를 다시 실행합니다.
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

