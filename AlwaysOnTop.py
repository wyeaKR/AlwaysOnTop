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

# --- 프로그램 정보 ---
CURRENT_VERSION = "1.0.1" # 버전을 v1.0.1로 업데이트
# 중요: GitHub ID를 wyeaKR로 수정했습니다.
GITHUB_API_URL = "https://api.github.com/repos/wyeaKR/AlwaysOnTop/releases/latest"

# --- 상태 관리를 위한 전역 변수 ---
running = False
hwnd = None
pin_thread = None

def check_for_updates():
    """GitHub에서 최신 버전을 확인하고 상태를 반환합니다. (latest, update_needed, error)"""
    try:
        response = requests.get(GITHUB_API_URL, timeout=5)
        response.raise_for_status() # HTTP 에러가 있을 경우 예외 발생

        latest_version_tag = response.json()['tag_name']
        latest_version = latest_version_tag.lstrip('v') # 'v' 접두사 제거

        # --- 수정된 부분: 버전을 숫자로 변환하여 정확하게 비교 ---
        current_parts = list(map(int, CURRENT_VERSION.split('.')))
        latest_parts = list(map(int, latest_version.split('.')))

        # 더 긴 버전 형식에 맞춰 길이를 통일 (예: 1.0 -> 1.0.0)
        while len(current_parts) < len(latest_parts):
            current_parts.append(0)
        while len(latest_parts) < len(current_parts):
            latest_parts.append(0)

        if latest_parts > current_parts:
            return 'update_needed'
        else:
            return 'latest'

    except requests.exceptions.RequestException as e:
        # 인터넷 연결이 없거나 API 주소가 잘못된 경우
        print(f"업데이트 확인 실패: {e}")
        return 'error'

def show_force_update_popup():
    """강제 업데이트 안내 팝업을 띄우고 홈페이지로 이동시킵니다."""
    message = "최신 버전이 아니라서 실행이 불가능합니다.\n홈페이지로 이동하여 최신 버전을 다운로드해주세요."
    ctypes.windll.user32.MessageBoxW(
        0,
        message,
        "업데이트 필요",
        win32con.MB_OK | win32con.MB_ICONEXCLAMATION | win32con.MB_TOPMOST
    )
    webbrowser.open_new_tab("https://www.wyea.info")

def show_connection_error_popup():
    """인터넷 연결 오류 팝업을 띄웁니다."""
    message = "최신 버전인지 확인이 불가능합니다.\n인터넷 연결 후 사용 바랍니다."
    ctypes.windll.user32.MessageBoxW(
        0,
        message,
        "인터넷 연결 확인",
        win32con.MB_OK | win32con.MB_ICONWARNING | win32con.MB_TOPMOST
    )

def is_admin():
    """현재 스크립트가 관리자 권한으로 실행되고 있는지 확인합니다."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def refresh_window_list():
    """현재 열려있는 모든 창의 목록을 가져와 콤보 박스를 업데이트합니다."""
    windows = [title for title in gw.getAllTitles() if title.strip()]
    combo['values'] = windows
    if windows:
        combo.current(0)
    status_label.config(text=f"{len(windows)}개의 창 감지됨")

def pin_window():
    """선택한 창을 감시하며 항상 최상위에 있도록 유지합니다."""
    global running, hwnd, pin_thread

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
    hwnd = win._hWnd

    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.1)

    running = True
    status_label.config(text=f"'{win.title}' 최상위 유지 중 (Alt+0 또는 Alt+Esc로 해제)")

    def keep_on_top():
        while running:
            try:
                if not win32gui.IsWindow(hwnd):
                    break
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST,
                                      0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            except win32gui.error:
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
    """단축키를 등록하여 언제든 고정을 해제할 수 있게 합니다."""
    keyboard.add_hotkey("alt+esc", unpin_window)
    keyboard.add_hotkey("alt+0", unpin_window)

def open_link(event):
    """지정된 URL을 웹 브라우저에서 엽니다."""
    webbrowser.open_new_tab("https://www.wyea.info")

# --- 메인 실행 로직 ---
if is_admin():
    # 실행 전, 강제로 업데이트 확인
    update_status = check_for_updates()

    if update_status == 'update_needed':
        show_force_update_popup()
        sys.exit() # 업데이트가 필요하면 프로그램 종료
    elif update_status == 'error':
        show_connection_error_popup()
        sys.exit() # 확인 실패 시 프로그램 종료

    # 업데이트가 필요 없을 경우에만 프로그램 실행
    root = tk.Tk()
    root.withdraw() # 메인 창을 먼저 숨김

    # 시작 팝업
    ctypes.windll.user32.MessageBoxW(
        0,
        "Created by WYEA",
        "AlwaysOnTop",
        win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST
    )

    root.deiconify() # 팝업 확인 후 메인 창 표시
    root.title("AlwaysOnTop v1.0.1") # 창 제목도 업데이트

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

    root.resizable(False, False)

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
        """프로그램 종료 시 메인 창을 먼저 숨기고 사용자에게 홈페이지 방문 여부를 묻습니다."""
        root.withdraw()
        result = ctypes.windll.user32.MessageBoxW(
            0,
            "개발자의 최신 소식 확인 및 프로그램 피드백 전달해주실래요?",
            "프로그램 종료",
            win32con.MB_YESNO | win32con.MB_ICONQUESTION | win32con.MB_TOPMOST
        )
        if result == win32con.IDYES:
            webbrowser.open_new_tab("https://www.wyea.info")
        root.destroy()

    root.after(100, prevent_minimize)
    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()
else:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

