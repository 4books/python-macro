import os
import json
import time
import threading
import schedule
import datetime
import pyautogui
import keyboard
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict, Tuple, Optional
import win32api
import win32con
import win32gui
import ctypes
from ctypes import wintypes
import logging
import traceback

# 화면 안전장치 비활성화
pyautogui.FAILSAFE = False

# 로깅 설정
class DebugLogger:
    def __init__(self):
        self.log_file = os.path.join(os.path.expanduser("~"), "PyMacro", "debug.log")
        os.makedirs(os.path.dirname(self.log_file), exist_ok=True)
        
        # 로거 설정
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
            handlers=[
                logging.FileHandler(self.log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
    def debug(self, msg):
        self.logger.debug(msg)
        
    def info(self, msg):
        self.logger.info(msg)
        
    def warning(self, msg):
        self.logger.warning(msg)
        
    def error(self, msg):
        self.logger.error(msg)
        
    def exception(self, msg):
        self.logger.exception(msg)

class MacroRecorder:
    def __init__(self):
        # 로거 초기화
        self.debug_logger = DebugLogger()
        self.log = self.debug_logger
        
        self.log.info("=" * 50)
        self.log.info("Python 매크로 프로그램 시작")
        self.log.info("=" * 50)
        
        try:
            # 시작 시 관리자 권한 강제 요청
            self.ensure_admin_privileges()
            
            # 기본 디렉토리 설정
            self.base_dir = os.path.join(os.path.expanduser("~"), "PyMacro")
            self.macros_dir = os.path.join(self.base_dir, "macros")
            self.schedules_file = os.path.join(self.base_dir, "schedules.json")
            
            self.log.info(f"기본 디렉토리: {self.base_dir}")
            self.log.info(f"매크로 디렉토리: {self.macros_dir}")
            self.log.info(f"스케줄 파일: {self.schedules_file}")
            
            # 디렉토리 생성
            os.makedirs(self.macros_dir, exist_ok=True)
            
            # 스케줄 관련 락 (데드락 방지)
            self.schedule_lock = threading.Lock()
            self.gui_update_lock = threading.Lock()
            
            # 매크로 및 스케줄 데이터
            self.recorded_macros = []
            self.schedules = []
            self.load_macros()
            self.load_schedules()
            
            # 녹화 상태 변수
            self.is_recording = False
            self.current_macro = None
            self.recording_thread = None
            self.current_events = []
            
            # 후킹 관련
            self.mouse_hook = None
            self.keyboard_hook = None
            self.global_keyboard_hook = None  # F11/F12용 전역 후킹
            self.start_time = 0
            self.hook_method = "none"  # 사용된 후킹 방법 추적
            
            # 스케줄 실행 스레드
            self.schedule_thread = None
            self.is_schedule_running = False
            
            # GUI 초기화
            self.root = None
            self.selected_macro = None
            
            # 전역 키보드 후킹 먼저 설정 (F11/F12용)
            self.setup_global_hotkeys()
            
            self.init_gui()
            
        except Exception as e:
            self.log.exception(f"초기화 중 오류 발생: {str(e)}")
            raise

    def ensure_admin_privileges(self):
        """관리자 권한 강제 요청"""
        try:
            import ctypes, sys
            if not ctypes.windll.shell32.IsUserAnAdmin():
                self.log.info("관리자 권한이 필요합니다. 재시작 중...")
                # UAC 프롬프트와 함께 관리자 권한으로 재실행
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, " ".join(sys.argv), None, 1
                )
                sys.exit()
            self.log.info("관리자 권한으로 실행 중...")
            return True
        except Exception as e:
            self.log.error(f"관리자 권한 요청 실패: {str(e)}")
            return False

    def setup_global_hotkeys(self):
        """F11/F12 전역 핫키 설정"""
        try:
            self.log.info("전역 핫키 설정 시작 (F11: 녹화 시작/중지, F12: 녹화 중지)")
            
            # 방법 1: pynput 사용 (우선)
            try:
                from pynput import keyboard as pynput_keyboard
                
                def on_hotkey_press(key):
                    try:
                        if hasattr(key, 'name'):
                            key_name = key.name
                        else:
                            key_name = str(key)
                        
                        if key_name == 'f11':
                            self.log.info("F11 키 감지 - 녹화 토글")
                            self.toggle_recording_hotkey()
                        elif key_name == 'f12':
                            self.log.info("F12 키 감지 - 녹화 중지")
                            self.stop_recording_hotkey()
                            
                    except Exception as e:
                        self.log.error(f"핫키 처리 오류: {str(e)}")
                
                def on_hotkey_release(key):
                    pass  # 키 릴리즈는 무시
                
                # 전역 키보드 리스너 시작
                self.global_keyboard_listener = pynput_keyboard.Listener(
                    on_press=on_hotkey_press,
                    on_release=on_hotkey_release
                )
                self.global_keyboard_listener.start()
                
                self.log.info("pynput 전역 핫키 설정 성공!")
                return True
                
            except ImportError:
                self.log.warning("pynput이 설치되지 않음. keyboard 라이브러리 시도 중...")
            
            # 방법 2: keyboard 라이브러리 사용
            try:
                def on_f11():
                    self.log.info("F11 키 감지 - 녹화 토글")
                    self.toggle_recording_hotkey()
                
                def on_f12():
                    self.log.info("F12 키 감지 - 녹화 중지")
                    self.stop_recording_hotkey()
                
                # 핫키 등록
                keyboard.add_hotkey('f11', on_f11)
                keyboard.add_hotkey('f12', on_f12)
                
                self.log.info("keyboard 라이브러리 전역 핫키 설정 성공!")
                return True
                
            except Exception as e:
                self.log.error(f"keyboard 라이브러리 핫키 설정 실패: {str(e)}")
            
            # 방법 3: Win32 API 사용 (최후의 수단)
            try:
                self.setup_win32_hotkeys()
                self.log.info("Win32 API 전역 핫키 설정 성공!")
                return True
                
            except Exception as e:
                self.log.error(f"Win32 API 핫키 설정 실패: {str(e)}")
            
            self.log.warning("모든 전역 핫키 설정 방법 실패")
            return False
            
        except Exception as e:
            self.log.exception(f"전역 핫키 설정 중 오류: {str(e)}")
            return False

    def setup_win32_hotkeys(self):
        """Win32 API를 사용한 전역 핫키 설정"""
        try:
            # RegisterHotKey를 사용한 전역 핫키 등록
            import ctypes
            from ctypes import wintypes
            
            # 상수 정의
            MOD_NONE = 0x0000
            WM_HOTKEY = 0x0312
            VK_F11 = 0x7A
            VK_F12 = 0x7B
            
            # 핫키 ID
            HOTKEY_F11 = 1
            HOTKEY_F12 = 2
            
            # 가상의 윈도우 핸들 (메시지 처리용)
            self.hotkey_window = None
            
            # 핫키 등록
            result1 = ctypes.windll.user32.RegisterHotKeyW(None, HOTKEY_F11, MOD_NONE, VK_F11)
            result2 = ctypes.windll.user32.RegisterHotKeyW(None, HOTKEY_F12, MOD_NONE, VK_F12)
            
            if not result1 or not result2:
                raise Exception("핫키 등록 실패")
            
            # 메시지 루프 스레드 시작
            self.hotkey_thread = threading.Thread(target=self._win32_hotkey_loop, daemon=True)
            self.hotkey_thread.start()
            
            self.log.info("Win32 API 핫키 등록 완료")
            
        except Exception as e:
            self.log.error(f"Win32 핫키 설정 실패: {str(e)}")
            raise

    def _win32_hotkey_loop(self):
        """Win32 핫키 메시지 루프"""
        try:
            import ctypes
            from ctypes import wintypes
            
            class MSG(ctypes.Structure):
                _fields_ = [
                    ("hwnd", wintypes.HWND),
                    ("message", wintypes.UINT),
                    ("wParam", wintypes.WPARAM),
                    ("lParam", wintypes.LPARAM),
                    ("time", wintypes.DWORD),
                    ("pt", wintypes.POINT)
                ]
            
            msg = MSG()
            
            while True:
                # 메시지 대기
                bRet = ctypes.windll.user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                
                if bRet == 0:  # WM_QUIT
                    break
                elif bRet == -1:  # 오류
                    self.log.error("핫키 메시지 루프 오류")
                    break
                
                # 핫키 메시지 처리
                if msg.message == 0x0312:  # WM_HOTKEY
                    hotkey_id = msg.wParam
                    if hotkey_id == 1:  # F11
                        self.log.info("Win32 F11 키 감지 - 녹화 토글")
                        threading.Thread(target=self.toggle_recording_hotkey, daemon=True).start()
                    elif hotkey_id == 2:  # F12
                        self.log.info("Win32 F12 키 감지 - 녹화 중지")
                        threading.Thread(target=self.stop_recording_hotkey, daemon=True).start()
                
                # 메시지 처리
                ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
                ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))
                
        except Exception as e:
            self.log.exception(f"Win32 핫키 루프 오류: {str(e)}")

    def toggle_recording_hotkey(self):
        """F11 키로 녹화 시작/중지 토글"""
        try:
            if self.is_recording:
                self.log.info("핫키로 녹화 중지")
                self.stop_recording()
            else:
                self.log.info("핫키로 녹화 시작")
                # 기본 매크로 이름 생성
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                macro_name = f"매크로_{timestamp}"
                
                # GUI에서 매크로 이름 설정
                if self.root and self.root.winfo_exists():
                    self.safe_gui_update(lambda: self.macro_name_entry.delete(0, tk.END))
                    self.safe_gui_update(lambda: self.macro_name_entry.insert(0, macro_name))
                
                self.start_recording(macro_name)
                
                # 상태 메시지 표시
                self.show_recording_notification("녹화 시작", f"매크로 '{macro_name}' 녹화를 시작합니다.\nF11: 녹화 중지, F12: 강제 중지")
                
        except Exception as e:
            self.log.exception(f"핫키 녹화 토글 오류: {str(e)}")

    def stop_recording_hotkey(self):
        """F12 키로 녹화 중지"""
        try:
            if self.is_recording:
                self.log.info("핫키로 녹화 중지")
                self.stop_recording()
                self.show_recording_notification("녹화 중지", "매크로 녹화가 중지되었습니다.")
            else:
                self.log.debug("녹화 중이 아님 - F12 키 무시")
                
        except Exception as e:
            self.log.exception(f"핫키 녹화 중지 오류: {str(e)}")

    def show_recording_notification(self, title, message):
        """녹화 상태 알림 표시 (토스트 알림)"""
        try:
            # 별도 스레드에서 알림 창 표시
            def show_notification():
                try:
                    if self.root and self.root.winfo_exists():
                        # 임시 알림 창 생성
                        notification = tk.Toplevel(self.root)
                        notification.title(title)
                        notification.geometry("400x150")
                        notification.resizable(False, False)
                        
                        # 항상 위에 표시
                        notification.attributes("-topmost", True)
                        
                        # 화면 중앙에 배치
                        notification.update_idletasks()
                        x = (notification.winfo_screenwidth() // 2) - (400 // 2)
                        y = (notification.winfo_screenheight() // 2) - (150 // 2)
                        notification.geometry(f"400x150+{x}+{y}")
                        
                        # 내용 표시
                        ttk.Label(notification, text=title, font=("", 12, "bold")).pack(pady=10)
                        ttk.Label(notification, text=message, justify=tk.CENTER).pack(pady=10)
                        
                        ttk.Button(notification, text="확인", 
                                 command=notification.destroy).pack(pady=10)
                        
                        # 3초 후 자동 닫기
                        notification.after(3000, notification.destroy)
                        
                        self.log.debug(f"알림 표시: {title}")
                        
                except Exception as e:
                    self.log.error(f"알림 표시 오류: {str(e)}")
            
            if self.root and self.root.winfo_exists():
                self.root.after(0, show_notification)
                
        except Exception as e:
            self.log.error(f"알림 표시 실패: {str(e)}")

    def cleanup_global_hotkeys(self):
        """전역 핫키 정리"""
        try:
            self.log.info("전역 핫키 정리 시작")
            
            # pynput 리스너 정리
            if hasattr(self, 'global_keyboard_listener'):
                try:
                    self.global_keyboard_listener.stop()
                    self.log.debug("pynput 전역 키보드 리스너 정리 완료")
                except:
                    pass
            
            # keyboard 라이브러리 핫키 정리
            try:
                keyboard.remove_all_hotkeys()
                self.log.debug("keyboard 라이브러리 핫키 정리 완료")
            except:
                pass
            
            # Win32 핫키 정리
            try:
                import ctypes
                ctypes.windll.user32.UnregisterHotKey(None, 1)  # F11
                ctypes.windll.user32.UnregisterHotKey(None, 2)  # F12
                self.log.debug("Win32 핫키 정리 완료")
            except:
                pass
            
            self.log.info("전역 핫키 정리 완료")
            
        except Exception as e:
            self.log.error(f"전역 핫키 정리 오류: {str(e)}")

    def setup_pynput_hooks(self):
        """pynput 라이브러리를 사용한 후킹 (가장 안정적)"""
        try:
            self.log.debug("pynput 후킹 설정 시작")
            from pynput import mouse, keyboard as pynput_keyboard
            
            def on_move(x, y):
                if self.is_recording:
                    current_time = time.time() - self.start_time
                    self.current_events.append({
                        'type': 'mouse_move', 'x': x, 'y': y, 'time': current_time
                    })

            def on_click(x, y, button, pressed):
                if self.is_recording:
                    current_time = time.time() - self.start_time
                    button_name = str(button).split('.')[-1].lower()
                    event_type = 'mouse_down' if pressed else 'mouse_up'
                    self.current_events.append({
                        'type': event_type, 'button': button_name, 
                        'x': x, 'y': y, 'time': current_time
                    })

            def on_press(key):
                if self.is_recording:
                    current_time = time.time() - self.start_time
                    try:
                        key_name = key.char if hasattr(key, 'char') and key.char else str(key).split('.')[-1].lower()
                    except:
                        key_name = str(key).split('.')[-1].lower()
                    
                    # F11/F12는 녹화에서 제외 (전역 핫키로 처리)
                    if key_name not in ['f11', 'f12']:
                        self.current_events.append({
                            'type': 'key_down', 'key': key_name, 'time': current_time
                        })

            def on_release(key):
                if self.is_recording:
                    current_time = time.time() - self.start_time
                    try:
                        key_name = key.char if hasattr(key, 'char') and key.char else str(key).split('.')[-1].lower()
                    except:
                        key_name = str(key).split('.')[-1].lower()
                    
                    # F11/F12는 녹화에서 제외
                    if key_name not in ['f11', 'f12']:
                        self.current_events.append({
                            'type': 'key_up', 'key': key_name, 'time': current_time
                        })

            # 리스너 생성
            self.mouse_listener = mouse.Listener(on_move=on_move, on_click=on_click)
            self.keyboard_listener = pynput_keyboard.Listener(on_press=on_press, on_release=on_release)
            
            # 리스너 시작
            self.mouse_listener.start()
            self.keyboard_listener.start()
            
            self.hook_method = "pynput"
            self.log.info("pynput 후킹 설정 성공!")
            return True
            
        except ImportError:
            self.log.warning("pynput 라이브러리가 설치되지 않았습니다.")
            return False
        except Exception as e:
            self.log.error(f"pynput 후킹 설정 실패: {str(e)}")
            return False

    def setup_keyboard_hook(self):
        """keyboard 라이브러리를 사용한 후킹"""
        try:
            self.log.debug("keyboard 라이브러리 후킹 설정 시작")
            def on_key_event(e):
                if not self.is_recording:
                    return
                
                current_time = time.time() - self.start_time
                event_type = 'key_down' if e.event_type == keyboard.KEY_DOWN else 'key_up'
                key_name = e.name if hasattr(e, 'name') else str(e)
                
                # F11/F12는 녹화에서 제외
                if key_name.lower() not in ['f11', 'f12']:
                    self.current_events.append({
                        'type': event_type, 'key': key_name.lower(), 'time': current_time
                    })
            
            # 키보드 후킹
            keyboard.hook(on_key_event)
            
            self.hook_method = "keyboard"
            self.log.info("keyboard 라이브러리 후킹 설정 성공!")
            return True
            
        except Exception as e:
            self.log.error(f"keyboard 라이브러리 후킹 설정 실패: {str(e)}")
            return False

    def setup_win32_polling(self):
        """win32api를 사용한 폴링 방식 (최후의 수단)"""
        try:
            self.hook_method = "win32_polling"
            self.log.info("win32 폴링 방식 사용 (제한적 기능)")
            return True
        except Exception as e:
            self.log.error(f"win32 폴링 설정 실패: {str(e)}")
            return False

    def setup_recording_method(self):
        """여러 방법 시도하여 가장 적합한 녹화 방법 설정"""
        methods = [
            ("pynput 라이브러리", self.setup_pynput_hooks),
            ("keyboard 라이브러리", self.setup_keyboard_hook),
            ("win32 폴링", self.setup_win32_polling)
        ]
        
        for method_name, method_func in methods:
            self.log.debug(f"{method_name} 시도 중...")
            try:
                if method_func():
                    self.log.info(f"{method_name} 설정 성공!")
                    return True
            except Exception as e:
                self.log.error(f"{method_name} 설정 중 예외 발생: {str(e)}")
        
        self.log.error("모든 녹화 방법 설정 실패!")
        return False

    def start_polling_recording(self):
        """폴링 방식 녹화 (win32api 사용)"""
        self.log.debug("폴링 방식 녹화 시작")
        last_x, last_y = win32api.GetCursorPos()
        last_buttons = {
            'left': False,
            'right': False,
            'middle': False
        }
        
        while self.is_recording:
            try:
                current_time = time.time() - self.start_time
                
                # 마우스 위치 확인
                current_x, current_y = win32api.GetCursorPos()
                if (current_x, current_y) != (last_x, last_y):
                    self.current_events.append({
                        'type': 'mouse_move', 'x': current_x, 'y': current_y, 'time': current_time
                    })
                    last_x, last_y = current_x, current_y
                
                # 마우스 버튼 상태 확인
                current_buttons = {
                    'left': win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000 != 0,
                    'right': win32api.GetAsyncKeyState(win32con.VK_RBUTTON) & 0x8000 != 0,
                    'middle': win32api.GetAsyncKeyState(win32con.VK_MBUTTON) & 0x8000 != 0
                }
                
                # 버튼 상태 변화 감지
                for button, current_state in current_buttons.items():
                    if current_state != last_buttons[button]:
                        event_type = 'mouse_down' if current_state else 'mouse_up'
                        self.current_events.append({
                            'type': event_type, 'button': button,
                            'x': current_x, 'y': current_y, 'time': current_time
                        })
                        last_buttons[button] = current_state
                
                # CPU 사용량 조절
                time.sleep(0.01)
                
            except Exception as e:
                self.log.error(f"폴링 녹화 오류: {str(e)}")
                break
        
        self.log.debug("폴링 방식 녹화 종료")

    def stop_all_hooks(self):
        """모든 후킹 해제"""
        try:
            self.log.debug(f"후킹 해제 시작 (방법: {self.hook_method})")
            if self.hook_method == "pynput":
                if hasattr(self, 'mouse_listener'):
                    self.mouse_listener.stop()
                if hasattr(self, 'keyboard_listener'):
                    self.keyboard_listener.stop()
            elif self.hook_method == "keyboard":
                # 전역 핫키는 유지하고 녹화용 후킹만 해제
                keyboard.unhook_all()
                # 전역 핫키 재설정
                self.setup_global_hotkeys()
            elif self.hook_method == "win32_polling":
                pass  # 폴링은 스레드 종료로 자동 해제
                
            self.hook_method = "none"
            self.log.info("모든 후킹 해제 완료")
            
        except Exception as e:
            self.log.error(f"후킹 해제 오류: {str(e)}")

    def advanced_click_methods(self, x, y, button="left", action="click"):
        """고급 클릭 방법들"""
        success = False
        
        # 방법 1: SetCursorPos + mouse_event (가장 기본적이고 안정적)
        try:
            win32api.SetCursorPos((x, y))
            time.sleep(0.05)
            
            if button == "left":
                if action in ["click", "down"]:
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    time.sleep(0.05)
                if action in ["click", "up"]:
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            elif button == "right":
                if action in ["click", "down"]:
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                    time.sleep(0.05)
                if action in ["click", "up"]:
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
            elif button == "middle":
                if action in ["click", "down"]:
                    win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
                    time.sleep(0.05)
                if action in ["click", "up"]:
                    win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
            
            success = True
            self.log.debug(f"기본 클릭 성공: {x}, {y}, {button}")
            
        except Exception as e:
            self.log.error(f"기본 클릭 실패: {str(e)}")
        
        if success:
            return True
        
        # 방법 2: SendInput API 사용
        try:
            # INPUT 구조체 정의
            class MOUSEINPUT(ctypes.Structure):
                _fields_ = [
                    ("dx", ctypes.c_long),
                    ("dy", ctypes.c_long),
                    ("mouseData", wintypes.DWORD),
                    ("dwFlags", wintypes.DWORD),
                    ("time", wintypes.DWORD),
                    ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG))
                ]

            class INPUT_UNION(ctypes.Union):
                _fields_ = [("mi", MOUSEINPUT)]

            class INPUT(ctypes.Structure):
                _fields_ = [("type", wintypes.DWORD), ("union", INPUT_UNION)]

            # 화면 좌표를 정규화
            screen_width = ctypes.windll.user32.GetSystemMetrics(0)
            screen_height = ctypes.windll.user32.GetSystemMetrics(1)
            
            normalized_x = int(65535 * (x / screen_width))
            normalized_y = int(65535 * (y / screen_height))
            
            inputs = []
            
            # 마우스 이동
            move_input = INPUT()
            move_input.type = 0  # INPUT_MOUSE
            move_input.union.mi.dx = normalized_x
            move_input.union.mi.dy = normalized_y
            move_input.union.mi.mouseData = 0
            move_input.union.mi.dwFlags = 0x0001 | 0x8000  # MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE
            move_input.union.mi.time = 0
            move_input.union.mi.dwExtraInfo = ctypes.pointer(ctypes.c_ulong(0))
            inputs.append(move_input)
            
            # 버튼 플래그 설정
            if button == "left":
                down_flag = 0x0002  # MOUSEEVENTF_LEFTDOWN
                up_flag = 0x0004    # MOUSEEVENTF_LEFTUP
            elif button == "right":
                down_flag = 0x0008  # MOUSEEVENTF_RIGHTDOWN
                up_flag = 0x0010    # MOUSEEVENTF_RIGHTUP
            elif button == "middle":
                down_flag = 0x0020  # MOUSEEVENTF_MIDDLEDOWN
                up_flag = 0x0040    # MOUSEEVENTF_MIDDLEUP
            else:
                return False
            
            # 클릭 이벤트 추가
            if action in ["click", "down"]:
                down_input = INPUT()
                down_input.type = 0
                down_input.union.mi.dx = normalized_x
                down_input.union.mi.dy = normalized_y
                down_input.union.mi.mouseData = 0
                down_input.union.mi.dwFlags = down_flag
                down_input.union.mi.time = 0
                down_input.union.mi.dwExtraInfo = ctypes.pointer(ctypes.c_ulong(0))
                inputs.append(down_input)
            
            if action in ["click", "up"]:
                up_input = INPUT()
                up_input.type = 0
                up_input.union.mi.dx = normalized_x
                up_input.union.mi.dy = normalized_y
                up_input.union.mi.mouseData = 0
                up_input.union.mi.dwFlags = up_flag
                up_input.union.mi.time = 0
                up_input.union.mi.dwExtraInfo = ctypes.pointer(ctypes.c_ulong(0))
                inputs.append(up_input)
            
            # 입력 전송
            input_array = (INPUT * len(inputs))(*inputs)
            result = ctypes.windll.user32.SendInput(len(inputs), ctypes.pointer(input_array), ctypes.sizeof(INPUT))
            
            success = (result == len(inputs))
            if success:
                self.log.debug(f"SendInput 클릭 성공: {x}, {y}, {button}")
            else:
                self.log.warning(f"SendInput 클릭 실패: {x}, {y}, {button}")
                
        except Exception as e:
            self.log.error(f"SendInput 방법 실패: {str(e)}")
        
        if success:
            return True
        
        # 방법 3: PyAutoGUI 사용 (최후의 수단)
        try:
            pyautogui.FAILSAFE = False
            if action == "click":
                pyautogui.click(x=x, y=y, button=button)
            elif action == "down":
                pyautogui.mouseDown(x=x, y=y, button=button)
            elif action == "up":
                pyautogui.mouseUp(x=x, y=y, button=button)
            
            success = True
            self.log.debug(f"PyAutoGUI 클릭 성공: {x}, {y}, {button}")
            
        except Exception as e:
            self.log.error(f"PyAutoGUI 클릭 실패: {str(e)}")
        
        return success

    def start_recording(self, macro_name: str):
        """매크로 녹화 시작"""
        try:
            self.log.info(f"매크로 녹화 시작 요청: {macro_name}")
            
            if self.is_recording:
                self.log.warning("이미 녹화 중입니다.")
                return
                
            self.is_recording = True
            self.current_macro = macro_name
            self.current_events = []
            self.start_time = time.time()
            
            # 녹화 방법 설정
            if not self.setup_recording_method():
                self.log.error("녹화 방법 설정 실패")
                messagebox.showerror("오류", 
                    "녹화 방법 설정에 실패했습니다.\n\n해결 방법:\n"
                    "1. 안티바이러스 프로그램을 일시 중지해보세요\n"
                    "2. Windows Defender 실시간 보호를 일시 중지해보세요\n"
                    "3. 다음 명령어로 라이브러리를 설치해보세요:\n"
                    "   pip install pynput\n"
                    "4. 프로그램을 관리자 권한으로 재실행해보세요")
                self.is_recording = False
                return
            
            # GUI 업데이트
            self.log.debug("녹화 시작 GUI 업데이트")
            self.safe_gui_update(lambda: self.update_recording_buttons_gui(True))
            
            # win32 폴링 방식인 경우 별도 스레드 시작
            if self.hook_method == "win32_polling":
                self.recording_thread = threading.Thread(target=self.start_polling_recording)
                self.recording_thread.daemon = True
                self.recording_thread.start()
                self.log.debug("폴링 스레드 시작")
            
            self.log.info(f"매크로 녹화 시작 완료: {macro_name}")
            
        except Exception as e:
            self.log.exception(f"녹화 시작 중 오류: {str(e)}")
            self.is_recording = False

    def stop_recording(self):
        """매크로 녹화 중지"""
        try:
            self.log.info("매크로 녹화 중지 요청")
            
            if not self.is_recording:
                self.log.warning("녹화 중이 아닙니다.")
                return
                
            self.is_recording = False
            
            # 후킹 해제
            self.stop_all_hooks()
            
            # 녹화 스레드 종료 대기
            if self.recording_thread and self.recording_thread.is_alive():
                self.log.debug("녹화 스레드 종료 대기 중...")
                self.recording_thread.join(1)
            
            # 매크로 저장
            if self.current_events:
                self.save_recorded_macro()
                self.log.info(f"녹화된 이벤트 수: {len(self.current_events)}")
            else:
                self.log.warning("녹화된 이벤트가 없습니다.")
            
            # 매크로 목록 새로고침
            self.load_macros()
            self.safe_gui_update(self.update_macro_list)
            
            # GUI 업데이트
            self.log.debug("녹화 중지 GUI 업데이트")
            self.safe_gui_update(lambda: self.update_recording_buttons_gui(False))
            
            self.log.info("매크로 녹화 중지 완료")
            
        except Exception as e:
            self.log.exception(f"녹화 중지 중 오류: {str(e)}")

    def safe_gui_update(self, update_func):
        """안전한 GUI 업데이트 (데드락 방지)"""
        try:
            if self.root and self.root.winfo_exists():
                # 메인 스레드에서 실행되는지 확인
                import threading
                if threading.current_thread() == threading.main_thread():
                    # 메인 스레드면 직접 실행
                    update_func()
                else:
                    # 다른 스레드면 after로 예약
                    self.root.after(0, update_func)
                self.log.debug("GUI 업데이트 완료")
            else:
                self.log.warning("GUI가 존재하지 않음 - 업데이트 건너뜀")
        except Exception as e:
            self.log.error(f"GUI 업데이트 오류: {str(e)}")

    def update_scheduler_buttons_gui(self, is_running):
        """스케줄러 버튼 상태 업데이트"""
        try:
            self.log.debug(f"스케줄러 버튼 상태 업데이트 시작: is_running={is_running}")
            
            if not self.root or not self.root.winfo_exists():
                self.log.warning("GUI가 존재하지 않음 - 버튼 업데이트 실패")
                return
            
            if is_running:
                # 스케줄러 실행 중
                self.scheduler_status.set("스케줄러 실행 중")
                self.start_sched_btn.config(state=tk.DISABLED)
                self.stop_sched_btn.config(state=tk.NORMAL)
                self.log.debug("스케줄러 실행 상태로 GUI 업데이트 완료")
            else:
                # 스케줄러 중지됨
                self.scheduler_status.set("스케줄러 중지됨")
                self.start_sched_btn.config(state=tk.NORMAL)
                self.stop_sched_btn.config(state=tk.DISABLED)
                self.log.debug("스케줄러 중지 상태로 GUI 업데이트 완료")
                
        except Exception as e:
            self.log.exception(f"스케줄러 버튼 상태 업데이트 오류: {str(e)}")

    def update_recording_buttons_gui(self, is_recording):
        """녹화 버튼 상태 업데이트"""
        try:
            self.log.debug(f"녹화 버튼 상태 업데이트 시작: is_recording={is_recording}")
            
            if not self.root or not self.root.winfo_exists():
                self.log.warning("GUI가 존재하지 않음 - 버튼 업데이트 실패")
                return
            
            if is_recording:
                # 녹화 중
                self.recording_status.set(f"녹화 중... (방법: {self.hook_method}) - F11/F12로 중지")
                self.start_rec_btn.config(state=tk.DISABLED)
                self.stop_rec_btn.config(state=tk.NORMAL)
                self.log.debug("녹화 중 상태로 GUI 업데이트 완료")
            else:
                # 녹화 중지됨
                self.recording_status.set("녹화 중지됨")
                self.start_rec_btn.config(state=tk.NORMAL)
                self.stop_rec_btn.config(state=tk.DISABLED)
                self.log.debug("녹화 중지 상태로 GUI 업데이트 완료")
                
        except Exception as e:
            self.log.exception(f"녹화 버튼 상태 업데이트 오류: {str(e)}")

    def play_macro(self, macro_path: str):
        """매크로 재생"""
        try:
            self.log.info(f"매크로 재생 시작: {macro_path}")
            
            # 매크로 파일 로드
            with open(macro_path, "r", encoding="utf-8") as f:
                macro_data = json.load(f)
            
            events = macro_data["events"]
            if not events:
                self.log.warning("재생할 이벤트가 없습니다.")
                messagebox.showinfo("알림", "재생할 이벤트가 없습니다.")
                return
            
            # 상태 업데이트
            macro_name = macro_data["name"]
            self.safe_gui_update(lambda: setattr(self.status_var, 'value', f"매크로 '{macro_name}' 실행 중..."))
            
<<<<<<< HEAD
            # 이벤트를 시간순으로 정렬
=======
            # 마우스 이동 시간을 직접 설정 (0.005초로 설정)
            mouse_move_duration = 0.005
            
            # 키보드 입력 사이의 지연 시간 (키보드 입력을 더 안정적으로)
            keyboard_delay = 0.1  # 100ms 지연
            
            # 대기 시간 압축을 제거하여 실제 기록된 시간대로 실행
            # speed_multiplier를 1.0으로 설정하면 원래 속도대로 실행됨
            speed_multiplier = 1.0  
            
            # 최대 대기 시간 (이전에는 0.1초로 제한했으나 제한 제거)
            max_wait_time = None  # 제한 없음
            
            # 키보드 상태 추적을 위한 딕셔너리
            pressed_keys = set()
            
            # 특수 키 매핑
            special_keys = {
                'shift': 'shift',
                'ctrl': 'ctrl',
                'alt': 'alt',
                'hangul': 'hangul',  # 한/영 키
                'han_yeong': 'hangul',  # 한/영 키 대체 이름
                'hanyeong': 'hangul',   # 한/영 키 대체 이름
                'ralt': 'hangul',  # 오른쪽 Alt 키
                'rctrl': 'ctrlright',  # 오른쪽 Ctrl 키
                'rshift': 'shiftright',  # 오른쪽 Shift 키
                'lalt': 'alt',  # 왼쪽 Alt 키
                'lctrl': 'ctrl',  # 왼쪽 Ctrl 키
                'lshift': 'shift',  # 왼쪽 Shift 키
                'capslock': 'capslock',  # CapsLock 키
                'esc': 'escape',  # Escape 키
                'space': 'space',  # Space 키
                'tab': 'tab',  # Tab 키
                'enter': 'enter',  # Enter 키
                'backspace': 'backspace',  # Backspace 키
                'delete': 'delete',  # Delete 키
                'insert': 'insert',  # Insert 키
                'home': 'home',  # Home 키
                'end': 'end',  # End 키
                'pageup': 'pageup',  # Page Up 키
                'pagedown': 'pagedown',  # Page Down 키
                'up': 'up',  # 위쪽 화살표
                'down': 'down',  # 아래쪽 화살표
                'left': 'left',  # 왼쪽 화살표
                'right': 'right',  # 오른쪽 화살표
            }
            
            # 마우스 위치 최적화를 위한 변수
            last_mouse_x, last_mouse_y = None, None
            
            # 마우스/키보드 이벤트 처리 (최적화 제거하고 원래 순서대로 실행)
            # 모든 이벤트를 시간순으로 정렬
>>>>>>> 6f0e4cc36dc834c8abdfd7317afaa57a72fdbb86
            sorted_events = sorted(events, key=lambda e: e["time"])
            
            if sorted_events:
                start_time = sorted_events[0]["time"]
                last_event_time = start_time
                
                self.log.info(f"매크로 실행 시작: {len(sorted_events)}개 이벤트")
                
                # 이벤트 처리 루프
                for i, event in enumerate(sorted_events):
                    if not self.root or not self.root.winfo_exists():
                        break
                    
                    # 대기 시간 계산 및 적용
                    wait_time = event["time"] - last_event_time
                    if wait_time > 0.01:
                        time.sleep(wait_time)
                    
                    # 이벤트 처리
                    event_type = event["type"]
                    
                    if event_type == "mouse_move":
                        x, y = event["x"], event["y"]
                        try:
                            win32api.SetCursorPos((x, y))
                        except:
                            pass
                    
                    elif event_type in ["mouse_down", "mouse_up"]:
                        x, y = event["x"], event["y"]
                        button = event["button"]
                        action = "down" if event_type == "mouse_down" else "up"
                        
                        # 고급 클릭 방법 사용
                        if not self.advanced_click_methods(x, y, button, action):
                            self.log.warning(f"클릭 실패: {x}, {y}, {button}, {action}")
                    
                    elif event_type == "key_down":
                        key = event["key"].lower()
                        key_code = self.get_virtual_keycode(key)
                        
                        if key_code is not None:
                            try:
                                win32api.keybd_event(key_code, 0, 0, 0)
                            except:
                                pass
                    
                    elif event_type == "key_up":
                        key = event["key"].lower()
                        key_code = self.get_virtual_keycode(key)
                        
                        if key_code is not None:
                            try:
                                win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                            except:
                                pass
                    
                    last_event_time = event["time"]
                    
                    # 진행률 표시 (매 10개 이벤트마다)
                    if i % 10 == 0:
                        progress = int((i / len(sorted_events)) * 100)
                        self.safe_gui_update(lambda p=progress, n=macro_name: 
                            setattr(self.status_var, 'value', f"매크로 '{n}' 실행 중... ({p}%)"))
            
            # 상태 업데이트
            self.safe_gui_update(lambda: setattr(self.status_var, 'value', f"매크로 '{macro_data['name']}' 실행 완료"))
            self.log.info("매크로 실행 완료")
        
        except Exception as e:
            error_msg = f"매크로 실행 오류: {str(e)}"
            self.log.exception(error_msg)
            self.safe_gui_update(lambda: setattr(self.status_var, 'value', error_msg))

    def get_virtual_keycode(self, key: str):
        """키 이름에서 가상 키코드 가져오기"""
        special_keys = {
            'shift': win32con.VK_SHIFT,
            'ctrl': win32con.VK_CONTROL,
            'alt': win32con.VK_MENU,
            'hangul': 0x15,
            'han_yeong': 0x15,
            'hanyeong': 0x15,
            'ralt': 0x15,
            'rctrl': win32con.VK_RCONTROL,
            'rshift': win32con.VK_RSHIFT,
            'lalt': win32con.VK_LMENU,
            'lctrl': win32con.VK_LCONTROL,
            'lshift': win32con.VK_LSHIFT,
            'capslock': win32con.VK_CAPITAL,
            'esc': win32con.VK_ESCAPE,
            'escape': win32con.VK_ESCAPE,
            'space': win32con.VK_SPACE,
            'tab': win32con.VK_TAB,
            'enter': win32con.VK_RETURN,
            'backspace': win32con.VK_BACK,
            'delete': win32con.VK_DELETE,
            'insert': win32con.VK_INSERT,
            'home': win32con.VK_HOME,
            'end': win32con.VK_END,
            'pageup': win32con.VK_PRIOR,
            'pagedown': win32con.VK_NEXT,
            'page_up': win32con.VK_PRIOR,
            'page_down': win32con.VK_NEXT,
            'up': win32con.VK_UP,
            'down': win32con.VK_DOWN,
            'left': win32con.VK_LEFT,
            'right': win32con.VK_RIGHT,
        }
        
        if key in special_keys:
            return special_keys[key]
        
        if len(key) == 1:
            return win32api.VkKeyScan(key) & 0xff
        
        if key.startswith('f') and len(key) <= 3:
            try:
                num = int(key[1:])
                if 1 <= num <= 12:
                    return win32con.VK_F1 + (num - 1)
            except:
                pass
        
        return None

    def save_recorded_macro(self):
        """녹화된 매크로를 파일로 저장"""
        try:
            self.log.debug(f"매크로 저장 시작: {self.current_macro}")
            
            if not self.current_events:
                self.log.warning("저장할 이벤트가 없습니다.")
                return
                
            macro_data = {
                "name": self.current_macro,
                "created": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "events": self.current_events,
                "hook_method": self.hook_method
            }
            
            filename = f"{self.current_macro.replace(' ', '_')}.json"
            file_path = os.path.join(self.macros_dir, filename)
            
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(macro_data, f, ensure_ascii=False, indent=2)
            
            self.log.info(f"매크로 저장 완료: {file_path}")
            
        except Exception as e:
            self.log.exception(f"매크로 저장 중 오류: {str(e)}")

    def load_macros(self):
        """저장된 매크로 파일 목록 로드"""
        try:
            self.log.debug("매크로 목록 로드 시작")
            self.recorded_macros = []
            
            if os.path.exists(self.macros_dir):
                for file in os.listdir(self.macros_dir):
                    if file.endswith(".json"):
                        try:
                            macro_path = os.path.join(self.macros_dir, file)
                            macro_name = file[:-5]  # .json 제거
                            
                            created_time = datetime.datetime.fromtimestamp(
                                os.path.getctime(macro_path)
                            ).strftime("%Y-%m-%d %H:%M:%S")
                            
                            self.recorded_macros.append({
                                "name": macro_name,
                                "file": file,
                                "created": created_time,
                                "path": macro_path
                            })
                        except Exception as e:
                            self.log.error(f"매크로 파일 로드 오류 ({file}): {str(e)}")
            
            self.log.info(f"매크로 {len(self.recorded_macros)}개 로드 완료")
            
        except Exception as e:
            self.log.exception(f"매크로 목록 로드 중 오류: {str(e)}")
    
    def load_schedules(self):
        """저장된 스케줄 목록 로드"""
        try:
            self.log.debug("스케줄 목록 로드 시작")
            
            if os.path.exists(self.schedules_file):
                try:
                    with open(self.schedules_file, "r", encoding="utf-8") as f:
                        self.schedules = json.load(f)
                    self.log.info(f"스케줄 {len(self.schedules)}개 로드 완료")
                except json.JSONDecodeError as e:
                    self.log.error(f"스케줄 파일 JSON 파싱 오류: {str(e)}")
                    self.schedules = []
                except Exception as e:
                    self.log.error(f"스케줄 파일 읽기 오류: {str(e)}")
                    self.schedules = []
            else:
                self.schedules = []
                self.log.info("스케줄 파일이 없습니다. 빈 목록으로 초기화.")
                
        except Exception as e:
            self.log.exception(f"스케줄 목록 로드 중 오류: {str(e)}")
            self.schedules = []
    
    def save_schedules(self):
        """스케줄 정보 저장"""
        try:
            self.log.debug("스케줄 저장 시작")
            
            # 백업 생성
            if os.path.exists(self.schedules_file):
                backup_file = self.schedules_file + ".backup"
                try:
                    import shutil
                    shutil.copy2(self.schedules_file, backup_file)
                    self.log.debug(f"스케줄 백업 생성: {backup_file}")
                except Exception as e:
                    self.log.warning(f"스케줄 백업 생성 실패: {str(e)}")
            
            # 임시 파일에 먼저 저장
            temp_file = self.schedules_file + ".tmp"
            with open(temp_file, "w", encoding="utf-8") as f:
                json.dump(self.schedules, f, ensure_ascii=False, indent=2)
            
            # 원자적 교체
            if os.path.exists(self.schedules_file):
                os.remove(self.schedules_file)
            os.rename(temp_file, self.schedules_file)
            
            self.log.info(f"스케줄 저장 완료: {len(self.schedules)}개")
            
        except Exception as e:
            self.log.exception(f"스케줄 저장 중 오류: {str(e)}")
            # 임시 파일 정리
            temp_file = self.schedules_file + ".tmp"
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass
    
    def delete_macro(self, macro_path: str):
        """매크로 삭제"""
        try:
            self.log.info(f"매크로 삭제 요청: {macro_path}")
            
            if os.path.exists(macro_path):
                os.remove(macro_path)
                self.log.info(f"매크로 파일 삭제 완료: {macro_path}")
                
            # 매크로가 삭제되었으므로 관련 스케줄도 삭제
            macro_name = os.path.basename(macro_path)[:-5]  # .json 제거
            old_count = len(self.schedules)
            
            with self.schedule_lock:
                self.schedules = [s for s in self.schedules 
                                if s["macro"] != macro_name]
                
            new_count = len(self.schedules)
            if old_count != new_count:
                self.log.info(f"관련 스케줄 {old_count - new_count}개 삭제")
                self.save_schedules()
            
            # 목록 새로고침
            self.load_macros()
            self.safe_gui_update(self.update_macro_list)
            self.safe_gui_update(self.update_schedule_list)
            
            self.log.info("매크로 삭제 완료")
            return True
            
        except Exception as e:
            error_msg = f"매크로 삭제 오류: {str(e)}"
            self.log.exception(error_msg)
            messagebox.showerror("오류", error_msg)
            return False
    
    def add_schedule(self, macro_name: str, time_str: str):
        """스케줄 추가"""
        try:
            self.log.info(f"스케줄 추가 요청: {macro_name} at {time_str}")
            
            # 시간 형식 확인 (HH:MM)
            try:
                hour, minute = map(int, time_str.split(":"))
                if not (0 <= hour < 24 and 0 <= minute < 60):
                    raise ValueError("시간 범위가 올바르지 않습니다.")
                self.log.debug(f"시간 형식 검증 완료: {hour:02d}:{minute:02d}")
            except ValueError as e:
                error_msg = f"시간 형식이 올바르지 않습니다: {str(e)}"
                self.log.error(error_msg)
                messagebox.showerror("오류", error_msg)
                return False
                
            # 스케줄 ID 생성 (현재 시간 + 랜덤 요소)
            import random
            schedule_id = f"{int(time.time())}_{random.randint(1000, 9999)}"
            self.log.debug(f"스케줄 ID 생성: {schedule_id}")
            
            # 새 스케줄 생성
            new_schedule = {
                "id": schedule_id,
                "macro": macro_name,
                "time": time_str,
                "created": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # 스케줄 목록에 안전하게 추가
            with self.schedule_lock:
                self.schedules.append(new_schedule)
                self.log.debug(f"스케줄 목록에 추가: {new_schedule}")
                
                # 파일 저장
                self.save_schedules()
            
            # 실행 중인 경우 스케줄 갱신
            if self.is_schedule_running:
                self.log.debug("스케줄러 실행 중이므로 스케줄 갱신")
                threading.Thread(target=self.update_scheduler_safe, daemon=True).start()
            
            # 목록 새로고침
            self.safe_gui_update(self.update_schedule_list)
            
            self.log.info(f"스케줄 추가 완료: {macro_name} at {time_str}")
            return True
            
        except Exception as e:
            error_msg = f"스케줄 추가 오류: {str(e)}"
            self.log.exception(error_msg)
            messagebox.showerror("오류", error_msg)
            return False
    
    def delete_schedule(self, schedule_id: str):
        """스케줄 삭제"""
        try:
            self.log.info(f"스케줄 삭제 요청: {schedule_id}")
            
            # ID로 스케줄 찾기 및 삭제
            old_count = len(self.schedules)
            with self.schedule_lock:
                self.schedules = [s for s in self.schedules if s["id"] != schedule_id]
                new_count = len(self.schedules)
                
                if old_count == new_count:
                    self.log.warning(f"삭제할 스케줄을 찾을 수 없음: {schedule_id}")
                    return False
                
                self.log.debug(f"스케줄 삭제됨: {old_count} -> {new_count}")
                
                # 파일 저장
                self.save_schedules()
            
            # 실행 중인 경우 스케줄 갱신
            if self.is_schedule_running:
                self.log.debug("스케줄러 실행 중이므로 스케줄 갱신")
                threading.Thread(target=self.update_scheduler_safe, daemon=True).start()
            
            # 목록 새로고침
            self.safe_gui_update(self.update_schedule_list)
            
            self.log.info(f"스케줄 삭제 완료: {schedule_id}")
            return True
            
        except Exception as e:
            error_msg = f"스케줄 삭제 오류: {str(e)}"
            self.log.exception(error_msg)
            messagebox.showerror("오류", error_msg)
            return False
    
    def start_scheduler(self):
        """스케줄러 시작"""
        try:
            self.log.info("스케줄러 시작 요청")
            
            if self.is_schedule_running:
                self.log.warning("스케줄러가 이미 실행 중입니다.")
                return
                
            self.log.debug("스케줄러 상태 변경: False -> True")
            self.is_schedule_running = True
            
            # GUI 업데이트 먼저 실행
            self.log.debug("스케줄러 GUI 업데이트 시작")
            self.safe_gui_update(lambda: self.update_scheduler_buttons_gui(True))
            
            # 스케줄 설정
            self.log.debug("스케줄러 갱신 스레드 시작")
            threading.Thread(target=self.update_scheduler_safe, daemon=True).start()
            
            # 스케줄러 스레드 시작
            self.log.debug("스케줄러 메인 스레드 시작")
            self.schedule_thread = threading.Thread(target=self._run_scheduler)
            self.schedule_thread.daemon = True
            self.schedule_thread.start()
            
            self.log.info("스케줄러 시작 완료")
            
        except Exception as e:
            self.log.exception(f"스케줄러 시작 중 오류: {str(e)}")
            self.is_schedule_running = False
            # 오류 시 GUI도 원상복구
            self.safe_gui_update(lambda: self.update_scheduler_buttons_gui(False))
    
    def stop_scheduler(self):
        """스케줄러 중지"""
        try:
            self.log.info("스케줄러 중지 요청")
            
            if not self.is_schedule_running:
                self.log.warning("스케줄러가 실행 중이 아닙니다.")
                return
                
            self.log.debug("스케줄러 상태 변경: True -> False")
            self.is_schedule_running = False
            
            # GUI 업데이트 먼저 실행
            self.log.debug("스케줄러 중지 GUI 업데이트 시작")
            self.safe_gui_update(lambda: self.update_scheduler_buttons_gui(False))
            
            # 모든 스케줄 작업 취소
            try:
                schedule.clear()
                self.log.debug("스케줄 작업 모두 취소")
            except Exception as e:
                self.log.error(f"스케줄 작업 취소 오류: {str(e)}")
            
            # 스케줄러 스레드 종료 대기
            if self.schedule_thread and self.schedule_thread.is_alive():
                self.log.debug("스케줄러 스레드 종료 대기 중...")
                self.schedule_thread.join(timeout=2)
                if self.schedule_thread.is_alive():
                    self.log.warning("스케줄러 스레드가 정상 종료되지 않음")
                else:
                    self.log.debug("스케줄러 스레드 종료 완료")
            
            self.log.info("스케줄러 중지 완료")
            
        except Exception as e:
            self.log.exception(f"스케줄러 중지 중 오류: {str(e)}")
    
    def update_scheduler_safe(self):
        """안전한 스케줄 작업 갱신"""
        try:
            self.log.debug("스케줄러 갱신 시작")
            
            # 기존 스케줄 모두 취소
            schedule.clear()
            self.log.debug("기존 스케줄 모두 취소")
            
            # 스케줄 목록 복사 (동시성 문제 방지)
            with self.schedule_lock:
                schedules_copy = self.schedules.copy()
            
            # 각 스케줄 등록
            registered_count = 0
            for sched in schedules_copy:
                try:
                    macro_name = sched["macro"]
                    time_str = sched["time"]
                    
                    self.log.debug(f"스케줄 등록 시도: {macro_name} at {time_str}")
                    
                    # 매크로 파일 경로 찾기
                    macro_path = None
                    for macro in self.recorded_macros:
                        if macro["name"] == macro_name:
                            macro_path = macro["path"]
                            break
                    
                    if macro_path and os.path.exists(macro_path):
                        # 시간 파싱
                        try:
                            hour, minute = map(int, time_str.split(":"))
                            
                            # 스케줄 등록 (매일 반복)
                            schedule.every().day.at(time_str).do(
                                self.play_macro_scheduled, macro_path=macro_path
                            )
                            registered_count += 1
                            self.log.debug(f"스케줄 등록 완료: {macro_name} at {time_str}")
                            
                        except ValueError as e:
                            self.log.error(f"시간 파싱 오류 ({time_str}): {str(e)}")
                    else:
                        self.log.warning(f"매크로 파일을 찾을 수 없음: {macro_name}")
                        
                except Exception as e:
                    self.log.error(f"개별 스케줄 등록 오류: {str(e)}")
            
            self.log.info(f"스케줄 갱신 완료: {registered_count}/{len(schedules_copy)}개 등록")
            
        except Exception as e:
            self.log.exception(f"스케줄러 갱신 중 오류: {str(e)}")
    
    def update_scheduler(self):
        """스케줄 작업 갱신 (기존 호환성)"""
        threading.Thread(target=self.update_scheduler_safe, daemon=True).start()
    
    def play_macro_scheduled(self, macro_path: str):
        """스케줄에 의해 매크로 실행"""
        try:
            self.log.info(f"스케줄된 매크로 실행: {macro_path}")
            
            # 별도 스레드에서 매크로 실행
            threading.Thread(
                target=self.play_macro, 
                args=(macro_path,),
                daemon=True
            ).start()
            
            return True
            
        except Exception as e:
            self.log.exception(f"스케줄된 매크로 실행 오류: {str(e)}")
            return False
    
    def _run_scheduler(self):
        """스케줄러 실행 루프"""
        self.log.info("스케줄러 루프 시작")
        
        while self.is_schedule_running:
            try:
                # 예약된 작업 실행
                schedule.run_pending()
                
                # CPU 사용량 감소
                time.sleep(1)
                
            except Exception as e:
                self.log.error(f"스케줄러 루프 오류: {str(e)}")
                time.sleep(5)  # 오류 시 잠시 대기
        
        self.log.info("스케줄러 루프 종료")
    
    def update_macro_list(self):
        """매크로 목록 업데이트"""
        try:
            if not self.root or not self.root.winfo_exists():
                return
                
            self.log.debug("매크로 목록 GUI 업데이트")
            
            # 테이블 초기화
            for item in self.macro_treeview.get_children():
                self.macro_treeview.delete(item)
                
            # 매크로 목록 데이터 추가
            for macro in self.recorded_macros:
                self.macro_treeview.insert("", tk.END, values=(
                    macro["name"],
                    macro["created"]
                ))
            
            self.log.debug(f"매크로 목록 업데이트 완료: {len(self.recorded_macros)}개")
            
        except Exception as e:
            self.log.error(f"매크로 목록 업데이트 오류: {str(e)}")
    
    def update_schedule_list(self):
        """스케줄 목록 업데이트"""
        try:
            if not self.root or not self.root.winfo_exists():
                return
                
            self.log.debug("스케줄 목록 GUI 업데이트")
            
            # 테이블 초기화
            for item in self.schedule_treeview.get_children():
                self.schedule_treeview.delete(item)
                
            # 스케줄 목록 데이터 추가
            with self.schedule_lock:
                for sched in self.schedules:
                    self.schedule_treeview.insert("", tk.END, values=(
                        sched["macro"],
                        sched["time"],
                        sched["created"],
                        sched["id"]
                    ))
            
            self.log.debug(f"스케줄 목록 업데이트 완료: {len(self.schedules)}개")
            
        except Exception as e:
            self.log.error(f"스케줄 목록 업데이트 오류: {str(e)}")
    
    def show_debug_info(self):
        """디버그 정보 표시"""
        try:
            info = f"""
디버그 정보:
- 관리자 권한: {'예' if ctypes.windll.shell32.IsUserAnAdmin() else '아니오'}
- 현재 후킹 방법: {self.hook_method}
- 녹화된 이벤트 수: {len(self.current_events) if self.current_events else 0}
- 스케줄 수: {len(self.schedules)}
- 스케줄러 실행 중: {'예' if self.is_schedule_running else '아니오'}
- 녹화 중: {'예' if self.is_recording else '아니오'}
- Python 버전: {os.sys.version}
- 운영체제: {os.name}
- 로그 파일: {self.log.log_file}

핫키 상태:
- F11: 녹화 시작/중지
- F12: 녹화 중지

라이브러리 상태:
- pywin32: 설치됨
- keyboard: {'설치됨' if 'keyboard' in os.sys.modules else '확인 불가'}
- pyautogui: {'설치됨' if 'pyautogui' in os.sys.modules else '확인 불가'}

pynput 설치 확인:
"""
            try:
                import pynput
                info += "- pynput: 설치됨\n"
            except ImportError:
                info += "- pynput: 설치되지 않음 (권장: pip install pynput)\n"
            
            self.log.info("디버그 정보 표시됨")
            messagebox.showinfo("디버그 정보", info)
            
        except Exception as e:
            self.log.exception(f"디버그 정보 표시 오류: {str(e)}")

    def show_logs(self):
        """로그 창 표시"""
        try:
            log_window = tk.Toplevel(self.root)
            log_window.title("로그 보기")
            log_window.geometry("800x600")
            
            # 텍스트 위젯과 스크롤바
            text_frame = ttk.Frame(log_window)
            text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            scrollbar = ttk.Scrollbar(text_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            text_widget = tk.Text(text_frame, yscrollcommand=scrollbar.set, wrap=tk.WORD)
            text_widget.pack(fill=tk.BOTH, expand=True)
            scrollbar.config(command=text_widget.yview)
            
            # 로그 파일 내용 읽기
            try:
                if os.path.exists(self.log.log_file):
                    with open(self.log.log_file, 'r', encoding='utf-8') as f:
                        log_content = f.read()
                        text_widget.insert(tk.END, log_content)
                    # 마지막 줄로 스크롤
                    text_widget.see(tk.END)
                else:
                    text_widget.insert(tk.END, "로그 파일이 존재하지 않습니다.")
            except Exception as e:
                text_widget.insert(tk.END, f"로그 파일 읽기 오류: {str(e)}")
            
            # 새로고침 버튼
            button_frame = ttk.Frame(log_window)
            button_frame.pack(fill=tk.X, padx=10, pady=5)
            
            def refresh_logs():
                text_widget.delete(1.0, tk.END)
                try:
                    with open(self.log.log_file, 'r', encoding='utf-8') as f:
                        log_content = f.read()
                        text_widget.insert(tk.END, log_content)
                    text_widget.see(tk.END)
                except Exception as e:
                    text_widget.insert(tk.END, f"로그 파일 읽기 오류: {str(e)}")
            
            ttk.Button(button_frame, text="새로고침", command=refresh_logs).pack(side=tk.LEFT)
            ttk.Button(button_frame, text="닫기", command=log_window.destroy).pack(side=tk.RIGHT)
            
        except Exception as e:
            self.log.exception(f"로그 창 표시 오류: {str(e)}")

    def init_gui(self):
        """GUI 초기화"""
        try:
            self.log.info("GUI 초기화 시작")
            
            self.root = tk.Tk()
            self.root.title("Python 매크로 프로그램 (관리자 권한) - F11: 녹화 토글, F12: 중지")
            self.root.geometry("900x700")
            
            # 변수 초기화
            self.recording_status = tk.StringVar(value="상태: 대기 중 (F11: 녹화 시작)")
            self.scheduler_status = tk.StringVar(value="상태: 중지됨")
            self.status_var = tk.StringVar(value="프로그램 상태: 준비됨 (관리자 권한)")
            self.selected_macro_var = tk.StringVar(value="없음")
            
            # 탭 생성
            notebook = ttk.Notebook(self.root)
            notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 매크로 관리 탭
            macro_tab = ttk.Frame(notebook)
            notebook.add(macro_tab, text="매크로 관리")
            
            # 스케줄 관리 탭
            schedule_tab = ttk.Frame(notebook)
            notebook.add(schedule_tab, text="스케줄 관리")
            
            # 매크로 관리 탭 내용
            macro_name_frame = ttk.Frame(macro_tab)
            macro_name_frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Label(macro_name_frame, text="매크로 이름:").pack(side=tk.LEFT, padx=5)
            self.macro_name_entry = ttk.Entry(macro_name_frame, width=30)
            self.macro_name_entry.pack(side=tk.LEFT, padx=5)
            
            # 녹화 제어 버튼
            record_frame = ttk.Frame(macro_tab)
            record_frame.pack(fill=tk.X, padx=5, pady=5)
            
            self.start_rec_btn = ttk.Button(record_frame, text="녹화 시작 (F11)", 
                                          command=self.on_start_record)
            self.start_rec_btn.pack(side=tk.LEFT, padx=5)
            self.log.debug("녹화 시작 버튼 생성 완료")
            
            self.stop_rec_btn = ttk.Button(record_frame, text="녹화 중지 (F12)", 
                                         command=self.stop_recording, state=tk.DISABLED)
            self.stop_rec_btn.pack(side=tk.LEFT, padx=5)
            self.log.debug("녹화 중지 버튼 생성 완료")
            
            ttk.Button(record_frame, text="디버그 정보", 
                     command=self.show_debug_info).pack(side=tk.LEFT, padx=5)
            
            ttk.Button(record_frame, text="로그 보기", 
                     command=self.show_logs).pack(side=tk.LEFT, padx=5)
            
            ttk.Label(record_frame, textvariable=self.recording_status).pack(side=tk.LEFT, padx=20)
            
            # 핫키 안내 텍스트
            hotkey_frame = ttk.Frame(macro_tab)
            hotkey_frame.pack(fill=tk.X, padx=5, pady=5)
            
            hotkey_text = ttk.Label(hotkey_frame, 
                text="🔥 전역 핫키: F11 = 녹화 시작/중지, F12 = 녹화 중지 (어느 창에서나 작동)",
                foreground="red", font=("", 9, "bold"))
            hotkey_text.pack()
            
            # 도움말 텍스트
            help_frame = ttk.Frame(macro_tab)
            help_frame.pack(fill=tk.X, padx=5, pady=5)
            
            help_text = ttk.Label(help_frame, 
                text="녹화가 안 되는 경우: 1) 안티바이러스 일시 중지 2) pip install pynput 실행 3) 프로그램 재시작",
                foreground="blue", font=("", 8))
            help_text.pack()
            
            # 매크로 목록
            ttk.Label(macro_tab, text="녹화된 매크로:").pack(anchor=tk.W, padx=5, pady=5)
            
            macro_list_frame = ttk.Frame(macro_tab)
            macro_list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # 스크롤바
            scrollbar = ttk.Scrollbar(macro_list_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Treeview 생성
            self.macro_treeview = ttk.Treeview(macro_list_frame, columns=("name", "created"),
                                             show="headings", yscrollcommand=scrollbar.set)
            self.macro_treeview.pack(fill=tk.BOTH, expand=True)
            scrollbar.config(command=self.macro_treeview.yview)
            
            # 열 설정
            self.macro_treeview.heading("name", text="매크로 이름")
            self.macro_treeview.heading("created", text="생성 시간")
            self.macro_treeview.column("name", width=200)
            self.macro_treeview.column("created", width=150)
            
            # 매크로 선택 이벤트
            self.macro_treeview.bind("<<TreeviewSelect>>", self.on_macro_selected)
            
            # 매크로 제어 버튼
            macro_control_frame = ttk.Frame(macro_tab)
            macro_control_frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Button(macro_control_frame, text="매크로 실행", 
                     command=self.on_play_macro).pack(side=tk.LEFT, padx=5)
            ttk.Button(macro_control_frame, text="매크로 삭제", 
                     command=self.on_delete_macro).pack(side=tk.LEFT, padx=5)
                    
            # 스케줄 관리 탭 내용
            schedule_macro_frame = ttk.Frame(schedule_tab)
            schedule_macro_frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Label(schedule_macro_frame, text="선택한 매크로:").pack(side=tk.LEFT, padx=5)
            ttk.Label(schedule_macro_frame, textvariable=self.selected_macro_var, width=30).pack(side=tk.LEFT, padx=5)
            
            schedule_time_frame = ttk.Frame(schedule_tab)
            schedule_time_frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Label(schedule_time_frame, text="실행 시간 (HH:MM):").pack(side=tk.LEFT, padx=5)
            self.schedule_time_entry = ttk.Entry(schedule_time_frame, width=10)
            self.schedule_time_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Button(schedule_time_frame, text="스케줄 추가", 
                     command=self.on_add_schedule).pack(side=tk.LEFT, padx=20)
            
            # 스케줄 목록
            ttk.Label(schedule_tab, text="등록된 스케줄:").pack(anchor=tk.W, padx=5, pady=5)
            
            schedule_list_frame = ttk.Frame(schedule_tab)
            schedule_list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # 스크롤바
            schedule_scrollbar = ttk.Scrollbar(schedule_list_frame)
            schedule_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Treeview 생성
            self.schedule_treeview = ttk.Treeview(schedule_list_frame, 
                                               columns=("macro", "time", "created", "id"),
                                               show="headings", 
                                               yscrollcommand=schedule_scrollbar.set)
            self.schedule_treeview.pack(fill=tk.BOTH, expand=True)
            schedule_scrollbar.config(command=self.schedule_treeview.yview)
            
            # 열 설정
            self.schedule_treeview.heading("macro", text="매크로 이름")
            self.schedule_treeview.heading("time", text="실행 시간")
            self.schedule_treeview.heading("created", text="생성 시간")
            self.schedule_treeview.heading("id", text="ID")
            
            self.schedule_treeview.column("macro", width=200)
            self.schedule_treeview.column("time", width=100)
            self.schedule_treeview.column("created", width=150)
            self.schedule_treeview.column("id", width=0, stretch=tk.NO)  # ID 열 숨김
            
            # 스케줄 제어 버튼
            schedule_control_frame = ttk.Frame(schedule_tab)
            schedule_control_frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Button(schedule_control_frame, text="스케줄 삭제", 
                     command=self.on_delete_schedule).pack(side=tk.LEFT, padx=5)
            
            schedule_status_frame = ttk.Frame(schedule_tab)
            schedule_status_frame.pack(fill=tk.X, padx=5, pady=5)
            
            self.start_sched_btn = ttk.Button(schedule_status_frame, text="스케줄러 시작", 
                                            command=self.start_scheduler)
            self.start_sched_btn.pack(side=tk.LEFT, padx=5)
            self.log.debug("스케줄러 시작 버튼 생성 완료")
            
            self.stop_sched_btn = ttk.Button(schedule_status_frame, text="스케줄러 중지", 
                                           command=self.stop_scheduler, state=tk.DISABLED)
            self.stop_sched_btn.pack(side=tk.LEFT, padx=5)
            self.log.debug("스케줄러 중지 버튼 생성 완료")
            
            ttk.Label(schedule_status_frame, textvariable=self.scheduler_status).pack(side=tk.LEFT, padx=20)
            self.log.debug("스케줄러 상태 라벨 생성 완료")
            
            # 상태 표시줄
            status_frame = ttk.Frame(self.root)
            status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
            
            ttk.Label(status_frame, textvariable=self.status_var, anchor=tk.W).pack(side=tk.LEFT)
            ttk.Button(status_frame, text="종료", command=self.on_exit).pack(side=tk.RIGHT)
            
            # 초기 데이터 로드
            self.update_macro_list()
            self.update_schedule_list()
            
            # 윈도우 종료 이벤트 처리
            self.root.protocol("WM_DELETE_WINDOW", self.on_exit)
            
            self.log.info("GUI 초기화 완료")
            
            # 시작 알림 표시
            self.show_recording_notification("매크로 프로그램 시작", 
                "전역 핫키가 활성화되었습니다!\nF11: 녹화 시작/중지\nF12: 녹화 중지")
            
            # 메인 루프 시작
            self.root.mainloop()
            
        except Exception as e:
            self.log.exception(f"GUI 초기화 중 오류: {str(e)}")
            raise
    
    # 이벤트 핸들러
    def on_start_record(self):
        """녹화 시작 버튼 이벤트"""
        try:
            self.log.debug("녹화 시작 버튼 클릭됨")
            macro_name = self.macro_name_entry.get().strip()
            if not macro_name:
                # 기본 이름 생성
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                macro_name = f"매크로_{timestamp}"
                self.macro_name_entry.insert(0, macro_name)
                self.log.debug(f"자동 매크로 이름 생성: {macro_name}")
                
            self.start_recording(macro_name)
        except Exception as e:
            self.log.exception(f"녹화 시작 버튼 이벤트 오류: {str(e)}")
    
    def on_macro_selected(self, event):
        """매크로 목록에서 선택했을 때"""
        try:
            self.log.debug("매크로 선택 이벤트 발생")
            selection = self.macro_treeview.selection()
            if not selection:
                self.log.debug("선택된 매크로 없음")
                return
                
            item = selection[0]
            values = self.macro_treeview.item(item, "values")
            if not values:
                self.log.debug("매크로 값 없음")
                return
                
            macro_name = values[0]
            # 선택된 매크로 찾기
            for macro in self.recorded_macros:
                if macro["name"] == macro_name:
                    self.selected_macro = macro
                    self.selected_macro_var.set(macro_name)
                    self.log.debug(f"매크로 선택됨: {macro_name}")
                    break
        except Exception as e:
            self.log.exception(f"매크로 선택 이벤트 오류: {str(e)}")

    def on_play_macro(self):
        """매크로 실행 버튼 이벤트"""
        try:
            self.log.debug("매크로 실행 버튼 클릭됨")
            selection = self.macro_treeview.selection()
            if not selection:
                self.log.warning("실행할 매크로가 선택되지 않음")
                messagebox.showerror("오류", "실행할 매크로를 선택하세요.")
                return
                
            item = selection[0]
            values = self.macro_treeview.item(item, "values")
            macro_name = values[0]
            
            # 선택된 매크로 찾기
            for macro in self.recorded_macros:
                if macro["name"] == macro_name:
                    self.log.info(f"매크로 실행 시작: {macro_name}")
                    threading.Thread(
                        target=self.play_macro,
                        args=(macro["path"],),
                        daemon=True
                    ).start()
                    break
        except Exception as e:
            self.log.exception(f"매크로 실행 버튼 이벤트 오류: {str(e)}")
    
    def on_delete_macro(self):
        """매크로 삭제 버튼 이벤트"""
        try:
            self.log.debug("매크로 삭제 버튼 클릭됨")
            selection = self.macro_treeview.selection()
            if not selection:
                self.log.warning("삭제할 매크로가 선택되지 않음")
                messagebox.showerror("오류", "삭제할 매크로를 선택하세요.")
                return
                
            item = selection[0]
            values = self.macro_treeview.item(item, "values")
            macro_name = values[0]
            
            # 확인 메시지
            if not messagebox.askyesno("확인", f"매크로 '{macro_name}'를 삭제하시겠습니까?"):
                self.log.debug(f"매크로 삭제 취소: {macro_name}")
                return
                
            # 선택된 매크로 찾기
            for macro in self.recorded_macros:
                if macro["name"] == macro_name:
                    self.log.info(f"매크로 삭제 실행: {macro_name}")
                    self.delete_macro(macro["path"])
                    break
        except Exception as e:
            self.log.exception(f"매크로 삭제 버튼 이벤트 오류: {str(e)}")
    
    def on_add_schedule(self):
        """스케줄 추가 버튼 이벤트"""
        try:
            self.log.debug("스케줄 추가 버튼 클릭됨")
            
            selected_macro_name = self.selected_macro_var.get()
            if selected_macro_name == "없음":
                self.log.warning("스케줄 추가 시 매크로가 선택되지 않음")
                messagebox.showerror("오류", "매크로를 먼저 선택하세요.")
                return
                
            time_str = self.schedule_time_entry.get().strip()
            if not time_str:
                self.log.warning("스케줄 추가 시 시간이 입력되지 않음")
                messagebox.showerror("오류", "실행 시간을 입력하세요. (HH:MM)")
                return
                
            # 시간 형식 확인
            try:
                hour, minute = map(int, time_str.split(":"))
                if not (0 <= hour < 24 and 0 <= minute < 60):
                    raise ValueError()
                self.log.debug(f"시간 형식 검증 통과: {time_str}")
            except:
                self.log.error(f"잘못된 시간 형식: {time_str}")
                messagebox.showerror("오류", "시간 형식이 올바르지 않습니다. (HH:MM)")
                return
            
            self.log.info(f"스케줄 추가 시도: {selected_macro_name} at {time_str}")
            if self.add_schedule(selected_macro_name, time_str):
                messagebox.showinfo("알림", f"스케줄이 추가되었습니다: {selected_macro_name} - {time_str}")
                # 입력 필드 초기화
                self.schedule_time_entry.delete(0, tk.END)
                self.log.debug("스케줄 추가 완료 및 입력 필드 초기화")
                
        except Exception as e:
            self.log.exception(f"스케줄 추가 버튼 이벤트 오류: {str(e)}")
    
    def on_delete_schedule(self):
        """스케줄 삭제 버튼 이벤트"""
        try:
            self.log.debug("스케줄 삭제 버튼 클릭됨")
            
            selection = self.schedule_treeview.selection()
            if not selection:
                self.log.warning("삭제할 스케줄이 선택되지 않음")
                messagebox.showerror("오류", "삭제할 스케줄을 선택하세요.")
                return
                
            item = selection[0]
            values = self.schedule_treeview.item(item, "values")
            schedule_id = values[3]  # ID는 4번째 컬럼
            
            # 확인 메시지
            if not messagebox.askyesno("확인", "선택한 스케줄을 삭제하시겠습니까?"):
                self.log.debug(f"스케줄 삭제 취소: {schedule_id}")
                return
                
            self.log.info(f"스케줄 삭제 실행: {schedule_id}")
            if self.delete_schedule(schedule_id):
                messagebox.showinfo("알림", "스케줄이 삭제되었습니다.")
                
        except Exception as e:
            self.log.exception(f"스케줄 삭제 버튼 이벤트 오류: {str(e)}")
    
    def on_exit(self):
        """종료 시 정리 작업"""
        try:
            self.log.info("프로그램 종료 요청")
            
            self.log.debug("녹화 중지 중...")
            self.stop_recording()
            
            self.log.debug("스케줄러 중지 중...")
            self.stop_scheduler()
            
            self.log.debug("후킹 해제 중...")
            self.stop_all_hooks()
            
            self.log.debug("전역 핫키 정리 중...")
            self.cleanup_global_hotkeys()
            
            if self.root:
                self.log.debug("GUI 창 닫는 중...")
                self.root.destroy()
                
            self.log.info("프로그램 종료 완료")
            
        except Exception as e:
            self.log.exception(f"프로그램 종료 중 오류: {str(e)}")


if __name__ == "__main__":
    try:
        print("=" * 50)
        print("Python 매크로 프로그램 시작")
        print("핫키: F11 = 녹화 시작/중지, F12 = 녹화 중지")
        print("로그 위치: ~/PyMacro/debug.log")
        print("=" * 50)
        app = MacroRecorder()
    except SystemExit:
        print("관리자 권한으로 재시작 중...")
        pass  # 관리자 권한 재시작 시 정상 종료
    except Exception as e:
        print(f"프로그램 시작 오류: {str(e)}")
        traceback.print_exc()
        input("엔터를 눌러 종료...")