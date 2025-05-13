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

# 화면 안전장치 비활성화 (실수로 마우스가 화면 구석으로 이동하는 경우 오류 방지)
pyautogui.FAILSAFE = False

class MacroRecorder:
    def __init__(self):
        # 기본 디렉토리 설정
        self.base_dir = os.path.join(os.path.expanduser("~"), "PyMacro")
        self.macros_dir = os.path.join(self.base_dir, "macros")
        self.schedules_file = os.path.join(self.base_dir, "schedules.json")
        
        # 디렉토리 생성
        os.makedirs(self.macros_dir, exist_ok=True)
        
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
        
        # 스케줄 실행 스레드
        self.schedule_thread = None
        self.is_schedule_running = False
        
        # GUI 초기화
        self.root = None
        self.init_gui()

    def load_macros(self):
        """저장된 매크로 파일 목록 로드"""
        self.recorded_macros = []
        if os.path.exists(self.macros_dir):
            for file in os.listdir(self.macros_dir):
                if file.endswith(".json"):
                    macro_path = os.path.join(self.macros_dir, file)
                    macro_name = file[:-5]  # .json 제거
                    
                    # 매크로 생성 시간 가져오기
                    created_time = datetime.datetime.fromtimestamp(
                        os.path.getctime(macro_path)
                    ).strftime("%Y-%m-%d %H:%M:%S")
                    
                    self.recorded_macros.append({
                        "name": macro_name,
                        "file": file,
                        "created": created_time,
                        "path": macro_path
                    })
    
    def load_schedules(self):
        """저장된 스케줄 목록 로드"""
        if os.path.exists(self.schedules_file):
            try:
                with open(self.schedules_file, "r", encoding="utf-8") as f:
                    self.schedules = json.load(f)
            except:
                self.schedules = []
        else:
            self.schedules = []
    
    def save_schedules(self):
        """스케줄 정보 저장"""
        with open(self.schedules_file, "w", encoding="utf-8") as f:
            json.dump(self.schedules, f, ensure_ascii=False, indent=2)
    
    def start_recording(self, macro_name: str):
        """매크로 녹화 시작"""
        if self.is_recording:
            return
            
        self.is_recording = True
        self.current_macro = macro_name
        self.current_events = []
        
        # 녹화 시작 시간 기록
        start_time = time.time()
        
        # 녹화 스레드 시작
        self.recording_thread = threading.Thread(
            target=self._record_events, 
            args=(start_time,)
        )
        self.recording_thread.daemon = True
        self.recording_thread.start()
        
        # GUI 업데이트
        self.recording_status.set("녹화 중...")
        self.start_rec_btn["state"] = tk.DISABLED
        self.stop_rec_btn["state"] = tk.NORMAL
    
    def stop_recording(self):
        """매크로 녹화 중지"""
        if not self.is_recording:
            return
            
        self.is_recording = False
        
        # 녹화 스레드가 종료될 때까지 대기
        if self.recording_thread and self.recording_thread.is_alive():
            self.recording_thread.join(1)
        
        # 매크로 저장
        self.save_recorded_macro()
        
        # 매크로 목록 새로고침
        self.load_macros()
        self.update_macro_list()
        
        # GUI 업데이트
        self.recording_status.set("녹화 중지됨")
        self.start_rec_btn["state"] = tk.NORMAL
        self.stop_rec_btn["state"] = tk.DISABLED
    
    def _record_events(self, start_time: float):
        """이벤트 녹화 루프"""
        # 마우스 이벤트를 위한 초기 위치
        last_x, last_y = pyautogui.position()
        
        # 마우스 버튼 상태 추적 (초기에는 모두 눌려있지 않음)
        left_pressed = False
        right_pressed = False
        middle_pressed = False
        
        # 키보드 이벤트를 위한 후크 설정
        keyboard_events = []
        
        def on_keyboard_event(e):
            if not self.is_recording:
                return
                
            # 특수 키 처리
            if hasattr(e, 'name'):
                key_name = e.name
            else:
                key_name = e.char
                
            event_type = 'key_down' if e.event_type == keyboard.KEY_DOWN else 'key_up'
            
            # 이벤트 시간과 함께 기록
            keyboard_events.append({
                'type': event_type,
                'key': key_name,
                'time': time.time() - start_time
            })
        
        # 키보드 후크 설정
        keyboard.hook(on_keyboard_event)
        
        # 마우스 클릭 감지를 위한 함수 설정
        try:
            # pynput으로 마우스 이벤트 감지 (대안 라이브러리)
            from pynput import mouse as pynput_mouse
            
            def on_click(x, y, button, pressed):
                if not self.is_recording:
                    return
                
                # 버튼 타입 변환
                button_name = 'left'
                if str(button) == 'Button.middle':
                    button_name = 'middle'
                elif str(button) == 'Button.right':
                    button_name = 'right'
                
                # 이벤트 시간과 함께 기록
                event_type = 'mouse_down' if pressed else 'mouse_up'
                self.current_events.append({
                    'type': event_type,
                    'button': button_name,
                    'x': x,
                    'y': y,
                    'time': time.time() - start_time
                })
            
            # 마우스 리스너 시작
            mouse_listener = pynput_mouse.Listener(on_click=on_click)
            mouse_listener.start()
            using_pynput = True
            
        except ImportError:
            using_pynput = False
            print("pynput 라이브러리가 설치되지 않았습니다. pip install pynput으로 설치하세요.")
            print("대체 방법으로 win32api를 사용합니다.")
            
            # win32api를 사용하여 마우스 감지 (윈도우 전용)
            try:
                import win32api
                import win32con
                win32_available = True
            except ImportError:
                win32_available = False
                print("win32api도 사용할 수 없습니다. 마우스 클릭 감지가 제한됩니다.")
        
        try:
            # 마우스 이벤트 녹화 루프
            while self.is_recording:
                # 현재 마우스 위치
                current_x, current_y = pyautogui.position()
                current_time = time.time() - start_time
                
                # 마우스가 이동했는지 확인
                if (current_x, current_y) != (last_x, last_y):
                    self.current_events.append({
                        'type': 'mouse_move',
                        'x': current_x,
                        'y': current_y,
                        'time': current_time
                    })
                    last_x, last_y = current_x, current_y
                
                # pynput을 사용하지 않고 win32api가 사용 가능한 경우에만 직접 감지
                if not using_pynput and win32_available:
                    # 현재 마우스 버튼 상태 확인
                    current_left = win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000 != 0
                    current_right = win32api.GetAsyncKeyState(win32con.VK_RBUTTON) & 0x8000 != 0
                    current_middle = win32api.GetAsyncKeyState(win32con.VK_MBUTTON) & 0x8000 != 0
                    
                    # 상태가 변경된 경우만 이벤트 추가
                    if current_left != left_pressed:
                        left_pressed = current_left
                        self.current_events.append({
                            'type': 'mouse_down' if left_pressed else 'mouse_up',
                            'button': 'left',
                            'x': current_x,
                            'y': current_y,
                            'time': current_time
                        })
                        
                    if current_right != right_pressed:
                        right_pressed = current_right
                        self.current_events.append({
                            'type': 'mouse_down' if right_pressed else 'mouse_up',
                            'button': 'right',
                            'x': current_x,
                            'y': current_y,
                            'time': current_time
                        })
                        
                    if current_middle != middle_pressed:
                        middle_pressed = current_middle
                        self.current_events.append({
                            'type': 'mouse_down' if middle_pressed else 'mouse_up',
                            'button': 'middle',
                            'x': current_x,
                            'y': current_y,
                            'time': current_time
                        })
                
                # CPU 사용량 감소
                time.sleep(0.01)
        finally:
            # 키보드 후크 해제
            keyboard.unhook_all()
            
            # pynput 사용 시 리스너 중지
            if using_pynput:
                mouse_listener.stop()
            
            # 키보드 이벤트를 현재 이벤트 목록에 추가
            self.current_events.extend(keyboard_events)
            
            # 이벤트 시간 기준으로 정렬
            self.current_events.sort(key=lambda x: x['time'])

    def save_recorded_macro(self):
        """녹화된 매크로를 파일로 저장"""
        if not self.current_events:
            return
            
        macro_data = {
            "name": self.current_macro,
            "created": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "events": self.current_events
        }
        
        # 파일명 생성 (공백은 밑줄로 대체)
        filename = f"{self.current_macro.replace(' ', '_')}.json"
        file_path = os.path.join(self.macros_dir, filename)
        
        # 매크로 저장
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(macro_data, f, ensure_ascii=False, indent=2)
    
    def play_macro(self, macro_path: str):
        """매크로 재생 (이벤트 없는 시간 포함)"""
        try:
            # 매크로 파일 로드
            with open(macro_path, "r", encoding="utf-8") as f:
                macro_data = json.load(f)
            
            events = macro_data["events"]
            if not events:
                return
            
            # 상태 업데이트
            self.status_var.set(f"매크로 '{macro_data['name']}' 실행 중...")
            
            # 마우스 이동 시간을 직접 설정 (0.005초로 설정)
            mouse_move_duration = 0.005
            
            # 키보드 입력 사이의 지연 시간 (키보드 입력을 더 안정적으로)
            keyboard_delay = 0.05  # 50ms 지연
            
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
            sorted_events = sorted(events, key=lambda e: e["time"])
            
            # 시작 시간과 마지막 이벤트 시간 설정
            if sorted_events:
                start_time = sorted_events[0]["time"]
                last_event_time = start_time
                last_keyboard_time = 0  # 마지막 키보드 이벤트 시간
            
                # 이벤트 처리 루프
                for event in sorted_events:
                    # 실행 취소 버튼이 눌렸는지 확인
                    if not self.root or not self.root.winfo_exists():
                        break
                    
                    # 이벤트 사이의 원래 대기 시간 계산 (압축하지 않음)
                    wait_time = event["time"] - last_event_time
                    
                    # 대기 시간이 너무 길면 사용자에게 알림 (옵션)
                    if wait_time > 5.0:
                        self.status_var.set(f"매크로 '{macro_data['name']}' - {wait_time:.1f}초 대기 중...")
                    
                    # 모든 대기 시간 유지 (이전 이벤트와의 시간 간격대로 대기)
                    if wait_time > 0.01:  # 10ms 이상의 대기 시간만 적용
                        time.sleep(wait_time)
                    
                    # 이벤트 처리
                    event_type = event["type"]
                    
                    if event_type == "mouse_move":
                        x, y = event["x"], event["y"]
                        
                        # 위치가 충분히 다른 경우 이동
                        if last_mouse_x is None or last_mouse_y is None or \
                        abs(last_mouse_x - x) > 5 or abs(last_mouse_y - y) > 5:
                            # 지정된 지속 시간으로 이동
                            pyautogui.moveTo(x, y, duration=mouse_move_duration)
                            last_mouse_x, last_mouse_y = x, y
                    
                    elif event_type == "mouse_click":
                        button = event["button"]
                        x, y = event["x"], event["y"]
                        
                        # 현재 위치와 다른 경우만 이동
                        if last_mouse_x is None or last_mouse_y is None or \
                        abs(last_mouse_x - x) > 5 or abs(last_mouse_y - y) > 5:
                            pyautogui.moveTo(x, y, duration=mouse_move_duration)
                        
                        pyautogui.click(button=button)
                        last_mouse_x, last_mouse_y = x, y
                    
                    elif event_type == "mouse_down":
                        button = event["button"]
                        x, y = event["x"], event["y"]
                        
                        # 현재 위치와 다른 경우만 이동
                        if last_mouse_x is None or last_mouse_y is None or \
                        abs(last_mouse_x - x) > 5 or abs(last_mouse_y - y) > 5:
                            pyautogui.moveTo(x, y, duration=mouse_move_duration)
                        
                        pyautogui.mouseDown(button=button)
                        last_mouse_x, last_mouse_y = x, y
                    
                    elif event_type == "mouse_up":
                        button = event["button"]
                        x, y = event["x"], event["y"]
                        
                        # 현재 위치와 다른 경우만 이동
                        if last_mouse_x is None or last_mouse_y is None or \
                        abs(last_mouse_x - x) > 5 or abs(last_mouse_y - y) > 5:
                            pyautogui.moveTo(x, y, duration=mouse_move_duration)
                        
                        pyautogui.mouseUp(button=button)
                        last_mouse_x, last_mouse_y = x, y
                    
                    elif event_type == "key_down":
                        # 키보드 이벤트 사이에 최소 지연 시간 적용
                        current_time = time.time()
                        time_since_last_keyboard = current_time - last_keyboard_time
                        
                        if time_since_last_keyboard < keyboard_delay:
                            # 마지막 키보드 이벤트와의 간격이 너무 짧으면 대기
                            time.sleep(keyboard_delay - time_since_last_keyboard)
                        
                        key = event["key"].lower()  # 소문자로 변환하여 일관성 유지
                        
                        # 특수 키 처리
                        if key in special_keys:
                            key = special_keys[key]
                        
                        # 한/영 키 특별 처리
                        if key == 'hangul':
                            try:
                                # PyAutoGUI를 사용하여 한/영 키 시뮬레이션
                                pyautogui.press('hangul')
                            except:
                                # PyAutoGUI에서 지원하지 않는 경우 대체 방법
                                try:
                                    # 윈도우에서는 alt+한/영 조합이 한/영 전환과 동일
                                    pyautogui.hotkey('alt', 'hangul')
                                except:
                                    # 마지막 대안으로 직접 키코드 전송 시도
                                    try:
                                        import win32api
                                        import win32con
                                        # 한/영 키 누름 (키코드 21)
                                        win32api.keybd_event(21, 0, 0, 0)
                                        time.sleep(0.05)
                                        # 한/영 키 뗌
                                        win32api.keybd_event(21, 0, win32con.KEYEVENTF_KEYUP, 0)
                                    except:
                                        print("한/영 키 처리 실패")
                        else:
                            # 일반 키 처리
                            try:
                                # 이미 눌린 키는 다시 누르지 않음
                                if key not in pressed_keys:
                                    keyboard.press(key)
                                    pressed_keys.add(key)
                            except Exception as e:
                                print(f"키 입력 오류 ({key}): {str(e)}")
                                # PyAutoGUI로 대체 시도
                                try:
                                    pyautogui.keyDown(key)
                                except:
                                    print(f"PyAutoGUI 키 입력 실패: {key}")
                        
                        last_keyboard_time = time.time()
                    
                    elif event_type == "key_up":
                        # 키보드 이벤트 사이에 최소 지연 시간 적용
                        current_time = time.time()
                        time_since_last_keyboard = current_time - last_keyboard_time
                        
                        if time_since_last_keyboard < keyboard_delay:
                            # 마지막 키보드 이벤트와의 간격이 너무 짧으면 대기
                            time.sleep(keyboard_delay - time_since_last_keyboard)
                        
                        key = event["key"].lower()  # 소문자로 변환하여 일관성 유지
                        
                        # 특수 키 처리
                        if key in special_keys:
                            key = special_keys[key]
                        
                        # 한/영 키는 key_up 이벤트에서는 별도 처리 필요 없음 (이미 press에서 처리됨)
                        if key != 'hangul':
                            try:
                                if key in pressed_keys:
                                    keyboard.release(key)
                                    pressed_keys.remove(key)
                            except Exception as e:
                                print(f"키 해제 오류 ({key}): {str(e)}")
                                # PyAutoGUI로 대체 시도
                                try:
                                    pyautogui.keyUp(key)
                                except:
                                    print(f"PyAutoGUI 키 해제 실패: {key}")
                        
                        last_keyboard_time = time.time()
                    
                    # 현재 이벤트 시간을 마지막 이벤트 시간으로 업데이트
                    last_event_time = event["time"]
                
                # 모든 키 상태 원상복구 (안전을 위해)
                for key in pressed_keys:
                    try:
                        keyboard.release(key)
                    except:
                        try:
                            pyautogui.keyUp(key)
                        except:
                            pass
            
            # 상태 업데이트
            self.status_var.set(f"매크로 '{macro_data['name']}' 실행 완료")
        
        except Exception as e:
            self.status_var.set(f"매크로 실행 오류: {str(e)}")

    def delete_macro(self, macro_path: str):
        """매크로 삭제"""
        try:
            if os.path.exists(macro_path):
                os.remove(macro_path)
                
            # 매크로가 삭제되었으므로 관련 스케줄도 삭제
            macro_name = os.path.basename(macro_path)[:-5]  # .json 제거
            self.schedules = [s for s in self.schedules 
                            if s["macro"] != macro_name]
            self.save_schedules()
            
            # 목록 새로고침
            self.load_macros()
            self.update_macro_list()
            self.update_schedule_list()
            
            return True
        except Exception as e:
            messagebox.showerror("오류", f"매크로 삭제 오류: {str(e)}")
            return False
    
    def add_schedule(self, macro_name: str, time_str: str):
        """스케줄 추가"""
        try:
            # 시간 형식 확인 (HH:MM)
            hour, minute = map(int, time_str.split(":"))
            if not (0 <= hour < 24 and 0 <= minute < 60):
                raise ValueError("시간 형식이 올바르지 않습니다.")
                
            # 스케줄 ID 생성 (단순히 현재 시간 기반)
            schedule_id = str(int(time.time()))
            
            # 새 스케줄 생성
            new_schedule = {
                "id": schedule_id,
                "macro": macro_name,
                "time": time_str,
                "created": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # 스케줄 목록에 추가
            self.schedules.append(new_schedule)
            self.save_schedules()
            
            # 실행 중인 경우 스케줄 갱신
            if self.is_schedule_running:
                self.update_scheduler()
            
            # 목록 새로고침
            self.update_schedule_list()
            
            return True
        except Exception as e:
            messagebox.showerror("오류", f"스케줄 추가 오류: {str(e)}")
            return False
    
    def delete_schedule(self, schedule_id: str):
        """스케줄 삭제"""
        try:
            # ID로 스케줄 찾기
            self.schedules = [s for s in self.schedules if s["id"] != schedule_id]
            self.save_schedules()
            
            # 실행 중인 경우 스케줄 갱신
            if self.is_schedule_running:
                self.update_scheduler()
            
            # 목록 새로고침
            self.update_schedule_list()
            
            return True
        except Exception as e:
            messagebox.showerror("오류", f"스케줄 삭제 오류: {str(e)}")
            return False
    
    def start_scheduler(self):
        """스케줄러 시작"""
        if self.is_schedule_running:
            return
            
        self.is_schedule_running = True
        
        # 스케줄 설정
        self.update_scheduler()
        
        # 스케줄러 스레드 시작
        self.schedule_thread = threading.Thread(target=self._run_scheduler)
        self.schedule_thread.daemon = True
        self.schedule_thread.start()
        
        # GUI 업데이트
        self.scheduler_status.set("스케줄러 실행 중")
        self.start_sched_btn["state"] = tk.DISABLED
        self.stop_sched_btn["state"] = tk.NORMAL
    
    def stop_scheduler(self):
        """스케줄러 중지"""
        if not self.is_schedule_running:
            return
            
        self.is_schedule_running = False
        
        # 모든 스케줄 작업 취소
        schedule.clear()
        
        # GUI 업데이트
        self.scheduler_status.set("스케줄러 중지됨")
        self.start_sched_btn["state"] = tk.NORMAL
        self.stop_sched_btn["state"] = tk.DISABLED
    
    def update_scheduler(self):
        """스케줄 작업 갱신"""
        # 기존 스케줄 모두 취소
        schedule.clear()
        
        # 각 스케줄 등록
        for sched in self.schedules:
            macro_name = sched["macro"]
            time_str = sched["time"]
            
            # 매크로 파일 경로 찾기
            macro_path = None
            for macro in self.recorded_macros:
                if macro["name"] == macro_name:
                    macro_path = macro["path"]
                    break
            
            if macro_path and os.path.exists(macro_path):
                # 시간 파싱
                hour, minute = map(int, time_str.split(":"))
                
                # 스케줄 등록 (매일 반복)
                schedule.every().day.at(time_str).do(
                    self.play_macro_scheduled, macro_path=macro_path
                )
    
    def play_macro_scheduled(self, macro_path: str):
        """스케줄에 의해 매크로 실행"""
        # 별도 스레드에서 매크로 실행
        threading.Thread(
            target=self.play_macro, 
            args=(macro_path,)
        ).start()
        
        # 다음 실행을 위해 True 반환 (schedule 라이브러리 요구사항)
        return True
    
    def _run_scheduler(self):
        """스케줄러 실행 루프"""
        while self.is_schedule_running:
            try:
                # 예약된 작업 실행
                schedule.run_pending()
                
                # CPU 사용량 감소
                time.sleep(1)
            except Exception as e:
                print(f"스케줄러 오류: {str(e)}")
    
    def update_macro_list(self):
        """매크로 목록 업데이트"""
        if not self.root or not self.root.winfo_exists():
            return
            
        # 테이블 초기화
        for item in self.macro_treeview.get_children():
            self.macro_treeview.delete(item)
            
        # 매크로 목록 데이터 추가
        for macro in self.recorded_macros:
            self.macro_treeview.insert("", tk.END, values=(
                macro["name"],
                macro["created"]
            ))
    
    def update_schedule_list(self):
        """스케줄 목록 업데이트"""
        if not self.root or not self.root.winfo_exists():
            return
            
        # 테이블 초기화
        for item in self.schedule_treeview.get_children():
            self.schedule_treeview.delete(item)
            
        # 스케줄 목록 데이터 추가
        for sched in self.schedules:
            self.schedule_treeview.insert("", tk.END, values=(
                sched["macro"],
                sched["time"],
                sched["created"],
                sched["id"]
            ))
    
    def init_gui(self):
        """GUI 초기화"""
        self.root = tk.Tk()
        self.root.title("Python 매크로 프로그램")
        self.root.geometry("800x600")
        
        # 변수 초기화
        self.recording_status = tk.StringVar(value="상태: 대기 중")
        self.scheduler_status = tk.StringVar(value="상태: 중지됨")
        self.status_var = tk.StringVar(value="프로그램 상태: 준비됨")
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
        
        self.start_rec_btn = ttk.Button(record_frame, text="녹화 시작", 
                                      command=self.on_start_record)
        self.start_rec_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_rec_btn = ttk.Button(record_frame, text="녹화 중지", 
                                     command=self.stop_recording, state=tk.DISABLED)
        self.stop_rec_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(record_frame, textvariable=self.recording_status).pack(side=tk.LEFT, padx=20)
        
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
        
        self.stop_sched_btn = ttk.Button(schedule_status_frame, text="스케줄러 중지", 
                                       command=self.stop_scheduler, state=tk.DISABLED)
        self.stop_sched_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(schedule_status_frame, textvariable=self.scheduler_status).pack(side=tk.LEFT, padx=20)
        
        # 상태 표시줄
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
        
        ttk.Label(status_frame, textvariable=self.status_var, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Button(status_frame, text="종료", command=self.on_exit).pack(side=tk.RIGHT)
        
        # 초기 데이터 로드
        self.update_macro_list()
        self.update_schedule_list()
        
        # 주기적 업데이트 및 이벤트 처리
        self.schedule_periodic_updates()
        
        # 윈도우 종료 이벤트 처리
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)
        
        # 메인 루프 시작
        self.root.mainloop()
    
    def schedule_periodic_updates(self):
        """주기적 업데이트 설정"""
        if self.root and self.root.winfo_exists():
            # 100ms마다 이벤트 루프 처리
            self.root.after(100, self.schedule_periodic_updates)
    
    # 이벤트 핸들러
    def on_start_record(self):
        """녹화 시작 버튼 이벤트"""
        macro_name = self.macro_name_entry.get().strip()
        if not macro_name:
            messagebox.showerror("오류", "매크로 이름을 입력하세요.")
            return
        self.start_recording(macro_name)
    
    def on_macro_selected(self, event):
        """매크로 목록에서 선택했을 때"""
        selection = self.macro_treeview.selection()
        if not selection:
            return
            
        item = selection[0]
        values = self.macro_treeview.item(item, "values")
        if not values:
            return
            
        macro_name = values[0]
        # 선택된 매크로 찾기
        for macro in self.recorded_macros:
            if macro["name"] == macro_name:
                self.selected_macro = macro
                self.selected_macro_var.set(macro_name)
                break

    def on_play_macro(self):
        """매크로 실행 버튼 이벤트"""
        selection = self.macro_treeview.selection()
        if not selection:
            messagebox.showerror("오류", "실행할 매크로를 선택하세요.")
            return
            
        item = selection[0]
        values = self.macro_treeview.item(item, "values")
        macro_name = values[0]
        
        # 선택된 매크로 찾기
        for macro in self.recorded_macros:
            if macro["name"] == macro_name:
                threading.Thread(
                    target=self.play_macro,
                    args=(macro["path"],)
                ).start()
                break
    
    def on_delete_macro(self):
        """매크로 삭제 버튼 이벤트"""
        selection = self.macro_treeview.selection()
        if not selection:
            messagebox.showerror("오류", "삭제할 매크로를 선택하세요.")
            return
            
        item = selection[0]
        values = self.macro_treeview.item(item, "values")
        macro_name = values[0]
        
        # 확인 메시지
        if not messagebox.askyesno("확인", f"매크로 '{macro_name}'를 삭제하시겠습니까?"):
            return
            
        # 선택된 매크로 찾기
        for macro in self.recorded_macros:
            if macro["name"] == macro_name:
                self.delete_macro(macro["path"])
                break
    
    def on_add_schedule(self):
        """스케줄 추가 버튼 이벤트"""
        selected_macro_name = self.selected_macro_var.get()
        if selected_macro_name == "없음":
            messagebox.showerror("오류", "매크로를 먼저 선택하세요.")
            return
            
        time_str = self.schedule_time_entry.get().strip()
        if not time_str:
            messagebox.showerror("오류", "실행 시간을 입력하세요. (HH:MM)")
            return
            
        # 시간 형식 확인
        try:
            hour, minute = map(int, time_str.split(":"))
            if not (0 <= hour < 24 and 0 <= minute < 60):
                raise ValueError()
        except:
            messagebox.showerror("오류", "시간 형식이 올바르지 않습니다. (HH:MM)")
            return
        
        if self.add_schedule(selected_macro_name, time_str):
            messagebox.showinfo("알림", f"스케줄이 추가되었습니다: {selected_macro_name} - {time_str}")
    
    def on_delete_schedule(self):
        """스케줄 삭제 버튼 이벤트"""
        selection = self.schedule_treeview.selection()
        if not selection:
            messagebox.showerror("오류", "삭제할 스케줄을 선택하세요.")
            return
            
        item = selection[0]
        values = self.schedule_treeview.item(item, "values")
        schedule_id = values[3]  # ID는 4번째 컬럼
        
        # 확인 메시지
        if not messagebox.askyesno("확인", "선택한 스케줄을 삭제하시겠습니까?"):
            return
            
        if self.delete_schedule(schedule_id):
            messagebox.showinfo("알림", "스케줄이 삭제되었습니다.")
    
    def on_exit(self):
        """종료 버튼 이벤트"""
        # 녹화/스케줄러 중지
        self.stop_recording()
        self.stop_scheduler()
        
        # 윈도우 종료
        if self.root:
            self.root.destroy()


if __name__ == "__main__":
    # 프로그램 실행
    app = MacroRecorder()