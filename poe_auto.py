import tkinter as tk
from tkinter import messagebox
import time
import json
import os
import threading
import keyboard
import mouse
import pygetwindow as gw
from PIL import ImageGrab, Image, ImageTk

class HardwareLevelDragMacro:
    def __init__(self):
        self.start_pos = None
        self.end_pos = None
        self.grid_width = 12
        self.grid_height = 5
        self.excluded_cells = []
        self.config_file = "hardware_drag_macro_config.json"
        self.screenshot = None
        self.tk_image = None
        self.is_running = False
        self.dragging = False
        
        # 기본 설정 로드
        self.load_config()
        
        # GUI 생성
        self.create_gui()
    
    def load_config(self):
        """설정 파일 로드"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.start_pos = config.get('start_pos')
                    self.end_pos = config.get('end_pos')
                    self.excluded_cells = config.get('excluded_cells', [])
            except Exception as e:
                print(f"설정 로드 오류: {e}")
    
    def save_config(self):
        """설정 파일 저장"""
        config = {
            'start_pos': self.start_pos,
            'end_pos': self.end_pos,
            'excluded_cells': self.excluded_cells
        }
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            print(f"설정 저장 오류: {e}")
    
    def find_path_of_exile_window(self):
        """Path of Exile 창 찾기"""
        try:
            # 'Path of Exile' 창 검색 (대소문자 무시)
            poe_windows = [w for w in gw.getWindowsWithTitle('Path of Exile') if w.title.lower().startswith('path of exile')]
            
            if not poe_windows:
                messagebox.showwarning("경고", "Path of Exile 창을 찾을 수 없습니다.")
                return False
            
            # 첫 번째 창 활성화
            poe_window = poe_windows[0]
            poe_window.activate()
            
            # 창이 최소화되어 있다면 복원
            if poe_window.isMinimized:
                poe_window.restore()
            
            return True
        except Exception as e:
            messagebox.showerror("오류", f"창 찾기 중 오류 발생: {e}")
            return False
    
    def run_macro(self):
        """매크로 실행"""
        if not self.start_pos or not self.end_pos:
            messagebox.showwarning("경고", "영역을 먼저 선택해주세요.")
            return
        
        if self.is_running:
            return
        
        # Path of Exile 창 찾기 및 활성화
        if not self.find_path_of_exile_window():
            return
        
        # 실행 상태 설정
        self.is_running = True
        self.run_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # 실행 스레드 시작
        macro_thread = threading.Thread(target=self._run_macro_thread, daemon=True)
        macro_thread.start()
    
    def create_gui(self):
        """GUI 생성"""
        self.root = tk.Tk()
        self.root.title("하드웨어 수준 Path of Exile 매크로")
        self.root.geometry("800x600")
        
        # 상단 프레임
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.status_label = tk.Label(top_frame, text="영역을 선택하려면 '영역 선택' 버튼을 클릭하세요")
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        self.select_btn = tk.Button(top_frame, text="영역 선택", command=self.select_area)
        self.select_btn.pack(side=tk.RIGHT, padx=5)
        
        # 좌표 정보
        coords_frame = tk.Frame(self.root)
        coords_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(coords_frame, text="시작 좌표:").grid(row=0, column=0, sticky=tk.W)
        self.start_pos_label = tk.Label(coords_frame, text=str(self.start_pos) if self.start_pos else "미설정")
        self.start_pos_label.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        tk.Label(coords_frame, text="끝 좌표:").grid(row=0, column=2, sticky=tk.W, padx=10)
        self.end_pos_label = tk.Label(coords_frame, text=str(self.end_pos) if self.end_pos else "미설정")
        self.end_pos_label.grid(row=0, column=3, sticky=tk.W, padx=5)
        
        # 클릭 설정
        settings_frame = tk.Frame(self.root)
        settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(settings_frame, text="클릭 간격(초):").grid(row=0, column=0, sticky=tk.W)
        self.click_delay = tk.DoubleVar(value=0.1)
        tk.Entry(settings_frame, textvariable=self.click_delay, width=5).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        self.use_ctrl_click = tk.BooleanVar(value=True)
        tk.Checkbutton(settings_frame, text="Ctrl 키 유지", variable=self.use_ctrl_click).grid(row=0, column=2, sticky=tk.W, padx=10)
        
        # 실행 버튼
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.run_btn = tk.Button(button_frame, text="매크로 실행 (F6)", command=self.run_macro)
        self.run_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = tk.Button(button_frame, text="매크로 중지 (F7)", command=self.stop_macro, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # 제외할 셀 목록
        excluded_frame = tk.Frame(self.root)
        excluded_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(excluded_frame, text="제외할 셀:").grid(row=0, column=0, sticky=tk.W)
        self.excluded_label = tk.Label(excluded_frame, text=str(self.excluded_cells))
        self.excluded_label.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        tk.Button(excluded_frame, text="목록 초기화", command=self.clear_excluded).grid(row=0, column=2, padx=5)
        
        # 메인 프레임 (스크린샷 및 그리드)
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 캔버스 (스크린샷 및 그리드 표시)
        self.canvas = tk.Canvas(self.main_frame, bg="white", bd=2, relief=tk.SUNKEN)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        
        # 단축키 등록
        keyboard.add_hotkey('f6', self.run_macro)
        keyboard.add_hotkey('f7', self.stop_macro)
        
        # 초기 상태 업데이트
        self.update_canvas()
        
        self.root.mainloop()
    
    def select_area(self):
        """영역 선택 모드 시작"""
        self.status_label.config(text="화면에서 드래그하여 영역을 선택하세요...")
        
        # 현재 창 숨기기
        self.root.withdraw()
        
        # 전체 화면 오버레이 창 생성
        self.overlay = tk.Toplevel()
        self.overlay.attributes('-fullscreen', True)
        self.overlay.attributes('-alpha', 0.3)
        self.overlay.attributes('-topmost', True)
        self.overlay.configure(bg='black')
        
        # 오버레이 캔버스
        self.overlay_canvas = tk.Canvas(self.overlay, bg="black", highlightthickness=0)
        self.overlay_canvas.pack(fill=tk.BOTH, expand=True)
        
        # 안내 텍스트
        self.overlay_canvas.create_text(
            self.overlay.winfo_screenwidth() // 2,
            self.overlay.winfo_screenheight() // 2,
            text="드래그하여 영역을 선택하세요. ESC를 눌러 취소합니다.",
            fill="white", font=("Arial", 16)
        )
        
        # 드래그 변수
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.drag_rect = None
        self.dragging = False
        
        # 이벤트 바인딩
        self.overlay.bind("<ButtonPress-1>", self.on_drag_start)
        self.overlay.bind("<B1-Motion>", self.on_drag_motion)
        self.overlay.bind("<ButtonRelease-1>", self.on_drag_release)
        self.overlay.bind("<Escape>", lambda e: self.cancel_selection())
        
        # 스크린샷
        self.full_screenshot = ImageGrab.grab()
    
    def on_drag_start(self, event):
        """드래그 시작"""
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self.dragging = True
        
        # 이전 사각형 삭제
        if self.drag_rect:
            self.overlay_canvas.delete(self.drag_rect)
    
    def on_drag_motion(self, event):
        """드래그 중"""
        if not self.dragging:
            return
        
        # 이전 사각형 삭제
        if self.drag_rect:
            self.overlay_canvas.delete(self.drag_rect)
        
        # 새 사각형 그리기
        self.drag_rect = self.overlay_canvas.create_rectangle(
            self.drag_start_x, self.drag_start_y,
            event.x, event.y,
            outline="red", width=2
        )
    
    def on_drag_release(self, event):
        """드래그 완료"""
        if not self.dragging:
            return
            
        self.dragging = False
        
        # 좌표 계산 (시작점이 항상 좌상단, 끝점이 항상 우하단)
        start_x = min(self.drag_start_x, event.x)
        start_y = min(self.drag_start_y, event.y)
        end_x = max(self.drag_start_x, event.x)
        end_y = max(self.drag_start_y, event.y)
        
        # 너무 작은 영역은 무시
        if end_x - start_x < 10 or end_y - start_y < 10:
            messagebox.showwarning("경고", "선택한 영역이 너무 작습니다. 다시 시도하세요.")
            return
        
        # 좌표 저장
        self.start_pos = (start_x, start_y)
        self.end_pos = (end_x, end_y)
        
        # 스크린샷 잘라내기
        self.screenshot = self.full_screenshot.crop((start_x, start_y, end_x, end_y))
        
        # 오버레이 창 닫기
        self.overlay.destroy()
        
        # 메인 창 복원
        self.root.deiconify()
        self.root.focus_force()
        
        # 상태 업데이트
        self.status_label.config(text="영역 선택 완료")
        self.start_pos_label.config(text=str(self.start_pos))
        self.end_pos_label.config(text=str(self.end_pos))
        
        # 설정 저장
        self.save_config()
        
        # 캔버스 업데이트
        self.update_canvas()
    
    def cancel_selection(self):
        """선택 취소"""
        if hasattr(self, 'overlay') and self.overlay.winfo_exists():
            self.overlay.destroy()
        self.root.deiconify()
        self.status_label.config(text="영역 선택 취소됨")
    
    def update_canvas(self):
        """캔버스 업데이트 (스크린샷 및 그리드)"""
        self.canvas.delete("all")
        
        if not self.start_pos or not self.end_pos or not self.screenshot:
            return
        
        try:
            # 캔버스 크기
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            if canvas_width <= 1 or canvas_height <= 1:
                canvas_width = 600
                canvas_height = 400
            
            # 스크린샷 크기
            img_width = self.end_pos[0] - self.start_pos[0]
            img_height = self.end_pos[1] - self.start_pos[1]
            
            # 비율 계산
            scale = min(canvas_width / img_width, canvas_height / img_height)
            new_width = int(img_width * scale)
            new_height = int(img_height * scale)
            
            # 이미지 리사이징 및 표시
            resized_img = self.screenshot.resize((new_width, new_height), Image.LANCZOS)
            self.tk_image = ImageTk.PhotoImage(resized_img)
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_image)
            
            # 그리드 그리기
            cell_width = new_width / self.grid_width
            cell_height = new_height / self.grid_height
            
            for y in range(self.grid_height):
                for x in range(self.grid_width):
                    x1 = x * cell_width
                    y1 = y * cell_height
                    x2 = (x + 1) * cell_width
                    y2 = (y + 1) * cell_height
                    
                    # 제외된 셀인지 확인
                    is_excluded = (x, y) in self.excluded_cells
                    
                    # 셀 사각형 그리기
                    if is_excluded:
                        # 제외된 셀은 반투명 회색으로 표시
                        rect_id = self.canvas.create_rectangle(
                            x1, y1, x2, y2,
                            fill="gray", stipple="gray50",
                            outline="red", width=1
                        )
                    else:
                        # 일반 셀은 테두리만 표시
                        rect_id = self.canvas.create_rectangle(
                            x1, y1, x2, y2,
                            outline="blue", width=1
                        )
                    
                    # 셀 좌표 텍스트
                    self.canvas.create_text(
                        x1 + cell_width/2, y1 + cell_height/2,
                        text=f"{x},{y}",
                        fill="red" if is_excluded else "white"
                    )
        except Exception as e:
            print(f"캔버스 업데이트 오류: {e}")
    
    def on_canvas_click(self, event):
        """캔버스 클릭 처리 (셀 선택/해제)"""
        if not self.start_pos or not self.end_pos:
            return
        
        try:
            # 캔버스 크기
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            if canvas_width <= 1 or canvas_height <= 1:
                canvas_width = 600
                canvas_height = 400
            
            # 스크린샷 크기
            img_width = self.end_pos[0] - self.start_pos[0]
            img_height = self.end_pos[1] - self.start_pos[1]
            
            # 비율 계산
            scale = min(canvas_width / img_width, canvas_height / img_height)
            new_width = int(img_width * scale)
            new_height = int(img_height * scale)
            
            # 셀 크기
            cell_width = new_width / self.grid_width
            cell_height = new_height / self.grid_height
            
            # 클릭한 셀 계산
            cell_x = int(event.x / cell_width)
            cell_y = int(event.y / cell_height)
            
            # 유효한 셀인지 확인
            if 0 <= cell_x < self.grid_width and 0 <= cell_y < self.grid_height:
                # 셀 상태 토글
                cell = (cell_x, cell_y)
                if cell in self.excluded_cells:
                    self.excluded_cells.remove(cell)
                else:
                    self.excluded_cells.append(cell)
                
                # 표시 업데이트
                self.excluded_label.config(text=str(self.excluded_cells))
                self.save_config()
                self.update_canvas()
        except Exception as e:
            print(f"캔버스 클릭 처리 오류: {e}")
    
    def clear_excluded(self):
        """제외 목록 초기화"""
        self.excluded_cells = []
        self.excluded_label.config(text="[]")
        self.save_config()
        self.update_canvas()
    
    def stop_macro(self):
        """매크로 중지"""
        if not self.is_running:
            return
            
        self.is_running = False
        self.status_label.config(text="매크로 중지됨")
        
        # Ctrl 키 해제
        try:
            keyboard.release('ctrl')
        except:
            pass
            
        # 버튼 상태 업데이트
        self.run_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
    
    def _run_macro_thread(self):
        """매크로 실행 (별도 스레드)"""
        self.status_label.config(text="매크로 실행 중...")
        
        # 창 최소화
        self.root.withdraw()
        
        # 잠시 대기 (창 전환용)
        time.sleep(0.5)
        
        try:
            # Ctrl 키 해제 (이전에 눌려있을 수 있음)
            keyboard.release('ctrl')
            
            # 박스 크기
            box_width = self.end_pos[0] - self.start_pos[0]
            box_height = self.end_pos[1] - self.start_pos[1]
            
            # 셀 크기
            cell_width = box_width / self.grid_width
            cell_height = box_height / self.grid_height
            
            # 설정
            delay = self.click_delay.get()
            use_ctrl = self.use_ctrl_click.get()
            
            # Ctrl 키 누르기
            if use_ctrl:
                keyboard.press('ctrl')
                time.sleep(0.1)  # 키 입력 안정화를 위한 짧은 대기
            
            try:
                # 각 셀 순회하며 클릭
                for y in range(self.grid_height):
                    for x in range(self.grid_width):
                        # 실행 중지 확인
                        if not self.is_running:
                            return
                            
                        # 제외된 셀 건너뛰기
                        if (x, y) in self.excluded_cells:
                            continue
                        
                        # 클릭 좌표 계산 (셀 중앙)
                        click_x = int(self.start_pos[0] + (x + 0.5) * cell_width)
                        click_y = int(self.start_pos[1] + (y + 0.5) * cell_height)
                        
                        # 현재 마우스 위치 저장
                        original_pos = mouse.get_position()
                        
                        # 하드웨어 수준 마우스 이동 및 클릭
                        mouse.move(click_x, click_y)
                        time.sleep(0.02)  # 마우스 이동 안정화
                        mouse.press(button='left')
                        time.sleep(0.02)  # 클릭 다운 유지
                        mouse.release(button='left')
                        
                        # 지연
                        if delay > 0:
                            time.sleep(delay)
            finally:
                # Ctrl 키 해제
                if use_ctrl:
                    keyboard.release('ctrl')
            
            self.status_label.config(text="매크로 실행 완료")
        except Exception as e:
            self.status_label.config(text=f"오류 발생: {str(e)}")
            print(f"매크로 실행 오류: {e}")
            
            # 오류 발생 시에도 Ctrl 키 해제
            try:
                keyboard.release('ctrl')
            except:
                pass
        finally:
            # UI 상태 복원
            self.is_running = False
            self.root.deiconify()
            self.root.focus_force()
            self.run_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

if __name__ == "__main__":
    try:
        print("하드웨어 수준 Path of Exile 클릭 매크로를 시작합니다...")
        print("F6: 매크로 실행, F7: 매크로 중지")
        HardwareLevelDragMacro()
    except Exception as e:
        # 오류 로깅
        import traceback
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(f"오류: {str(e)}\n\n{traceback.format_exc()}")
        
        # 오류 메시지 표시
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("오류", f"프로그램 실행 중 오류가 발생했습니다: {str(e)}\n자세한 내용은 error_log.txt 파일을 확인하세요.")
        root.destroy()
