import tkinter as tk
from tkinter import messagebox, filedialog
import time
import json
import os
import threading
import keyboard
import mouse
import random
import pygetwindow as gw
from PIL import ImageGrab, Image, ImageTk
import numpy as np

class HardwareLevelDragMacro:
    def __init__(self):
        self.start_pos = None
        self.end_pos = None
        self.grid_width = 12
        self.grid_height = 5
        self.excluded_cells = []
        self.config_file = "hardware_drag_macro_config.json"
        self.inventory_image_path = None
        self.initial_screenshot = None
        self.macro_screenshot = None
        self.initial_tk_image = None
        self.macro_tk_image = None
        self.is_running = False
        self.dragging = False
        self.similarity_threshold = 50  # 이미지 유사성 임계값 (낮을수록 더 엄격함)
        
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
                    self.inventory_image_path = config.get('inventory_image_path')
                    self.similarity_threshold = config.get('similarity_threshold', 50)
            except Exception as e:
                print(f"설정 로드 오류: {e}")
    
    def save_config(self):
        """설정 파일 저장"""
        config = {
            'start_pos': self.start_pos,
            'end_pos': self.end_pos,
            'excluded_cells': self.excluded_cells,
            'inventory_image_path': self.inventory_image_path,
            'similarity_threshold': self.similarity_threshold
        }
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            print(f"설정 저장 오류: {e}")
            
    def create_gui(self):
        """GUI 생성"""
        self.root = tk.Tk()
        self.root.title("하드웨어 수준 Path of Exile 매크로")
        self.root.geometry("500x700")  # 창 크기 줄임
        
        # 인벤토리 이미지 선택 프레임
        image_select_frame = tk.Frame(self.root)
        image_select_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.initial_image_label = tk.Label(image_select_frame, text="빈 인벤토리 이미지를 선택하세요")
        self.initial_image_label.pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            image_select_frame, 
            text="영역 선택", 
            command=self.select_area
        ).pack(side=tk.RIGHT, padx=5)
        
        # 초기 이미지 캔버스
        self.initial_canvas = tk.Canvas(
            self.root, 
            bg="black", 
            width=400, 
            height=166, 
            bd=2, 
            relief=tk.SUNKEN
        )
        self.initial_canvas.pack(padx=10, pady=5)
        
        # 영역 선택 정보
        coords_frame = tk.Frame(self.root)
        coords_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(coords_frame, text="시작 좌표:").grid(row=0, column=0, sticky=tk.W)
        self.start_pos_label = tk.Label(coords_frame, text=str(self.start_pos) if self.start_pos else "미설정")
        self.start_pos_label.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        tk.Label(coords_frame, text="끝 좌표:").grid(row=0, column=2, sticky=tk.W, padx=10)
        self.end_pos_label = tk.Label(coords_frame, text=str(self.end_pos) if self.end_pos else "미설정")
        self.end_pos_label.grid(row=0, column=3, sticky=tk.W, padx=5)
        
        # 매크로 스크린샷 캔버스
        self.macro_canvas = tk.Canvas(
            self.root, 
            bg="black", 
            width=400, 
            height=166, 
            bd=2, 
            relief=tk.SUNKEN
        )
        self.macro_canvas.pack(padx=10, pady=5)
        
        # 상태 정보 레이블
        self.status_label = tk.Label(self.root, text="빈 인벤토리 영역을 선택하세요")
        self.status_label.pack(padx=10, pady=5)
        
        # 클릭 설정
        settings_frame = tk.Frame(self.root)
        settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(settings_frame, text="클릭 간격(초):").grid(row=0, column=0, sticky=tk.W)
        self.click_delay = tk.DoubleVar(value=0.1)
        tk.Entry(settings_frame, textvariable=self.click_delay, width=5).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        self.use_ctrl_click = tk.BooleanVar(value=True)
        tk.Checkbutton(settings_frame, text="Ctrl 키 유지", variable=self.use_ctrl_click).grid(row=0, column=2, sticky=tk.W, padx=10)
        
        # 창 최소화 설정 추가
        self.minimize_window = tk.BooleanVar(value=False)
        tk.Checkbutton(settings_frame, text="실행 시 창 최소화", variable=self.minimize_window).grid(row=0, column=3, sticky=tk.W, padx=10)
        
        # 이미지 비교 설정 프레임
        compare_frame = tk.Frame(self.root)
        compare_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.detect_items = tk.BooleanVar(value=True)
        tk.Checkbutton(compare_frame, text="아이템 감지 (빈 인벤토리와 비교)", variable=self.detect_items).grid(row=0, column=0, sticky=tk.W)
        
        tk.Label(compare_frame, text="유사도 임계값:").grid(row=0, column=1, sticky=tk.W, padx=10)
        self.threshold_slider = tk.Scale(compare_frame, from_=0, to=100, orient=tk.HORIZONTAL, 
                                        variable=tk.IntVar(value=self.similarity_threshold))
        self.threshold_slider.grid(row=0, column=2, sticky=tk.W)
        self.threshold_slider.set(self.similarity_threshold)
        
        # 실행 버튼
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.run_btn = tk.Button(button_frame, text="매크로 실행 (F6)", command=self.run_macro)
        self.run_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = tk.Button(
            button_frame, 
            text="매크로 중지 (F7)", 
            command=self.stop_macro, 
            state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # 제외할 셀 목록
        excluded_frame = tk.Frame(self.root)
        excluded_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(excluded_frame, text="제외할 셀:").grid(row=0, column=0, sticky=tk.W)
        self.excluded_label = tk.Label(excluded_frame, text=str(self.excluded_cells))
        self.excluded_label.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        tk.Button(excluded_frame, text="목록 초기화", command=self.clear_excluded).grid(row=0, column=2, padx=5)
        
        # 단축키 등록
        keyboard.add_hotkey('f6', self.run_macro)
        keyboard.add_hotkey('f7', self.stop_macro)
        
        self.root.mainloop()
    
    def select_area(self):
        """영역 선택 모드 시작"""
        self.status_label.config(text="인벤토리에서 드래그하여 영역을 선택하세요...")
        
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
        
        # 전역 키보드 이벤트 추가 (keyboard 라이브러리 사용)
        # ESC 키를 감지하는 별도의 이벤트 핸들러 추가
        def esc_handler():
            print("ESC 키가 눌림")  # 디버깅용
            self.overlay.destroy()
            self.root.deiconify()
            self.root.focus_force()
            self.status_label.config(text="영역 선택 취소됨")
            
        # 기존 ESC 핫키 제거 후 다시 등록
        try:
            keyboard.remove_hotkey('esc')
        except:
            pass
        keyboard.add_hotkey('esc', esc_handler)
        
        # 이벤트 바인딩
        self.overlay.bind("<ButtonPress-1>", self.on_drag_start)
        self.overlay.bind("<B1-Motion>", self.on_drag_motion)
        self.overlay.bind("<ButtonRelease-1>", self.on_drag_release)
        
        # 추가: 오버레이 창에 키보드 이벤트도 바인딩 (belt and suspenders 방식)
        self.overlay.bind("<Escape>", lambda e: esc_handler())
        
        # 창에 포커스 설정
        self.overlay.focus_set()

    def cancel_selection(self, event=None):
        """선택 취소"""
        print("cancel_selection 호출됨")  # 디버깅용
        if hasattr(self, 'overlay') and self.overlay.winfo_exists():
            self.overlay.destroy()
        self.root.deiconify()
        self.root.focus_force()
        self.status_label.config(text="영역 선택 취소됨")
        
    # 기존 코드 계속...    
  

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
    
        # 제외할 셀 목록 초기화 (중복 제거)
        self.excluded_cells = []
        self.excluded_label.config(text="[]")
    
        # 스크린샷 캡처
        self.initial_screenshot = ImageGrab.grab((start_x, start_y, end_x, end_y))
    
        # 캔버스 크기
        canvas_width = self.initial_canvas.winfo_width()
        canvas_height = self.initial_canvas.winfo_height()
    
        # 이미지 리사이징
        resized_img = self.initial_screenshot.resize(
            (canvas_width, canvas_height), 
            Image.LANCZOS
        )
    
        # Tkinter 이미지로 변환
        self.initial_tk_image = ImageTk.PhotoImage(resized_img)
    
        # 캔버스에 이미지 표시
        self.initial_canvas.delete("all")
        self.initial_canvas.create_image(
            0, 0, 
            anchor=tk.NW, 
            image=self.initial_tk_image
        )
    
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
    
        # 캔버스 업데이트 및 그리드 표시
        self.update_canvas()
        
    def on_canvas_click(self, event):
        """캔버스 클릭 처리 (셀 선택/해제)"""
        if not self.start_pos or not self.end_pos:
            return
        
        try:
            # 캔버스 크기
            canvas_width = self.initial_canvas.winfo_width()
            canvas_height = self.initial_canvas.winfo_height()
            
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

    def update_canvas(self):
        """캔버스 업데이트 (스크린샷 및 그리드)"""
        if not self.start_pos or not self.end_pos or self.initial_screenshot is None:
            return
        
        try:
            # 캔버스 크기
            canvas_width = self.initial_canvas.winfo_width()
            canvas_height = self.initial_canvas.winfo_height()
            
            # 스크린샷 크기
            img_width = self.end_pos[0] - self.start_pos[0]
            img_height = self.end_pos[1] - self.start_pos[1]
            
            # 비율 계산
            scale = min(canvas_width / img_width, canvas_height / img_height)
            new_width = int(img_width * scale)
            new_height = int(img_height * scale)
            
            # 이미지 리사이징
            resized_img = self.initial_screenshot.resize((new_width, new_height), Image.LANCZOS)
            
            # Tkinter 이미지로 변환
            self.initial_tk_image = ImageTk.PhotoImage(resized_img)
            
            # 캔버스 초기화
            self.initial_canvas.delete("all")
            
            # 이미지 표시
            self.initial_canvas.create_image(0, 0, anchor=tk.NW, image=self.initial_tk_image)
            
            # 셀 크기
            cell_width = new_width / self.grid_width
            cell_height = new_height / self.grid_height
            
            # 그리드 그리기
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
                        rect_id = self.initial_canvas.create_rectangle(
                            x1, y1, x2, y2,
                            fill="gray", stipple="gray50",
                            outline="red", width=1
                        )
                    else:
                        # 일반 셀은 테두리만 표시
                        rect_id = self.initial_canvas.create_rectangle(
                            x1, y1, x2, y2,
                            outline="blue", width=1
                        )
                    
                    # 셀 좌표 텍스트
                    self.initial_canvas.create_text(
                        x1 + cell_width/2, y1 + cell_height/2,
                        text=f"{x},{y}",
                        fill="red" if is_excluded else "white"
                    )
            
            # 캔버스에 클릭 이벤트 바인딩
            self.initial_canvas.bind("<Button-1>", self.on_canvas_click)
        
        except Exception as e:
            print(f"캔버스 업데이트 오류: {e}")
    

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
    
    def compare_cell_images(self, initial_img, current_img, cell_x, cell_y, cell_width, cell_height):
        """두 이미지의 특정 셀을 비교하여 아이템 감지 (매우 단순화된 버전)"""
        try:
            # 셀 이미지 추출
            cell_left = int(cell_x * cell_width)
            cell_top = int(cell_y * cell_height)
            cell_right = int((cell_x + 1) * cell_width)
            cell_bottom = int((cell_y + 1) * cell_height)
            
            # 초기 이미지에서 셀 추출
            initial_cell = initial_img.crop((cell_left, cell_top, cell_right, cell_bottom))
            
            # 현재 이미지에서 셀 추출 
            current_cell = current_img.crop((cell_left, cell_top, cell_right, cell_bottom))
            
            # 흑백 이미지로 변환
            initial_gray = initial_cell.convert('L')
            current_gray = current_cell.convert('L')
            
            # NumPy 배열로 변환
            initial_array = np.array(initial_gray)
            current_array = np.array(current_gray)
            
            # 간단한 밝기 임계값
            bright_threshold = 50
            
            # 밝은 픽셀 수 계산
            initial_bright = np.sum(initial_array > bright_threshold)
            current_bright = np.sum(current_array > bright_threshold)
            
            # 총 픽셀 수
            total_pixels = initial_array.size
            
            # 밝은 픽셀 비율
            initial_ratio = initial_bright / total_pixels
            current_ratio = current_bright / total_pixels
            
            # 차이 계산
            diff = current_ratio - initial_ratio
            
            # 로그 출력
            print(f"셀({cell_x},{cell_y}) - 초기: {initial_ratio:.3f}, 현재: {current_ratio:.3f}, 차이: {diff:.3f}")
            
            # 아주 단순한 조건: 현재 밝은 픽셀이 20% 이상이고, 차이가 0.15 이상인 경우 아이템으로 간주
            # 로그를 보니 대부분의 셀에서 현재 값이 0.25 이상이고 차이가 0.2 이상
            has_item = current_ratio > 0.2 and diff > 0.15
            
            # 유사도 반환 (0: 아이템 있음, 100: 아이템 없음)
            return 0 if has_item else 100
            
        except Exception as e:
            print(f"이미지 비교 오류: {e}")
            return 100  # 오류 발생 시 아이템 없다고 간주 (변경됨!)
            
    def run_macro(self):
        """매크로 실행"""
        if not self.start_pos or not self.end_pos:
            messagebox.showwarning("경고", "영역을 먼저 선택해주세요.")
            return
        
        if self.is_running:
            return
        
        # 임계값 설정 저장
        self.similarity_threshold = self.threshold_slider.get()
        self.save_config()
        
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
        
    def _run_macro_thread(self):
        """매크로 실행 (별도 스레드)"""
        self.status_label.config(text="매크로 실행 중...")
        
        # 창 최소화 (선택적)
        if self.minimize_window.get():
            self.root.withdraw()
        
        # 잠시 대기 (창 전환용)
        time.sleep(0.5)
        
        try:
            # Ctrl 키 해제 (이전에 눌려있을 수 있음)
            keyboard.release('ctrl')
            
            # 새 스크린샷 캡처
            macro_screenshot = ImageGrab.grab((self.start_pos[0], self.start_pos[1], self.end_pos[0], self.end_pos[1]))
            
            # 매크로 캔버스 크기
            canvas_width = self.macro_canvas.winfo_width()
            canvas_height = self.macro_canvas.winfo_height()
            
            # 이미지 리사이징
            resized_img = macro_screenshot.resize(
                (canvas_width, canvas_height), 
                Image.LANCZOS
            )
            
            # Tkinter 이미지로 변환
            self.macro_tk_image = ImageTk.PhotoImage(resized_img)
            
            # 기존 캔버스 내용 지우기
            self.macro_canvas.delete("all")
            
            # 이미지 표시
            self.macro_canvas.create_image(
                0, 0, 
                anchor=tk.NW, 
                image=self.macro_tk_image
            )
            
            # 박스 크기
            box_width = self.end_pos[0] - self.start_pos[0]
            box_height = self.end_pos[1] - self.start_pos[1]
            
            # 셀 크기
            cell_width = box_width / self.grid_width
            cell_height = box_height / self.grid_height
            
            # 설정
            delay = self.click_delay.get()
            use_ctrl = self.use_ctrl_click.get()
            detect_items = self.detect_items.get()
            
            # 아이템이 있는 셀 목록 (비교 모드에서 사용)
            item_cells = []
            
            # 아이템 감지 모드일 경우 셀별 비교 수행
            if detect_items and self.initial_screenshot is not None:
                print("아이템 감지 모드 활성화됨")
                # 셀별 비교
                for y in range(self.grid_height):
                    for x in range(self.grid_width):
                        # 제외된 셀 건너뛰기
                        if (x, y) in self.excluded_cells:
                            print(f"셀({x},{y}) - 제외됨")
                            continue
                        
                        # 빈 인벤토리와 현재 인벤토리의 셀 이미지 비교
                        similarity = self.compare_cell_images(
                            self.initial_screenshot, 
                            macro_screenshot, 
                            x, y, 
                            cell_width, cell_height
                        )
                        
                        # 유사도가 임계값보다 낮으면 아이템이 있다고 판단
                        if similarity < 50:  # 유사도가 낮으면 = 아이템이 있다고 판단
                            item_cells.append((x, y))
                            print(f"셀({x},{y}) - 아이템 감지됨!")
                
                # 감지된 아이템 표시 (중요: 이 로그 메시지 확인!)
                print(f"총 {len(item_cells)}개 셀에서 아이템 감지됨: {item_cells}")
                self.status_label.config(text=f"아이템 감지: {len(item_cells)}개 셀")
            
            # Ctrl 키 누르기
            if use_ctrl:
                keyboard.press('ctrl')
                time.sleep(0.1)  # 키 입력 안정화를 위한 짧은 대기
            
            try:
                # 클릭 로직
                if detect_items and self.initial_screenshot is not None:
                    if item_cells:
                        # 아이템 감지 모드: 아이템이 있는 셀만 클릭
                        print(f"아이템이 있는 {len(item_cells)}개 셀만 클릭합니다.")
                        for x, y in item_cells:
                            # 실행 중지 확인
                            if not self.is_running:
                                return
                                
                            # 셀의 좌상단 좌표 계산
                            cell_base_x = int(self.start_pos[0] + x * cell_width)
                            cell_base_y = int(self.start_pos[1] + y * cell_height)
                            
                            # 랜덤 클릭 지점 계산
                            click_x, click_y = self._calculate_random_click_point(
                                cell_base_x, cell_base_y, cell_width, cell_height
                            )
                            
                            # 하드웨어 수준 마우스 이동 및 클릭
                            mouse.move(click_x, click_y)
                            time.sleep(0.02)  # 마우스 이동 안정화
                            mouse.press(button='left')
                            time.sleep(0.02)  # 클릭 다운 유지
                            mouse.release(button='left')
                            
                            # 매크로 캔버스에 클릭 표시
                            if not self.minimize_window.get():
                                # 캔버스 상의 좌표 계산
                                canvas_x = int((x * cell_width) * (canvas_width / box_width))
                                canvas_y = int((y * cell_height) * (canvas_height / box_height))
                                canvas_cell_w = int(cell_width * (canvas_width / box_width))
                                canvas_cell_h = int(cell_height * (canvas_height / box_height))
                                
                                # 클릭한 셀 표시 (빨간색 테두리)
                                self.macro_canvas.create_rectangle(
                                    canvas_x, canvas_y, 
                                    canvas_x + canvas_cell_w, canvas_y + canvas_cell_h,
                                    outline="red", width=2
                                )
                            
                            # 지연
                            if delay > 0:
                                time.sleep(delay)
                    else:
                        print("감지된 아이템이 없습니다. 클릭을 실행하지 않습니다.")
                else:
                    # 기존 방식: 모든 셀 순회하며 클릭 (제외된 셀은 건너뜀)
                    print("일반 모드: 모든 셀을 클릭합니다.")
                    for y in range(self.grid_height):
                        for x in range(self.grid_width):
                            # 실행 중지 확인
                            if not self.is_running:
                                return
                                
                            # 제외된 셀 건너뛰기
                            if (x, y) in self.excluded_cells:
                                continue
                            
                            # 셀의 좌상단 좌표 계산
                            cell_base_x = int(self.start_pos[0] + x * cell_width)
                            cell_base_y = int(self.start_pos[1] + y * cell_height)
                            
                            # 랜덤 클릭 지점 계산
                            click_x, click_y = self._calculate_random_click_point(
                                cell_base_x, cell_base_y, cell_width, cell_height
                            )
                            
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
            
            # 아이템 감지 모드인 경우 결과 표시
            if detect_items and self.initial_screenshot is not None:
                self.status_label.config(text=f"매크로 실행 완료 - {len(item_cells)}개 셀 클릭됨")
            else:
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
            if self.minimize_window.get():
                self.root.deiconify()
            self.root.focus_force()
            self.run_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            
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
    
    def _calculate_random_click_point(self, base_x, base_y, cell_width, cell_height):
        """
        각 셀의 중앙을 기준으로 랜덤한 클릭 지점 계산
        
        :param base_x: 셀의 기준 x 좌표 (좌상단)
        :param base_y: 셀의 기준 y 좌표 (좌상단)
        :param cell_width: 셀의 너비
        :param cell_height: 셀의 높이
        :return: 랜덤한 클릭 x, y 좌표
        """
        # 셀 중앙 계산
        center_x = base_x + cell_width / 2
        center_y = base_y + cell_height / 2
        
        # 랜덤 오차 범위 설정 (셀 크기의 일정 비율)
        x_offset_range = cell_width * 0.2  # 너비의 20% 이내
        y_offset_range = cell_height * 0.2  # 높이의 20% 이내
        
        # 랜덤 오프셋 생성
        x_offset = random.uniform(-x_offset_range, x_offset_range)
        y_offset = random.uniform(-y_offset_range, y_offset_range)
        
        # 랜덤 클릭 지점 계산
        click_x = int(center_x + x_offset)
        click_y = int(center_y + y_offset)
        
        return click_x, click_y
    
    def clear_excluded(self):
        """제외 목록 초기화"""
        self.excluded_cells = []
        self.excluded_label.config(text="[]")
        self.save_config()

# 메인 실행 부분
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
