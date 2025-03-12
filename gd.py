import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import requests
from bs4 import BeautifulSoup
import time
import threading
import json
import datetime
import os
from plyer import notification
from PIL import Image, ImageTk
import io
import re
import traceback

class ProductMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("G-DRAGON 재고 알림 프로그램")
        self.root.geometry("900x800")
        self.root.iconbitmap("icon.ico") if os.path.exists("icon.ico") else None
        
        # 상품 정보를 저장할 변수들
        self.products = []
        self.previous_products = []
        self.running = False
        self.thread = None
        self.notification_enabled = tk.BooleanVar(value=True)
        self.check_interval = tk.IntVar(value=60)  # 기본 1분(60초)
        self.console_autoscroll = tk.BooleanVar(value=True)
        self.last_updated = "아직 확인하지 않음"
        
        # 카카오톡 인증 관련 변수
        self.kakao_token = tk.StringVar()
        self.kakao_refresh_token = tk.StringVar()
        self.kakao_token_expires_at = tk.IntVar(value=0)
        self.kakao_client_id = tk.StringVar()
        self.kakao_client_secret = tk.StringVar()
        self.kakao_auth_code = tk.StringVar()
        self.redirect_uri = tk.StringVar(value="https://eateapple.xyz/kakao-redirect.html")
        
        # 프로그램 스타일 설정
        self.setup_styles()
        
        self.load_settings()
        self.create_ui()
    
    def setup_styles(self):
        """스타일 설정"""
        style = ttk.Style()
        
        # 기본 스타일 설정
        style.configure("TFrame", background="#f5f5f5")
        style.configure("TLabel", background="#f5f5f5", font=("맑은 고딕", 9))
        style.configure("TButton", font=("맑은 고딕", 9), padding=2)
        
        # 헤더 스타일
        style.configure("Treeview.Heading", font=("맑은 고딕", 9, "bold"))
        
        # 트리뷰 스타일
        style.configure("Treeview", 
                        background="#ffffff",
                        fieldbackground="#ffffff", 
                        font=("맑은 고딕", 9))
        
        # 버튼 스타일
        style.configure("GD.TButton", 
                       background="#6255f6", 
                       foreground="#ffffff",
                       font=("맑은 고딕", 9, "bold"))
        
        # 노트북 스타일
        style.configure("TNotebook", background="#f5f5f5")
        style.configure("TNotebook.Tab", font=("맑은 고딕", 9, "bold"), padding=[10, 2])
        
    def load_settings(self):
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    self.check_interval.set(settings.get("check_interval", 60))
                    self.notification_enabled.set(settings.get("notification_enabled", True))
                    self.kakao_token.set(settings.get("kakao_token", ""))
                    self.kakao_refresh_token.set(settings.get("kakao_refresh_token", ""))
                    self.kakao_token_expires_at.set(settings.get("kakao_token_expires_at", 0))
                    self.kakao_client_id.set(settings.get("kakao_client_id", ""))
                    self.kakao_client_secret.set(settings.get("kakao_client_secret", ""))
                    self.kakao_auth_code.set(settings.get("kakao_auth_code", ""))
        except Exception as e:
            self.log_message(f"설정 파일 로드 실패: {e}")
        
    def save_settings(self):
        try:
            settings = {
                "check_interval": self.check_interval.get(),
                "notification_enabled": self.notification_enabled.get(),
                "kakao_token": self.kakao_token.get(),
                "kakao_refresh_token": self.kakao_refresh_token.get(),
                "kakao_token_expires_at": self.kakao_token_expires_at.get(),
                "kakao_client_id": self.kakao_client_id.get(),
                "kakao_client_secret": self.kakao_client_secret.get(),
                "kakao_auth_code": self.kakao_auth_code.get()
            }
            with open("settings.json", "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("알림", "설정이 저장되었습니다.")
        except Exception as e:
            self.log_message(f"설정 파일 저장 실패: {e}")
            messagebox.showerror("오류", f"설정 저장 중 오류 발생: {e}")
            
    def update_token_display(self):
        """토큰 정보를 마스킹해서 표시"""
        if not self.kakao_token.get():
            self.masked_token_var.set("설정되지 않음")
            return
            
        token = self.kakao_token.get()
        if len(token) > 10:
            # 앞의 5자와 뒤의 5자만 표시, 나머지는 *로 마스킹
            masked = token[:5] + "*" * (len(token) - 10) + token[-5:]
            self.masked_token_var.set(masked)
        else:
            self.masked_token_var.set("*" * len(token))
            
    def open_auth_page(self):
        """인증 페이지 열기"""
        if not self.kakao_client_id.get():
            messagebox.showwarning("경고", "REST API 키를 먼저 입력해주세요.")
            return
            
        # 카카오 인증 URL 생성 - 권한 추가 (talk_message)
        auth_url = (
            f"https://kauth.kakao.com/oauth/authorize"
            f"?response_type=code"
            f"&client_id={self.kakao_client_id.get()}"
            f"&redirect_uri={self.redirect_uri.get()}"
            f"&scope=profile_nickname,talk_message,friends"  # 필요한 권한 명시
        )
        
        # 브라우저로 인증 URL 열기
        import webbrowser
        webbrowser.open(auth_url)
        
        # 안내 메시지 표시
        messagebox.showinfo("안내", 
            "브라우저에서 카카오 로그인 후 나타나는 인증 코드를 복사하여 '인증 코드' 입력란에 붙여넣기 해주세요.\n\n" +
            "※ 권한 요청 화면에서 '카카오톡 메시지 전송' 권한을 반드시 허용해주세요.")
            
    def get_token_from_auth_code(self):
        """인증 코드로부터 토큰 발급"""
        if not self.kakao_client_id.get() or not self.kakao_auth_code.get():
            messagebox.showwarning("경고", "REST API 키와 인증 코드가 필요합니다.")
            return
            
        try:
            # 토큰 발급 요청 데이터
            token_data = {
                "grant_type": "authorization_code",
                "client_id": self.kakao_client_id.get(),
                "redirect_uri": self.redirect_uri.get(),
                "code": self.kakao_auth_code.get()
            }
            
            # 클라이언트 시크릿이 있으면 추가
            if self.kakao_client_secret.get():
                token_data["client_secret"] = self.kakao_client_secret.get()
            
            # 토큰 발급 요청
            response = requests.post(
                "https://kauth.kakao.com/oauth/token",
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                data=token_data
            )
            
            if response.status_code == 200:
                token_info = response.json()
                
                # 토큰 정보 저장
                self.kakao_token.set(token_info.get("access_token", ""))
                self.kakao_refresh_token.set(token_info.get("refresh_token", ""))
                
                # 만료 시간 계산 (현재 시간 + expires_in - 여유 시간 10분)
                current_time = int(time.time())
                expires_in = token_info.get("expires_in", 21600)  # 기본 6시간
                self.kakao_token_expires_at.set(current_time + expires_in - 600)
                
                # 설정 저장
                self.save_settings()
                
                # 화면 업데이트
                self.update_token_status()
                self.update_token_display()
                
                messagebox.showinfo("성공", "카카오톡 토큰이 발급되었습니다.")
                self.log_message("카카오톡 토큰 발급 성공")
            else:
                error_msg = f"토큰 발급 실패: {response.status_code}\n{response.text}"
                messagebox.showerror("오류", error_msg)
                self.log_message(error_msg)
                
        except Exception as e:
            error_msg = f"토큰 발급 중 오류 발생: {str(e)}"
            messagebox.showerror("오류", error_msg)
            self.log_message(error_msg)
    
    def check_and_refresh_token(self):
        """토큰 만료 확인 및 리프레시"""
        # 리프레시 토큰이 없거나 클라이언트 ID가 없으면 토큰 갱신 불가
        if not self.kakao_refresh_token.get() or not self.kakao_client_id.get():
            return False
        
        # 현재 시간이 만료 시간 이후인지 확인
        current_time = int(time.time())
        if current_time < self.kakao_token_expires_at.get():
            # 토큰이 아직 유효함
            return True
            
        try:
            # 리프레시 토큰으로 새 액세스 토큰 요청
            refresh_data = {
                "grant_type": "refresh_token",
                "client_id": self.kakao_client_id.get(),
                "refresh_token": self.kakao_refresh_token.get()
            }
            
            # 클라이언트 시크릿이 있으면 추가
            if self.kakao_client_secret.get():
                refresh_data["client_secret"] = self.kakao_client_secret.get()
            
            response = requests.post(
                "https://kauth.kakao.com/oauth/token",
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                data=refresh_data
            )
            
            if response.status_code == 200:
                token_info = response.json()
                self.kakao_token.set(token_info.get("access_token", ""))
                
                # 새 리프레시 토큰이 있으면 업데이트
                if "refresh_token" in token_info:
                    self.kakao_refresh_token.set(token_info.get("refresh_token"))
                
                # 만료 시간 계산 (현재 시간 + expires_in - 여유 시간 10분)
                expires_in = token_info.get("expires_in", 21600)  # 기본 6시간
                self.kakao_token_expires_at.set(current_time + expires_in - 600)
                
                # 설정 저장
                self.save_settings()
                
                self.log_message("카카오톡 액세스 토큰이 자동으로 갱신되었습니다.")
                return True
            else:
                self.log_message(f"카카오톡 토큰 갱신 실패: {response.status_code}, {response.text}")
                return False
                
        except Exception as e:
            self.log_message(f"카카오톡 토큰 갱신 중 오류: {e}")
            return False
    
    def create_ui(self):
        # 메인 프레임과 탭 컨트롤 생성
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 탭 1: 상품 모니터링
        self.monitor_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.monitor_tab, text="상품 모니터링")
        
        # 탭 2: 설정
        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="설정")
        
        # 모니터링 탭 UI 생성
        self.create_monitor_tab()
        
        # 설정 탭 UI 생성
        self.create_settings_tab()
        
        # 상태 표시줄
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.status_label = ttk.Label(status_frame, text=f"상태: 대기 중 | 마지막 업데이트: {self.last_updated}")
        self.status_label.pack(side=tk.LEFT)
        
        # 테마 설정 (지드래곤 컨셉)
        self.root.configure(bg="#000000")
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0")
        style.configure("TButton", background="#f0f0f0")
        
    def create_monitor_tab(self):
        control_frame = ttk.Frame(self.monitor_tab)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 시작/중지 버튼
        self.start_button = ttk.Button(control_frame, text="모니터링 시작", command=self.start_monitoring)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="모니터링 중지", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # 지금 확인 버튼
        self.check_now_button = ttk.Button(control_frame, text="지금 확인", command=self.check_now)
        self.check_now_button.pack(side=tk.LEFT, padx=5)
        
        # 상품 목록 표시 영역
        list_frame = ttk.Frame(self.monitor_tab)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 트리뷰 (표 형태로 상품 목록 표시)
        columns = ("id", "name", "price", "status", "url")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        self.tree.heading("id", text="상품 ID")
        self.tree.heading("name", text="상품명")
        self.tree.heading("price", text="가격")
        self.tree.heading("status", text="상태")
        self.tree.heading("url", text="링크")
        
        self.tree.column("id", width=50)
        self.tree.column("name", width=300)
        self.tree.column("price", width=100)
        self.tree.column("status", width=100)
        self.tree.column("url", width=100)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 트리뷰 항목 더블클릭 이벤트 바인딩
        self.tree.bind("<Double-1>", self.on_item_double_click)
        
        # 콘솔 출력
        console_frame = ttk.LabelFrame(self.monitor_tab, text="로그")
        console_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=10)
        
        self.console = scrolledtext.ScrolledText(console_frame, height=10, wrap=tk.WORD)
        self.console.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.console.configure(state=tk.DISABLED)
        
    def create_settings_tab(self):
        settings_frame = ttk.LabelFrame(self.settings_tab, text="모니터링 설정")
        settings_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 체크 간격 설정
        interval_frame = ttk.Frame(settings_frame)
        interval_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(interval_frame, text="확인 간격 (초):").pack(side=tk.LEFT, padx=5)
        interval_spinner = ttk.Spinbox(interval_frame, from_=10, to=3600, textvariable=self.check_interval, width=10)
        interval_spinner.pack(side=tk.LEFT, padx=5)
        
        # 알림 설정
        notification_frame = ttk.Frame(settings_frame)
        notification_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Checkbutton(notification_frame, text="데스크톱 알림 사용", variable=self.notification_enabled).pack(side=tk.LEFT, padx=5)
        
        # 카카오톡 알림 설정
        kakao_frame = ttk.LabelFrame(self.settings_tab, text="카카오톡 알림 설정")
        kakao_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 1단계: API 키 및 시크릿 입력
        step1_frame = ttk.LabelFrame(kakao_frame, text="Step 1: API 키 설정")
        step1_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # API 키 입력
        api_frame = ttk.Frame(step1_frame)
        api_frame.pack(fill=tk.X, padx=5, pady=3)
        ttk.Label(api_frame, text="REST API 키:").pack(side=tk.LEFT, padx=5)
        api_entry = ttk.Entry(api_frame, textvariable=self.kakao_client_id, width=40)
        api_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 클라이언트 시크릿 입력
        secret_frame = ttk.Frame(step1_frame)
        secret_frame.pack(fill=tk.X, padx=5, pady=3)
        ttk.Label(secret_frame, text="Client Secret:").pack(side=tk.LEFT, padx=5)
        secret_entry = ttk.Entry(secret_frame, textvariable=self.kakao_client_secret, width=40)
        secret_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 인증키 받기 버튼
        get_auth_button = ttk.Button(step1_frame, text="인증키 받기", command=self.open_auth_page)
        get_auth_button.pack(anchor=tk.E, padx=10, pady=5)
        
        # 2단계: 인증 코드 입력
        step2_frame = ttk.LabelFrame(kakao_frame, text="Step 2: 인증 코드 입력")
        step2_frame.pack(fill=tk.X, padx=10, pady=5)
        
        auth_code_frame = ttk.Frame(step2_frame)
        auth_code_frame.pack(fill=tk.X, padx=5, pady=3)
        ttk.Label(auth_code_frame, text="인증 코드:").pack(side=tk.LEFT, padx=5)
        auth_code_entry = ttk.Entry(auth_code_frame, textvariable=self.kakao_auth_code, width=40)
        auth_code_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 3단계: 토큰 발급
        step3_frame = ttk.LabelFrame(kakao_frame, text="Step 3: 토큰 발급")
        step3_frame.pack(fill=tk.X, padx=10, pady=5)
        
        get_token_button = ttk.Button(step3_frame, text="인증하기", command=self.get_token_from_auth_code)
        get_token_button.pack(anchor=tk.E, padx=10, pady=5)
        
        # 토큰 정보 표시 영역
        token_info_frame = ttk.LabelFrame(kakao_frame, text="토큰 정보")
        token_info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 액세스 토큰 표시
        access_token_frame = ttk.Frame(token_info_frame)
        access_token_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(access_token_frame, text="액세스 토큰:").pack(side=tk.LEFT, padx=5)
        
        # 마스킹된 토큰 표시
        self.masked_token_var = tk.StringVar(value="설정되지 않음")
        masked_label = ttk.Label(access_token_frame, textvariable=self.masked_token_var)
        masked_label.pack(side=tk.LEFT, padx=5)
        
        # 토큰 상태 표시
        token_status_frame = ttk.Frame(token_info_frame)
        token_status_frame.pack(fill=tk.X, padx=5, pady=2)
        self.token_status_var = tk.StringVar(value="토큰 상태: 설정되지 않음")
        token_status_label = ttk.Label(token_status_frame, textvariable=self.token_status_var)
        token_status_label.pack(anchor=tk.W, padx=5)
        
        # 토큰 테스트 및 수동 갱신 버튼
        token_button_frame = ttk.Frame(kakao_frame)
        token_button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        test_token_button = ttk.Button(token_button_frame, text="카카오톡 테스트 전송", command=self.test_kakao_message)
        test_token_button.pack(side=tk.LEFT, padx=5, pady=5)
        
        refresh_token_button = ttk.Button(token_button_frame, text="토큰 수동 갱신", command=self.manual_refresh_token)
        refresh_token_button.pack(side=tk.LEFT, padx=5, pady=5)
        
        info_button = ttk.Button(token_button_frame, text="설정 방법 안내", command=self.show_token_info)
        info_button.pack(side=tk.RIGHT, padx=5, pady=5)
        
        self.update_token_display()
        token_status_label = ttk.Label(token_info_frame, textvariable=self.token_status_var, foreground="#666666")
        token_status_label.pack(anchor=tk.W, padx=5, pady=2)
        
        # 저장 버튼
        save_button = ttk.Button(self.settings_tab, text="설정 저장", command=self.save_settings)
        save_button.pack(anchor=tk.E, padx=10, pady=10)
        
        # 현재 토큰 상태 업데이트
        self.update_token_status()
        
        # 친구 목록 섹션 - create_settings_tab 메서드 내에 추가
        friend_frame = ttk.LabelFrame(kakao_frame, text="친구에게 메시지 보내기")
        friend_frame.pack(fill=tk.X, padx=10, pady=5)
    
        # 친구 목록 저장 변수
        self.friends_list = []
    
        # 친구 목록 새로고침 버튼
        refresh_friends_button = ttk.Button(friend_frame, text="친구 목록 새로고침", command=self.refresh_friends_list)
        refresh_friends_button.pack(anchor=tk.W, padx=10, pady=5)
    
        # 친구 목록 콤보박스
        self.friends_var = tk.StringVar()
        self.friends_combobox = ttk.Combobox(friend_frame, textvariable=self.friends_var, state="readonly")
        self.friends_combobox.pack(fill=tk.X, padx=10, pady=5)
    
        # 친구에게 테스트 메시지 보내기 버튼
        send_to_friend_button = ttk.Button(friend_frame, text="선택한 친구에게 테스트 메시지 보내기", command=self.test_friend_message)
        send_to_friend_button.pack(anchor=tk.E, padx=10, pady=5)
        
    def log_message(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        self.console.configure(state=tk.NORMAL)
        self.console.insert(tk.END, log_message)
        if self.console_autoscroll.get():
            self.console.see(tk.END)
        self.console.configure(state=tk.DISABLED)
        
    def start_monitoring(self):
        if not self.running:
            self.running = True
            self.start_button.configure(state=tk.DISABLED)
            self.stop_button.configure(state=tk.NORMAL)
            self.log_message("모니터링을 시작합니다.")
            
            # 초기 상품 정보 가져오기
            self.check_products()
            
            # 모니터링 스레드 시작
            self.thread = threading.Thread(target=self.monitoring_thread, daemon=True)
            self.thread.start()
    
    def stop_monitoring(self):
        if self.running:
            self.running = False
            self.start_button.configure(state=tk.NORMAL)
            self.stop_button.configure(state=tk.DISABLED)
            self.log_message("모니터링을 중지합니다.")
    
    def check_now(self):
        threading.Thread(target=self.check_products, daemon=True).start()
        
    def check_products(self):
        try:
            # 실제 URL에서 HTML 가져오기
            url = "https://withmuulive.com/product/list.html?cate_no=53"
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            }
            
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            # HTML 파싱
            soup = BeautifulSoup(response.text, "html.parser")
            
            # 상품 목록 가져오기 - 정확한 CSS 선택자 사용
            product_list = soup.select("ul.prdList.grid2 > li.xans-record-")
            
            self.log_message(f"웹페이지에서 {len(product_list)}개의 상품을 발견했습니다.")
            
            # 이전 상태 저장
            self.previous_products = self.products.copy() if self.products else []
            
            # 상품 정보 업데이트
            self.products = []
            
            for product in product_list:
                try:
                    # 상품 ID
                    product_id = product.get("id", "").replace("anchorBoxId_", "")
                    
                    # 상품명 파싱: strong.name 전체 텍스트 가져오기
                    product_name = "이름 없음"
                    name_strong = product.select_one("div.description strong.name")
                    if name_strong:
                        # strong 태그 안의 모든 텍스트 가져오기 (alt 속성 백업)
                        product_name = name_strong.get_text(strip=True)
                        if not product_name:
                            img_tag = product.select_one("div.thumbnail a img")
                            if img_tag and img_tag.has_attr("alt"):
                                product_name = img_tag["alt"]
                    
                    # 가격 파싱 - HTML 구조에 직접 맞춤
                    product_price = "가격 정보 없음"
                    spec_list = product.select_one("ul.spec")
                    if spec_list:
                        price_spans = spec_list.select("li span")
                        for span in price_spans:
                            if "원" in span.text:
                                product_price = span.text.strip()
                                break
                    
                    # 품절 여부 확인
                    sold_out_tag = product.select_one("div.icon img[alt='품절']")
                    status = "품절" if sold_out_tag else "구매가능"
                    
                    # 상품 URL
                    url_tag = product.select_one("div.thumbnail a")
                    product_url = "https://withmuulive.com" + url_tag["href"] if url_tag and "href" in url_tag.attrs else ""
                    
                    # 이미지 URL
                    img_tag = product.select_one("div.thumbnail a img")
                    img_url = img_tag["src"] if img_tag and "src" in img_tag.attrs else ""
                    
                    self.products.append({
                        "id": product_id,
                        "name": product_name,
                        "price": product_price,
                        "status": status,
                        "url": product_url,
                        "img_url": img_url
                    })
                    
                    self.log_message(f"상품 정보: ID={product_id}, 이름={product_name}, 가격={product_price}, 상태={status}")
                    
                except Exception as e:
                    self.log_message(f"상품 정보 파싱 오류: {e}")
                    self.log_message(traceback.format_exc())
            
            # UI 업데이트
            self.update_product_list()
            
            # 변경 사항 확인 및 알림
            self.check_for_changes()
            
            # 상태 표시줄 업데이트
            self.last_updated = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.status_label.configure(text=f"상태: 모니터링 중 | 마지막 업데이트: {self.last_updated}")
            
            self.log_message(f"상품 정보를 업데이트했습니다. 총 {len(self.products)}개 상품이 있습니다.")
            
        except Exception as e:
            self.log_message(f"상품 정보 가져오기 실패: {e}")
    
    def update_product_list(self):
        # 트리뷰 초기화
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 상품 정보 추가
        for product in self.products:
            self.tree.insert("", tk.END, values=(
                product["id"],
                product["name"],
                product["price"],
                product["status"],
                "바로가기"
            ))
    
    def on_item_double_click(self, event):
        """아이템 더블 클릭 시 URL 열기"""
        # 클릭된 아이템 가져오기
        item = self.tree.identify_row(event.y)
        if not item:
            return
            
        # 클릭된 열(컬럼) 가져오기
        column = self.tree.identify_column(event.x)
        col_no = int(column.replace('#', ''))
        
        # 해당 아이템의 값 가져오기
        values = self.tree.item(item, 'values')
        
        # URL 열이거나 다른 열이지만 URL이 있는 경우
        if col_no == 5 or (values and len(values) > 4):  # URL 컬럼(5번째) 또는 다른 컬럼
            product_id = values[0]  # 상품 ID
            
            # ID로 상품 찾기
            for product in self.products:
                if product["id"] == product_id:
                    import webbrowser
                    webbrowser.open(product["url"])
                    self.log_message(f"상품 페이지 열기: {product['name']}")
                    break
    
    def test_kakao_message(self):
        """카카오톡 테스트 메시지 보내기"""
        if not self.kakao_token.get():
            messagebox.showwarning("알림", "카카오톡 액세스 토큰이 설정되지 않았습니다.\n설정 탭에서 먼저 토큰을 입력해주세요.")
            return
        
        try:
            message = "이것은 G-DRAGON 재고 알리미 테스트 메시지입니다.\n정상적으로 카카오톡 메시지가 전송되었습니다!"
            
            headers = {
                "Authorization": f"Bearer {self.kakao_token.get()}",
                "Content-Type": "application/x-www-form-urlencoded"
            }
            
            template_object = {
                "object_type": "text",
                "text": message,
                "link": {
                    "web_url": "https://withmuulive.com/product/list.html?cate_no=53",
                    "mobile_web_url": "https://withmuulive.com/product/list.html?cate_no=53"
                },
                "button_title": "상품 페이지 가기"
            }
            
            data = {
                "template_object": json.dumps(template_object)
            }
            
            response = requests.post(
                "https://kapi.kakao.com/v2/api/talk/memo/default/send",
                headers=headers,
                data=data
            )
            
            if response.status_code == 200:
                messagebox.showinfo("성공", "카카오톡 테스트 메시지가 성공적으로 전송되었습니다.")
                self.log_message("카카오톡 테스트 메시지 전송 성공")
            else:
                messagebox.showerror("오류", f"카카오톡 메시지 전송 실패: {response.status_code}\n{response.text}")
                self.log_message(f"카카오톡 테스트 메시지 전송 실패: {response.status_code} {response.text}")
                
        except Exception as e:
            messagebox.showerror("오류", f"카카오톡 테스트 메시지 전송 중 오류 발생: {str(e)}")
            self.log_message(f"카카오톡 테스트 메시지 전송 오류: {e}")
    
    def check_for_changes(self):
        if not self.previous_products:
            return
        
        # 각 상품에 대해 변경 사항 확인
        for new_product in self.products:
            # 이전 상태에서 같은 ID의 상품 찾기
            old_product = next((p for p in self.previous_products if p["id"] == new_product["id"]), None)
            
            if old_product:
                # 품절 상태가 변경된 경우
                if old_product["status"] != new_product["status"]:
                    self.log_message(f"[변경 감지] {new_product['name']} - {old_product['status']} → {new_product['status']}")
                    
                    if old_product["status"] == "품절" and new_product["status"] == "구매가능":
                        # 품절 → 구매가능으로 변경된 경우 알림
                        message = f"{new_product['name']}이(가) 이제 구매 가능합니다!"
                        self.send_notification(message, new_product["url"])
                    
                    elif old_product["status"] == "구매가능" and new_product["status"] == "품절":
                        # 구매가능 → 품절로 변경된 경우 알림
                        message = f"{new_product['name']}이(가) 품절되었습니다."
                        self.send_notification(message, new_product["url"])
    
    def get_friends(self):
        """카카오톡 친구 목록 가져오기"""
        if not self.kakao_token.get():
            self.log_message("토큰이 설정되지 않았습니다.")
            return []
        
        try:
            headers = {
                "Authorization": f"Bearer {self.kakao_token.get()}"
            }
            response = requests.get(
                "https://kapi.kakao.com/v1/api/talk/friends",
                headers=headers
            )
            
            if response.status_code == 200:
                friends = response.json().get("elements", [])
                self.log_message(f"친구 목록 {len(friends)}명을 가져왔습니다.")
                return friends
            else:
                self.log_message(f"친구 목록 가져오기 실패: {response.status_code} {response.text}")
                return []
        except Exception as e:
            self.log_message(f"친구 목록 가져오기 오류: {e}")
            return []
    
    def send_to_friend(self, friend_id, message, url):
        """친구에게 카카오톡 메시지 보내기"""
        try:
            headers = {
                "Authorization": f"Bearer {self.kakao_token.get()}",
                "Content-Type": "application/x-www-form-urlencoded"
            }
            
            data = {
                "receiver_uuids": f'["{friend_id}"]',
                "template_object": json.dumps({
                    "object_type": "text",
                    "text": message,
                    "link": {
                        "web_url": url,
                        "mobile_web_url": url
                    },
                    "button_title": "상품 보기"
                })
            }
            
            self.log_message(f"요청 데이터: {data}")
            
            response = requests.post(
                "https://kapi.kakao.com/v1/api/talk/friends/message/default/send",
                headers=headers,
                data=data
            )
            
            self.log_message(f"응답 상태: {response.status_code}")
            self.log_message(f"응답 내용: {response.text}")
            
            if response.status_code == 200:
                self.log_message(f"친구에게 카카오톡 메시지가 전송되었습니다.")
                return True
            else:
                self.log_message(f"친구에게 메시지 전송 실패: {response.status_code} {response.text}")
                return False
        except Exception as e:
            self.log_message(f"친구에게 메시지 전송 오류: {e}")
            return False
    
    def refresh_friends_list(self):
        """친구 목록 갱신"""
        friends = self.get_friends()
        if friends:
            self.friends_list = friends  # 친구 정보 저장
            self.friends_combobox['values'] = [f"{friend['profile_nickname']} ({friend['uuid']})" for friend in friends]
            if len(friends) > 0:
                self.friends_combobox.current(0)  # 첫 번째 친구 선택
            messagebox.showinfo("성공", f"친구 목록을 갱신했습니다. ({len(friends)}명)")
        else:
            messagebox.showwarning("알림", "친구 목록을 가져오지 못했습니다.")
    
    def test_friend_message(self):
        """선택한 친구에게 테스트 메시지 보내기"""
        if not self.friends_var.get():
            messagebox.showwarning("알림", "친구를 선택해주세요.")
            return
        
        # 선택된 친구 UUID 추출
        selected = self.friends_var.get()
        uuid = selected.split("(")[1].split(")")[0]  # 형식: "닉네임 (UUID)"에서 UUID 추출
        
        result = self.send_to_friend(
            uuid, 
            "이것은 G-DRAGON 재고 알리미 테스트 메시지입니다.\n정상적으로 카카오톡 메시지가 전송되었습니다!",
            "https://withmuulive.com/product/list.html?cate_no=53"
        )
        
        if result:
            messagebox.showinfo("성공", "선택한 친구에게 테스트 메시지가 전송되었습니다.")
        else:
            messagebox.showerror("실패", "친구에게 메시지 전송에 실패했습니다.")
    
    def send_notification(self, message, url):
        # 데스크톱 알림
        if self.notification_enabled.get():
            try:
                notification.notify(
                    title="G-DRAGON 상품 알림",
                    message=message,
                    app_name="G-DRAGON 재고 알림",
                    timeout=10
                )
            except Exception as e:
                self.log_message(f"데스크톱 알림 전송 실패: {e}")
        
        # 카카오톡 알림
        if self.kakao_token.get():
            try:
                self.send_kakao_message(message, url)
            except Exception as e:
                self.log_message(f"카카오톡 알림 전송 실패: {e}")
 
        # 친구에게 알림 (설정된 경우)
        if hasattr(self, 'friends_var') and self.friends_var.get():
            try:
                selected = self.friends_var.get()
                uuid = selected.split("(")[1].split(")")[0]
                self.send_to_friend(uuid, message, url)
            except Exception as e:
                self.log_message(f"친구에게 알림 전송 실패: {e}")
    
    def send_kakao_message(self, message, url):
        try:
            headers = {
                "Authorization": f"Bearer {self.kakao_token.get()}",
                "Content-Type": "application/x-www-form-urlencoded"
            }
            
            template_object = {
                "object_type": "text",
                "text": message,
                "link": {
                    "web_url": url,
                    "mobile_web_url": url
                },
                "button_title": "상품 보기"
            }
            
            data = {
                "template_object": json.dumps(template_object)
            }
            
            response = requests.post(
                "https://kapi.kakao.com/v2/api/talk/memo/default/send",
                headers=headers,
                data=data
            )
            
            if response.status_code == 200:
                self.log_message("카카오톡 알림이 전송되었습니다.")
            else:
                self.log_message(f"카카오톡 알림 전송 실패: {response.status_code} {response.text}")
                
        except Exception as e:
            self.log_message(f"카카오톡 알림 전송 오류: {e}")
    
    def monitoring_thread(self):
        while self.running:
            # 설정된 간격만큼 대기
            for i in range(self.check_interval.get()):
                if not self.running:
                    break
                time.sleep(1)
            
            if self.running:
                self.check_products()
    
    def update_token_status(self):
        """토큰 상태 업데이트 및 표시"""
        if not self.kakao_token.get():
            self.token_status_var.set("토큰 상태: 설정되지 않음")
            return
            
        current_time = int(time.time())
        expires_at = self.kakao_token_expires_at.get()
        
        if expires_at == 0:
            self.token_status_var.set("토큰 상태: 유효기간 정보 없음")
        elif current_time > expires_at:
            self.token_status_var.set("토큰 상태: 만료됨 (갱신 필요)")
        else:
            # 남은 시간 계산 (시:분:초)
            remaining = expires_at - current_time
            hours = remaining // 3600
            minutes = (remaining % 3600) // 60
            seconds = remaining % 60
            
            self.token_status_var.set(f"토큰 상태: 유효함 (남은 시간: {hours:02d}:{minutes:02d}:{seconds:02d})")
    
    def manual_refresh_token(self):
        """수동으로 토큰 갱신"""
        if not self.kakao_refresh_token.get() or not self.kakao_client_id.get():
            messagebox.showwarning("알림", "리프레시 토큰과 REST API 키가 필요합니다.")
            return
            
        try:
            refresh_data = {
                "grant_type": "refresh_token",
                "client_id": self.kakao_client_id.get(),
                "refresh_token": self.kakao_refresh_token.get()
            }
            
            # 클라이언트 시크릿이 있으면 추가
            if self.kakao_client_secret.get():
                refresh_data["client_secret"] = self.kakao_client_secret.get()
            
            response = requests.post(
                "https://kauth.kakao.com/oauth/token",
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                data=refresh_data
            )
            
            if response.status_code == 200:
                token_info = response.json()
                
                # 토큰 정보 업데이트
                self.kakao_token.set(token_info.get("access_token", ""))
                
                # 새 리프레시 토큰이 있으면 업데이트
                if "refresh_token" in token_info:
                    self.kakao_refresh_token.set(token_info.get("refresh_token"))
                
                # 만료 시간 계산 (현재 시간 + expires_in - 여유 시간 10분)
                current_time = int(time.time())
                expires_in = token_info.get("expires_in", 21600)  # 기본 6시간
                self.kakao_token_expires_at.set(current_time + expires_in - 600)
                
                # 설정 저장
                self.save_settings()
                
                # 상태 업데이트
                self.update_token_status()
                
                messagebox.showinfo("성공", "카카오톡 액세스 토큰이 갱신되었습니다.")
                self.log_message("카카오톡 액세스 토큰 수동 갱신 성공")
            else:
                error_msg = f"토큰 갱신 실패: {response.status_code}\n{response.text}"
                messagebox.showerror("오류", error_msg)
                self.log_message(error_msg)
                
        except Exception as e:
            error_msg = f"토큰 갱신 중 오류 발생: {str(e)}"
            messagebox.showerror("오류", error_msg)
            self.log_message(error_msg)
    
    def show_token_info(self):
        info = """
카카오톡 토큰 설정 방법:

1. 카카오 개발자 사이트에 접속: https://developers.kakao.com
2. 로그인 후 '내 애플리케이션' 메뉴로 이동
3. '애플리케이션 추가하기' 버튼 클릭
4. 앱 이름 입력 (예: G-DRAGON 재고 알리미)
5. 앱 생성 후 '플랫폼' 메뉴에서 'Web' 플랫폼 등록
   - 사이트 도메인에 https://developers.kakao.com 입력
6. '카카오 로그인' 메뉴에서 '활성화 설정' ON으로 변경
7. '동의항목' 메뉴에서 필요한 동의항목 설정
   - '카카오톡 메시지 전송' 권한 필요 (필수 동의항목으로 설정)
8. '보안' 메뉴에서 'Client Secret' 발급 (코드 선택)
9. '요약 정보'에서 'REST API 키' 확인

10. 아래 URL을 웹 브라우저에 입력하여 인증 진행:
    https://kauth.kakao.com/oauth/authorize?response_type=code&client_id=[REST API 키]&redirect_uri=https://developers.kakao.com/tool/demo/oauth

11. 인증 후 리다이렉트 페이지에서 '인증코드 확인' 에서 코드를 복사
12. 아래 명령어를 명령 프롬프트에서 실행 (또는 해당 요청을 보낼 수 있는 도구 사용):
    curl -v -X POST "https://kauth.kakao.com/oauth/token" \\
    -H "Content-Type: application/x-www-form-urlencoded" \\
    -d "grant_type=authorization_code" \\
    -d "client_id=[REST API 키]" \\
    -d "client_secret=[Client Secret]" \\
    -d "redirect_uri=https://developers.kakao.com/tool/demo/oauth" \\
    -d "code=[인증 코드]"

13. 응답에서 다음 값들을 복사하여 설정에 입력:
   - "access_token": 액세스 토큰 필드에 입력
   - "refresh_token": 리프레시 토큰 필드에 입력
   - REST API 키: REST API 키 필드에 입력
   - Client Secret: Client Secret 필드에 입력

14. '설정 저장' 버튼을 클릭하여 저장

※ 리프레시 토큰과 API 키를 설정하면 액세스 토큰이 자동으로 갱신됩니다.
※ 액세스 토큰의 유효 기간은 약 6시간입니다.

팀원에게 보내기
https://kauth.kakao.com/oauth/authorize?response_type=code&client_id=e6f5886762862a7042ae5b867b615b6a&redirect_uri=https://eateapple.xyz/kakao-redirect.html&scope=profile_nickname,talk_message,friends
위의 URL을 팀원이 접속해서 권한동의를해야함.
        """
        
        info_window = tk.Toplevel(self.root)
        info_window.title("카카오톡 토큰 설정 방법")
        info_window.geometry("700x600")
        
        info_text = scrolledtext.ScrolledText(info_window, wrap=tk.WORD, font=("맑은 고딕", 10))
        info_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        info_text.insert(tk.END, info)
        info_text.configure(state=tk.DISABLED)

def main():
    root = tk.Tk()
    app = ProductMonitor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
