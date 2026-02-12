import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip
import re


class ScrewInputApp:
    def __init__(self, root):
        self.root = root
        self.root.title("automatic screw input program")
        self.root.geometry("800x900")
        self.root.resizable(True, True)  # 창 크기 조정 가능
        self.root.minsize(700, 800)  # 최소 크기 설정
        
        # 창 드래그 기능을 위한 변수
        self.start_x = 0
        self.start_y = 0
        
        # 상태 변수 (라디오 버튼용)
        self.customer_var = tk.StringVar(value="")  # "SM" 또는 "OTHER"
        self.screw_type_var = tk.StringVar(value="")  # "INNER" 또는 "OUTER"
        self.gauge_var = tk.StringVar(value="")  # "NORMAL" 또는 "BEFORE_PLATING"
        self.plating_value_var = tk.StringVar(value="")  # 도금 값
        
        # 나사 데이터 저장소
        # 구조: {나사이름: [{'type': 'OUTER'/'INNER', 'gauge': 'NORMAL'/'BEFORE_PLATING', 'major': {...}, 'pitch': {...}, 'minor': {...}}, ...]}
        # 같은 나사 이름이 여러 게이지에 있을 수 있음
        self.screw_data = {}
        
        # SM, 외경나사/내경나사, 일반게이지/도금 전 게이지 데이터 초기화
        self.init_screw_data()
        
        # 보유 게이지 목록 (이미지 기반, Ring Gauge) - 정규화하여 저장
        raw_inventory_list = [
            "M7X0.75", "4-48 UNF-2A", "4-40 UNC-2A", "3-56 UNF-2A", "3/8-32 UNEF-2A", 
            "1/2-28 UNEF-2A", "1/4-36 UNS-2A", "10-32 UNF-2A", "10-48 UNS-2A", "10-36 UNS-2A", 
            "5/16-28 UN-2A", "5/16-32 UNEF-2A", "7/16-28 UNEF-2A", "0-80 UNF-2A", "1-72 UNF-2A", 
            "2-56 UNC-2A", "12-56 UNS-2A", "12-32 UNEF-2A", "12-40 UNS-2A", 
            "M2X0.4", "M3X0.35", "M2.6X0.45", "M10X0.75", "M14X1.0", 
            "2-56 UNC", "M5X0.5", "M3X0.5", "M4X0.7", "M20X1.0", "M29X1.5", 
            "7/16-28 UNEF", "5/8-24 UNEF", "1/4-36 UNS", "M5.5X0.5", "M8X0.75", 
            "M2.5X0.45", "M8X1.0", "9/32-40 UNS", "M6X0.5",
            
            # Plug Gauge 추가 목록
            "8-64 UNS-2B", "8-32 UNC-2B", "6-40 UNF-2B", "1/4-36 UNS-2B", "1/2-28 UNEF-2B", 
            "3/8-32 UNEF-2B", "10-32 UNF-2B", "12-56 UNS-2B", "10-36 UNS-2B", "5/16-18 UNC-2B", 
            "7/16-28 UNEF-2B", "5/16-32 UNEF-2B", "2-56 UNC-2B", "1-72 UNF-2B", "3-56 UNF-2B", 
            "0-80 UNF-2B", "9/16-24 UNEF-2B", "9/16-28 UN-2B", "7/16-32 UN-2B", "1-64 UNC-2B", 
            "00-90 UNS-2B", "0.1510-56 UNS-2B", "3/8-24 UNF-2B",
            "0-80 UNF", "10-36 UNS", "1-72 UNF", "6-40 UNF", "2-56 UNC", "M18X1.0", 
            "5/8-24 UNEF", "1/4-36 UNS", "7/16-28 UNEF", "9/32-40 UNS", "10-56 UNS",
            "4-48 UNF-2A", "8-56 UNS-3B", "10-48 UNS-2B", "1/4-28 UNF-2B", "9/32-40 UNS-2B",
            "1/2-28 UNEF-2B", "11/16-24 UNEF-2B", "4-40 UNC-2B", "6-32 UNC-2B", "5/16-18 UNC-2B",
            
            "4-40 UNC-2A", "6-32 UNC-2A", "6-40 UNF-2A", "8-56 UNS-3A", "1/4-20 UNC-2A",
            "1/4-28 UNF-2A", "9/32-40 UNS-2A", "3/8-32 UNEF-2A", "11/16-24 UNEF-2A"
        ]
        self.inventory = set()
        for transform_name in raw_inventory_list:
            self.inventory.add(self.normalize_screw_name(transform_name))
        
        self.setup_ui()
        
    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목 (드래그 가능 표시)
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(title_frame, text="나사 자동 입력 프로그램", 
                               font=("맑은 고딕", 16, "bold"))
        title_label.pack(side=tk.LEFT)
        
        # 드래그 가능 표시
        drag_hint_label = ttk.Label(title_frame, text="(드래그로 이동 가능)", 
                                   font=("맑은 고딕", 9), foreground="gray")
        drag_hint_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # 제목 바 드래그 바인딩
        title_frame.bind("<Button-1>", self.start_drag)
        title_frame.bind("<B1-Motion>", self.on_drag)
        title_label.bind("<Button-1>", self.start_drag)
        title_label.bind("<B1-Motion>", self.on_drag)
        drag_hint_label.bind("<Button-1>", self.start_drag)
        drag_hint_label.bind("<B1-Motion>", self.on_drag)
        
        # 나사 이름 입력 섹션
        input_frame = ttk.LabelFrame(main_frame, text="나사 이름 입력", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Entry와 Listbox를 담을 프레임
        entry_listbox_frame = ttk.Frame(input_frame)
        entry_listbox_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.screw_name_entry = ttk.Entry(entry_listbox_frame, font=("맑은 고딕", 12))
        self.screw_name_entry.pack(fill=tk.X)
        self.screw_name_entry.bind("<Return>", self.on_enter_pressed)
        self.screw_name_entry.bind("<KeyPress-Return>", self.on_enter_pressed)
        self.screw_name_entry.bind("<Button-1>", self.on_entry_click)
        # 나사 이름 입력 시 필터링 및 미리보기 업데이트
        self.screw_name_entry.bind("<KeyRelease>", self.on_entry_key_release_with_preview)
        # 포커스 설정
        self.screw_name_entry.focus_set()
        
        # 드롭다운 리스트박스
        listbox_frame = ttk.Frame(entry_listbox_frame)
        listbox_frame.pack(fill=tk.X)
        
        self.screw_listbox = tk.Listbox(listbox_frame, font=("맑은 고딕", 11), height=6)
        self.screw_listbox.pack(fill=tk.X)
        self.screw_listbox.bind("<Double-Button-1>", self.on_listbox_select)
        self.screw_listbox.bind("<Button-1>", self.on_listbox_select)
        self.screw_listbox.pack_forget()  # 초기에는 숨김
        
        self.listbox_visible = False
        
        # 변환 버튼
        button_frame = ttk.Frame(input_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.convert_button = ttk.Button(button_frame, text="변환", 
                                         command=self.copy_to_clipboard,
                                         width=20)
        self.convert_button.pack()
        
        # --------------------------------------------------------------------------
        # 옵션 선택 & 비고 (통합 프레임) - 회색 박스(LabelFrame) 제거
        # --------------------------------------------------------------------------
        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 1행: [고객사: SM/그외] [나사종류: 내경/외경] [게이지: 일반/도금전]
        row1_frame = ttk.Frame(options_frame)
        row1_frame.pack(fill=tk.X, pady=5)
        
        # 고객사 라디오버튼
        ttk.Label(row1_frame, text="고객사:", font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        self.sm_radio = ttk.Radiobutton(row1_frame, text="SM", variable=self.customer_var, value="SM", command=self.on_selection_changed)
        self.sm_radio.pack(side=tk.LEFT, padx=(0, 10))
        self.other_radio = ttk.Radiobutton(row1_frame, text="그외 고객사", variable=self.customer_var, value="OTHER", command=self.on_selection_changed)
        self.other_radio.pack(side=tk.LEFT, padx=(0, 20))
        
        # 나사 종류 라디오버튼
        ttk.Label(row1_frame, text="종류:", font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        self.inner_radio = ttk.Radiobutton(row1_frame, text="내경나사", variable=self.screw_type_var, value="INNER", command=self.on_selection_changed)
        self.inner_radio.pack(side=tk.LEFT, padx=(0, 10))
        self.outer_radio = ttk.Radiobutton(row1_frame, text="외경나사", variable=self.screw_type_var, value="OUTER", command=self.on_selection_changed)
        self.outer_radio.pack(side=tk.LEFT, padx=(0, 20))
        
        # 게이지 라디오버튼
        ttk.Label(row1_frame, text="게이지:", font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        self.normal_gauge_radio = ttk.Radiobutton(row1_frame, text="일반", variable=self.gauge_var, value="NORMAL", command=self.on_selection_changed)
        self.normal_gauge_radio.pack(side=tk.LEFT, padx=(0, 10))
        self.before_plating_radio = ttk.Radiobutton(row1_frame, text="도금 전", variable=self.gauge_var, value="BEFORE_PLATING", command=self.on_selection_changed)
        self.before_plating_radio.pack(side=tk.LEFT, padx=(0, 10))

        # 2행: [비고 입력창] [보유 현황 표시]
        row2_frame = ttk.Frame(options_frame)
        row2_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(row2_frame, text="보유 나사게이지 확인 검색 창:", font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        self.memo_entry = ttk.Entry(row2_frame, width=30, font=("맑은 고딕", 10))
        self.memo_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 예시 텍스트 추가
        ttk.Label(row2_frame, text="(예 : 1/4-36 UNS-2A)", font=("맑은 고딕", 9), foreground="gray").pack(side=tk.LEFT, padx=(0, 20))
        
        # 엔터 키 바인딩 (팝업 확인)
        self.memo_entry.bind("<Return>", self.check_inventory_popup)
        self.memo_entry.bind("<KeyPress-Return>", self.check_inventory_popup)
        
        self.inventory_status_label = ttk.Label(row2_frame, text="대기 중", font=("맑은 고딕", 10))
        self.inventory_status_label.pack(side=tk.LEFT)
        
        # 도금 값 입력 섹션
        plating_frame = ttk.LabelFrame(main_frame, text="도금 값 입력", padding="10")
        plating_frame.pack(fill=tk.X, pady=(0, 20))
        
        plating_inner_frame = ttk.Frame(plating_frame)
        plating_inner_frame.pack(fill=tk.X)
        
        ttk.Label(plating_inner_frame, text="도금 값:", font=("맑은 고딕", 10)).pack(side=tk.LEFT, padx=(0, 10))
        
        self.plating_entry = ttk.Entry(plating_inner_frame, textvariable=self.plating_value_var, 
                                       font=("맑은 고딕", 12), width=15)
        self.plating_entry.pack(side=tk.LEFT, padx=(0, 5))
        self.plating_entry.bind("<Return>", self.on_enter_pressed)
        self.plating_entry.bind("<KeyPress-Return>", self.on_enter_pressed)
        
        ttk.Label(plating_inner_frame, text="㎛", font=("맑은 고딕", 10)).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(plating_inner_frame, text="(그외 고객사 + 도금 전 게이지에서 사용)", 
                 font=("맑은 고딕", 9), foreground="gray").pack(side=tk.LEFT)
        
        # 상태 표시 레이블
        self.status_label = ttk.Label(main_frame, text="상태: 대기 중", 
                                     font=("맑은 고딕", 9), foreground="gray")
        self.status_label.pack(pady=(10, 5))
        
        # 나사 출력 결과 표시 섹션
        output_frame = ttk.LabelFrame(main_frame, text="나사 출력 결과", padding="10")
        output_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 출력 결과를 표시할 Text 위젯
        self.output_text = tk.Text(output_frame, font=("맑은 고딕", 11), 
                                   wrap=tk.WORD, height=8, state=tk.DISABLED)
        self.output_text.pack(fill=tk.BOTH, expand=True)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(output_frame, orient=tk.VERTICAL, command=self.output_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.output_text.config(yscrollcommand=scrollbar.set)
    
    def init_screw_data(self):
        """SM/그 외 고객사, 외경나사/내경나사, 일반게이지/도금 전 게이지 데이터 초기화"""
        
        # 외경나사 데이터 - 일반게이지 (OUTER 타입)
        outer_screw_list_normal = [
            # 나사이름, Major Min, Major Max, Pitch Min, Pitch Max, Minor Min (없으면 None), Minor Max
            ("1/4-36 UNS-2A", "6.188", "6.327", "5.792", "5.869", None, "5.461"),
            ("8-64 UNS-2A", "4.054", "4.417", "3.833", "3.891", "3.577", "3.718"),
            ("3/8-24 UNF-2A", "9.315", "9.497", "8.713", "8.808", None, "8.237"),
            ("10-32 UNF-2A", "4.644", "4.792", "4.199", "4.267", None, "3.820"),
            ("10-36 UNS-2A", "4.664", "4.803", "4.270", "4.345", None, "3.962"),
            ("12-56 UNS-2A", "5.365", "5.468", "5.111", "5.173", None, "4.930"),
            ("5-44 UNF-2A", "3.036", "3.157", "2.718", "2.781", None, "2.468"),
            ("10-56 UNS-2A", "4.705", "4.808", "4.451", "4.513", None, "4.269"),
            ("1-72 UNF-2A", "1.750", "1.839", "1.562", "1.611", None, "1.419"),
            ("4-48 UNF-2A", "2.713", "2.627", "2.424", "2.484", None, "2.176"),
            ("10-48 UNS-2A", "4.692", "4.805", "4.397", "4.462", None, "4.155"),
            ("10-32 UNF-2A", "4.651", "4.803", "4.212", "4.287", None, "3.830"),
            ("2-56 UNC-2A", "2.055", "2.153", "1.801", "1.844", None, "1.597"),
            ("8-36 UNF-2A", "4.005", "4.145", "3.616", "3.688", None, "3.304"),
            ("7/16-28 UNEF-2A", "10.920", "11.084", "10.404", "10.495", None, "9.748"),
            ("1/2-28 UNEF-2A", "12.507", "12.672", "11.989", "12.082", None, "11.559"),
            ("3-56 UNF-2A", "2.392", "2.497", "2.146", "2.203", None, "1.958"),
            ("0-80 UNF-2A", "1.430", "1.511", "1.260", "1.305", None, "1.133"),
            ("1/4-18 UNEF-2A", "31.490", "31.711", "30.670", "30.794", None, "30.032"),
            ("5/16-32 UNEF-2A", "7.760", "7.912", "7.316", "7.396", None, "6.967"),
            ("5/8-24 UNEF-2A", "15.661", "15.845", "15.055", "15.156", None, "14.584"),
            ("5/16-36 UNS-2A", "7.774", "7.914", "7.378", "7.457", None, "7.073"),
            ("9/16-24 UNEF-2A", "14.074", "14.257", "13.469", "13.568", None, "12.997"),
            ("12-48 UNS-2A", "5.351", "5.466", "5.057", "5.123", None, "4.836"),
            ("4-40 UNC-2A", "2.695", "2.845", "2.433", "2.517", "2.156", "2.385"),
            ("6-32 UNC-2A", "3.332", "3.505", "2.990", "3.084", "2.642", "2.896"), 
            ("6-40 UNF-2A", "3.356", "3.505", "3.094", "3.180", "2.819", "3.023"),
            ("8-56 UNS-3A", "4.000", "4.166", "3.871", "3.932", "3.683", "3.785"),
            ("1/4-20 UNC-2A", "6.160", "6.350", "5.525", "5.649", "4.978", "5.258"),
            ("1/4-28 UNF-2A", "6.185", "6.350", "5.761", "5.870", "5.359", "5.588"),
            ("9/32-40 UNS-2A", "6.980", "7.142", "6.731", "6.828", "6.452", "6.604"),
            ("3/8-32 UNEF-2A", "9.363", "9.525", "9.009", "9.121", "8.661", "8.865"),
            ("11/16-24 UNEF-2A", "17.290", "17.463", "16.774", "16.906", "16.307", "16.561"),
        ]
        
        # 그 외 고객사 - 일반게이지 - 외경나사 데이터 (OUTER 타입)
        # 올바른 데이터 구조: Major Min, Major Max, Pitch Min, Pitch Max, Minor Max (Minor Min 없음)
        other_outer_screw_list_normal = [
            # 나사이름, Major Min, Major Max, Pitch Min, Pitch Max, Minor Min (없으면 None), Minor Max
            ("0-80 UNF-2A", "1.431", "1.511", "1.260", "1.305", None, "1.132"),
            ("1/4-36 UNS-2A", "6.188", "6.327", "5.792", "5.869", None, "5.488"),
            ("5/16-32 UNEF-2A", "7.760", "7.912", "7.316", "7.396", None, "6.697"),
            ("8-36 UNF-2A", "4.026", "4.165", "3.656", "3.708", None, "3.324"),
            ("4-40 UNC-2A", "2.695", "2.824", "2.350", "2.413", None, "2.067"),
            ("7/16-28 UNEF-2A", "10.920", "11.084", "10.404", "10.495", None, "9.972"),
            ("6-80 UNS-2A", "3.405", "3.485", "3.232", "3.278", None, "3.095"),
            ("12-36 UNS-2A", "5.324", "5.463", "4.931", "5.006", None, "4.625"),
            ("5/8-24 UNEF-2A", "15.662", "15.844", "15.055", "15.156", None, "14.546"),
            ("6-40 UNF-2A", "3.356", "3.484", "3.008", "3.073", None, "2.727"),
            ("10-36 UNS-2A", "4.664", "4.803", "4.270", "4.345", None, "3.962"),
            ("10-48 UNS-2A", "4.692", "4.805", "4.397", "4.462", None, "4.175"),
            ("5/16-40 UNS-2A", "7.786", "7.914", "7.430", "7.503", None, "7.157"),
            ("10-32 UNF-2A", "4.651", "4.803", "4.212", "4.287", None, "3.858"),
            ("0.8mm UNM-2A", "0.771", "0.800", "0.647", "0.670", None, "0.571"),
            ("8-64 UNS-2A", "4.054", "4.147", "3.833", "3.891", "3.577", "3.718"),
            ("5-40 UNC-2A", "3.156", "3.175", "2.744", "2.764", None, "2.375"),
            ("6-32 UNC-2A", "3.364", "3.505", "2.899", "2.969", None, "2.748"),
            ("8-32 UNC-2A", "4.024", "4.166", "3.554", "3.627", None, "3.409"),
            ("10-24 UNC-2A", "4.659", "4.826", "4.027", "4.112", None, "3.896"),
            ("12-24 UNC-2A", "5.457", "5.486", "4.766", "4.798", None, "4.074"),
            ("1/4-20 UNC-2A", "6.168", "6.350", "5.404", "5.497", None, "5.034"),
            ("1/4-28 UNF-2A", "6.186", "6.350", "5.655", "5.735", None, "5.352"),
            ("5/16-18 UNC-2A", "7.908", "7.938", "6.988", "7.021", None, "6.259"),
            ("5/16-24 UNF-2A", "7.911", "7.938", "7.220", "7.249", None, "6.695"),
            ("3/8-16 UNC-2A", "9.493", "9.525", "8.459", "8.494", None, "7.633"),
            ("4-48 UNF-2A", "2.714", "2.827", "2.424", "2.484", None, "2.195"),
            ("3-56 UNF-2A", "2.394", "2.497", "2.147", "2.202", None, "1.961"),
            ("3/8-32 UNEF-2A", "9.363", "9.525", "8.899", "8.984", None, "8.494"),
            ("1/2-28 UNEF-2A", "12.538", "12.700", "11.990", "12.083", None, "11.521"),
            ("1-72 UNF-2A", "1.766", "1.839", "1.576", "1.610", None, "1.422"),
            ("2-56 UNC-2A", "2.066", "2.169", "1.822", "1.864", None, "1.628"),
            ("12-56 UNS-2A", "5.345", "5.469", "5.101", "5.156", None, "4.935"),
            ("12-32 UNEF-2A", "5.325", "5.486", "4.870", "4.948", None, "4.458"),
            ("12-40 UNS-2A", "5.350", "5.486", "4.974", "5.050", None, "4.661"),
            ("9/32-40 UNS-2A", "7.006", "7.144", "6.630", "6.708", None, "6.317"),
            ("5/16-28 UN-2A", "7.748", "7.912", "7.237", "7.323", None, "6.761"),
            ("12-24 UNC-2A", "5.335", "5.486", "4.700", "4.765", None, "4.097"),
            ("M7X0.75", "6.838", "6.978", "6.391", "6.491", None, "6.166"),
            ("M2X0.4", "1.886", "1.981", "1.654", "1.721", None, "1.509"),
            ("M3X0.35", "2.896", "2.981", "2.687", "2.754", None, "2.571"),
            ("M2.6X0.45", "2.480", "2.580", "2.217", "2.288", None, "2.113"),
            ("M10X0.75", "9.838", "9.978", "9.391", "9.491", None, "9.166"),
            ("M14X1.0", "13.790", "13.974", "13.210", "13.320", None, "12.891"),
            ("M5X0.5", "4.826", "4.976", "4.596", "4.670", None, "4.456"),
            ("M3X0.5", "2.874", "2.980", "2.580", "2.655", None, "2.459"),
            ("M4X0.7", "3.838", "3.978", "3.433", "3.523", None, "3.242"),
            ("M20X1.0", "19.794", "19.974", "19.206", "19.324", None, "18.891"),
            ("M29X1.5", "28.732", "28.952", "27.872", "28.026", None, "27.327"),
            ("M5.5X0.5", "5.374", "5.490", "5.060", "5.175", None, "4.939"),
            ("M8X0.75", "7.838", "7.978", "7.410", "7.510", None, "7.166"),
            ("M2.5X0.45", "2.380", "2.480", "2.117", "2.188", None, "2.013"),
            ("M8X1.0", "7.794", "7.974", "7.234", "7.324", None, "6.891"),
            ("M6X0.5", "5.874", "5.978", "5.580", "5.655", None, "5.459"),
        ]
        
        # 외경나사 데이터 - 도금 전 게이지 (OUTER 타입)
        outer_screw_list_before_plating = [
            # 나사이름, Major Min, Major Max, Pitch Min, Pitch Max, Minor Min (없으면 None), Minor Max
            ("1/4-36 UNS-2A", "6.185", "6.314", "5.787", "5.844", None, "5.448"),
            ("10-56 UNS-2A", "4.703", "4.753", "4.449", "4.479", None, "4.266"),
            ("3/8-32 UNEF-2A", "9.332", "9.476", "8.868", "8.938", None, "8.503"),
            ("5/8-24 UNEF-2A", "15.647", "15.821", "15.025", "15.110", None, "14.523"),
            ("10-32 UNF-2A", "4.644", "4.792", "4.199", "4.267", None, "3.820"),
            ("000-120 UNS-2A", "0.826", "0.840", "0.674", "0.703", None, None),
            ("4-40 UNC-2A", "2.685", "2.809", "2.330", "2.382", None, "2.029"),
            ("6-40 UNF-2A", "3.348", "3.474", "2.995", "3.053", None, "2.694"),
            ("7/16-28 UNEF-2A", "10.905", "11.061", "10.374", "10.449", None, "9.949"),
            ("1/2-28 UNEF-2A", "12.492", "12.649", "11.959", "12.037", None, "11.536"),
            ("0-80 UNF-2A", "1.430", "1.511", "1.260", "1.305", None, "1.133"),
            ("2-56 UNC-2A", "2.055", "2.153", "1.801", "1.844", None, "1.597"),
            ("5/16-32 UNEF-2A", "7.753", "7.901", "7.303", "7.376", None, "6.929"),
            ("8-64 UNS-2A", "4.044", "4.142", "3.813", "3.881", "3.666", "3.713"),
            ("12-36 UNS-2A", "5.317", "5.448", "4.923", "4.991", None, "4.617"),
            ("10-36 UNS-2A", "4.664", "4.803", "4.270", "4.345", None, "3.962"),
            ("11/16-24 UNEF-2A", "17.234", "17.409", "16.612", "16.697", None, "16.111"),
            ("9/16-24 UNEF-2A", "14.074", "14.257", "13.469", "13.568", None, "12.997"),
            ("5/16-18 UNF-2A", "7.907", "7.686", "6.990", "6.888", None, "6.225"),
            ("2-56 UNC-2A", "2.055", "2.154", "1.801", "1.844", None, "1.598"),
            ("4-40 UNC-2A", "2.685", "2.809", "2.330", "2.383", None, "2.029"),
            ("4-48 UNF-2A", "2.703", "2.812", "2.403", "2.454", None, "2.162"),
            ("6-32 UNC-2A", "3.325", "3.475", "2.885", "2.949", None, "2.502"),
            ("6-40 UNF-2A", "3.348", "3.475", "2.995", "3.053", None, "2.695"),
            ("8-56 UNS-3A", "4.034", "4.183", "3.802", "3.840", None, "3.627"),
            ("10-48 UNS-2A", "4.661", "4.796", "4.384", "4.442", None, "4.145"),
            ("1/4-20 UNC-2A", "6.002", "6.307", "5.334", "5.466", None, "4.750"),
            ("1/4-28 UNF-2A", "6.066", "6.309", "5.588", "5.705", None, "5.197"),
            ("1/4-36 UNS-2A", "6.185", "6.314", "5.786", "5.845", None, "5.448"),
            ("9/32-40 UNS-2A", "6.988", "7.107", "6.627", "6.683", None, "6.330"),
            ("5/16-32 UNEF-2A", "7.752", "7.902", "7.303", "7.376", None, "6.929"),
            ("3/8-32 UNEF-2A", "9.332", "9.477", "8.867", "8.938", None, "8.504"),
            ("7/16-28 UNEF-2A", "10.904", "11.062", "10.373", "10.449", None, "9.949"),
            ("1/2-28 UNEF-2A", "12.492", "12.649", "11.958", "12.035", None, "11.537"),
            ("5/8-24 UNEF-2A", "15.646", "15.822", "15.024", "15.110", None, "14.524"),
            ("11/16-24 UNEF-2A", "17.234", "17.409", "16.612", "16.698", None, "16.111"),
        ]
        
        # 내경나사 데이터 - 일반게이지 (INNER 타입)
        inner_screw_list_normal = [
            # 나사이름, Minor Min, Minor Max, Pitch Min, Pitch Max, Major Min
            ("0-80 UNF-2B", "1.182", "1.305", "1.305", "1.376", "1.524"),
            ("1/4-36 UNS-2B", "5.601", "5.742", "5.742", "5.999", "6.363"),
            ("12-56 UNS-2B", "5.004", "5.105", "5.105", "5.273", "5.487"),
            ("3/8-24 UNF-2B", "8.383", "8.636", "8.636", "8.961", "9.526"),
            ("5/16-32 UNEF-2B", "7.087", "7.264", "7.264", "7.528", "7.938"),
            ("5-44 UNF-2B", "2.550", "2.740", "2.740", "2.880", "3.175"),
            ("6-40 UNF-2B", "2.820", "3.022", "3.022", "3.180", "3.506"),
            ("1-72 UNF-2B", "1.473", "1.612", "1.612", "1.689", "1.854"),
            ("2-56 UNC-2B", "1.695", "1.871", "1.871", "1.960", "2.185"),
            ("10-56 UNS-2B", "4.344", "4.445", "4.445", "4.612", "4.826"),
            ("10-36 UNS-2B", "4.064", "4.216", "4.216", "4.467", "4.826"),
            ("9/16-24 UNEF-2B", "13.132", "13.385", "13.385", "13.728", "14.288"),
            ("1/4-20 UNC-2B", "4.979", "5.257", "5.257", "5.648", "6.350"),
            ("5/8-24 UNEF-2B", "14.755", "15.001", "15.001", "15.349", "15.897"),
            ("3/8-32 UNEF-2B", "8.661", "8.865", "8.865", "9.122", "9.525"),
            ("7/16-28 UNEF-2B", "10.135", "10.337", "10.337", "10.640", "11.113"),
            ("8-64 UNF-2B", "3.734", "3.835", "3.835", "3.985", "4.165"),
            ("4-40 UNC-2B", "2.156", "2.385", "2.433", "2.517", "2.845"),
            ("4-48 UNF-2B", "2.271", "2.459", "2.502", "2.581", "2.845"),
            ("6-32 UNC-2B", "2.642", "2.896", "2.990", "3.084", "3.505"),
            ("8-56 UNS-3B", "3.683", "3.785", "3.871", "3.932", "4.166"),
            ("10-32 UNF-2B", "3.962", "4.166", "4.310", "4.409", "4.826"),
            ("10-48 UNS-2B", "4.242", "4.369", "4.483", "4.569", "4.826"),
            ("1/4-28 UNF-2B", "5.359", "5.588", "5.761", "5.870", "6.350"),
            ("9/32-40 UNS-2B", "6.452", "6.604", "6.731", "6.828", "7.142"),
            ("1/2-28 UNEF-2B", "11.709", "11.938", "12.111", "12.233", "12.700"),
            ("11/16-24 UNEF-2B", "16.307", "16.561", "16.774", "16.906", "17.463"),
        ]
        
        # 그 외 고객사 - 일반게이지 - 내경나사 데이터 (INNER 타입)
        other_inner_screw_list_normal = [
            # 나사이름, Minor Min, Minor Max, Pitch Min, Pitch Max, Major Min
            ("0-80 UNF-2B", "1.181", "1.306", "1.318", "1.377", "1.524"),
            ("8-36 UNF-2B", "3.404", "3.596", "3.709", "3.776", "4.166"),
            ("5/16-32 UNEF-2B", "7.087", "7.264", "7.422", "7.528", "7.938"),
            ("1/4-36 UNS-2B", "5.588", "5.740", "5.893", "5.994", "6.350"),
            ("7/16-28 UNEF-2B", "10.135", "10.337", "10.524", "10.640", "11.113"),
            ("12-36 UNS-2B", "4.725", "4.876", "5.030", "5.128", "5.487"),
            ("6-40 UNF-2B", "2.819", "3.022", "3.096", "3.180", "3.505"),
            ("5/8-24 UNEF-2B", "14.733", "14.986", "15.187", "15.318", "15.876"),
            ("1-64 UNC-2B", "1.425", "1.582", "1.598", "1.663", "1.855"),
            ("10-36 UNS-2B", "4.064", "4.216", "4.369", "4.487", "4.826"),
            ("10-48 UNS-2B", "4.242", "4.368", "4.484", "4.569", "4.826"),
            ("4-40 UNC-2B", "2.157", "2.385", "2.434", "2.517", "2.845"),
            ("10-32 UNF-2B", "3.963", "4.165", "4.311", "4.409", "4.827"),
            ("0.8mm UNM-2B", "0.609", "0.684", "0.671", "0.694", "0.801"),
            ("5/16-28 UNF-2B", "6.782", "7.035", "7.250", "7.371", "7.938"),
            ("5-40 UNC-2B", "2.488", "2.697", "2.765", "2.847", "3.176"),
            ("6-32 UNC-2B", "2.643", "2.896", "2.991", "3.084", "3.506"),
            ("8-32 UNC-2B", "3.303", "3.531", "3.651", "3.746", "4.167"),
            ("10-24 UNC-2B", "3.684", "3.962", "4.139", "4.249", "4.827"),
            ("12-24 UNC-2B", "4.344", "4.597", "4.799", "4.910", "5.487"),
            ("1/4-20 UNC-2B", "4.979", "5.258", "5.525", "5.649", "6.351"),
            ("1/4-28 UNF-2B", "5.360", "5.588", "5.762", "5.870", "6.351"),
            ("5/16-18 UNC-2B", "6.402", "6.731", "7.022", "7.155", "7.939"),
            ("5/16-24 UNF-2B", "6.783", "7.036", "7.250", "7.371", "7.939"),
            ("3/8-16 UNC-2B", "7.799", "8.153", "8.495", "8.639", "9.526"),
            ("8-64 UNS-2B", "3.736", "3.835", "3.835", "3.891", "4.166"),
            ("7/16-32 UN-2B", "10.340", "10.516", "10.605", "10.693", "11.113"),
            ("9/16-28 UN-2B", "13.312", "13.513", "13.691", "13.785", "14.288"),
            ("0.1510-56 UNS-2B", "3.279", "3.391", "3.531", "3.597", "3.835"),
            ("3/8-24 UNF-2B", "8.385", "8.636", "8.636", "8.961", "9.525"),
            
            ("M7X0.75", "6.188", "6.378", "6.513", "6.645", "7.000"),
            ("M20X1.0", "18.917", "19.153", "19.350", "19.510", "20.000"),
            ("M2X0.4", "1.567", "1.679", "1.740", "1.830", "2.000"),
            ("M2.6X0.45", "2.113", "2.238", "2.308", "2.393", "2.600"),
            ("M3X0.35", "2.621", "2.766", "2.773", "2.900", "3.000"),
            ("M10X0.75", "9.188", "9.378", "9.513", "9.645", "10.000"),
            ("M14X1.0", "12.917", "13.153", "13.350", "13.510", "14.000"),
            ("M4X0.7", "3.242", "3.422", "3.545", "3.663", "4.000"),
            ("M5X0.5", "4.459", "4.599", "4.675", "4.770", "5.000"),
            ("M3X0.5", "2.459", "2.599", "2.675", "2.775", "3.000"),
            ("M18X1.0", "16.917", "17.153", "17.350", "17.510", "18.000"),
            ("M29X1.5", "27.376", "27.676", "28.026", "28.216", "29.000"),
            ("M5.5X0.5", "4.959", "5.099", "5.175", "5.295", "5.500"),
            ("M8X0.75", "7.188", "7.378", "7.513", "7.645", "8.000"),
            ("M8X1.0", "6.917", "7.153", "7.350", "7.500", "8.000"),
            ("M2.5X0.45", "2.013", "2.138", "2.208", "2.303", "2.500"),
            ("M6X0.5", "5.459", "5.599", "5.675", "5.787", "6.000"),
        ]
        
        # 내경나사 데이터 - 도금 전 게이지 (INNER 타입)
        inner_screw_list_before_plating = [
            # 나사이름, Minor Min, Minor Max, Pitch Min, Pitch Max, Major Min
            ("1/4-36 UNS-2B", "5.601", "5.742", "5.919", "5.999", "6.363"),
            ("10-36 UNS-2B", "4.064", "4.217", "4.368", "4.445", "4.826"),
            ("4-40 UNC-2B", "2.172", "2.395", "2.464", "2.537", "2.861"),
            ("2-56 UNC-2B", "1.710", "1.882", "1.921", "1.981", "2.200"),
            ("7/16-28 UNEF-2B", "10.158", "10.172", "10.569", "10.670", "11.136"),
            ("5/16-32 UNEF-2B", "7.097", "7.272", "7.443", "7.541", "7.948"),
            ("6-40 UNF-2B", "2.830", "3.030", "3.115", "3.192", "3.516"),
            ("1/2-28 UNEF-2B", "11.733", "11.953", "12.157", "12.263", "12.723"),
            ("11/16-24 UNEF-2B", "16.330", "16.576", "16.820", "16.885", "17.488"),
            ("6-32 UNC-2B", "2.652", "2.903", "3.010", "3.096", "3.516"),
            ("5/8-24 UNEF-2B", "14.732", "14.986", "15.187", "15.318", "15.875"),
            ("1-72 UNF-2B", "1.473", "1.612", "1.625", "1.689", "1.854"),
            ("10-32 UNF-2B", "3.972", "4.173", "4.330", "4.422", "4.836"),
            ("12-56 UNS-2B", "5.003", "5.105", "5.191", "5.273", "5.486"),
            ("3/8-32 UNEF-2B", "8.685", "8.879", "9.506", "9.151", "9.547"),
            ("4-48 UNF-2A", "2.286", "2.477", "2.532", "2.601", "2.860"),
            ("8-56 UNS-3B", "3.698", "3.795", "3.901", "3.952", "4.181"),
            ("10-48 UNS-2B", "4.252", "4.376", "4.503", "4.582", "4.836"),
            ("1/4-28 UNF-2B", "5.375", "5.598", "5.791", "5.890", "6.365"),
            ("9/32-40 UNS-2B", "6.464", "6.607", "6.756", "6.833", "7.155"),
            ("1/4-36 UNS-2B", "5.588", "5.740", "5.893", "5.994", "6.350"),
            ("5/16-18 UNC-2B", "6.402", "6.731", "7.022", "7.155", "7.939"),
            ("10-36 UNS-2B", "4.064", "4.216", "4.369", "4.487", "4.826"),
            ("7/16-28 UNEF-2B", "10.135", "10.337", "10.524", "10.640", "11.113"),
            ("0-80 UNF-2B", "1.181", "1.306", "1.318", "1.377", "1.524"),
        ]
        
        # 외경나사 데이터 처리 - 일반게이지
        for name, major_min, major_max, pitch_min, pitch_max, minor_min, minor_max in outer_screw_list_normal:
            data = {'type': 'OUTER', 'gauge': 'NORMAL'}
            
            # Major 데이터
            if major_min and major_max:
                data['major'] = {'min': major_min, 'max': major_max}
            elif major_min:
                data['major'] = {'min': major_min}
            elif major_max:
                data['major'] = {'max': major_max}
            
            # Pitch 데이터
            if pitch_min and pitch_max:
                data['pitch'] = {'min': pitch_min, 'max': pitch_max}
            elif pitch_min:
                data['pitch'] = {'min': pitch_min}
            elif pitch_max:
                data['pitch'] = {'max': pitch_max}
            
            # Minor 데이터
            if minor_min and minor_max:
                data['minor'] = {'min': minor_min, 'max': minor_max}
            elif minor_min:
                data['minor'] = {'min': minor_min}
            elif minor_max:
                data['minor'] = {'max': minor_max}
            
            # 같은 이름이 여러 게이지에 있을 수 있으므로 리스트로 저장
            if name not in self.screw_data:
                self.screw_data[name] = []
            self.screw_data[name].append(data)
        
        # 외경나사 데이터 처리 - 도금 전 게이지
        for name, major_min, major_max, pitch_min, pitch_max, minor_min, minor_max in outer_screw_list_before_plating:
            data = {'type': 'OUTER', 'gauge': 'BEFORE_PLATING'}
            
            # Major 데이터
            if major_min and major_max:
                data['major'] = {'min': major_min, 'max': major_max}
            elif major_min:
                data['major'] = {'min': major_min}
            elif major_max:
                data['major'] = {'max': major_max}
            
            # Pitch 데이터
            if pitch_min and pitch_max:
                data['pitch'] = {'min': pitch_min, 'max': pitch_max}
            elif pitch_min:
                data['pitch'] = {'min': pitch_min}
            elif pitch_max:
                data['pitch'] = {'max': pitch_max}
            
            # Minor 데이터
            if minor_min and minor_max:
                data['minor'] = {'min': minor_min, 'max': minor_max}
            elif minor_min:
                data['minor'] = {'min': minor_min}
            elif minor_max:
                data['minor'] = {'max': minor_max}
            
            # 같은 이름이 여러 게이지에 있을 수 있으므로 리스트로 저장
            if name not in self.screw_data:
                self.screw_data[name] = []
            self.screw_data[name].append(data)
        
        # 그 외 고객사 - 일반게이지 - 외경나사 데이터 처리
        # 올바른 데이터 구조: Major Min, Major Max, Pitch Min, Pitch Max, Minor Min (없으면 None), Minor Max
        for name, major_min, major_max, pitch_min, pitch_max, minor_min, minor_max in other_outer_screw_list_normal:
            data = {'type': 'OUTER', 'gauge': 'NORMAL', 'customer': 'OTHER'}
            
            # Major 데이터
            if major_min and major_max:
                data['major'] = {'min': major_min, 'max': major_max}
            elif major_min:
                data['major'] = {'min': major_min}
            elif major_max:
                data['major'] = {'max': major_max}
            
            # Pitch 데이터
            if pitch_min and pitch_max:
                data['pitch'] = {'min': pitch_min, 'max': pitch_max}
            elif pitch_min:
                data['pitch'] = {'min': pitch_min}
            elif pitch_max:
                data['pitch'] = {'max': pitch_max}
            
            # Minor 데이터
            if minor_min and minor_max:
                data['minor'] = {'min': minor_min, 'max': minor_max}
            elif minor_min:
                data['minor'] = {'min': minor_min}
            elif minor_max:
                data['minor'] = {'max': minor_max}
            
            # 같은 이름이 여러 게이지/고객사에 있을 수 있으므로 리스트로 저장
            if name not in self.screw_data:
                self.screw_data[name] = []
            self.screw_data[name].append(data)
        
        # 내경나사 데이터 처리 - 일반게이지
        for name, minor_min, minor_max, pitch_min, pitch_max, major_min in inner_screw_list_normal:
            data = {'type': 'INNER', 'gauge': 'NORMAL'}
            
            # Minor 데이터
            if minor_min and minor_max:
                data['minor'] = {'min': minor_min, 'max': minor_max}
            elif minor_min:
                data['minor'] = {'min': minor_min}
            elif minor_max:
                data['minor'] = {'max': minor_max}
            
            # Pitch 데이터
            if pitch_min and pitch_max:
                data['pitch'] = {'min': pitch_min, 'max': pitch_max}
            elif pitch_min:
                data['pitch'] = {'min': pitch_min}
            elif pitch_max:
                data['pitch'] = {'max': pitch_max}
            
            # Major 데이터 (내경나사는 Min만 있음)
            if major_min:
                data['major'] = {'min': major_min}
            
            # 같은 이름이 여러 게이지에 있을 수 있으므로 리스트로 저장
            if name not in self.screw_data:
                self.screw_data[name] = []
            self.screw_data[name].append(data)
        
        # 그 외 고객사 - 일반게이지 - 내경나사 데이터 처리
        for name, minor_min, minor_max, pitch_min, pitch_max, major_min in other_inner_screw_list_normal:
            data = {'type': 'INNER', 'gauge': 'NORMAL', 'customer': 'OTHER'}
            
            # Minor 데이터
            if minor_min and minor_max:
                data['minor'] = {'min': minor_min, 'max': minor_max}
            elif minor_min:
                data['minor'] = {'min': minor_min}
            elif minor_max:
                data['minor'] = {'max': minor_max}
            
            # Pitch 데이터
            if pitch_min and pitch_max:
                data['pitch'] = {'min': pitch_min, 'max': pitch_max}
            elif pitch_min:
                data['pitch'] = {'min': pitch_min}
            elif pitch_max:
                data['pitch'] = {'max': pitch_max}
            
            # Major 데이터 (내경나사는 Min만 있음)
            if major_min:
                data['major'] = {'min': major_min}
            
            # 같은 이름이 여러 게이지/고객사에 있을 수 있으므로 리스트로 저장
            if name not in self.screw_data:
                self.screw_data[name] = []
            self.screw_data[name].append(data)
        
        # 내경나사 데이터 처리 - 도금 전 게이지
        for name, minor_min, minor_max, pitch_min, pitch_max, major_min in inner_screw_list_before_plating:
            data = {'type': 'INNER', 'gauge': 'BEFORE_PLATING'}
            
            # Minor 데이터
            if minor_min and minor_max:
                data['minor'] = {'min': minor_min, 'max': minor_max}
            elif minor_min:
                data['minor'] = {'min': minor_min}
            elif minor_max:
                data['minor'] = {'max': minor_max}
            
            # Pitch 데이터
            if pitch_min and pitch_max:
                data['pitch'] = {'min': pitch_min, 'max': pitch_max}
            elif pitch_min:
                data['pitch'] = {'min': pitch_min}
            elif pitch_max:
                data['pitch'] = {'max': pitch_max}
            
            # Major 데이터 (내경나사는 Min만 있음)
            if major_min:
                data['major'] = {'min': major_min}
            
            # 같은 이름이 여러 게이지에 있을 수 있으므로 리스트로 저장
            if name not in self.screw_data:
                self.screw_data[name] = []
            self.screw_data[name].append(data)
    
    def normalize_screw_name(self, name):
        """나사 이름 정규화 (-2A, -2B 제거, # 제거)"""
        if not name:
            return ""
        normalized = str(name).strip()
        # # 문자 제거
        normalized = normalized.replace('#', '')
        # -2A, -2B, -2a, -2b 제거 (대소문자 구분 없이)
        normalized = re.sub(r'-2[ABab]', '', normalized, flags=re.IGNORECASE)
        return normalized.strip()
    
    def get_filtered_screw_list(self):
        """선택된 조건에 맞는 나사 목록 반환"""
        customer = self.customer_var.get()
        screw_type = self.screw_type_var.get()
        gauge = self.gauge_var.get()
        
        # 조건이 모두 선택되지 않았으면 빈 리스트 반환
        if not customer or not screw_type or not gauge:
            return []
        
        # 고객사/게이지 조합 허용 확인
        allowed = False
        if customer == "SM" and gauge in ("NORMAL", "BEFORE_PLATING"):
            allowed = True
        elif customer == "OTHER" and gauge in ("NORMAL", "BEFORE_PLATING"):
            allowed = True
        
        if not allowed:
            return []
        
        # 그 외 고객사 + 도금 전 게이지인 경우, 일반게이지 데이터를 표시
        target_gauge = gauge
        if customer == "OTHER" and gauge == "BEFORE_PLATING":
            target_gauge = "NORMAL"
        
        # 조건에 맞는 나사 목록 수집
        filtered_screws = []
        for db_name in self.screw_data.keys():
            data_list = self.screw_data[db_name]
            for data in data_list:
                if data.get('type') == screw_type and data.get('gauge') == target_gauge:
                    data_customer = data.get('customer', 'SM')
                    if customer == data_customer:
                        # 중복 제거
                        if db_name not in filtered_screws:
                            filtered_screws.append(db_name)
                        break
        
        # 정렬
        filtered_screws.sort()
        return filtered_screws
    
    def update_screw_listbox(self):
        """드롭다운 리스트박스 업데이트"""
        filtered_screws = self.get_filtered_screw_list()
        
        # 리스트박스 내용 지우기
        self.screw_listbox.delete(0, tk.END)
        
        if filtered_screws:
            for screw_name in filtered_screws:
                self.screw_listbox.insert(tk.END, screw_name)
    
    def on_entry_click(self, event=None):
        """Entry 클릭 시 드롭다운 표시"""
        customer = self.customer_var.get()
        screw_type = self.screw_type_var.get()
        gauge = self.gauge_var.get()
        
        # 조건이 모두 선택되었을 때만 드롭다운 표시
        if customer and screw_type and gauge:
            self.update_screw_listbox()
            if self.screw_listbox.size() > 0:
                self.screw_listbox.pack(fill=tk.X)
                self.listbox_visible = True
        else:
            self.screw_listbox.pack_forget()
            self.listbox_visible = False
    
    def on_entry_key_release(self, event=None):
        """Entry 키 입력 시 필터링"""
        if self.listbox_visible:
            # 입력된 텍스트로 필터링
            search_text = self.screw_name_entry.get().strip().lower()
            filtered_screws = self.get_filtered_screw_list()
            
            # 입력된 텍스트로 필터링
            if search_text:
                filtered_screws = [s for s in filtered_screws if search_text in s.lower()]
            
            # 리스트박스 업데이트
            self.screw_listbox.delete(0, tk.END)
            if filtered_screws:
                for screw_name in filtered_screws:
                    self.screw_listbox.insert(tk.END, screw_name)
    
    def on_entry_key_release_with_preview(self, event=None):
        """Entry 키 입력 시 필터링 및 미리보기 업데이트"""
        # 필터링 처리
        self.on_entry_key_release(event)
        # 미리보기 업데이트
        self.update_output_preview()
    
    def on_listbox_select(self, event=None):
        """리스트박스에서 나사 선택 시"""
        selection = self.screw_listbox.curselection()
        if selection:
            selected_screw = self.screw_listbox.get(selection[0])
            self.screw_name_entry.delete(0, tk.END)
            self.screw_name_entry.insert(0, selected_screw)
            self.screw_listbox.pack_forget()
            self.listbox_visible = False
            # 복사 실행
            self.copy_to_clipboard()
    
    def find_screw_name(self, input_name, screw_type=None):
        """입력된 이름으로 데이터베이스에서 실제 나사 이름 찾기 (나사 종류 고려)"""
        normalized_input = self.normalize_screw_name(input_name)
        
        # 정확히 일치하는 경우
        if input_name in self.screw_data:
            data_list = self.screw_data[input_name]
            # 나사 종류가 지정되어 있으면 일치하는지 확인
            if screw_type is None:
                return input_name
            for data in data_list:
                if data.get('type') == screw_type:
                    return input_name
        
        # 정규화된 이름으로 매칭
        for db_name in self.screw_data.keys():
            normalized_db = self.normalize_screw_name(db_name)
            if normalized_input == normalized_db:
                data_list = self.screw_data[db_name]
                # 나사 종류가 지정되어 있으면 일치하는지 확인
                if screw_type is None:
                    return db_name
                for data in data_list:
                    if data.get('type') == screw_type:
                        return db_name
        
        return None
    
    def get_screw_specs(self, screw_name):
        """나사 이름으로 사양 정보 가져오기 (선택된 옵션에 따라 필터링)"""
        customer = self.customer_var.get()
        screw_type = self.screw_type_var.get()
        gauge = self.gauge_var.get()
        
        # 고객사/게이지 조합 허용
        allowed = False
        if customer == "SM" and gauge in ("NORMAL", "BEFORE_PLATING"):
            allowed = True  # SM 일반/도금전 둘 다 허용
        elif customer == "OTHER" and gauge in ("NORMAL", "BEFORE_PLATING"):
            allowed = True  # 그외 고객사 일반/도금전 둘 다 허용
        
        if not allowed:
            return None, None
        
        # 나사 종류가 선택되지 않았으면 None 반환
        if not screw_type:
            return None, None
        
        # 디버깅: 데이터베이스에 있는 나사 이름 확인
        normalized_input = self.normalize_screw_name(screw_name)
        matching_screws = []
        for db_name in self.screw_data.keys():
            normalized_db = self.normalize_screw_name(db_name)
            if normalized_input == normalized_db:
                data_list = self.screw_data[db_name]
                for data in data_list:
                    matching_screws.append(f"{db_name} (type: {data.get('type')}, gauge: {data.get('gauge')})")
        
        # 실제 나사 이름 찾기 (-2A, -2B 무시, 나사 종류도 고려)
        actual_screw_name = self.find_screw_name(screw_name, screw_type)
        
        if not actual_screw_name:
            # 디버깅 정보 출력
            if matching_screws:
                print(f"디버그: '{screw_name}' 입력, 나사 종류: {screw_type}, 게이지: {gauge}")
                print(f"  → 매칭되는 나사: {', '.join(matching_screws)}")
                print(f"  → 선택된 나사 종류({screw_type})와 일치하는 것이 없음")
            return None, None
        
        # 게이지와 나사 종류에 맞는 데이터 찾기
        if actual_screw_name in self.screw_data:
            data_list = self.screw_data[actual_screw_name]
            
            # 그 외 고객사 + 도금 전 게이지인 경우, 일반 게이지 데이터를 가져와서 계산
            if customer == "OTHER" and gauge == "BEFORE_PLATING":
                # 일반 게이지 데이터 찾기
                normal_data = None
                for data in data_list:
                    if data.get('type') == screw_type and data.get('gauge') == 'NORMAL':
                        data_customer = data.get('customer', 'SM')
                        if data_customer == 'OTHER':
                            normal_data = data
                            break
                
                if normal_data:
                    # 일반 게이지 데이터를 기반으로 도금 전 게이지 계산
                    return actual_screw_name, normal_data  # 계산은 format_value에서 수행
            
            # 일반적인 경우: 게이지와 나사 종류가 모두 일치하는 데이터 확인
            for data in data_list:
                if data.get('type') == screw_type and data.get('gauge') == gauge:
                    # 고객사 확인 (데이터에 customer 필드가 있으면 확인, 없으면 SM로 간주)
                    data_customer = data.get('customer', 'SM')
                    if customer == data_customer:
                        return actual_screw_name, data
        
        return None, None
    
    def apply_plating_value(self, value_dict, screw_type, plating_value_um):
        """도금 전 게이지 계산 (그 외 고객사 + 도금 전 게이지)
        
        계산 방법:
        - Min: (그 외 고객사, 일반 게이지 Min) - 입력된 도금 값 / 1000
        - Max: (그 외 고객사 Min + 그 외 고객사 Max) / 2 - 입력된 도금 값 * 2 / 1000
        """
        if not value_dict or not isinstance(value_dict, dict) or not plating_value_um:
            return value_dict
        
        try:
            # ㎛를 mm로 변환 (1㎛ = 0.001mm, /1000)
            plating_value_mm = float(plating_value_um) / 1000.0
            
            result = {}
            
            # Min 계산: (일반 게이지 Min) - 도금 값 / 1000
            if 'min' in value_dict and value_dict['min']:
                try:
                    normal_min = float(value_dict['min'])
                    before_plating_min = normal_min - plating_value_mm
                    result['min'] = f"{before_plating_min:.3f}".rstrip('0').rstrip('.')
                except (ValueError, TypeError):
                    result['min'] = value_dict['min']
            
            # Max 계산: (일반 게이지 Min + Max) / 2 - 도금 값 * 2 / 1000
            if 'max' in value_dict and value_dict['max']:
                try:
                    normal_min = float(value_dict.get('min', value_dict['max']))
                    normal_max = float(value_dict['max'])
                    # (Min + Max) / 2 계산
                    average = (normal_min + normal_max) / 2.0
                    # 도금 값 * 2 / 1000 빼기
                    before_plating_max = average - (plating_value_mm * 2)
                    result['max'] = f"{before_plating_max:.3f}".rstrip('0').rstrip('.')
                except (ValueError, TypeError):
                    result['max'] = value_dict['max']
            elif 'min' in value_dict and value_dict['min']:
                # Max가 없고 Min만 있는 경우
                try:
                    normal_min = float(value_dict['min'])
                    before_plating_max = normal_min - (plating_value_mm * 2)
                    result['max'] = f"{before_plating_max:.3f}".rstrip('0').rstrip('.')
                except (ValueError, TypeError):
                    pass
            
            return result
        except (ValueError, TypeError):
            # 도금값이 숫자가 아니면 원본 반환
            return value_dict
    
    def format_value(self, value_dict, label, screw_type=None, plating_value_um=None):
        """값 딕셔너리를 포맷팅하여 문자열 반환 (Ø ~ Ø 형식, 요청에 따라 접미사 처리)"""
        if not value_dict or not isinstance(value_dict, dict):
            return None
        
        # 도금값 적용 (조건: 그외 고객사 + 도금 전 게이지)
        customer = self.customer_var.get()
        gauge = self.gauge_var.get()
        if customer == "OTHER" and gauge == "BEFORE_PLATING" and plating_value_um:
            value_dict = self.apply_plating_value(value_dict, screw_type, plating_value_um)
        
        min_val = value_dict.get('min', '')
        max_val = value_dict.get('max', '')
        
        # 값이 문자열인 경우 공백 제거
        if isinstance(min_val, str):
            min_val = min_val.strip()
        if isinstance(max_val, str):
            max_val = max_val.strip()
        
        # 빈 문자열 체크
        if not min_val:
            min_val = None
        if not max_val:
            max_val = None
        
        # 요청 접미사 처리
        suffix = ""
        if label == "Minor" and screw_type == "OUTER":
            # 외경(-2A): Minor에 Max.를 붙임 (띄어쓰기 포함)
            suffix = " Max."
        if label == "Major" and screw_type == "INNER":
            # 내경(-2B): Major에 Min.를 붙임 (띄어쓰기 포함)
            suffix = " Min."
        
        # Min/Max 출력
        if min_val and max_val:
            # 범위가 있을 때는 접미사는 붙이지 않고 범위 표시
            return f"{label} : Ø{min_val} ~ Ø{max_val}"
        elif max_val:
            return f"{label} : Ø{max_val}{suffix}"
        elif min_val:
            return f"{label} : Ø{min_val}{suffix}"
        
        return None
    
    def format_screw_text(self, screw_name, specs=None, actual_screw_name=None):
        """나사 정보를 지정된 형식으로 포맷팅"""
        lines = []
        
        # 첫 번째 줄: 나사 이름 (실제 데이터베이스 이름 사용, -2A/-2B 포함)
        if actual_screw_name:
            # 실제 이름에 -2A/-2B가 있으면 그대로 사용
            display_name = actual_screw_name
        else:
            # 없으면 입력한 이름에 나사 종류에 따라 추가
            if specs and specs.get('type') == 'OUTER':
                # -2A가 없으면 추가
                if '-2A' not in screw_name.upper() and '-2B' not in screw_name.upper():
                    display_name = f"{screw_name}-2A"
                else:
                    display_name = screw_name
            elif specs and specs.get('type') == 'INNER':
                # -2B가 없으면 추가
                if '-2A' not in screw_name.upper() and '-2B' not in screw_name.upper():
                    display_name = f"{screw_name}-2B"
                else:
                    display_name = screw_name
            else:
                display_name = screw_name
        
        lines.append(display_name)
        
        # 두 번째 줄: 게이지 정보
        gauge = self.gauge_var.get()
        if gauge == "BEFORE_PLATING":
            lines.append("(도금 전 게이지 사용)")
        elif gauge == "NORMAL":
            lines.append("(일반 게이지 사용)")
        
        # 나머지 줄: 사양 정보 (나사 종류에 따라 순서 다름)
        if specs:
            screw_type = specs.get('type', '')
            customer = self.customer_var.get()
            gauge = self.gauge_var.get()
            
            # 도금값 가져오기 (그외 고객사 + 도금 전 게이지일 때만 사용)
            plating_value_um = None
            if customer == "OTHER" and gauge == "BEFORE_PLATING":
                plating_value_str = self.plating_value_var.get().strip()
                if plating_value_str:
                    try:
                        plating_value_um = float(plating_value_str)
                    except (ValueError, TypeError):
                        plating_value_um = None
            
            if screw_type == 'OUTER':
                # 외경나사: Major, Pitch, Minor 순서
                # Minor는 Max. 접미사, Major는 Min/Max 둘 다 있으면 범위, 하나만 있으면 단일
                # 외경나사는 Minor 값이 없으면 제외하고 출력
                if 'major' in specs:
                    major_str = self.format_value(specs['major'], "Major", screw_type, plating_value_um)
                    if major_str:
                        lines.append(major_str)
                
                if 'pitch' in specs:
                    pitch_str = self.format_value(specs['pitch'], "Pitch", screw_type, plating_value_um)
                    if pitch_str:
                        lines.append(pitch_str)
                
                # 외경나사: Minor 값이 있을 때만 출력
                if 'minor' in specs and specs['minor']:
                    minor_str = self.format_value(specs['minor'], "Minor", screw_type, plating_value_um)
                    if minor_str:
                        lines.append(minor_str)
            
            elif screw_type == 'INNER':
                # 내경나사: Minor, Pitch, Major 순서
                # 내경나사는 Major 값이 없으면 제외하고 출력
                if 'minor' in specs:
                    minor_str = self.format_value(specs['minor'], "Minor", screw_type, plating_value_um)
                    if minor_str:
                        lines.append(minor_str)
                
                if 'pitch' in specs:
                    pitch_str = self.format_value(specs['pitch'], "Pitch", screw_type, plating_value_um)
                    if pitch_str:
                        lines.append(pitch_str)
                
                # 내경나사: Major 값이 있을 때만 출력
                if 'major' in specs and specs['major']:
                    major_str = self.format_value(specs['major'], "Major", screw_type, plating_value_um)
                    if major_str:
                        lines.append(major_str)
            else:
                # 타입이 없을 경우 기본 순서 (Minor, Pitch, Major)
                # 도금값 가져오기 (그외 고객사 + 도금 전 게이지일 때만 사용)
                plating_value_um = None
                if customer == "OTHER" and gauge == "BEFORE_PLATING":
                    plating_value_str = self.plating_value_var.get().strip()
                    if plating_value_str:
                        try:
                            plating_value_um = float(plating_value_str)
                        except (ValueError, TypeError):
                            plating_value_um = None
                
                if 'minor' in specs:
                    minor_str = self.format_value(specs['minor'], "Minor", None, plating_value_um)
                    if minor_str:
                        lines.append(minor_str)
                
                if 'pitch' in specs:
                    pitch_str = self.format_value(specs['pitch'], "Pitch", None, plating_value_um)
                    if pitch_str:
                        lines.append(pitch_str)
                
                if 'major' in specs:
                    major_str = self.format_value(specs['major'], "Major", None, plating_value_um)
                    if major_str:
                        lines.append(major_str)
        
        return '\n'.join(lines)
    
    def on_selection_changed(self):
        """라디오 버튼 선택이 변경되었을 때 호출"""
        # 드롭다운 업데이트
        if self.listbox_visible:
            self.update_screw_listbox()
            if self.screw_listbox.size() == 0:
                self.screw_listbox.pack_forget()
                self.listbox_visible = False
        # 출력 결과 미리보기 업데이트 (복사는 하지 않음)
        self.update_output_preview()
    
    def on_enter_pressed(self, event=None):
        """엔터 키를 눌렀을 때 클립보드에 복사"""
        if event:
            event.widget = self.screw_name_entry
        self.copy_to_clipboard()
    
    def update_output_preview(self):
        """선택사항 변경 시 출력 결과 미리보기만 업데이트 (복사는 하지 않음)"""
        screw_name = self.screw_name_entry.get().strip()
        
        if not screw_name:
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(1.0, "나사 이름을 입력하세요.")
            self.output_text.config(state=tk.DISABLED)
            return
        
        try:
            # 나사 사양 정보 가져오기
            actual_screw_name, specs = self.get_screw_specs(screw_name)
            
            # 사양 정보가 없으면 빈 화면
            if not specs:
                self.output_text.config(state=tk.NORMAL)
                self.output_text.delete(1.0, tk.END)
                self.output_text.insert(1.0, "해당 조건에 맞는 나사 데이터를 찾을 수 없습니다.")
                self.output_text.config(state=tk.DISABLED)
                return
            
            # 지정된 형식으로 텍스트 포맷팅
            formatted_text = self.format_screw_text(screw_name, specs, actual_screw_name)
            
            # 보유 현황 업데이트 (제거됨: 별도 검색창에서만 확인)
            
            # 출력 결과 표시만 (복사는 하지 않음)
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(1.0, formatted_text)
            self.output_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(1.0, f"오류 발생: {str(e)}")
            self.output_text.config(state=tk.DISABLED)
    
    def check_inventory_popup(self, event=None):
        """보유 나사게이지 확인 검색 창 엔터 입력 시 팝업 표시"""
        search_text = self.memo_entry.get().strip()
        
        if not search_text:
            return
            
        # 입력된 이름을 정규화하여 검색
        normalized_input = self.normalize_screw_name(search_text)
        
        if normalized_input in self.inventory:
            self.inventory_status_label.config(text="보유 게이지 (O)", foreground="green")
            messagebox.showinfo("확인 결과", "보유중인 나사 게이지 입니다.")
        else:
            self.inventory_status_label.config(text="미보유 게이지 (X)", foreground="red")
            messagebox.showwarning("확인 결과", "보유중이지 않은 나사 게이지 입니다.")
    
    def copy_to_clipboard(self):
        """나사 이름을 지정된 형식으로 클립보드에 복사"""
        screw_name = self.screw_name_entry.get().strip()
        
        if not screw_name:
            self.status_label.config(text="상태: 나사 이름을 입력하세요", foreground="orange")
            return
        
        try:
            # 나사 사양 정보 가져오기
            actual_screw_name, specs = self.get_screw_specs(screw_name)
            
            # 디버깅: 저장된 데이터 확인
            if actual_screw_name and actual_screw_name in self.screw_data:
                stored_data_list = self.screw_data[actual_screw_name]
                print(f"디버그: 저장된 데이터 확인 - '{actual_screw_name}'")
                for idx, stored_data in enumerate(stored_data_list):
                    print(f"  → 데이터 {idx+1}: Minor={stored_data.get('minor')}, Pitch={stored_data.get('pitch')}, Major={stored_data.get('major')}, Type={stored_data.get('type')}, Gauge={stored_data.get('gauge')}")
            
            # 사양 정보가 없으면 경고
            if not specs:
                customer = self.customer_var.get()
                gauge = self.gauge_var.get()
                screw_type = self.screw_type_var.get()
                
                if customer != "SM" or gauge != "NORMAL":
                    self.status_label.config(text="상태: SM + 일반게이지를 선택해야 합니다", foreground="orange")
                    return
                
                if not screw_type:
                    self.status_label.config(text="상태: 나사 종류를 선택하세요", foreground="orange")
                    return
                
                # 디버깅: 매칭되는 나사 확인
                normalized_input = self.normalize_screw_name(screw_name)
                matching_screws = []
                for db_name in self.screw_data.keys():
                    normalized_db = self.normalize_screw_name(db_name)
                    if normalized_input == normalized_db:
                        data_list = self.screw_data[db_name]
                        for data in data_list:
                            matching_screws.append(f"{db_name} (type: {data.get('type')}, gauge: {data.get('gauge')})")
                
                if matching_screws:
                    self.status_label.config(
                        text=f"상태: '{screw_name}' ({screw_type}) 데이터를 찾을 수 없습니다. 매칭: {', '.join(matching_screws)}", 
                        foreground="orange")
                else:
                    self.status_label.config(text=f"상태: '{screw_name}' 데이터를 찾을 수 없습니다", foreground="orange")
                return
            
            # 지정된 형식으로 텍스트 포맷팅 (실제 나사 이름 사용하여 -2A/-2B 포함)
            formatted_text = self.format_screw_text(screw_name, specs, actual_screw_name)
            
            # 출력 결과 표시
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(1.0, formatted_text)
            self.output_text.config(state=tk.DISABLED)
            
            # 클립보드에 복사
            pyperclip.copy(formatted_text)
            
            # 복사 완료 알림창 표시
            messagebox.showinfo("복사 완료", f"복사 되었습니다! {actual_screw_name}")
            
            # 상태 표시
            status_parts = []
            customer = self.customer_var.get()
            screw_type = self.screw_type_var.get()
            gauge = self.gauge_var.get()
            
            if customer == "SM":
                status_parts.append("SM")
            elif customer == "OTHER":
                status_parts.append("그외 고객사")
            
            if screw_type == "INNER":
                status_parts.append("내경나사")
            elif screw_type == "OUTER":
                status_parts.append("외경나사")
            
            if gauge == "NORMAL":
                status_parts.append("일반게이지")
            elif gauge == "BEFORE_PLATING":
                status_parts.append("도금 전 게이지")
            
            status_text = f"상태: 복사 완료 - {', '.join(status_parts) if status_parts else '옵션 미선택'}"
            self.status_label.config(text=status_text, foreground="green")
            
        except Exception as e:
            self.status_label.config(text=f"상태: 오류 발생 - {str(e)}", foreground="red")
            messagebox.showerror("오류", f"클립보드 복사 중 오류가 발생했습니다:\n{str(e)}")

    def start_drag(self, event):
        """드래그 시작 위치 저장"""
        self.start_x = event.x_root
        self.start_y = event.y_root
    
    def on_drag(self, event):
        """창 드래그 이동"""
        x = self.root.winfo_x() + (event.x_root - self.start_x)
        y = self.root.winfo_y() + (event.y_root - self.start_y)
        self.root.geometry(f"+{x}+{y}")
        self.start_x = event.x_root
        self.start_y = event.y_root

def run_program():
    root = tk.Tk()
    ScrewInputApp(root)  # app 변수 할당 제거 (사용되지 않음)
    root.mainloop()

if __name__ == "__main__":
    run_program()

