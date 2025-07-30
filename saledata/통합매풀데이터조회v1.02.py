import os
import pickle
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from tksheet import Sheet
import threading
import time

# 상수 선언
DB_DIRS = [r"\\NAS451\team451\DB", r"D:\hSync\Coding\DB"]
PICKLE_FILE = r"\통합매출데이터.pickle"
EXCEL_FILE = r"\통합매출데이터.xlsx"
MENU_NAMES = ["상품별 판매집계", "판매처별 판매집계", "년도별 판매집계"]

class SaleDateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("통합매출데이터 조회 v1.02")
        self.root.geometry("1800x900+60+60")
        self.root.state("zoomed")
        self.root.config(background="#EEEEEE")

        self.data = pd.DataFrame()
        self.filtered_data = pd.DataFrame()

        # UI 구성
        self._build_ui()

    # 에러 메시지 출력 함수
    def show_error(self, msg, err):
        messagebox.showerror('에러', f'{msg}\n{err}')

    # UI 세팅
    def _build_ui(self):
        # 전체 프레임
        self.main_frame = tk.Frame(self.root, bg="#EEEEEE")
        self.main_frame.pack(fill="both", expand=True)
        # 좌측 메뉴 프레임
        self.menu_frame = tk.Frame(self.main_frame, bg="#DDDDDD", width=150)
        self.menu_frame.pack(side="left", fill="y")
        self.menu_frame.pack_propagate(False)
        # 우측 디스플레이 프레임
        self.container_frame = tk.Frame(self.main_frame, bg="#FFFFFF")
        self.container_frame.pack(side="top", fill="both", expand=True)
        # 메뉴 타이틀
        self.title_label = tk.Label(self.menu_frame, width=22, height=4, text="판매데이터", font=("맑은 고딕", 12), bg="#FFFFFF")
        self.title_label.pack(pady=5, padx=5)
        # 메뉴 버튼 생성
        self.menu_buttons = []
        self._create_menu_buttons()
        # 디스플레이 라벨
        self.display = tk.Frame(self.container_frame, bg="#FFFFFF")
        self.display.pack(expand=True)

    # 메뉴 버튼 생성 함수
    def _create_menu_buttons(self):
        for idx, name in enumerate(MENU_NAMES):
            btn = tk.Button(self.menu_frame, text=name, width=22, justify="left", relief="flat", bg="#333333", fg="#FFFFFF",
                            command=lambda n=name: self.show_display(n))
            btn.pack(pady=5, padx=5, fill="x")
            self.menu_buttons.append(btn)

    def show_display(self, menu_name):
        btn_relief = "flat"
        # Function Frame (검색 옵션, 버튼)
        if menu_name == "상품별 판매집계":
            self.top_frame = tk.Frame(self.container_frame)
            self.top_frame.pack(side="top", padx=5, pady=5, fill="x")
            self.top_frame.pack(padx=5, pady=5, fill="x")
            try:
                from tkcalendar import DateEntry
            except ImportError:
                self.show_error("tkcalendar 모듈이 설치되어 있지 않습니다.\n명령 프롬프트에서 'pip install tkcalendar'를 실행해 주세요.", "")
                self.main_frame.quit()
                return
            tk.Label(self.top_frame, text="기간별 조회:", font=("맑은 고딕", 10)).pack(side="left", padx=5, pady=5)
            self.period_start = DateEntry(self.top_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
            self.period_start.pack(side="left", padx=5, pady=5, ipady=2)
            tk.Label(self.top_frame, text="~").pack(side="left")
            self.period_end = DateEntry(self.top_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
            self.period_end.pack(side="left", padx=5, pady=5, ipady=2)
            tk.Button(self.top_frame, text="조회", width=10, relief=btn_relief, fg='#FFFFFF', bg='#007ACC',
                    command=self.show_product_sales_summary).pack(side="left", padx=5, pady=5)
            tk.Button(self.top_frame, text="초기화", width=12, relief=btn_relief, fg='#FFFFFF', bg='#333333',
                    command=self.reset).pack(side="left", padx=5, pady=5)
            
            self.contant_frame= tk.Frame(self.container_frame)
            self.contant_frame.pack(padx=5, pady=5, fill="both", expand=True)
        
        elif menu_name == "판매처별 판매집계":
            self.contant_frame= tk.Frame(self.container_frame)
            self.contant_frame.pack(padx=5, pady=5, fill="both", expand=True)
        
        else:
            self.contant_frame= tk.Frame(self.container_frame)
            self.contant_frame.pack(padx=5, pady=5, fill="both", expand=True)

    # 데이터 로드
    def load_data(self):
        try:
            db_dir = next((d for d in DB_DIRS if os.path.isdir(d)), None)
            if not db_dir:
                raise FileNotFoundError('NAS451서버 또는 로컬 DB 폴더를 찾을 수 없습니다.')
            pickle_path = db_dir + PICKLE_FILE
            excel_path = db_dir + EXCEL_FILE
            if os.path.exists(pickle_path):
                with open(pickle_path, 'rb') as fr:
                    self.data = pickle.load(fr)
            else:
                self.data = pd.read_excel(excel_path, sheet_name="DB", header=1)
        except Exception as e:
            self.show_error('상품정보 파일을 불러올 수 없습니다!', e)
            self.data = pd.DataFrame()
        self.data = self.data.astype(str, errors="ignore").fillna("")
        # self.filtered_data = self.data.copy()
        # self.update_sheet(self.data)

    # 데이터 로드 스레드 (UI 멈춤 방지)
    def load_data_threaded(self):
        self.start_progress()
        threading.Thread(target=self.load_data, daemon=True).start()
    # filter_by_period를 버튼에서 호출할 때 인자 없이 호출하므로, start_date와 end_date를 Entry에서 읽어오도록 오버로딩된 메서드를 추가합니다.

    # 시트 출력 함수 (tksheet)
    def update_sheet(self, df):
        # 이전 Sheet Frame 초기화
        for widget in self.frm_sheet.winfo_children():
            widget.destroy()
        if df.empty:
            tk.Label(self.frm_sheet, text="조회할 데이터가 없습니다.", fg="red").pack(padx=10, pady=10)
            return

        sheet = Sheet(self.frm_sheet, data=df.values.tolist(), headers=list(df.columns),
                      header_height=25, header_fg="#FFFFFF", header_bg="#333333")
        sheet.header_font(('NanumGothic', 10, 'normal'))
        sheet.font(('NanumGothic', 9, 'normal'))
        sheet.table_align(align="left")
        sheet.set_all_column_widths(width=None)
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)

    # 진행바 애니메이션 (UI 스레드-safe)
    def start_progress(self):
        self.progress_var.set(0)
        def animate(i=0):
            if i > 100:
                return
            self.progress_var.set(i)
            self.frm_layout.update_idletasks()
            self.frm_layout.after(8, animate, i+2)
        animate()

if __name__ == "__main__":
    root = tk.Tk()
    app = SaleDateApp(root)
    root.mainloop()
