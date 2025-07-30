import os
import pickle
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from tksheet import Sheet
import threading
import time

class SaleDateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("통합매출데이터 조회 v1.01")
        self.root.geometry("1800x900+60+60")
        self.root.state("zoomed")
        self.root.config(background="#EEEEEE")

        self.data = pd.DataFrame()
        self.filtered_data = pd.DataFrame()

        # UI 구성
        self._build_ui()

        # 데이터 로드
        self.load_data_threaded()

    # UI 세팅
    def _build_ui(self):
        btn_relief = "flat"

        # Layout Frame (설명, 종료 버튼)
        self.frm_layout = tk.Frame(self.root)
        self.frm_layout.pack(padx=10, pady=2, fill="both", expand=True)
        
        # Top Frame (설명, 종료 버튼)
        self.frm_header = tk.Frame(self.frm_layout, background="#FFFFFF")
        self.frm_header.pack(padx=2, pady=2, fill="x")

        desc = "1. 검색유형을 선택하고 검색 정보를 입력해 주세요.   2. 소스데이터 : /nas451/DB/통합매출데이터.xlsx"
        tk.Label(self.frm_header, text=desc, height=2, justify="left", bg='#FFFFFF').pack(side="left", padx=5, pady=5)
        tk.Button(self.frm_header, text="프로그램 종료", width=16, relief=btn_relief, fg='#FFFFFF', bg='#333333',
                  command=self.frm_layout.quit).pack(side="right", padx=5, pady=5)
        tk.Button(self.frm_header, text="등록상품 업데이트", width=20, relief=btn_relief, fg='#FFFFFF', bg='#333333',
                  command=self.update_data_threaded).pack(side="right", padx=5, pady=5)
        
        # Progress Bar
        self.frm_progress = tk.Frame(self.frm_layout, background="#FFFFFF")
        self.frm_progress.pack(fill="x", padx=2, pady=2)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.frm_progress, maximum=100, variable=self.progress_var, mode='determinate')
        self.progress_bar.pack(padx=5, fill="x")

        # Main Container (좌:메뉴창, 우: 데이터시트)
        self.frm_container = tk.Frame(self.frm_layout, background="#FFFFFF")
        self.frm_container.pack(padx=2, pady=2, fill="both", expand=True)

        self.frm_aside = tk.Frame(self.frm_container, background="#FFFFFF")
        self.frm_aside.pack(side="left", padx=0, pady=0, fill="y")

        self.frm_menu = tk.LabelFrame(self.frm_aside)
        self.frm_menu.pack(side="left", padx=5, pady=5, fill="both")
        
        tk.Button(self.frm_menu, text="상품별 판매집계", width=25, relief=btn_relief, fg='#FFFFFF', bg='#333333',
                  command=self.reset).pack(padx=5, pady=5)
        tk.Button(self.frm_menu, text="판매처별 판매집계", width=25, relief=btn_relief, fg='#FFFFFF', bg='#333333',
                  command=self.reset).pack(padx=5, pady=5)
        tk.Button(self.frm_menu, text="년도별 판매집계", width=25, relief=btn_relief, fg='#FFFFFF', bg='#333333',
                  command=self.reset).pack(padx=5, pady=5)        

        # Main Container (우: 데이터시트)
        self.frm_content = tk.Frame(self.frm_container, background="#FFFFFF")
        self.frm_content.pack(side="left", padx=0, pady=0, fill="both", expand=True)

        # Function Frame (검색 옵션, 버튼)
        self.frm_function = tk.LabelFrame(self.frm_content)
        self.frm_function.pack(padx=5, pady=5, fill="x")
        
        # 달력 위젯을 사용하여 기간을 선택할 수 있도록 tkcalendar의 DateEntry를 사용합니다.
        # tkcalendar가 설치되어 있어야 하며, 설치가 안 되어 있다면 pip install tkcalendar 필요
        try:
            from tkcalendar import DateEntry
        except ImportError:
            messagebox.showerror("오류", "tkcalendar 모듈이 설치되어 있지 않습니다.\n명령 프롬프트에서 'pip install tkcalendar'를 실행해 주세요.")
            self.frm_layout.quit()
            return

        tk.Label(self.frm_function, text="기간별 조회:", font=("맑은 고딕", 10)).pack(side="left", padx=5, pady=5)
        self.period_start = DateEntry(self.frm_function, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.period_start.pack(side="left", padx=5, pady=5, ipady=2)
        tk.Label(self.frm_function, text="~").pack(side="left")
        self.period_end = DateEntry(self.frm_function, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.period_end.pack(side="left", padx=5, pady=5, ipady=2)
         
        tk.Button(self.frm_function, text="조회", width=10, relief=btn_relief, fg='#FFFFFF', bg='#007ACC',
                  command=self.show_product_sales_summary).pack(side="left", padx=5, pady=5)
        tk.Button(self.frm_function, text="초기화", width=12, relief=btn_relief, fg='#FFFFFF', bg='#333333',
                  command=self.reset).pack(side="left", padx=5, pady=5)
        tk.Button(self.frm_function, text="전체 등록상품 보기", width=20, relief=btn_relief, fg='#FFFFFF', bg='#333333',
                  command=self.show_all_data).pack(side="right", padx=5, pady=5)

        self.frm_sheet = tk.LabelFrame(self.frm_content)
        self.frm_sheet.pack(side="left", padx=5, pady=5, fill="both", expand=True)

    # 데이터 로드
    def load_data(self):
        try:
            dir_path1 = r"\\NAS451\team451\DB"
            dir_path2 = r"D:\hSync\Coding\DB"
            pickle_file = r"\통합매출데이터.pickle"
            excel_file = r"\통합매출데이터.xlsx"

            if os.path.isdir(dir_path1):
                pickle_path = dir_path1 + pickle_file
                excel_path = dir_path1 + excel_file
            elif os.path.isdir(dir_path2):
                pickle_path = dir_path2 + pickle_file
                excel_path = dir_path2 + excel_file
            else:
                raise FileNotFoundError('NAS451서버 또는 로컬 DB 폴더를 찾을 수 없습니다.')

            # 피클이 있으면 피클 우선, 없으면 엑셀
            if os.path.exists(pickle_path):
                with open(pickle_path, 'rb') as fr:
                    self.data = pickle.load(fr)
            else:
                self.data = pd.read_excel(excel_path, sheet_name="DB", header=1)
        except Exception as e:
            messagebox.showerror('에러', f'상품정보 파일을 불러올 수 없습니다!\n{e}')
            self.data = pd.DataFrame()

        # 모든 컬럼 문자 처리
        self.data = self.data.astype(str, errors="ignore").fillna("")
        # self.filtered_data = self.data.copy()
        # self.update_sheet(self.data)

    # 데이터 로드 스레드 (UI 멈춤 방지)
    def load_data_threaded(self):
        self.start_progress()
        threading.Thread(target=self.load_data, daemon=True).start()
    # filter_by_period를 버튼에서 호출할 때 인자 없이 호출하므로, start_date와 end_date를 Entry에서 읽어오도록 오버로딩된 메서드를 추가합니다.
    
    def filter_by_period(self):
        """
        기간 Entry에서 값을 읽어와서 filter_by_period_core를 호출합니다.
        """
        start_date = self.period_start.get().strip()
        end_date = self.period_end.get().strip()
        if not start_date or not end_date:
            messagebox.showerror('에러', '시작일과 종료일을 모두 입력해 주세요.')
            return
        self.filter_by_period_core(start_date, end_date)

    def filter_by_period_core(self, start_date, end_date, date_column="날짜"):
        """
        지정한 기간(start_date ~ end_date) 내의 데이터만 필터링합니다.
        start_date, end_date: 'YYYY-MM-DD' 형식의 문자열
        date_column: 날짜가 저장된 컬럼명 (기본값: '날짜')
        """
        try:
            # 날짜 컬럼이 존재하는지 확인
            if date_column not in self.data.columns:
                messagebox.showerror('에러', f"'{date_column}' 컬럼이 존재하지 않습니다.")
                return

            # 날짜 컬럼을 datetime으로 변환 (오류 발생시 NaT)
            df = self.data.copy()
            df[date_column] = pd.to_datetime(df[date_column], errors='coerce')

            # 입력값을 datetime으로 변환
            start = pd.to_datetime(start_date)
            end = pd.to_datetime(end_date)

            # 기간 필터링
            mask = (df[date_column] >= start) & (df[date_column] <= end)
            filtered = df[mask].copy()

            # 다시 문자열로 변환 (UI 일관성)
            filtered = filtered.astype(str, errors="ignore").fillna("")

            self.filtered_data = filtered
            self.update_sheet(filtered)
        except Exception as e:
            messagebox.showerror('에러', f'기간검색 중 오류 발생: {e}')

    # 등록상품 업데이트 (엑셀 → 피클)
    def update_data(self):
        try:
            file_path = r'\\NAS451\team451\DB\통합매출데이터.xlsx'
            pickle_path = r'\\NAS451\team451\DB\통합매출데이터.pickle'
            df = pd.read_excel(file_path, sheet_name="DB")
            with open(pickle_path, 'wb') as fw:
                pickle.dump(df, fw)
            messagebox.showinfo('알림', '업데이트가 완료되었습니다')
            # 새로고침
            self.data = df.astype(str, errors="ignore").fillna("")
            self.filtered_data = self.data.copy()
            # self.update_sheet(self.data)
        except Exception as e:
            messagebox.showerror('에러', f'업데이트 실패: {e}')
    
    def show_product_sales_summary(self, start_date, end_date, product_code_column="상품코드", quantity_column="수량", date_column="날짜"):
        """
        상품별 판매집계: self.data에서 조회기간(start_date~end_date) 내 상품코드별 집계수량을 내림차순 정렬하여 sheet에 출력
        start_date, end_date: 'YYYY-MM-DD' 형식 문자열
        product_code_column: 상품코드 컬럼명
        quantity_column: 집계할 수량 컬럼명
        date_column: 날짜 컬럼명
        """
        try:
            # 날짜 컬럼이 존재하는지 확인
            if date_column not in self.data.columns:
                messagebox.showerror('에러', f"'{date_column}' 컬럼이 존재하지 않습니다.")
                return
            if product_code_column not in self.data.columns:
                messagebox.showerror('에러', f"'{product_code_column}' 컬럼이 존재하지 않습니다.")
                return
            if quantity_column not in self.data.columns:
                messagebox.showerror('에러', f"'{quantity_column}' 컬럼이 존재하지 않습니다.")
                return

            # 날짜 필터링
            df = self.data.copy()
            df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
            start = pd.to_datetime(start_date)
            end = pd.to_datetime(end_date)
            mask = (df[date_column] >= start) & (df[date_column] <= end)
            filtered = df[mask].copy()

            # 수량 컬럼 숫자 변환
            filtered[quantity_column] = pd.to_numeric(filtered[quantity_column], errors='coerce').fillna(0)

            # 상품코드별 집계
            summary = (
                filtered
                .groupby(product_code_column, as_index=False)[quantity_column]
                .sum()
                .sort_values(quantity_column, ascending=False)
            )

            # 문자열 변환 및 결측치 처리
            summary = summary.astype(str, errors="ignore").fillna("")

            self.filtered_data = summary
            self.update_sheet(self, summary)
        except Exception as e:
            messagebox.showerror('에러', f'상품별 판매집계 중 오류 발생: {e}')


    # 업데이트 스레드
    def update_data_threaded(self):
        self.start_progress()
        threading.Thread(target=self.update_data, daemon=True).start()

    # 전체 보기
    def show_all_data(self):
        self.filtered_data = self.data.copy()
        self.update_sheet(self.data)

    # 입력/검색 초기화
    def reset(self):
        # 오늘 날짜로 기간 초기화
        import datetime
        today = datetime.date.today()
        self.period_start.set_date(today)
        self.period_end.set_date(today)
        self.filtered_data = self.data.copy()
        self.update_sheet(self.data)

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