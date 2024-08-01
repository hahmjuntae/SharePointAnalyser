import tkinter as tk
from tkinter import filedialog, messagebox, ttk # tkinter 모듈 및 하위 모듈들 가져오기
import pandas as pd # pandas 라이브러리 가져오기
import json # json 모듈 가져오기
from datetime import datetime, timedelta # datetime 및 timedelta 가져오기
import tkinter.font as tkfont # tkinter의 폰트 모듈 가져오기
import locale # 로케일 설정을 위한 locale 모듈 가져오기

# 한글 로케일 설정
locale.setlocale(locale.LC_TIME, 'ko_KR.UTF-8')

class CSVViewerApp:
    def __init__(self, root):
        self.root = root # root 윈도우 설정
        self.root.title("Audit Log Analyser") # 윈도우 제목 설정
        self.root.geometry("1200x800") # 윈도우 크기 설정

        # 버튼 프레임 설정
        self.button_frame = tk.Frame(root)
        self.button_frame.pack(fill=tk.X, pady=10)

        # 파일 업로드 버튼
        self.upload_btn = tk.Button(self.button_frame, text="Upload", command=self.upload_file, font=("Arial", 12))
        self.upload_btn.pack(side=tk.LEFT, padx=10)

        # 파일 저장 버튼
        self.save_btn = tk.Button(self.button_frame, text="Export", command=self.save_file, font=("Arial", 12))
        self.save_btn.pack(side=tk.LEFT, padx=10)

        # 초기화 버튼
        self.reset_btn = tk.Button(self.button_frame, text="Reset", command=self.reset_file, font=("Arial", 12))
        self.reset_btn.pack(side=tk.RIGHT, padx=10)

        # 종료 버튼
        self.close_btn = tk.Button(self.button_frame, text="Close", command=root.quit, font=("Arial", 12))
        self.close_btn.pack(side=tk.RIGHT, padx=10)

        # 파일 이름 표시 레이블
        self.filename_label = tk.Label(root, text="", font=("Arial", 12))
        self.filename_label.pack(pady=5)

        # 트리뷰 프레임 설정
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)

        # 트리뷰 수직 스크롤바 설정
        self.tree_scroll_y = tk.Scrollbar(self.tree_frame)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # 트리뷰 수평 스크롤바 설정
        self.tree_scroll_x = tk.Scrollbar(self.tree_frame, orient='horizontal')
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # 트리뷰 설정
        self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        # 트리뷰 항목 선택 및 복사 이벤트 바인딩
        self.tree.bind("<ButtonRelease-1>", self.select_item)
        self.tree.bind("<Control-c>", self.copy_to_clipboard)

        self.filtered_df = None # 필터링된 데이터프레임 저장을 위한 변수 초기화

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")]) # 파일 선택 다이얼로그 열기
        if not file_path: # 파일을 선택하지 않은 경우
            return

        try:
            df = self.read_csv_with_encoding(file_path)
            df = self.extract_audit_data(df)  # 'AuditData' 열의 JSON 데이터 추출
            filtered_df = self.filter_data(df)  # 필요한 열만 선택하여 필터링
            filtered_df = self.format_creation_time(filtered_df)  # 'CreationTime' 열의 시간을 KST로 변환
            self.filtered_df = filtered_df  # 필터링된 데이터프레임 저장
            self.display_data(filtered_df)  # 데이터를 트리뷰에 표시
            self.filename_label.config(text=f"Uploaded file: {file_path}")  # 파일 이름 표시
        except Exception as e:  # 예외 처리
            messagebox.showerror("Error", f"An error occurred while reading the file: {e}")

    def reset_file(self):
        self.tree.delete(*self.tree.get_children())  # 트리뷰 초기화
        self.filename_label.config(text="")  # 파일 이름 레이블 초기화
        self.filtered_df = None  # 필터링된 데이터프레임 초기화

    def read_csv_with_encoding(self, file_path):
        # 다양한 인코딩 시도
        encodings = ['utf-8-sig', 'utf-8', 'latin1', 'cp949', 'ISO-8859-1']
        for enc in encodings:
            try:
                return pd.read_csv(file_path, encoding=enc)  # 인코딩을 시도하여 CSV 파일 읽기
            except UnicodeDecodeError:
                continue
        raise ValueError("Could not decode file with available encodings")  # 인코딩 실패 시 예외 발생

    def extract_audit_data(self, df):
        audit_data_list = []  # 감사 데이터 저장을 위한 리스트
        for _, row in df.iterrows():  # 각 행에 대해
            audit_data = json.loads(row['AuditData'])  # 'AuditData' 열의 JSON 데이터 추출
            audit_data_list.append(audit_data)  # 리스트에 추가
        audit_df = pd.DataFrame(audit_data_list)  # 리스트를 데이터프레임으로 변환
        return audit_df

    def filter_data(self, df):
        # 필요한 키들만 추출하고 순서 지정
        keys_to_extract = [
            "CreationTime", "Operation", "Workload", "SourceFileName", "SiteUrl", "SourceRelativeUrl", "Platform", "UserAgent", "UserId", "ClientIP"
        ]
        filtered_df = df[keys_to_extract]  # 필요한 열만 선택하여 필터링
        return filtered_df

    def format_creation_time(self, df):
        # CreationTime 포맷 변경 및 KST로 변환
        def convert_utc_to_kst(utc_time_str):
            utc_time = datetime.strptime(utc_time_str, '%Y-%m-%dT%H:%M:%S')  # UTC 시간 문자열을 datetime 객체로 변환
            kst_time = utc_time + timedelta(hours=9)  # KST로 변환 (UTC+9)
            return kst_time.strftime('%Y년 %m월 %d일(%a) %H:%M:%S').replace('Mon', '월').replace('Tue', '화').replace('Wed', '수').replace('Thu', '목').replace('Fri', '금').replace('Sat', '토').replace('Sun', '일')

        df['CreationTime'] = df['CreationTime'].apply(convert_utc_to_kst)  # 'CreationTime' 열의 각 값을 변환
        return df

    def display_data(self, df):
        if self.tree:  # 트리뷰가 이미 있는 경우
            self.tree.destroy()  # 기존 트리뷰 삭제

        self.tree = ttk.Treeview(self.tree_frame, columns=list(df.columns), yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        self.tree["show"] = "headings"  # 열 제목 표시

        for col in df.columns:  # 각 열에 대해
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_data(c, False))  # 열 제목 설정 및 정렬 기능 추가
            self.tree.column(col, width=tkfont.Font().measure(col))  # 열 너비 설정

        for index, row in df.iterrows():  # 각 행에 대해
            values = list(row)  # 행의 값을 리스트로 변환
            self.tree.insert("", index, values=values)  # 트리뷰에 행 추가
            for col_num, value in enumerate(values):  # 각 값에 대해
                col_width = tkfont.Font().measure(value)  # 값의 너비 측정
                if self.tree.column(df.columns[col_num], width=None) < col_width:  # 기존 열 너비보다 큰 경우
                    self.tree.column(df.columns[col_num], width=col_width)  # 열 너비 업데이트

    def sort_data(self, col, descending):
        data_list = [(self.tree.set(child, col), child) for child in self.tree.get_children('')]  # 각 행의 데이터를 리스트로 저장
        data_list.sort(reverse=descending)  # 리스트 정렬

        for index, (val, child) in enumerate(data_list):  # 정렬된 리스트에 대해
            self.tree.move(child, '', index)  # 트리뷰에서 행 이동

        self.tree.heading(col, command=lambda: self.sort_data(col, not descending))  # 정렬 방향 토글

    def save_file(self):
        if self.filtered_df is not None:  # 필터링된 데이터프레임이 있는 경우
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])  # 파일 저장 다이얼로그 열기
            if not file_path:  # 파일 경로를 선택하지 않은 경우
                return
            self.filtered_df.to_excel(file_path, index=False)  # 데이터프레임을 엑셀 파일로 저장
            messagebox.showinfo("Saved", f"Filtered data has been saved to {file_path}")  # 저장 성공 메시지
        else:
            messagebox.showwarning("No Data", "No data to save. Please upload a CSV file first.")  # 데이터 없음 경고 메시지

    def select_item(self, event):
        cur_item = self.tree.focus()  # 선택된 항목 가져오기
        if cur_item:  # 선택된 항목이 있는 경우
            print(self.tree.item(cur_item))  # 항목 정보 출력

    def copy_to_clipboard(self, event):
        cur_item = self.tree.focus()  # 선택된 항목 가져오기
        if cur_item:  # 선택된 항목이 있는 경우
            item_values = self.tree.item(cur_item)["values"]  # 항목 값 가져오기
            self.root.clipboard_clear()  # 클립보드 초기화
            self.root.clipboard_append("\t".join(map(str, item_values)))  # 클립보드에 값 추가
            self.root.update()  # 클립보드 업데이트

if __name__ == "__main__":
    root = tk.Tk()  # 루트 윈도우 생성
    app = CSVViewerApp(root)  # 앱 인스턴스 생성
    root.mainloop()  # 메인 루프 실행
