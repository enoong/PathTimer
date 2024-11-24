import ttkbootstrap as ttk
from tkinter import filedialog, messagebox
from PIL import ImageGrab
import time
import os
from datetime import datetime
import easyocr
import pandas as pd
import numpy as np
import ctypes


class ScreenCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Screen Capture Tool")

        # 버튼 글꼴 스타일 정의
        self.custom_font = ("Helvetica", 14)

        # 스타일 객체 생성 및 색상과 폰트 설정
        self.style = ttk.Style()

        # 노란색 버튼 스타일
        self.style.configure(
            "Yellow.TButton",  # 스타일 이름
            font=self.custom_font,  # 글꼴 설정
            background="#FFD700",  # 노란색 (Golden Yellow)
            foreground="black",    # 글자 색
            borderwidth=2,         # 테두리 두께
        )

        # 빨간색 버튼 스타일
        self.style.configure(
            "Red.TButton",
            font=self.custom_font,
            background="#FF4500",  # 빨간색 (OrangeRed)
            foreground="white",    # 글자 색
            borderwidth=2,
        )

        # 녹색 버튼 스타일
        self.style.configure(
            "Green.TButton",
            font=self.custom_font,
            background="#32CD32",  # 녹색 (LimeGreen)
            foreground="white",
            borderwidth=2,
        )

        # 파란색 버튼 스타일
        self.style.configure(
            "Blue.TButton",
            font=self.custom_font,
            background="#1E90FF",
            foreground="white",
            borderwidth=2,
        )


        # 열 배치 설정 (모든 열의 weight를 동일하게 설정하여 가로 길이를 균등하게 나누기)
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)

        # New와 Load 버튼 생성
        self.new_button = ttk.Button(root, text="New", width=6, command=self.create_new_excel, style="Blue.TButton")
        self.new_button.grid(row=0, column=0, pady=10)

        self.load_button = ttk.Button(root, text="Load", width=6, command=self.load_excel, style="Blue.TButton")
        self.load_button.grid(row=0, column=1, padx=(0, 20), pady=10)

        # 현재 불러온 엑셀 파일명을 표시할 레이블
        self.excel_label = ttk.Label(root, font=("Arial", 13))
        self.excel_label.grid(row=1, column=0, columnspan=2, pady=10)

        self.capture_button = ttk.Button(root, text="Capture", command=self.start_capture, style="Yellow.TButton")
        self.capture_button.grid(row=2, column=0, columnspan=2, pady=10)

        self.text_display = ttk.Text(root, height=1, width=16)
        self.text_display.configure(font=("Arial", 16, "bold"))
        self.text_display.grid(row=3, column=0, columnspan=2, pady=10)


        # 1st Start 버튼 생성 (중앙 배치를 위해 열을 나누어 설정)
        self.start_button = ttk.Button(root, text="1st Start", width=8, command=self.start_timer, style="Green.TButton")
        self.start_button.grid(row=4, column=0, pady=10)

        # 경과 시간 레이블 생성 (1st Start)
        self.elapsed_time_label = ttk.Label(root, text="00:00", font=("Arial", 20), anchor="center")
        self.elapsed_time_label.grid(row=4, column=1, pady=10)

        # 2nd Start 버튼 생성 (중앙 배치를 위해 열을 나누어 설정)
        self.second_start_button = ttk.Button(root, text="2nd Start", width=8, command=self.second_start_timer, style="Green.TButton")
        self.second_start_button.grid(row=5, column=0, pady=10)

        # 체크박스 추가
        self.always_on_top_var = ttk.BooleanVar(value=True)  # 초기값을 True로 설정
        self.always_on_top_checkbox = ttk.Checkbutton(
            root,
            text="Always on Top",
            variable=self.always_on_top_var,
            command=self.toggle_always_on_top
        )
        self.always_on_top_checkbox.grid(row=6, column=0, columnspan=2, pady=(0, 5))

        # 초기 상태를 항상 위로 설정
        self.root.attributes('-topmost', True)

        # 경과 시간 레이블 생성 (2nd Start)
        self.second_elapsed_time_label = ttk.Label(root, text="00:00", font=("Arial", 20), anchor="center")
        self.second_elapsed_time_label.grid(row=5, column=1, pady=10)

        # EasyOCR Reader 초기화
        self.reader = easyocr.Reader(['en'])

        # DPI 스케일링 보정 설정
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()

        # 전체 가상 스크린의 크기 가져오기
        self.screen_width = user32.GetSystemMetrics(78)  # SM_CXVIRTUALSCREEN
        self.screen_height = user32.GetSystemMetrics(79)  # SM_CYVIRTUALSCREEN
        self.screen_x = user32.GetSystemMetrics(76)  # SM_XVIRTUALSCREEN
        self.screen_y = user32.GetSystemMetrics(77)  # SM_YVIRTUALSCREEN

        self.excel_path = None
        self.start_time = None
        self.second_start_time = None
        self.current_id = None
        self.timer_running = False
        self.second_timer_running = False

    def create_new_excel(self):
        # 파일 저장 경로 선택
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            # 날짜 추가하여 파일명 생성
            timestamp = datetime.now().strftime("%Y_%m_%d")
            self.excel_path = f"{file_path[:-5]}_{timestamp}.xlsx"
            # 새로운 엑셀 파일 생성
            df = pd.DataFrame(
                columns=['ID', 'First_Start', 'First_End', 'Second_Start', 'Second_End', 'First_Elapsed_Time',
                         'Second_Elapsed_Time', 'Total_Elapsed_Time'])
            df.to_excel(self.excel_path, index=False)
            self.excel_label.config(text=f"{os.path.basename(self.excel_path)}")

    def load_excel(self):
        # 기존 엑셀 파일 불러오기
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                # 파일이 열려 있는 경우 예외 처리
                with open(file_path, 'a'):  # 파일을 열 수 있는지 확인
                    pass
                self.excel_path = file_path
                self.excel_label.config(text=f"{os.path.basename(self.excel_path)}")
            except PermissionError:
                messagebox.showerror("Error",
                                     "The selected Excel file is currently open. Please close it and try again.")

    def start_capture(self):
        # 창 최소화 후 캡처 시작
        self.root.iconify()
        self.root.after(500, self.create_overlay)

    def create_overlay(self):
        # 전체 가상 화면 오버레이 캔버스 생성
        self.overlay = ttk.Toplevel(self.root)
        self.overlay.geometry(f"{self.screen_width}x{self.screen_height}+{self.screen_x}+{self.screen_y}")
        self.overlay.attributes("-alpha", 0.3)  # 투명도 조절
        self.overlay.configure(bg='gray')
        self.overlay.bind("<ButtonPress-1>", self.start_selection)
        self.overlay.bind("<B1-Motion>", self.update_selection)
        self.overlay.bind("<ButtonRelease-1>", self.end_selection)

        self.canvas = ttk.Canvas(self.overlay, cursor='cross')
        self.canvas.pack(fill=ttk.BOTH, expand=True)

        self.start_x = None
        self.start_y = None
        self.rect_id = None

    def start_selection(self, event):
        # 드래그 시작 지점 저장 (절대 좌표로 변경, 화면 오프셋 고려)
        self.start_x = event.x_root
        self.start_y = event.y_root

    def update_selection(self, event):
        # 드래그하는 동안 직사각형 업데이트
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        end_x = event.x_root
        end_y = event.y_root
        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)
        self.rect_id = self.canvas.create_rectangle(x1 - self.overlay.winfo_rootx(), y1 - self.overlay.winfo_rooty(),
                                                    x2 - self.overlay.winfo_rootx(), y2 - self.overlay.winfo_rooty(),
                                                    outline='green', width=2)

    def end_selection(self, event):
        # 드래그 종료 후 영역 캡처
        end_x = event.x_root
        end_y = event.y_root

        # 실제 화면 좌표 계산
        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)

        # 오버레이 창 제거
        self.overlay.destroy()

        # 지정한 영역을 캡처
        time.sleep(0.1)  # 잠시 대기 후 캡처 수행
        self.root.deiconify()  # 루트 창을 캡처 전에 복원

        # 스크린 캡처 수행
        try:
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
        except Exception as e:
            print(f"Screenshot error: {e}")
            return

        # OCR 처리 및 나머지 코드는 동일
        text_results = self.reader.readtext(np.array(screenshot))
        text = " ".join([result[1] for result in text_results])

        self.text_display.delete(1.0, ttk.END)
        self.text_display.tag_configure("center", justify="center")
        self.text_display.insert(ttk.END, text)
        self.text_display.tag_add("center", "1.0", "end")

        if text.strip():
            self.current_id = text.strip()
            self.save_text_to_excel(self.current_id)

    def save_text_to_excel(self, text):
        if self.excel_path is None:
            return
        try:
            if not os.path.exists(self.excel_path):
                df = pd.DataFrame(columns=['ID', 'First_Start', 'First_End', 'Second_Start', 'Second_End', 'First_Elapsed_Time', 'Second_Elapsed_Time', 'Total_Elapsed_Time'])
            else:
                df = pd.read_excel(self.excel_path)

            if text not in df['ID'].values:
                new_row = {'ID': text, 'First_Start': None, 'First_End': None, 'Second_Start': None, 'Second_End': None, 'First_Elapsed_Time': None, 'Second_Elapsed_Time': None, 'Total_Elapsed_Time': None}
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

            df.to_excel(self.excel_path, index=False)
        except PermissionError:
            messagebox.showerror("Error", "The selected Excel file is currently open. Please close it and try again.")

    def start_timer(self):
        if not self.timer_running:
            self.start_time = datetime.now()
            # 버튼을 Stop 상태로 전환, bootstyle을 Danger로 변경
            self.start_button.config(text="1st Stop", style="Red.TButton")
            self.start_button.config(command=self.stop_timer)
            self.timer_running = True
            self.update_elapsed_time()

    def stop_timer(self):
        if self.timer_running:
            end_time = datetime.now()
            elapsed_time = end_time - self.start_time
            elapsed_minutes, elapsed_seconds = divmod(elapsed_time.total_seconds(), 60)
            elapsed_time_str = f"{int(elapsed_minutes):02}:{int(elapsed_seconds):02}"
            self.elapsed_time_label.config(text=f"{elapsed_time_str}")

            # 엑셀 파일에 Start, End, Elapsed Time 기록
            if self.excel_path and self.current_id:
                try:
                    df = pd.read_excel(self.excel_path)
                    if self.current_id in df['ID'].values:
                        idx = df.index[df['ID'] == self.current_id].tolist()[0]
                        df.at[idx, 'First_Start'] = self.start_time
                        df.at[idx, 'First_End'] = end_time
                        df.at[idx, 'First_Elapsed_Time'] = int(elapsed_time.total_seconds())  # 소수점 버림
                    df.to_excel(self.excel_path, index=False)
                except PermissionError:
                    messagebox.showerror("Error", "The selected Excel file is currently open. Please close it and try again.")

            self.timer_running = False
            # 버튼을 Start 상태로 전환, bootstyle을 Success로 변경
            self.start_button.config(text="1st Start", style="Green.TButton")
            self.start_button.config(command=self.start_timer)

    def second_start_timer(self):
        if not self.second_timer_running:
            self.second_start_time = datetime.now()
            # 버튼을 Stop 상태로 전환, bootstyle을 Danger로 변경
            self.second_start_button.config(text="2nd Stop", style="Red.TButton")
            self.second_start_button.config(command=self.second_stop_timer)
            self.second_timer_running = True
            self.update_second_elapsed_time()

    def second_stop_timer(self):
        if self.second_timer_running:
            end_time = datetime.now()
            elapsed_time = end_time - self.second_start_time
            elapsed_minutes, elapsed_seconds = divmod(elapsed_time.total_seconds(), 60)
            elapsed_time_str = f"{int(elapsed_minutes):02}:{int(elapsed_seconds):02}"
            self.second_elapsed_time_label.config(text=f"{elapsed_time_str}")

            # 엑셀 파일에 Second Start, Second End, Second Elapsed Time 기록
            if self.excel_path and self.current_id:
                try:
                    df = pd.read_excel(self.excel_path)
                    if self.current_id in df['ID'].values:
                        idx = df.index[df['ID'] == self.current_id].tolist()[0]
                        df.at[idx, 'Second_Start'] = self.second_start_time
                        df.at[idx, 'Second_End'] = end_time
                        df.at[idx, 'Second_Elapsed_Time'] = int(elapsed_time.total_seconds())
                        # 총 경과 시간 계산
                        first_elapsed = df.at[idx, 'First_Elapsed_Time']
                        if pd.isna(first_elapsed):
                            first_elapsed = 0
                        total_elapsed = first_elapsed + int(elapsed_time.total_seconds())
                        df.at[idx, 'Total_Elapsed_Time'] = total_elapsed
                    df.to_excel(self.excel_path, index=False)
                except PermissionError:
                    messagebox.showerror("Error", "The selected Excel file is currently open. Please close it and try again.")

            self.second_timer_running = False
            # 버튼을 Start 상태로 전환, bootstyle을 Success로 변경
            self.second_start_button.config(text="2nd Start", style="Green.TButton")
            self.second_start_button.config(command=self.second_start_timer)

    def update_elapsed_time(self):
        if self.timer_running:
            elapsed_time = datetime.now() - self.start_time
            elapsed_minutes, elapsed_seconds = divmod(elapsed_time.total_seconds(), 60)
            elapsed_time_str = f"{int(elapsed_minutes):02}:{int(elapsed_seconds):02}"
            self.elapsed_time_label.config(text=f"{elapsed_time_str}")
            self.root.after(1000, self.update_elapsed_time)

    def update_second_elapsed_time(self):
        if self.second_timer_running:
            elapsed_time = datetime.now() - self.second_start_time
            elapsed_minutes, elapsed_seconds = divmod(elapsed_time.total_seconds(), 60)
            elapsed_time_str = f"{int(elapsed_minutes):02}:{int(elapsed_seconds):02}"
            self.second_elapsed_time_label.config(text=f"{elapsed_time_str}")
            self.root.after(1000, self.update_second_elapsed_time)

    def toggle_always_on_top(self):
        """체크박스 상태에 따라 창을 항상 위에 표시하거나 해제"""
        if self.always_on_top_var.get():
            self.root.attributes('-topmost', True)
        else:
            self.root.attributes('-topmost', False)

if __name__ == "__main__":
    root = ttk.Window(themename="flatly")  # 또는 themename="flatly"
    app = ScreenCaptureApp(root)
    root.mainloop()
