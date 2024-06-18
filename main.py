import tkinter as tk
from tkinter import ttk
from tkinter import PhotoImage
import matplotlib.pyplot as plt
from  matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import os
import shutil
from datetime import datetime
from purchase import PurchaseFrame
from sales import SalesFrame
from backup import backup_excel_files
from control import update_inventory_purchase

class MainApplication(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.initUI()

    def initUI(self):
        self.parent.title("엑셀 데이터로 라인 차트 만들기")
        self.pack(fill=tk.BOTH, expand=1)
        self.create_frames()
        self.create_widgets()
        self.show_main_chart()

    def create_frames(self):
        self.top_frame = tk.Frame(self, height=50, bg="white")
        self.top_frame.pack(side=tk.TOP, fill=tk.X)

        self.middle_frame = tk.Frame(self, height=300, bg="white")
        self.middle_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.bottom_frame = tk.Frame(self, height=300, bg="lightgrey")
        self.bottom_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    def create_widgets(self):
        script_path = os.path.dirname(__file__)
        img_path = os.path.join(script_path, "images", "button_image.png")
        self.img = PhotoImage(file=img_path).subsample(2)
        
        self.button = tk.Button(self.top_frame, image=self.img, command=self.main_button_clicked, bd=0)
        self.button.pack(side=tk.LEFT, padx=10, pady=10)

        purchase_button = tk.Button(self.top_frame, text="매입 관리", command=self.purchase_button_clicked, font=("Nanum Gothic", 20), bd=0)
        purchase_button.pack(side=tk.LEFT, padx=10, pady=10)

        sales_button = tk.Button(self.top_frame, text="매출 관리", command=self.sales_button_clicked, font=("Nanum Gothic", 20), bd=0)
        sales_button.pack(side=tk.LEFT, padx=10, pady=10)

        setting_button = tk.Button(self.top_frame, text="Setting", command=self.setting_button_clicked, font=("Nanum Gothic", 20), bd=0)
        setting_button.pack(side=tk.LEFT, padx=10, pady=10)
      


    def main_button_clicked(self):
        self.clear_frame(self.middle_frame)
        self.show_main_chart()



    def show_main_chart(self):
        
        try:
            df = pd.read_excel('ps.xlsx', sheet_name='ps')
        except FileNotFoundError:
            print("파일을 찾을 수 없습니다: ps.xlsx")
            return
        except pd.errors.EmptyDataError:
            print("데이터가 비어있습니다: ps.xlsx")
            return

        plt.rcParams['font.family'] = 'Nanum Gothic'
        plt.rcParams['axes.unicode_minus'] = False

        # 데이터 전처리 및 차트 생성
        sales_data = df.loc[0, '1월':'12월']
        stock_data = df.loc[1, '1월':'12월']

        fig, ax = plt.subplots(figsize=(6, 3))
        ax.plot(sales_data.index, sales_data.values, marker="o", label='매출')
        ax.plot(stock_data.index, stock_data.values, marker="o", label='재고')
        ax.set_title("매출 차트")
        ax.legend()

        canvas = FigureCanvasTkAgg(fig, master=self.middle_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

    def update_inventory_and_ps(self):
        update_inventory_purchase(code, value)

    def purchase_button_clicked(self):
        self.clear_frame(self.middle_frame)
        PurchaseFrame(self.middle_frame)

    def sales_button_clicked(self):
        self.clear_frame(self.middle_frame)
        SalesFrame(self.middle_frame)

    def setting_button_clicked(self):
        self.clear_frame(self.middle_frame)
        # TODO: Add settings functionality

    def clear_frame(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

    def schedule_backup(self):
        backup_excel_files()
        self.parent.after(3600000, self.schedule_backup)

def main():
    root = tk.Tk()
    root.title("P.S ERP")
    root.geometry("1000x600")
    root.resizable(True, True)
    app = MainApplication(root)
    app.schedule_backup()
    root.protocol("WM_DELETE_WINDOW", lambda: on_closing(root))
    root.mainloop()

def on_closing(root):
    backup_excel_files()
    root.destroy()

if __name__ == '__main__':
    main()
