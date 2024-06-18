import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkinter import messagebox

class SalesFrame(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.pack(fill=tk.BOTH, expand=1)
        self.initUI()

    def initUI(self):
        top_frame = tk.Frame(self, bg="lightgrey")
        top_frame.pack(side=tk.TOP, fill=tk.X)

        add_button = tk.Button(top_frame, text="매출 추가", command=self.open_add_window, font=("Nanum Gothic", 20))
        add_button.pack(side=tk.LEFT, padx=10, pady=10)

        edit_button = tk.Button(top_frame, text="수정", command=self.edit_entry, font=("Nanum Gothic", 20))
        edit_button.pack(side=tk.LEFT, padx=10, pady=10)

        delete_button = tk.Button(top_frame, text="삭제", command=self.delete_entry, font=("Nanum Gothic", 20))
        delete_button.pack(side=tk.LEFT, padx=10, pady=10)

        self.tree = ttk.Treeview(self, columns=('상품코드', '상품명', '수량', '판매가', '판매 시간', '총합'), show='headings')
        self.tree.heading('상품코드', text='상품코드')
        self.tree.heading('상품명', text='상품명')
        self.tree.heading('수량', text='수량')
        self.tree.heading('판매가', text='판매가')
        self.tree.heading('판매 시간', text='판매 시간')
        self.tree.heading('총합', text='총합')

        for col in self.tree['columns']:
            self.tree.column(col, width=100, anchor='center')

        self.tree.pack(fill=tk.BOTH, expand=1)

        self.total_label = tk.Label(self, text="총합: 0 원", font=("Nanum Gothic", 20))
        self.total_label.pack(side=tk.BOTTOM, pady=10)

        self.load_data()

    def load_data(self):
        try:
            df = pd.read_excel('Sales.xlsx', sheet_name='sales')
        except FileNotFoundError:
            print("파일을 찾을 수 없습니다: Sales.xlsx")
            return
        except pd.errors.EmptyDataError:
            print("데이터가 비어있습니다: Sales.xlsx")
            return

        self.tree.delete(*self.tree.get_children())
        total_sales = 0
        for index, row in df.iterrows():
            self.tree.insert("", "end", values=(row['상품코드'], row['상품명'], row['수량'], row['판매가'], row['판매 시간'], row['총합']))
            total_sales += row['총합']
        self.total_label.config(text=f"총합: {total_sales:,} 원")

    def open_add_window(self):
        self.add_window = tk.Toplevel(self)
        self.add_window.title("매출 추가")
        self.add_window.geometry("400x400")

        items_df = pd.read_excel('items.xlsx')

        tk.Label(self.add_window, text="상품 선택").pack(pady=5)
        self.item_var = tk.StringVar(self.add_window)
        self.item_menu = ttk.Combobox(self.add_window, textvariable=self.item_var)
        self.item_menu['values'] = items_df['상품명'].tolist()
        self.item_menu.pack(pady=5)
        self.item_menu.bind("<<ComboboxSelected>>", lambda event: self.fill_item_info(items_df))

        labels = ['상품코드', '상품명', '판매가', '수량']
        self.entries = {}

        for label in labels:
            frame = tk.Frame(self.add_window)
            frame.pack(fill=tk.X)

            lbl = tk.Label(frame, text=label, width=15)
            lbl.pack(side=tk.LEFT, padx=5, pady=5)
            entry = tk.Entry(frame)
            entry.pack(side=tk.LEFT, padx=5, pady=5)
            self.entries[label] = entry

        self.total_label = tk.Label(self.add_window, text="매출 총액: 0 원")
        self.total_label.pack(pady=10)

        add_btn = tk.Button(self.add_window, text="추가", command=self.add_sales_data)
        add_btn.pack(pady=10)

    def fill_item_info(self, items_df):
        selected_item = self.item_var.get()
        item_info = items_df[items_df['상품명'] == selected_item].iloc[0]

        self.entries['상품코드'].delete(0, tk.END)
        self.entries['상품코드'].insert(0, item_info['상품코드'])

        self.entries['상품명'].delete(0, tk.END)
        self.entries['상품명'].insert(0, item_info['상품명'])

        self.entries['판매가'].delete(0, tk.END)
        self.entries['판매가'].insert(0, item_info['매입단가'])  # Assume 판매가 is same as 매입단가

        self.entries['수량'].delete(0, tk.END)
        self.entries['수량'].bind("<KeyRelease>", self.calculate_totals)

    def calculate_totals(self, event):
        try:
            sale_price = int(self.entries['판매가'].get())
            quantity = int(self.entries['수량'].get())
        except ValueError:
            sale_price = quantity = 0

        sale_total = sale_price * quantity
        self.total_label.config(text=f"매출 총액: {sale_total:,} 원")

    def add_sales_data(self):
        data = {label: entry.get() for label, entry in self.entries.items()}
        data['판매 시간'] = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        data['총합'] = int(data['판매가']) * int(data['수량'])

        try:
            df = pd.read_excel('Sales.xlsx', sheet_name='sales')
        except FileNotFoundError:
            print("파일을 찾을 수 없습니다: Sales.xlsx")
            df = pd.DataFrame(columns=['상품코드', '상품명', '수량', '판매가', '판매 시간', '총합'])
        except pd.errors.EmptyDataError:
            print("데이터가 비어있습니다: Sales.xlsx")
            df = pd.DataFrame(columns=['상품코드', '상품명', '수량', '판매가', '판매 시간', '총합'])

        new_data = {
            '상품코드': [int(data['상품코드'])],
            '상품명': [data['상품명']],
            '수량': [int(data['수량'])],
            '판매가': [int(data['판매가'])],
            '판매 시간': [data['판매 시간']],
            '총합': [data['총합']]
        }

        df = pd.concat([df, pd.DataFrame(new_data)], ignore_index=True)
        df.to_excel('Sales.xlsx', sheet_name='sales', index=False)
        self.update_inventory(int(data['상품코드']), int(data['수량']))
        self.load_data()
        self.add_window.destroy()

    def update_inventory(self, product_code, quantity_sold):
        try:
            inventory_df = pd.read_excel('inventory.xlsx', sheet_name='inventory')
        except FileNotFoundError:
            print("파일을 찾을 수 없습니다: inventory.xlsx")
            return
        except pd.errors.EmptyDataError:
            print("데이터가 비어있습니다: inventory.xlsx")
            return

        inventory_df.loc[inventory_df['상품코드'] == product_code, '수량'] -= quantity_sold
        inventory_df.to_excel('inventory.xlsx', sheet_name='inventory', index=False)

    def edit_entry(self):
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("경고", "수정할 항목을 선택하세요.")
            return

        item_values = self.tree.item(selected_item, 'values')
        self.open_add_window()
        for label, value in zip(['상품코드', '상품명', '수량', '판매가'], item_values):
            self.entries[label].insert(0, value)

        add_btn = self.add_window.winfo_children()[-1]
        add_btn.config(text="수정", command=lambda: self.update_data(selected_item))

    def update_data(self, item):
        data = {label: entry.get() for label, entry in self.entries.items()}
        data['판매 시간'] = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        data['총합'] = int(data['판매가']) * int(data['수량'])

        df = pd.read_excel('Sales.xlsx', sheet_name='sales')
        index = self.tree.index(item)

        df.loc[index] = [int(data['상품코드']), data['상품명'], int(data['수량']), int(data['판매가']), data['판매 시간'], data['총합']]
        df.to_excel('Sales.xlsx', sheet_name='sales', index=False)
        self.load_data()
        self.add_window.destroy()

    def delete_entry(self):
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("경고", "삭제할 항목을 선택하세요.")
            return

        result = messagebox.askquestion("확인", "정말 삭제하시겠습니까?", icon='warning')
        if result == 'yes':
            df = pd.read_excel('Sales.xlsx', sheet_name='sales')
            index = self.tree.index(selected_item)
            df = df.drop(index)
            df.to_excel('Sales.xlsx', sheet_name='sales', index=False)
            self.load_data()
