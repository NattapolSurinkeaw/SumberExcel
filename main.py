from tkinter import *
from tkinter import filedialog
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

root = Tk()
root.title("Python Sumber")

label1 = Label(text="Python Sumber").pack()

def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        label_text = "ไฟล์ที่ถูกเลือก : " + file_path
        label2 = Label(text=label_text)
        label2.pack()
        print("ไฟล์ที่ถูกเลือก:", file_path)
        process_excel_file(file_path)


def process_excel_file(file_path):
    df = pd.read_excel(file_path)

    df['sum'] = df['number'].astype(str).apply(lambda x: sum(map(int, x)))
    print(df)
    save_excel_file(df)
    return df

def save_excel_file(df):
    print(df)

    wb = Workbook()
    ws = wb.active

    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        wb.save(save_path)
        print(f"บันทึกไฟล์ Excel ที่: {save_path}")

button = Button(root, text="เลือกไฟล์ Excel", command=open_file_dialog)
button.pack()


root.geometry("400x200")
root.mainloop()
