import time

import customtkinter
import openpyxl

import tkinter as tk
from tkinter import filedialog



def combine_same_prefix_column(file_name: str):
    workbook = openpyxl.load_workbook(file_name,data_only=True)
    data_sheet = workbook['test_data']
    head_quarter_sheet = workbook['test_總公司']
    new_sheet_title = ['總公司']

    # get總公司名稱
    head_quarter_dict ={}
    for item in head_quarter_sheet['A']:
        if item.value !='總公司':
            head_quarter_dict[item.value] = {}
    print(head_quarter_dict.keys())

    # 合併資料
    length = len(data_sheet['A'])+1
    for prefix in head_quarter_dict.keys():
        for row_num in range(1,length):
            row = data_sheet[f'{row_num}']
            title = data_sheet[f'{1}']
            for col_num in range(len(row)):
                cell = row[col_num].value
                # 標籤欄
                if row_num ==1 :
                    if col_num != 0:
                        head_quarter_dict.get(prefix)[cell] = 0
                        if cell not in new_sheet_title:
                            new_sheet_title.append(cell)
                # 資料欄
                else:
                    if row[0].value.startswith(prefix) and col_num != 0 :
                        head_quarter_dict.get(prefix)[title[col_num].value] += cell
    print(head_quarter_dict)

    # 輸出資料
    export_sheet = workbook.create_sheet('輸出')
    export_sheet.append(new_sheet_title)
    final_data=[]
    for company,subdict in head_quarter_dict.items():
        final_data.append(company)
        for key,value in subdict.items():
            final_data.append(value)
        export_sheet.append(final_data)
        final_data.clear()

    workbook.save(file_name)

class CustomTKinterApp(customtkinter.CTk):
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 合併子公司 UI")

        self.file_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):

        self.frame = customtkinter.CTkFrame(self.root , width= 600 , height= 400)
        self.frame.pack(padx=10, pady = 10)

        self.sub_frame = customtkinter.CTkFrame(self.frame, width= 575 , height= 150 )
        self.sub_frame.grid(row= 0,column =0,columnspan= 2)
        self.sub_frame.pack(padx = 10, pady = 10)

        # Text box to display file path
        self.file_path_textbox = customtkinter.CTkEntry(self.sub_frame, textvariable=self.file_path, width=500)
        self.file_path_textbox.grid(row= 0, column = 1,columnspan= 2)
        self.file_path_textbox.pack(padx=10, pady=10)

        # Browse button
        self.browse_button = customtkinter.CTkButton(self.sub_frame, text="瀏覽檔案", command=self.browse_file)
        self.browse_button.pack(padx=10, pady=1,side= tk.RIGHT)

        # Submit button
        self.submit_button = customtkinter.CTkButton(self.frame, text="開始合併", command=self.submit_file)
        self.submit_button.pack(padx=20, pady=1,side = tk.RIGHT)

    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_path.set(file_path)

    def submit_file(self):
        file_name = self.file_path.get()
        if file_name:
            combine_same_prefix_column(file_name=file_name)
            # You can perform further actions with the file path here



if __name__ == '__main__':
    root = tk.Tk()
    app = CustomTKinterApp(root)
    root.mainloop()
