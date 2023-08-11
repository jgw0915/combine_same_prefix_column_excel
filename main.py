import threading
from datetime import datetime
from operator import *
import customtkinter
import openpyxl

import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog


def combine_same_prefix_column(file_name: str, data_sheet_name: str, company_sheet_name: str, transform_company_sheet_name:str, delete_sheet_name:str):
    try:
        workbook = openpyxl.load_workbook(file_name, data_only=True)
    except openpyxl.utils.exceptions.InvalidFileException:
        status_message = '無法取得Excel檔'
        return False, status_message
    except FileNotFoundError:
        status_message = '找不到Excel檔'
        return False, status_message
    try:
        data_sheet = workbook[data_sheet_name]
    except KeyError:
        status_message = '找不到資料列工作表'
        return False, status_message
    try:
        head_quarter_sheet = workbook[company_sheet_name]
    except KeyError:
        status_message = '找不到總公司工作表'
        return False, status_message
    try:
        transform_company_sheet = workbook[transform_company_sheet_name]
    except KeyError:
        status_message = '找不到公司名稱轉換工作表'
        return False, status_message
    try:
        delete_sheet = workbook[delete_sheet_name]
    except KeyError:
        status_message = '找不到沖銷公司工作表'
        return False, status_message

    new_sheet_title = [head_quarter_sheet['a1'].value]

    # get總公司名稱
    head_quarter_dict = {}
    for row in range(1, len(head_quarter_sheet['A'])+1):
        if row != 1:
            head_quarter_dict[str(head_quarter_sheet[f'a{row}'].value).rstrip(' ')] = {}
    print(head_quarter_dict.keys())

    # 轉換公司名稱
    origin_name_col = transform_company_sheet['a']
    transform_name_col = transform_company_sheet['b']
    transform_name_dict = {}
    for row_num in range(1,len(origin_name_col)) :
        transform_name_dict[str(origin_name_col[row_num].value).rstrip(' ')] = str(transform_name_col[row_num].value).rstrip(' ')
    print(transform_name_dict)
    data_company_col = data_sheet['a']
    for key,value in transform_name_dict.items():
        for row_num in range(1,len(data_company_col)):
            cell = str(data_company_col[row_num].value).rstrip()
            if cell == key:
                print(data_company_col[row_num].value + ',' + value)
                data_company_col[row_num].value = value

    # 沖銷公司名稱
    delete_company_col = delete_sheet['a']
    delete_company_list=[]
    for row in range(2,len(delete_company_col)):
        delete_company_list.append(str(delete_company_col[row].value))

    # 合併資料
    length = len(data_sheet['A'])+1
    for prefix in head_quarter_dict.keys():
        if prefix in delete_company_list:
            continue
        for row_num in range(1, length):
            row = data_sheet[f'{row_num}']
            title = data_sheet[f'{1}']
            for col_num in range(len(row)):
                cell = row[col_num].value
                # 標籤欄
                if row_num == 1:
                    if col_num != 0:
                        head_quarter_dict.get(prefix)[cell] = 0
                        if type(cell) == datetime:
                            cell = cell.strftime("%Y/%m")
                        if str(cell) not in new_sheet_title:
                            new_sheet_title.append(str(cell))
                # 資料欄
                else:
                    if row[0].value.startswith(prefix) and col_num != 0:
                        try:
                            print(row[0].value+',{}',format(cell))
                            head_quarter_dict.get(prefix)[title[col_num].value] += float(cell)
                        except TypeError:
                            head_quarter_dict.get(prefix)[title[col_num].value] += 0
                        except ValueError:
                            head_quarter_dict.get(prefix)[title[col_num].value] += 0
    print(head_quarter_dict)

    # 輸出資料
    export_sheet = workbook.create_sheet('輸出')
    export_sheet.append(new_sheet_title)
    sorted_data = []
    final_data = []
    # 排序資料
    for sorted_row in sorted(head_quarter_dict.items(), key=lambda x: x[1][new_sheet_title[1]],reverse=True):
        sorted_data.append(sorted_row)
    for tuple in sorted_data:
        final_data.append(tuple[0])
        for key, value in tuple[1].items():
            final_data.append(value)
        export_sheet.append(final_data)
        final_data.clear()
    try:
        status_message = '已完成 Excel 資料合併'
        workbook.save(file_name)
        return True, status_message
    except PermissionError:
        status_message = '無法儲存Excel檔(Excel檔可能已開啟)'
        return False, status_message


class SuccessDialog(simpledialog.Dialog):
    def __init__(self, parent, title):
        super().__init__(parent, title)
        self.textBox = None

    def body(self, frame):
        # print(type(frame)) # tkinter.Frame
        self.textBox = customtkinter.CTkLabel(frame, text='已完成 Excel 資料合併')
        self.textBox.pack(padx=10, pady=10)

        return frame


class ErrorDialog(simpledialog.Dialog):
    def __init__(self, parent, title):
        super().__init__(parent, title)
        self.textBox = None


    def body(self, frame):
        # print(type(frame)) # tkinter.Frame
        self.textBox = customtkinter.CTkLabel(frame, text='出現錯誤')
        self.textBox.pack(padx=10, pady=10)

        return frame


class CustomTKinterApp(customtkinter.CTk):
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 合併子公司 UI")
        self.root.iconbitmap('icon2.ico')

        self.data_sheet_name = tk.StringVar()
        self.company_reference_sheet_name = tk.StringVar()
        self.transform_sheet_name = tk.StringVar()
        self.delete_sheet_name = tk.StringVar()
        self.file_path = tk.StringVar()
        self.status_message = tk.StringVar()
        self.status_color = tk.StringVar()

        # Element
        self.sub_frame = None
        self.file_path_textbox = None
        self.browse_button = None
        self.data_sheet_name_label = None
        self.data_sheet_name_textbox = None
        self.company_reference_sheet_name_label = None
        self.company_reference_sheet_name_textbox = None
        self.submit_button = None
        self.message_label = None


        self.create_widgets()

    def create_widgets(self):

        self.frame = customtkinter.CTkFrame(self.root, width=700, height=400)
        self.frame.pack(padx=10, pady=10)

        self.sub_frame = customtkinter.CTkFrame(self.frame, width=625, height=250)
        self.sub_frame.pack(padx=10, pady=10)

        # Text box to display file path
        self.file_path_textbox = customtkinter.CTkEntry(self.sub_frame, textvariable=self.file_path, width=500)
        self.file_path_textbox.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.file_path_textbox.place(x=0, y=20)

        # Browse button
        self.browse_button = customtkinter.CTkButton(self.sub_frame, text="選取檔案", command=self.browse_file, width=100)
        self.browse_button.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.browse_button.place(x=510, y=20)

        # Label for enter data sheet name
        self.data_sheet_name_label = customtkinter.CTkLabel(self.sub_frame, text='請輸入資料列工作表名稱:')
        self.data_sheet_name_label.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.data_sheet_name_label.place(x=0, y=70)
        # Text box to enter data sheet name
        self.data_sheet_name_textbox = customtkinter.CTkEntry(self.sub_frame,
                                                              textvariable=self.data_sheet_name, width=250)
        self.data_sheet_name_textbox.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.data_sheet_name_textbox.place(x=150, y=70)

        # Label for enter company reference sheet name
        self.company_reference_sheet_name_label = customtkinter.CTkLabel(self.sub_frame, text='請輸入預合併公司工作表名稱:')
        self.company_reference_sheet_name_label.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.company_reference_sheet_name_label.place(x=0, y=120)

        # Text box to enter company reference sheet name
        self.company_reference_sheet_name_textbox = customtkinter.CTkEntry(
            self.sub_frame, textvariable=self.company_reference_sheet_name, width=250)
        self.company_reference_sheet_name_textbox.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.company_reference_sheet_name_textbox.place(x=180, y=120)

        self.transform_sheet_name_label = customtkinter.CTkLabel(self.sub_frame, text='請輸入轉換公司工作表名稱:')
        self.transform_sheet_name_label.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.transform_sheet_name_label.place(x=0, y=170)

        # Text box to enter company reference sheet name
        self.transform_sheet_name_textbox = customtkinter.CTkEntry(
            self.sub_frame, textvariable=self.transform_sheet_name, width=250)
        self.transform_sheet_name_textbox.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.transform_sheet_name_textbox.place(x=165, y=170)

        self.delete_sheet_name_label = customtkinter.CTkLabel(self.sub_frame, text='請輸入沖銷公司工作表名稱:')
        self.delete_sheet_name_label.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.delete_sheet_name_label.place(x=0, y=220)

        # Text box to enter company reference sheet name
        self.delete_sheet_name_textbox = customtkinter.CTkEntry(
            self.sub_frame, textvariable=self.delete_sheet_name, width=250)
        self.delete_sheet_name_textbox.place(relx=0, anchor='w')  # move the text to the left side of frame
        self.delete_sheet_name_textbox.place(x=165, y=220)

        # Submit button
        self.submit_button = customtkinter.CTkButton(self.frame, text="開始資料合併", command=self.start_thread, width=100)
        self.submit_button.pack(padx=20, pady=1, side=tk.RIGHT)

        self.message_label = customtkinter.CTkLabel(
            self.frame, textvariable=self.status_message, text_color=self.check_message_label_text_color())
        self.message_label.place(relx=0, anchor='e')  # move the text to the left side of frame
        self.message_label.place(x=520, y=285)


    def start_thread(self):
        try:
            thread = threading.Thread(target=self.submit_file)
            thread.start()
        except RuntimeError as e:
            print(e)
    def check_message_label_text_color(self):
        if self.status_message.get() == '資料處理中...':
            self.status_color.set('black')
        elif self.status_message.get() == '已完成 Excel 資料合併':
            self.status_color.set('green')
        else:
            self.status_color.set('dark red')
        return self.status_color.get()

    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_path.set(file_path)


    def submit_file(self):
        file_name = self.file_path.get()
        data_sheet_name = self.data_sheet_name_textbox.get()
        company_sheet_name = self.company_reference_sheet_name_textbox.get()
        transform_sheet_name = self.transform_sheet_name.get()
        delete_sheet_name = self.delete_sheet_name.get()
        if file_name:
            self.set_message_label(status_message='資料處理中...')
            function_success, status_message = combine_same_prefix_column(
                file_name=file_name, data_sheet_name=data_sheet_name,
                company_sheet_name=company_sheet_name,
                transform_company_sheet_name=transform_sheet_name,
                delete_sheet_name=delete_sheet_name)
            # button = customtkinter.CTkButton(self.root, text='Ok',width=50)
            new_root = customtkinter.CTk()
            new_root.withdraw()
            if function_success:
                SuccessDialog(new_root, title='Success')
                self.set_message_label(status_message)
            else:
                ErrorDialog(new_root, title='Failed')
                self.set_message_label(status_message)
            new_root.destroy()


            # You can perform further actions with the file path here

    def set_message_label(self, status_message):
        self.status_message.set(status_message)
        self.message_label.destroy()
        self.message_label = customtkinter.CTkLabel(self.frame, textvariable=self.status_message,
                                                    text_color=self.check_message_label_text_color())
        self.message_label.place(relx=0, anchor='e')  # move the text to the left side of frame
        self.message_label.place(x=520, y=285)

        self.submit_button.destroy()
        self.submit_button = customtkinter.CTkButton(self.frame, text="開始資料合併",
                                                     command=threading.Thread(target=self.submit_file).start,
                                                     width=100)
        self.submit_button.pack(padx=20, pady=1, side=tk.RIGHT)


if __name__ == '__main__':
    root = tk.Tk()
    app = CustomTKinterApp(root)
    root.mainloop()
