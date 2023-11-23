import threading
from datetime import datetime
import customtkinter
import openpyxl

from tkinter import ttk
import tkinter as tk
from tkinter import filedialog


def combine_same_prefix_column(file_name: str, data_sheet_name: str, transform_company_sheet_name:str, delete_sheet_name:str):
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
        transform_company_sheet = workbook[transform_company_sheet_name]
    except KeyError:
        status_message = '找不到公司名稱轉換工作表'
        return False, status_message
    try:
        delete_sheet = workbook[delete_sheet_name]
    except KeyError:
        status_message = '找不到沖銷公司工作表'
        return False, status_message

    new_sheet_title = []
    head_quarter_dict = {}

    # 取得標頭
    for col_num in range(len(data_sheet['1'])):
        cell = data_sheet['1'][col_num].value
        if type(cell) == datetime:
            cell = cell.strftime("%Y/%m")
        new_sheet_title.append(str(cell).rstrip(' '))
    print(new_sheet_title)

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
    for row in range(1,len(delete_company_col)):
        delete_company_list.append(str(delete_company_col[row].value).rstrip(' '))
    print(delete_company_list)

    # 合併資料
    row_length = len(data_sheet['A'])+1
    col_length = len(data_sheet['1'])
    for row_num in range(2,row_length):
        for col_num in range (1,col_length):
            row = data_sheet[f'{row_num}']
            company_name = str(row[0].value).rstrip(' ')
            if company_name in delete_company_list:
                continue
            if company_name not in head_quarter_dict.keys():
                head_quarter_dict[company_name] = {}
                for i in range(1,col_length):
                    head_quarter_dict.get(company_name)[new_sheet_title[i]] = 0
            cell = row[col_num].value
            try:
                print(company_name + f',{cell}')
                head_quarter_dict.get(company_name)[new_sheet_title[col_num]] += float(cell)
            except TypeError:
                head_quarter_dict.get(company_name)[new_sheet_title[col_num]] += 0
            except ValueError:
                head_quarter_dict.get(company_name)[new_sheet_title[col_num]] += 0


    # 輸出資料
    export_sheet = workbook.create_sheet('輸出')
    export_sheet.append(new_sheet_title)
    sorted_data = []
    final_data = []
    # 排序資料
    try:
        for sorted_row in sorted(head_quarter_dict.items(), key=lambda x: x[1][new_sheet_title[1]],reverse=True):
            sorted_data.append(sorted_row)
    except KeyError as e:
        return False,'排序時發生錯誤，請重新確認各工作表名稱和格式是否正確'
    for tuple in sorted_data:
        final_data.append(tuple[0])
        for key, value in tuple[1].items():
            final_data.append(value)
        print(final_data)
        export_sheet.append(final_data)
        final_data.clear()
    try:
        status_message = '已完成 Excel 資料合併'
        workbook.save(file_name)
        return True, status_message
    except PermissionError:
        status_message = '無法儲存Excel檔(Excel檔可能已開啟)'
        return False, status_message


class CustomTKinterApp(tk.Tk):
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 合併子公司 UI")
        self.root.iconbitmap('icon2.ico')
        self.style = ttk.Style(root)
        self.root.tk.call("source", "forest-light.tcl")
        self.root.tk.call("source", "forest-dark.tcl")
        self.style.theme_use("forest-light")

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
        self.submit_button = None
        self.message_label = None

        customtkinter.set_appearance_mode("light")

        self.create_widgets()

    def create_widgets(self):


        self.frame = ttk.Frame(self.root)
        self.frame.pack(padx=5, pady=5)

        self.sub_frame = ttk.Frame(self.frame)
        self.sub_frame.grid(row=0,padx=10,pady=10)

        self.sub_frame1 = ttk.Frame(self.sub_frame)
        self.sub_frame1.grid(row=0,sticky="w")

        self.sub_frame2 = ttk.Frame(self.sub_frame)
        self.sub_frame2.grid(row=1,sticky="w")

        self.sub_frame3 = ttk.Frame(self.sub_frame)
        self.sub_frame3.grid(row=2, sticky="E")

        # Text box to display file path
        self.file_path_textbox = ttk.Entry(self.sub_frame1, textvariable=self.file_path,width=82)
        self.file_path_textbox.grid(row=0,column=0,pady=5,padx=5,sticky='W')

        # Browse button
        self.browse_button = ttk.Button(self.sub_frame1, text="選取檔案", command=self.browse_file)
        self.browse_button.grid(row=0,column=1,pady=5,padx=5,sticky='W')

        # Label for enter data sheet name
        self.data_sheet_name_label = ttk.Label(self.sub_frame2, text='請輸入資料列工作表名稱:',width=28)
        self.data_sheet_name_label.grid(row=0,column=0,pady=5,sticky='W')

        # Text box to enter data sheet name
        self.data_sheet_name_textbox = ttk.Entry(self.sub_frame2,
                                                              textvariable=self.data_sheet_name,width=70)
        self.data_sheet_name_textbox.grid(row=0,column=1,pady=5,sticky='W')

        self.transform_sheet_name_label = ttk.Label(self.sub_frame2, text='請輸入轉換公司工作表名稱:')
        self.transform_sheet_name_label.grid(row=1,column=0,pady=5,sticky='W',columnspan=2)

        # Text box to enter company reference sheet name
        self.transform_sheet_name_textbox = ttk.Entry(
            self.sub_frame2, textvariable=self.transform_sheet_name, width=70)
        self.transform_sheet_name_textbox.grid(row=1,column=1,pady=5,sticky='W',columnspan=2)

        self.delete_sheet_name_label = ttk.Label(self.sub_frame2, text='請輸入沖銷公司工作表名稱:')
        self.delete_sheet_name_label.grid(row=2,column=0,pady=5,sticky='W',columnspan=2)

        # Text box to enter company reference sheet name
        self.delete_sheet_name_textbox = ttk.Entry(
            self.sub_frame2, textvariable=self.delete_sheet_name, width=70)
        self.delete_sheet_name_textbox.grid(row=2,column=1,pady=5,sticky='W',columnspan=2)

        # Submit button
        self.submit_button = ttk.Button(self.sub_frame3, text="開始資料合併", command=self.start_thread)
        # self.submit_button.pack(padx=20, pady=1, side=tk.RIGHT)
        self.submit_button.grid(row=0, column=1, padx=2,pady=5, sticky='E')

        self.message_label = ttk.Label(
            self.sub_frame3, textvariable=self.status_message)
        self.message_label.grid(row=0,column=0,pady=5,sticky='E')

        # text_color = self.check_message_label_text_color()


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
        transform_sheet_name = self.transform_sheet_name.get()
        delete_sheet_name = self.delete_sheet_name.get()
        if file_name:
            self.set_message_label(status_message='資料處理中...')
            function_success, status_message = combine_same_prefix_column(
                file_name=file_name, data_sheet_name=data_sheet_name,
                transform_company_sheet_name=transform_sheet_name,
                delete_sheet_name=delete_sheet_name)
            if function_success:
                self.set_message_label(status_message)
            else:
                self.set_message_label(status_message)


            # You can perform further actions with the file path here

    def set_message_label(self, status_message):
        self.status_message.set(status_message)
        self.message_label.destroy()
        # Submit button
        self.submit_button = ttk.Button(self.sub_frame3, text="開始資料合併", command=self.start_thread)
        self.submit_button.grid(row=0, column=1, padx=2, pady=5, sticky='E')

        self.message_label = ttk.Label(
            self.sub_frame3, textvariable=self.status_message, foreground=self.check_message_label_text_color())
        self.message_label.grid(row=0, column=0, pady=5, sticky='E')


if __name__ == '__main__':
    root = tk.Tk()
    app = CustomTKinterApp(root)
    root.mainloop()
