import openpyxl


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


if __name__ == '__main__':
    combine_same_prefix_column(file_name='combine_same_prefix_column_test.xlsx')
