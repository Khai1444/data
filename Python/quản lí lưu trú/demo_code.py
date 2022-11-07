import PySimpleGUI as sg
import pandas as pd

sg.theme("GreenMono")
print("Hello! Chào mừng đến hệ thống nhập liệu thông minh")

# Phần chức năng

layout = [
    [sg.Text("Công cụ quản lí lưu trú", background_color= "Green", text_color="Yellow", justification= "left")],
    [sg.Text('STT',size=(10,1)), sg.Input(key="STT", size=(7,1)), sg.Text('NGÀY ĐẾN',size=(10,1)), sg.Input(key="NGAY_DEN", size=(10,1))],
    [sg.Text('HỌ TÊN',size=(10,1)), sg.Input(key="HO_TEN", size=(50,1))],
    [sg.Text('GIỚI TÍNH', size=(10,1)), sg.Combo(['Nam',  'Nữ'], key='GIOI_TINH', size=(5,2)), sg.Text('NGÀY SINH',size=(10,1)), sg.Input(key="NGAY_SINH", size=(10,1)), sg.Text('CMND/CCCD', size=(11,1)), sg.Input(key='CMND_CCCD', size=(15,1))],
    [sg.Text('ĐĂNG KÍ THƯỜNG TRÚ',size=(20,1)), sg.Input(key="DANG_KI_THUONG_TRU", size=(60,1))],
    [sg.Text('CHỔ Ở HIỆN NAY',size=(20,1)), sg.Input(key="CHO_O_HIEN_NAY", size=(60,1))],
    [sg.Text('GHI CHÚ',size=(20,1)), sg.Multiline(key="GHI_CHU", size=(60,3))],
    [sg.Button("Save", key="Save", button_color='Blue'),sg.Button("Chỉnh sửa", key="Edit", button_color='Brown'), sg.Button("Refresh", key="Refresh", button_color='Black'), sg.Button("Delete", key="Delete", button_color='Violet',), sg.Button("Thoát", key="Exit", button_color= 'Red'), sg.Button("ⓘ Thông tin", key="THONG_TIN", button_color='Green')],

  #  [sg.Text("Copyright © Nguyen Van Khai. All rights reserved 2022", background_color= "Green", text_color="Yellow", justification= "Center")],
    ]
window = sg.Window("Nhập dữ liệu tự động từ Python",layout)

#window.read()

df = pd.read_excel("demo_code.xlsx")
#print(df)

def clear_input():
          for key in values:
                  window[key]("")
          return None

while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Exit":
                break

        if event == "Save":
                df = df.append(values, ignore_index= True)
                df.to_excel("demo_code.xlsx", index= False)
                sg.popup("Save thành công! Vui lòng nhấn (Ok) để nhập dữ liệu mới.")
                clear_input()

        if event == "Edit":
                if values["STT"] == "":
                        layout = [
                                [sg.Text("Vui lòng nhập (STT) để thay đổi thông tin.")],
                                [sg.Text('THÔNG TIN THAY ĐỔI', size=(18,1)), sg.InputText(key='THONG_TIN_THAY_DOI')],
                                [sg.Submit(), sg.Cancel()],
                        ]
                        window1 = sg.Window("Thông tin thay đổi", layout)
                        event1, value_id = window1.read()
                        THONG_TIN_THAY_DOI = int(value_id["THONG_TIN_THAY_DOI"])
                        indexa = df.loc[df['STT'] == THONG_TIN_THAY_DOI].index.to_list()[0]
                        df_id = df.iloc[indexa]
                        dicta = df_id.to_dict()
                        window1.close()
                        for key, value in dicta.items():
                                window[key].update(value)
                else:
                        indexa = df.loc[df['STT'] == THONG_TIN_THAY_DOI].index.to_list()[0]
                        header_list = list(df.columns.values)

                        for key in header_list:
                                df.loc[indexa, key] = values[key]

                        df.to_excel("demo_code.xlsx", index= False)
                        sg.popup("Chỉnh sửa thành công!")
                        clear_input()

        if event == "Delete":
                if values["STT"] == "":
                        layout = [
                                [sg.Text("Vui lòng nhập (STT) để xóa.")],
                                [sg.Text('THÔNG TIN THAY ĐỔI', size=(18,1)), sg.InputText(key='THONG_TIN_THAY_DOI')],
                                [sg.Submit(), sg.Cancel()],
                        ]
                        window1 = sg.Window("Xóa thông tin", layout)
                        event1, value_id = window1.read()
                        THONG_TIN_THAY_DOI = int(value_id["THONG_TIN_THAY_DOI"])
                        indexa = df.loc[df['STT'] == THONG_TIN_THAY_DOI].index.to_list()[0]
                        df_id = df.iloc[indexa]
                        dicta = df_id.to_dict()
                        window1.close()
                        for key, value in dicta.items():
                                window[key].update(value)
                else:
                        indexa = df.loc[df['STT'] == THONG_TIN_THAY_DOI].index.to_list()[0]
                        delete_df = df.drop(indexa)
                        delete_df.to_excel("demo_code.xlsx", index= False)
                        sg.popup("Xóa thành công!")
                        clear_input()

        if event == "Refresh":
                  clear_input()

        if event == "THONG_TIN":
                        sg.popup("Copyright © Nguyen Van Khai. All rights reserved 2022")
                        
