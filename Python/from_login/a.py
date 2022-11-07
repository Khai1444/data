import PySimpleGUI as sg
import pandas as pd
df = pd.read_excel("data_py.xlsx")
table_data = df.values.tolist()
table_header = df.columns.values.tolist()
sg.theme("GreenMono")
print("Hello! Chào mừng đến hệ thống nhập liệu thông minh")

# Phần chức năng

layout = [
    [sg.Text("Công cụ nhập liệu nhanh", background_color= "Green", text_color="Yellow", justification= "left")],
    [sg.Text('Emp ID',size=(10,1)), sg.Input(key="ID"), sg.Text('Department', size=(11,1)), sg.Combo(['LT', 'PLD', 'BT', 'PTT', '_'], key='Department', size=(58,5))],
    [sg.Text('Emp Name',size=(10,1)), sg.Input(key="Name"), sg.Text('Thành Phố', size=(11,1)), sg.Input(key='City', size=(60,1))],
    [sg.Text('Giới Tính', size=(10,1)), sg.Radio("Nam", "g", True, key="g1"), sg.Radio("Nữ", "g", key="g2")],
    [sg.Text('Đoàn Viên', size=(12,1)), sg.Radio("Có", "d", True, key="d1"), sg.Radio("Không", "d", key="d2")],
    [sg.Button("Save", key="Save", button_color='Blue'),sg.Button("Edit", key="Edit"), sg.Button("Refresh", key="Refresh", button_color='Black'), sg.Button("Delete", key="Delete", button_color='Yellow'), sg.Button("Exit", key="Exit", button_color= 'Red')],
    [sg.Table(values= table_data,
         headings= table_header,
         key="Table",
         row_height= 40,
         justification='center',
         expand_x= True,
         expand_y= True,
         )],
    [sg.Text("Copyright © Nguyen Van Khai. All rights reserved 2022", background_color= "Green", text_color="Yellow", justification= "Center")],
    ]
window = sg.Window("Nhập dữ liệu tự động từ Python",layout)


while True:
    event,values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Save":
        values["Giới Tính"] = "Nam" if values["g1"] else "Nữ"
        values["Đoàn Viên"] = "Có" if values["d1"] else "Không"
        del values["g1"]
        del values["g2"]
        del values["d1"]
        del values["d2"]
        del values["Table"]


                

        print(values)
        df = df.append(values, ignore_index = True)
        df.to_excel("data_py.xlsx", index = False)
        val = df.values.tolist()
        window["Table"].update(values = val)
    



                        





                        






