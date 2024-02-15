import PySimpleGUI as sg


sg.theme('SystemDefault')

layout = [sg.table]

#GUI表示実行部分
while True:
    # ウィンドウ表示
    event, values = window.read()



    #クローズボタンの処理
    if event is None:
        print('exit')
        break

window.close()