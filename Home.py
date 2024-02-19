import PySimpleGUI as sg
import os

os.chdir('C:\\Users\\wiiue\\JOEU-FM')


#先程確認して決めたテーマカラーをsg.themeで設定
sg.theme('SystemDefault')

layout = [

[
    sg.Text("明屋書店Excelデータ"),
    sg.InputText('ファイルを選択', key='-INPUTTEXT-', enable_events=True,size=(41,1)), 
    sg.FileBrowse(button_text='選択', font=('メイリオ',8), size=(5,1), key="-FILENAME-")
],

[
    sg.Text("ランキングフォーマット"),
    sg.InputText('ファイルを選択', key='-INPUTTEXT-', enable_events=True,size=(37,1)), 
    sg.FileBrowse(button_text='選択', font=('メイリオ',8), size=(5,1), key="-FILENAME-")
],

[
    sg.Text("楽曲データCSV"),
    sg.InputText('ファイルを選択', key='-INPUTTEXT-', enable_events=True,), 
    sg.FileBrowse(button_text='選択', font=('メイリオ',8), size=(5,1), key="-FILENAME-")
],

[   sg.Button('今週のデータを生成する',size=(30,3),key='今週データ生成')],

[   sg.Button('先週のデータを生成する',size=(30,3),key='先週データ生成')],

[   sg.Button('任意の週のデータを生成する',size=(30,3),key='任意週生成'),sg.Button('管理者用',size=(10,3),key='管理者'),sg.Button('ランキング修正',size=(12,3),key='ランキング修正')]


]


window = sg.Window('FMベストヒットランキング自動生成システム', layout, resizable=True)


if os.path.exists('test.db') == True:
   os.remove('test.db')

#GUI表示実行部分
while True:
    # ウィンドウ表示
    event, values = window.read()

    if event == '今週データ生成':
       import CreateDB
       import GetData
       GetData.GetThisWeekRank() 
       import ViewDeta
       import CreateExcel
       CreateExcel.MajicalExcel(GetData.GetThisWeekDate())
       import WriteCSV
       WriteCSV.WriteCSV(GetData.GetThisWeekDate())
    if event == '先週データ生成':
       import CreateDB
       import GetData
       GetData.GetLastWeekRank() 
       import ViewDeta
       import CreateExcel
       CreateExcel.MajicalExcel(GetData.GetThisWeekDate())
    if event == '任意週生成':
       break
    if event == '管理者':
       import CreateDB
       import AdminUser
    if event == 'ランキング修正':
       break


    #クローズボタンの処理
    if event is None:
      # print('exit')
        break

window.close()