import PySimpleGUI as sg
import os
import sys

# os.chdir('C:\\Users\\wiiue\\JOEU-FM\\')
os.chdir(os.path.dirname(sys.argv[0]))


#先程確認して決めたテーマカラーをsg.themeで設定
sg.theme('SystemDefault')

layout = [

[
    sg.Text("明屋書店Excelデータ"),
    sg.InputText('ファイルを選択', key='-HaruyaExcel-', enable_events=True,size=(41,1)), 
    sg.FileBrowse(button_text='選択', font=('メイリオ',8), size=(5,1), key="-HaruyaExcel-")
],

[   sg.Button('今週のデータを生成する',size=(30,3),key='今週データ生成')],

[   sg.Button('先週のデータを生成する',size=(30,3),key='先週データ生成')],

[   sg.Button('任意の週のデータを生成する',size=(30,3),key='任意週生成'),
    sg.Button('管理者用',size=(10,3),key='管理者'),
    sg.Button('ランキング修正',size=(12,3),key='ランキング修正')],
    

# [   sg.Button('オリコン週間\nランキング',size=(15,3),key='オリコン週間'),sg.Button('オリコンデジタル\nランキング',size=(15,3),key='オリコンデジタル'),sg.Button('ビルボードJAPAN\nHOT100',size=(15,3),key='ビルボード')]


]


window = sg.Window('FM Besthit Automatic Create System', layout, resizable=True)


if os.path.exists('test.db') == True:
   os.remove('test.db')

#GUI表示実行部分
while True:
    # ウィンドウ表示
    event, values = window.read()

    if event == '今週データ生成':
      
       HaruyaPath = values['-HaruyaExcel-']
       import Check
       import CreateDB
       import GetData
       GetData.GetThisWeekRank(HaruyaPath) 
       import ViewDeta
       import CreateExcel
       CreateExcel.MajicalExcel(GetData.GetThisWeekDate())
       import WriteCSV
       WriteCSV.WriteCSV(GetData.GetThisWeekDate())
       import ManuscriptGeneration

       window.close()
       
    if event == '先週データ生成':
       sg.popup_ok('このモードでは明屋書店のデータは取得しません')
       
       import CreateDB
       
       import GetData
       GetData.GetLastWeekRank() 
       import ViewDeta
       import OldCreateExcel
       OldCreateExcel.OldMajicalExcel(GetData.GetLastWeekDate())
       sg.popup('過去回のためDBには書き込みできません')

       window.close()

    if event == '任意週生成':
      sg.popup_ok('このモードでは明屋書店のデータは取得しません\n日付は2020年8月3日以降を入力してください')
      
      layout = [
      [sg.InputText(key='-input1-'), 
      sg.CalendarButton('Date', target='-input1-', format="%Y-%m-%d"),
      sg.Button('OK')]
      ]

      window = sg.Window('カレンダーからの入力', layout)

      while True:
         event, values = window.read()  # イベントの入力を待つ
       
         if event == sg.WINDOW_CLOSED:
            break
         elif event == 'OK':
            SelectDay = values['-input1-']
            if SelectDay:
               import CreateDB
               
               import GetData
               GetData.GetSelectWeekRank(SelectDay)
               import ViewDeta
               import OldCreateExcel
               OldCreateExcel.OldMajicalExcel(GetData.GetSelectWeekDate())
               sg.popup('過去回のためDBには書き込みできません')
            break  # 処理が終了したらループを抜ける

      window.close()

 

    if event == '管理者':
       
       import CreateDB
       
       import AdminUser

   
    if event == 'ランキング修正':
      sg.popup_ok('このモードでは明屋書店のデータは取得しません\n最新回以外のランキングは修正できません')
      layout= [[sg.Text("ランキングデータ"),
      sg.InputText('ファイルを選択', key='-HaruyaExcel-', enable_events=True,size=(41,1)), 
      sg.FileBrowse(button_text='選択', font=('メイリオ',8), size=(5,1), key="-RankExcel-"),
      sg.Button('OK')]
      ]
      window = sg.Window('ランキングデータ取得', layout)

      while True:
         event, values = window.read()  # イベントの入力を待つ
         if event == sg.WINDOW_CLOSED:
            break
         elif event == 'OK':
            FilePath = values['-RankExcel-']
            if FilePath:
               import CreateDB
               
               import RevisionRank
               RevisionRank.RevisionRank(FilePath)
            break  # 処理が終了したらループを抜ける

    #クローズボタンの処理
    if event is None:
      # print('exit')
        break

window.close()