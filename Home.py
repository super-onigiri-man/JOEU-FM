import PySimpleGUI as sg
import os
import sys
import datetime 

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


window = sg.Window('FM Besthit Automatic Create System', layout, resizable=False)

# DBがあったら消す（過去のデータの重複防止）
if os.path.exists('test.db') == True:
   os.remove('test.db')

# DBを生成する（CreateDB.py）
import CreateDB

#GUI表示実行部分
while True:
    # ウィンドウ表示
    event, values = window.read()
   

    if event == '今週データ生成':
      
      HaruyaPath = values['-HaruyaExcel-']
      if HaruyaPath ==  'ファイルを選択' :
          sg.popup_error('明屋書店のデータを追加してください',no_titlebar=True)
          continue

          
      else:
         import Check #ファイルが重複してないか確認
          #DB構成
         import GetData #データ取得
         GetData.GetThisWeekRank(HaruyaPath) 
         import ViewData #データ閲覧・編集
         import CreateExcel # ランキングをExcelに書き込み
         CreateExcel.MajicalExcel(GetData.OriconTodays())
         import WriteCSV #DBを元にCSVに書き込む
         WriteCSV.WriteCSV(GetData.OriconTodays())
         import ManuscriptGeneration #原稿を自動生成

      continue
       
    if event == '先週データ生成':
       sg.popup_ok('このモードでは明屋書店のデータは取得しません',no_titlebar=True)
        
       import GetData
       GetData.GetLastWeekRank() 
       import ViewData
       import OldCreateExcel
       OldCreateExcel.OldMajicalExcel(GetData.OriconLastWeek())
       sg.popup('過去回のためDBには書き込みしません',no_titlebar=True)

       break

    if event == '任意週生成':
      sg.popup_ok('このモードでは明屋書店のデータは取得しません\n日付は2020年8月3日以降を入力してください',no_titlebar=True)
      
      layout = [
      [sg.InputText(key='-input1-'), 
      sg.CalendarButton('日付選択', target='-input1-', format="%Y-%m-%d",no_titlebar=True),
      sg.Button('OK')]
      ]

      window = sg.Window('カレンダーからの入力', layout)

      while True:
         event, values = window.read()  # イベントの入力を待つ
       
         if event == sg.WINDOW_CLOSED:
            continue
         elif event == 'OK':
            SelectDay = datetime.datetime.strptime(values['-input1-'], '%Y-%m-%d')
            if SelectDay > datetime.datetime.today() or SelectDay < datetime.datetime(2020, 8, 3): 
               sg.popup('指定した日付ではランキング生成できません',no_titlebar=True)
               continue

            else:
               
               import GetData
               GetData.GetSelectWeekRank(SelectDay)
               import ViewData
               import OldCreateExcel
               OldCreateExcel.OldMajicalExcel(GetData.GetSelectWeekDate())
               sg.popup('過去回のためDBには書き込みしません',no_titlebar=True)
            break

      

 

    if event == '管理者':
       
       
       import AdminUser #管理者画面の設置

   
    if event == 'ランキング修正':
      sg.popup_ok('このモードでは明屋書店のデータは取得しません\n最新回以外のランキングは修正できません',no_titlebar=True)
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
               
               
               import RevisionRank
               RevisionRank.RevisionRank(FilePath)
               sg.popup('処理が完了しました')
            continue  # 処理が終了したらループを抜ける

    #クローズボタンの処理
    if event is None:
      # print('exit')
        break
    
    if event == sg.WINDOW_CLOSED:
        break

window.close()