import PySimpleGUI as sg
import os
import sys
import datetime
import webbrowser
import GetData


# カレントディレクトリをスクリプトのディレクトリに変更
os.chdir(os.path.dirname(sys.argv[0]))

# DBがなかったらCreateDBをインポートしてDBを作成
if os.path.exists('test.db'):
    import CreateDB

# テーマカラーを設定
sg.theme('SystemDefault')

# レイアウトを定義
layout = [
    [
        sg.Text("明屋書店Excelデータ"),
        sg.InputText('ファイルを選択', key='-HaruyaExcel-', enable_events=True, size=(48, 1)),
        sg.FileBrowse(button_text='選択', font=('メイリオ', 8), size=(5, 1), key="-HaruyaExcel-")
    ],
    [   sg.Checkbox('明屋書店Excelデータの新フォーマットを読み込む',key='format')],
    [sg.Button('今週のデータを生成する', size=(25, 3), key='今週データ生成')],
    [sg.Button('先週のデータを生成する', size=(25, 3), key='先週データ生成')],
    [
        sg.Button('任意の週のデータを生成する', size=(25, 3), key='任意週生成'),
        sg.Button('管理者用', size=(10, 3), key='管理者'),
        sg.Button('ランキング登録\nランキング修正', size=(12, 3), key='ランキング修正'),
        sg.Button('Webサイト\nランキング表示', size=(12, 3), key='ランキング表示')
    ]
]

# ウィンドウを作成
window = sg.Window('FM Besthit Automatic Create System', layout, resizable=False,icon='FM-BACS.ico')
GetData.WriteLog(0,"FM BACS（ホーム画面）起動")

# GUI表示実行部分
while True:
    event, values = window.read()

    if event in (sg.WINDOW_CLOSED, None):
        break

    if event == '今週データ生成':
        HaruyaPath = values['-HaruyaExcel-']
        if HaruyaPath == 'ファイルを選択':
            sg.popup_ok('明屋書店が選択されていません\n明屋書店以外のデータでランキングを生成します', no_titlebar=True)
            import Check
            GetData.NGetThisWeekRank()
            import ViewData
            if GetData.Flags()==True:
                import CreateExcel  # ランキングをExcelに書き込み
                CreateExcel.MajicalExcel(GetData.OriconTodays())
                import WriteCSV  # DBを元にCSVに書き込む
                WriteCSV.WriteCSV(GetData.OriconTodays())
                import ManuscriptGeneration  # 原稿を自動生成
            else:
                import OldCreateExcel
                OldCreateExcel.OldMajicalExcel(GetData.OriconTodays())

            break
        
        else:
            import Check  # ファイルが重複してないか確認
            if values['format'] == True:
                GetData.GetThisWeekRank(HaruyaPath,True) #新しいフォーマットを読みます
            else:
                GetData.GetThisWeekRank(HaruyaPath,False) #古いフォーマットを読みます
            import ViewData  # データ閲覧・編集
            if GetData.Flags() == True: # 先週のランキングを作ったか確認
                import CreateExcel  # ランキングをExcelに書き込み
                CreateExcel.MajicalExcel(GetData.OriconTodays())
                import WriteCSV  # DBを元にCSVに書き込む
                WriteCSV.WriteCSV(GetData.OriconTodays())
                import ManuscriptGeneration  # 原稿を自動生成
            else:
                import OldCreateExcel
                OldCreateExcel.OldMajicalExcel(GetData.OriconTodays())
            
            break

    if event == '先週データ生成':
      sg.popup_ok('このモードでは明屋書店のデータは取得しません', no_titlebar=True)
      GetData.GetLastWeekRank()
      import ViewData
      import OldCreateExcel
      OldCreateExcel.OldMajicalExcel(GetData.OriconLastWeek())
      sg.popup('過去回のためDBには書き込みしません', no_titlebar=True)
      break

    if event == '任意週生成':
        year2 = datetime.datetime.today() - datetime.timedelta(days=365 * 2)
        sg.popup_ok(f'このモードでは明屋書店のデータは取得しません\n日付は{year2.year}年{year2.month}月{year2.day}日以降を入力してください', no_titlebar=True)

        date_layout = [
            [sg.InputText(key='-input1-'),
             sg.CalendarButton('日付選択', target='-input1-', format="%Y-%m-%d", no_titlebar=True),
             sg.Button('OK')]
        ]

        date_window = sg.Window('カレンダーからの入力', date_layout)
        GetData.WriteLog(0,"任意週作成のカレンダーを起動")

        while True:
            date_event, date_values = date_window.read()
            if date_event in (sg.WINDOW_CLOSED, None):
                break
            elif date_event == 'OK' or date_event == '\r':
                try:
                    SelectDay = datetime.datetime.strptime(date_values['-input1-'], '%Y-%m-%d')
                except Exception as e:
                    sg.popup_error('入力された値が不正です。正しい日付を入力してください。',no_titlebar=True)
                    continue
                if SelectDay > datetime.datetime.today() or SelectDay < year2:
                    sg.popup('指定した日付ではランキング生成できません', no_titlebar=True)
                    continue
                try:
                    date_window.close()
                    GetData.WriteLog(1,"カレンダーで"+str(SelectDay)+"を選択")
                    GetData.WriteLog(5,"任意週作成のカレンダーを終了")
                    GetData.GetSelectWeekRank(SelectDay,False)
                    import ViewData
                    import OldCreateExcel
                    OldCreateExcel.OldMajicalExcel(GetData.GetSelectWeekDate())
                    sg.popup('過去回のためDBには書き込みしません', no_titlebar=True)
                except Exception as e:
                    sg.popup_error(f"エラーが発生しました: {e}", no_titlebar=True)
                break
        
    if event == '管理者':
       
        import AdminUser  # 管理者画面の設置
        import WriteCSV2 #DBからcsvに直接書き込む 
        break
        

    if event == 'ランキング修正':

        GetData.WriteLog(1,"ランキング修正を選択")

        last_number = GetData.GetLastNumber()
        rank_layout = [
            [sg.Text('No.'+str(last_number)+'より前のランキングデータは登録・修正できません')],
            [sg.Text("ランキングデータ"),
             sg.InputText('ファイルを選択', key='-HaruyaExcel-', enable_events=True, size=(41, 1)),
             sg.FileBrowse(button_text='選択', font=('メイリオ', 8), size=(5, 1), key="-RankExcel-"),
             sg.Button('OK')],
            [sg.Button('ロールバックする',button_color=('white','red'))]
        ]
        rank_window = sg.Window('ランキング登録・修正', rank_layout,icon='FM-BACS.ico')
        GetData.WriteLog(0,"ホーム：ランキング登録・修正画面を起動")

        while True:
            rank_event, rank_values = rank_window.read()
            if rank_event in (sg.WINDOW_CLOSED, None):
                
                break
            elif rank_event == 'OK':
                FilePath = rank_values['-RankExcel-']
                if FilePath:
                  import RevisionRank
                  if RevisionRank.ExcelCheck(FilePath): # Excelに日付などが入っているかのチェック（返り値bool）
                     RevisionRank.RevisionRank(FilePath) 
                     RevisionRank.CopyFile(FilePath)
                     break
                  else:
                      break
                  
            elif rank_event == 'ロールバックする':
                GetData.WriteLog(2,"ホーム：ロールバックを選択")
                result = sg.popup_ok_cancel('No.'+str(last_number)+'のランキングを削除します\nこの操作はOKを押すと取り消せません\n実行しますか？',no_titlebar=True)
                if result == 'OK':
                    GetData.WriteLog(1,"ホーム：ロールバック操作を承認")
                    import RollBack
                    break
                else:
                    GetData.WriteLog(1,"ホーム：ロールバック操作をキャンセル")
                    continue
        
        GetData.WriteLog(5,"ホーム：ランキング登録・修正画面を終了")
        rank_window.close()

    if event == 'ランキング表示':
        display_layout = [
            [sg.Button('オリコン週間\nランキング', size=(15, 3), key='オリコン週間'),
             sg.Button('オリコンデジタル\nランキング', size=(15, 3), key='オリコンデジタル'),
             sg.Button('ビルボードJAPAN\nHOT100', size=(15, 3), key='ビルボード')]
        ]

        display_window = sg.Window('ランキングサイト一覧', display_layout,icon='FM-BACS.ico')

        while True: #選択するとブラウザ表示（すでにブラウザが出てる場合は新しいタブで表示）
            display_event, display_values = display_window.read()
            if display_event in (sg.WINDOW_CLOSED, None):
                break
            elif display_event == 'オリコン週間':
                webbrowser.open(GetData.OriconWeekUrl(), new=0)
                break
            elif display_event == 'オリコンデジタル':
                webbrowser.open(GetData.OriconDigitalUrl(), new=0)
                break
            elif display_event == 'ビルボード':
                webbrowser.open(GetData.BillboardUrl(), new=0)
                break
        display_window.close()

if os.path.exists('Reload.txt'):
    os.remove('Reload.txt')

GetData.WriteLog(5,"ホーム：FM BACS（ホーム画面）から終了\n")
window.close()