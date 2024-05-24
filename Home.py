import PySimpleGUI as sg
import os
import sys
import datetime
import webbrowser
import sqlite3
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
    [sg.Button('今週のデータを生成する', size=(25, 3), key='今週データ生成')],
    [sg.Button('先週のデータを生成する', size=(25, 3), key='先週データ生成')],
    [
        sg.Button('任意の週のデータを生成する', size=(25, 3), key='任意週生成'),
        sg.Button('管理者用', size=(10, 3), key='管理者'),
        sg.Button('ランキング修正', size=(12, 3), key='ランキング修正'),
        sg.Button('Webサイト\nランキング表示', size=(12, 3), key='ランキング表示')
    ]
]

# ウィンドウを作成
window = sg.Window('FM Besthit Automatic Create System', layout, resizable=False)

# GUI表示実行部分
while True:
    event, values = window.read()

    if event in (sg.WINDOW_CLOSED, None):
        break

    if event == '今週データ生成':
        HaruyaPath = values['-HaruyaExcel-']
        if HaruyaPath == 'ファイルを選択':
            sg.popup_error('明屋書店のデータを追加してください', no_titlebar=True)
            continue
        else:
            import Check  # ファイルが重複してないか確認
            import GetData  # データ取得
            GetData.GetThisWeekRank(HaruyaPath)
            import ViewData  # データ閲覧・編集
            import CreateExcel  # ランキングをExcelに書き込み
            CreateExcel.MajicalExcel(GetData.OriconTodays())
            import WriteCSV  # DBを元にCSVに書き込む
            WriteCSV.WriteCSV(GetData.OriconTodays())
            import ManuscriptGeneration  # 原稿を自動生成
        
        break

    if event == '先週データ生成':
      sg.popup_ok('このモードでは明屋書店のデータは取得しません', no_titlebar=True)
      import GetData
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

        while True:
            date_event, date_values = date_window.read()
            if date_event in (sg.WINDOW_CLOSED, None):
                break
            elif date_event == 'OK' or date_event == '\r':
                SelectDay = datetime.datetime.strptime(date_values['-input1-'], '%Y-%m-%d')
                if SelectDay > datetime.datetime.today() or SelectDay < year2:
                    sg.popup('指定した日付ではランキング生成できません', no_titlebar=True)
                    continue
                try:
                    import GetData
                    GetData.GetSelectWeekRank(SelectDay)
                    import ViewData
                    import OldCreateExcel
                    OldCreateExcel.OldMajicalExcel(GetData.GetSelectWeekDate())
                    sg.popup('過去回のためDBには書き込みしません', no_titlebar=True)
                except Exception as e:
                    sg.popup_error(f"エラーが発生しました: {e}", no_titlebar=True)
                break
        date_window.close()

    if event == '管理者':
       
        import AdminUser  # 管理者画面の設置
        

    if event == 'ランキング修正':
        sg.popup_ok('このモードでは明屋書店のデータは取得しません\n最新回以外のランキングは修正できません', no_titlebar=True)
        rank_layout = [
            [sg.Text("ランキングデータ"),
             sg.InputText('ファイルを選択', key='-HaruyaExcel-', enable_events=True, size=(41, 1)),
             sg.FileBrowse(button_text='選択', font=('メイリオ', 8), size=(5, 1), key="-RankExcel-"),
             sg.Button('OK')]
        ]
        rank_window = sg.Window('ランキングデータ取得', rank_layout)

        while True:
            rank_event, rank_values = rank_window.read()
            if rank_event in (sg.WINDOW_CLOSED, None):
                break
            elif rank_event == 'OK':
                FilePath = rank_values['-RankExcel-']
                if FilePath:
                  import RevisionRank
                  RevisionRank.RevisionRank(FilePath)
                  break
        rank_window.close()

    if event == 'ランキング表示':
        display_layout = [
            [sg.Button('オリコン週間\nランキング', size=(15, 3), key='オリコン週間'),
             sg.Button('オリコンデジタル\nランキング', size=(15, 3), key='オリコンデジタル'),
             sg.Button('ビルボードJAPAN\nHOT100', size=(15, 3), key='ビルボード')]
        ]

        display_window = sg.Window('ランキングデータ取得', display_layout)

        while True:
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

window.close()