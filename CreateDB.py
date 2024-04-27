import csv
import sqlite3
import PySimpleGUI as sg
import requests

dbname = ('test.db')#データベース名.db拡張子で設定
conn = sqlite3.connect(dbname, isolation_level=None)#データベースを作成、自動コミット機能ON
cursor = conn.cursor() #カーソルオブジェクトを作成

try:
    layout = [
        [sg.Text('DB作成中...', size=(15, 1)), sg.ProgressBar(72, orientation='h', size=(20, 20), key='progressbar')],
        [sg.Button('読み込み中止')]
    ]

    window = sg.Window('DB作成中...', layout,finalize=True)

    def update_progress_bar(progress_bar, value):
        progress_bar.update_bar(value)
        window.refresh()

    while True:

        update_progress_bar(window['progressbar'], 0)
        
        sql = """CREATE TABLE music_master (
            Title TEXT,
            Artist TEXT,
            Score DOUBLE,
            Last_Rank INT,
            Last_Number INT,
            On_Chart INT,
            PRIMARY KEY (Title, Artist)
        );"""
        #↑DBのフォーマット設定（別のところに書いておきます）
        cursor.execute(sql)#executeコマンドでSQL文を実行
        conn.commit()#データベースにコミット(Excelでいう上書き保存。自動コミット設定なので不要だが一応・・)

        update_progress_bar(window['progressbar'], 36)

        with open('楽曲データ.csv',encoding='utf-8') as f:
            # CSVリーダーオブジェクトを作成
            csv_reader = csv.reader(f)
            for row in csv_reader:
                # テーブルにデータを挿入
                cursor.execute("INSERT INTO music_master VALUES (?,?,?,?,?,?)", (row[0], row[1],row[2],row[3],row[4],row[5]))

        update_progress_bar(window['progressbar'], 72)

        window.close()

        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'キャンセル':
                    break

except Exception as e:
        import traceback
        with open('error.log', 'a') as f:
            traceback.print_exc( file=f)
        sg.popup_error("DBを作成できませんでした。\n システムを終了します",title="エラー")
        
        exit()
            