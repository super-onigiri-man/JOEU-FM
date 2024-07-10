import csv
import sqlite3
import PySimpleGUI as sg
import traceback
import sys
import pandas as pd

dbname = 'test.db'  # データベース名
csv_file = '楽曲データ.csv'  # CSVファイル名

sg.theme('SystemDefault')

conn = sqlite3.connect(dbname, isolation_level=None)  # データベースに接続（自動コミット機能ON）
cursor = conn.cursor()  # カーソルオブジェクトを作成

def update_progress_bar(progress_bar, progmsg, value, msg):
    if progress_bar is not None and progmsg is not None:
        progress_bar.update_bar(value)
        progmsg.update(msg)
        window.refresh()

try:
    layout = [
        [sg.Text('読み込み中...', size=(15, 1)), sg.ProgressBar(3, orientation='h', size=(20, 20), key='progressbar')],
        [sg.Text(key='progmsg')]
    ]

    window = sg.Window('システム起動中', layout, finalize=True, icon='FM-BACS.ico')

    progress_bar = window['progressbar']
    progmsg = window['progmsg']

    if progress_bar is None or progmsg is None:
        raise ValueError("エラーです")

    # テーブル作成のSQL文
    sql = """CREATE TABLE IF NOT EXISTS music_master (
        Title TEXT,
        Artist TEXT,
        Score DOUBLE,
        Last_Rank INT,
        Last_Number INT,
        On_Chart INT,
        Unique_id TEXT,
        PRIMARY KEY (Unique_id)
    );"""
    
    # DBのフォーマット設定
    cursor.execute(sql)
    conn.commit()

    update_progress_bar(progress_bar, progmsg, 1, 'DB作成中')

    with open('楽曲データ.csv', encoding='utf-8') as f:
        next(f)
        csv_reader = csv.reader(f)
        for row in csv_reader:
            cursor.execute("INSERT OR REPLACE INTO music_master VALUES (?, ?, ?, ?, ?, ?, ?)", (row[0], row[1], row[2], row[3], row[4], row[5], row[6]))

    update_progress_bar(progress_bar, progmsg, 2, 'データ整合性確認中')

    cursor.execute('UPDATE music_master SET Score = 0;') #Scoreの値を全消し
    cursor.execute('''DELETE FROM music_master WHERE Last_Number = '' OR Last_Number = 0;''') #Last_Numberの値が0もしくは空文字の場合消す
    cursor.execute('''DELETE FROM music_master WHERE Last_Number = "ベストヒットえひめ";''')

    update_progress_bar(progress_bar, progmsg, 3, '完了')

    window.close()

except Exception as e:
    with open('error.log', 'a') as f:
        traceback.print_exc(file=f)
    sg.popup_error("DBを作成できませんでした。\n システムを終了します", title="エラー", no_titlebar=True)
    sys.exit()
