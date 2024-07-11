import sqlite3
import PySimpleGUI as sg
import traceback
import sys

dbname = 'test2.db'  # データベース名

sg.theme('SystemDefault')

conn = sqlite3.connect(dbname, isolation_level=None)  # データベースに接続（自動コミット機能ON）
cursor = conn.cursor()  # カーソルオブジェクトを作成

try:
    layout = [
        [sg.Text('データベースを作成中...', size=(30, 1))],
        [sg.ProgressBar(1, orientation='h', size=(20, 20), key='progressbar')],
        [sg.Text(key='progmsg')]
    ]

    window = sg.Window('システム起動中', layout, finalize=True)

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

    # 進捗バーの更新
    window['progressbar'].update_bar(1)
    window['progmsg'].update('データベース作成完了')

    # 進捗バーを表示するためにウィンドウを少し待つ
    # sg.popup("データベースの作成が完了しました。", title="完了", no_titlebar=True)

    window.close()

except Exception as e:
    # エラーログを出力
    with open('error.log', 'a') as f:
        traceback.print_exc(file=f)
    sg.popup_error("データベースを作成できませんでした。\n システムを終了します", title="エラー", no_titlebar=True)
    sys.exit()
finally:
    # 接続を閉じる
    conn.close()
