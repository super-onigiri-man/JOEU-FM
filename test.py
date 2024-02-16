import PySimpleGUI as sg
import pandas as pd
import sqlite3
import os

# データベースファイルのパス
db_file = r'C:\Users\wiiue\JOEU-FM\test.db'

# データベースファイルが存在するかどうかを確認
if not os.path.exists(db_file):
    print(f"Error: {db_file} が見つかりません")
    exit()

# データベースに接続
conn = sqlite3.connect(db_file, isolation_level=None)
cursor = conn.cursor()

# データの読み込み関数
def load_data():
    cursor.execute('''
        SELECT * FROM music_master
        WHERE Score IS NOT NULL AND Score != ''
        ORDER BY Score DESC
        LIMIT 20
    ''')
    return pd.DataFrame(cursor.fetchall(), columns=[description[0] for description in cursor.description])

# データを読み込む
df = load_data()

# PySimpleGUIのレイアウト
layout = [
    [sg.Table(values=df.values.tolist(), headings=df.columns.tolist(), auto_size_columns=True, enable_events=True, key='-TABLE-',
              display_row_numbers=False, justification='left', num_rows=min(25, len(df.head(20))))],
    [sg.Button('削除', size=(10,1), key='削除')]
]

# ウィンドウを作成
window = sg.Window('FMベストヒットランキング自動生成システム', layout, resizable=True)

# イベントループ
while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break
    elif event == '削除':
        selected_rows = values['-TABLE-']
        if selected_rows:
            # 選択された行の番号を取得
            selected_row_index = selected_rows[0]
            # 選択された行を削除
            df.drop(selected_row_index, inplace=True)
            # テーブルを更新
            window['-TABLE-'].update(values=df.values.tolist())

# ウィンドウを閉じる
window.close()
