import PySimpleGUI as sg
import pandas as pd
import sqlite3

dbname = (r'C:\Users\wiiue\JOEU-FM\test.db')
conn = sqlite3.connect(dbname, isolation_level=None)#データベースを作成、自動コミット機能ON
cursor = conn.cursor() #カーソルオブジェクトを作成

# サンプルのデータフレームを作成

# Scoreの高い順に上位25件を取得


# top_20_query = 
cursor.execute('''
    SELECT * FROM music_master
    WHERE Score IS NOT NULL AND Score != ''
    ORDER BY Score DESC
    LIMIT 20
''').fetchall()

# コミットして変更を保存
conn.commit()

# 整形後のデータ出力用
df = pd.read_sql('''
SELECT * FROM music_master
WHERE Score IS NOT NULL AND Score != ''
ORDER BY Score DESC
LIMIT 60''', conn)

def reload():
    cursor.execute('''
        SELECT * FROM music_master
        WHERE Score IS NOT NULL AND Score != ''
        ORDER BY Score DESC
        LIMIT 20
    ''').fetchall()

    # コミットして変更を保存
    conn.commit()

    # 整形後のデータ出力用
    df = pd.read_sql('''
        SELECT * FROM music_master
        WHERE Score IS NOT NULL AND Score != ''
        ORDER BY Score DESC
        LIMIT 60
    ''', conn)
    # df = df.head(20)


# df = df.head(20)

# データフレームをPySimpleGUIの表に変換
table_data = df.values.tolist()
header_list = ['楽曲名','アーティスト','得点','最終ランクイン回','チャートイン回数']
# PySimpleGUIのレイアウト
layout = [
    [sg.Table(values=table_data, headings=header_list, auto_size_columns=True,enable_events=True,key='-TABLE-',
              display_row_numbers=False, justification='left', num_rows=min(25, len(df.head(20))))],

    [sg.Button('削除',size=(10,1),key='削除'),sg.Button('更新',size=(10,1),key='更新'),sg.Button('Excel書き込み',size=(10,1),key='書き込み')]
]

# ウィンドウを作成
window = sg.Window('FMベストヒットランキング自動生成システム', layout,resizable=True)

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
            reload()
            # テーブルのデータを更新
            table_data = df.values.tolist()
            # テーブルを更新
            window['-TABLE-'].update(values=table_data)

# ウィンドウを閉じる
window.close()