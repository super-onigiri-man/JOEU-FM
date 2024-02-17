import PySimpleGUI as sg
import pandas as pd
import sqlite3

dbname = (r'C:\Users\wiiue\JOEU-FM\test.db')
conn = sqlite3.connect(dbname, isolation_level=None)#データベースを作成、自動コミット機能ON
cursor = conn.cursor() #カーソルオブジェクトを作成

# サンプルのデータフレームを作成

# Scoreの高い順に上位25件を取得


top_20_query = '''
    SELECT * FROM music_master
    WHERE Score IS NOT NULL AND Score != ''
    ORDER BY Score DESC
    LIMIT 20
'''

top_20_results = cursor.execute(top_20_query).fetchall()

# コミットして変更を保存
conn.commit()

# 整形後のデータ出力用
df = pd.read_sql('''
SELECT * FROM music_master
WHERE Score IS NOT NULL AND Score != ''
ORDER BY Score DESC
LIMIT 60''', conn)

def reload():
    global df
    global top_20_results

    top_20_query ='''
        SELECT * FROM music_master
        WHERE Score IS NOT NULL AND Score != ''
        ORDER BY Score DESC
        LIMIT 20
    '''

    top_20_results = cursor.execute(top_20_query).fetchall()

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

def deleterow(Title,Artist):
    params = (Title, Artist)
    cursor.execute("DELETE FROM music_master WHERE Title = ? AND Artist = ?;", params)
        
def updatescore(Title,Artist,Score):
    params = (Score,Title,Artist)
    cursor.execute("UPDATE music_master SET Score = ? WHERE Title = ? AND Artist= ?;",params)

def sortrankin():
    global df
    df = pd.read_sql("SELECT * FROM music_master ORDER BY On_Chart DESC;",conn)

def sortlastepisode():
    global df
    df = pd.read_sql("SELECT * FROM music_master ORDER BY Last_Number DESC;",conn)

# df = df.head(20)

# データフレームをPySimpleGUIの表に変換
table_data = df.values.tolist()
header_list = ['楽曲名','アーティスト','得点','前回の順位','前回ランクイン','ランクイン回数']
# PySimpleGUIのレイアウト
layout = [
    [sg.Table(values=table_data, headings=header_list, auto_size_columns=False,enable_events=True,key='-TABLE-',
              display_row_numbers=False, justification='left', num_rows=min(25, len(df.head(20))))],

    [sg.Button('削除',size=(10,3),key='削除'),
     sg.Button('ランクイン回数順で並び替え',size=(15,3),key='ランクイン'),
     sg.Button('最新回順で並び替え',size=(15,3),key='最終回'),
     sg.Button('追加',size=(8,3),key='追加'),
     sg.Button('更新',size=(8,3),key='更新')]
]

# ウィンドウを作成
window = sg.Window('FMベストヒットランキング自動生成システム', layout,resizable=True)

# イベントループ
while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break

    elif event == 'ランクイン':
        sortrankin()
        # reload()
        # テーブルのデータを更新
        table_data = df.values.tolist()
        # テーブルを更新
        window['-TABLE-'].update(values=table_data)

    elif event == '最終回':
        sortlastepisode()
        # reload()
        # テーブルのデータを更新
        table_data = df.values.tolist()
        # テーブルを更新
        window['-TABLE-'].update(values=table_data)
    


# ウィンドウを閉じる
window.close()