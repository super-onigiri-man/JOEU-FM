import PySimpleGUI as sg
import pandas as pd
import sqlite3
import sys

dbname = ('test.db')
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

cursor.execute("DELETE FROM music_master WHERE Artist = '';")

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

def deleteartist(Artist):
    params = (Artist,)
    cursor.execute("DELETE FROM music_master WHERE Artist = ?;", params)

        
def updatescore(Title,Artist,Score):
    params = (Score,Title,Artist)
    cursor.execute("UPDATE music_master SET Score = ? WHERE Title = ? AND Artist= ?;",params)

def update3score(Title,Artist):
    params = (Title,Artist)
    cursor.execute("UPDATE music_master SET Score = Score + 3 WHERE Title = ? AND Artist= ?;",params)

# df = df.head(20)

# データフレームをPySimpleGUIの表に変換
table_data = df.values.tolist()
header_list = ['楽曲名','アーティスト','得点','前回の順位','前回ランクイン','ランクイン回数']
# PySimpleGUIのレイアウト
layout = [
    [sg.Table(values=table_data, headings=header_list, auto_size_columns=False,enable_events=True,key='-TABLE-',
              display_row_numbers=False, justification='left', num_rows=min(25, len(df.head(20))))],

    [sg.Button('削除',size=(10,1),key='削除'),
     sg.Button('アーティスト名で削除',size=(18,1),key='アーティスト名で削除'),
     sg.Button('得点修正',size=(10,1),key='得点修正'),
     sg.Button('3点追加',size=(10,1),key = '3点追加'),
     sg.Button('Excel書き込み',size=(12,1),key='書き込み',button_color=('white', 'red'))
     ]
]

# ウィンドウを作成
window = sg.Window('FM Besthit Automatic Create System', layout,resizable=True)

# イベントループ
while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        sg.popup('ランキングは作成されませんでした\n作るには最初から操作をやり直してください')
        sys.exit()

    if event is None:
        sg.popup('ランキングは作成されませんでした\n作るには最初から操作をやり直してください')
        sys.exit()

    elif event == '削除':
        
        selected_rows = values['-TABLE-']
        if selected_rows:
            # 選択された行を取得
            selected_row_index = values['-TABLE-'][0]

            # 選択された行の情報を取得
            selected_row_Title = table_data[selected_row_index][0]
            # print(selected_row_Title)
            selected_row_Artist = table_data[selected_row_index][1]
            # print(selected_row_Artist)
            result = sg.popup_ok_cancel(selected_row_Title+'/'+selected_row_Artist+'を削除しますか？',title='削除確認')
            if result == 'OK':
                # sg.popup('削除しました')
                # 選択された行を削除
                deleterow(selected_row_Title,selected_row_Artist)
                reload()
                # テーブルのデータを更新
                table_data = df.values.tolist()
                # テーブルを更新
                window['-TABLE-'].update(values=table_data)
            elif result == 'Cancel':
                # sg.popup('キャンセルが選択されました')
                continue

    elif event == 'アーティスト名で削除':
        
        selected_rows = values['-TABLE-']
        if selected_rows:
            # 選択された行を取得
            selected_row_index = values['-TABLE-'][0]

            # 選択された行の情報を取得
            selected_row_Title = table_data[selected_row_index][0]
            # print(selected_row_Title)
            selected_row_Artist = table_data[selected_row_index][1]
            # print(selected_row_Artist)
            result = sg.popup_ok_cancel('【警告！】'+selected_row_Artist+'の曲をすべて削除しますか？',title='削除確認')
            if result == 'OK':
                # sg.popup('削除しました')
                # 選択された行を削除
                deleteartist(selected_row_Artist)
                reload()
                # テーブルのデータを更新
                table_data = df.values.tolist()
                # テーブルを更新
                window['-TABLE-'].update(values=table_data)
            elif result == 'Cancel':
                # sg.popup('キャンセルが選択されました')
                continue

    elif event ==  '得点修正':
        selected_rows = values['-TABLE-']
        if selected_rows:
            # 選択された行を取得
            selected_row_index = values['-TABLE-'][0]

            # 選択された行の情報を取得
            selected_row_Title = table_data[selected_row_index][0]
            # print(selected_row_Title)
            selected_row_Artist = table_data[selected_row_index][1]
            # print(selected_row_Artist)
            selected_row_Score = table_data[selected_row_index][2]

            NewScore = sg.popup_get_text('得点を入力してください', '得点修正')

            result2 = sg.popup_ok_cancel(selected_row_Title+'/'+selected_row_Artist+'の得点を\n'+str(selected_row_Score)+'点から'+str(NewScore)+'点に更新しますか？')
            if result2 == 'OK':
                # sg.popup('削除しました')
                # 選択された行を削除
                updatescore(selected_row_Title,selected_row_Artist,NewScore)
                reload()
                # テーブルのデータを更新
                table_data = df.values.tolist()
                # テーブルを更新
                window['-TABLE-'].update(values=table_data)
            elif result2 == 'Cancel':
                # sg.popup('キャンセルが選択されました')
                continue

    elif event ==  '3点追加':
        selected_rows = values['-TABLE-']
        if selected_rows:
            # 選択された行を取得
            selected_row_index = values['-TABLE-'][0]

            # 選択された行の情報を取得
            selected_row_Title = table_data[selected_row_index][0]
            # print(selected_row_Title)
            selected_row_Artist = table_data[selected_row_index][1]

            result2 = sg.popup_ok_cancel(selected_row_Title+'/'+selected_row_Artist+'の得点に＋３点しますか？')
            if result2 == 'OK':
                # sg.popup('削除しました')
                # 選択された行を削除
                update3score(selected_row_Title,selected_row_Artist)
                reload()
                # テーブルのデータを更新
                table_data = df.values.tolist()
                # テーブルを更新
                window['-TABLE-'].update(values=table_data)
            elif result2 == 'Cancel':
                # sg.popup('キャンセルが選択されました')
                continue

    elif event == '書き込み':
        break
        

# ウィンドウを閉じる
window.close() #ここを実行するとランキング生成に行きます