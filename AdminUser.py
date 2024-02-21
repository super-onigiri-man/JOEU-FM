import PySimpleGUI as sg
import pandas as pd
import sqlite3

dbname = ('C:\\Users\\wiiue\\JOEU-FM\\test.db')
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
        LIMIT 200
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

def addmusic(Title,Artist):
    params = (Title,Artist)
    cursor.execute('''INSERT INTO music_master
     (Title, Artist, Score, Last_Rank, Last_Number, On_chart)
     VALUES (?, ?, 0.0, 0, 0, 0)''',params)

def sorttitle():
    global df 
    df = pd.read_sql("SELECT * FROM music_master ORDER BY Title ASC;",conn)

def sortartist():
    global df 
    df = pd.read_sql("SELECT * FROM music_master ORDER BY Artist ASC;",conn)

def serchtitle(title):
    # クエリの実行
    cursor.execute("SELECT * FROM music_master WHERE Title LIKE ?;",title)
    # 結果の取得
    result = cursor.fetchone()

    return result
    # df = df.head(20)

# データフレームをPySimpleGUIの表に変換
table_data = df.values.tolist()
header_list = ['楽曲名','アーティスト','得点','前回の順位','前回ランクイン','ランクイン回数']
# PySimpleGUIのレイアウト
layout = [
    [sg.Table(values=table_data, headings=header_list, auto_size_columns=False,enable_events=True,key='-TABLE-',
              display_row_numbers=False, justification='left', num_rows=min(25, len(df.head(200))))],

    [sg.Button('削除',size=(10,3),key='削除'),
     sg.Button('追加',size=(10,3),key='追加'),
    #  sg.Button('曲名検索',size=(12,3),key='曲名検索'),
    #  sg.Button('アーティスト名検索',size=(18,3),key='アーティスト名検索')
    ],

    [sg.Button('曲名で\n並び替え',size = (15,3),key='曲名'),
    sg.Button('アーティスト名で\n並び替え',size = (15,3),key='アーティスト'),
    sg.Button('ランクイン回数順で\n並び替え',size=(15,3),key='ランクイン'),
    sg.Button('最新回順で\n並び替え',size=(15,3),key='最終回')
    ]
]

# ウィンドウを作成
window = sg.Window('管理者画面', layout,resizable=True)

# イベントループ
while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break

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
    
    elif event == '追加':
        NewTitle = sg.popup_get_text('追加したい曲名を入力してください', '曲名')
        if NewTitle == None:
            continue
        NewArtist = sg.popup_get_text(str(NewTitle)+'のアーティスト名を入力してください', 'アーティスト')
        if NewArtist == None:
            continue
        result = sg.popup_ok_cancel(NewTitle+'/'+NewArtist+'を追加しますか？',title='追加確認')
        if result == 'OK':
            addmusic(NewTitle,NewArtist)
            # reload()
            # テーブルのデータを更新
            table_data = df.values.tolist()
            # テーブルを更新
            window['-TABLE-'].update(values=table_data)
        elif result == 'Cancel':
            # sg.popup('キャンセルが選択されました')
            continue
        
    elif event == '曲名':
        sorttitle()
        # reload()
        # テーブルのデータを更新
        table_data = df.values.tolist()
        # テーブルを更新
        window['-TABLE-'].update(values=table_data)

    elif event == 'アーティスト':
        sortartist()
        # reload()
        # テーブルのデータを更新
        table_data = df.values.tolist()
        # テーブルを更新
        window['-TABLE-'].update(values=table_data)   

    # elif event == '曲名検索':
    #     SerchTitle = sg.popup_get_text('検索したい曲名を入力してください\n あいまい検索は最後に%をつけてください', '曲名')
    #     if SerchTitle == None:
    #         continue

    #     print(serchtitle(SerchTitle))


# ウィンドウを閉じる
window.close()