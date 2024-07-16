import PySimpleGUI as sg
import pandas as pd
import sqlite3
import os
import sys
import GetData

dbname = ('test.db')
conn = sqlite3.connect(dbname, isolation_level=None)#データベースを作成、自動コミット機能ON
cursor = conn.cursor() #カーソルオブジェクトを作成

# サンプルのデータフレームを作成

# Scoreの高い順に上位25件を取得


top_20_query = '''
    SELECT * FROM music_master
'''

top_20_results = cursor.execute(top_20_query).fetchall()

# コミットして変更を保存
conn.commit()

# 整形後のデータ出力用
df = pd.read_sql('''
SELECT * FROM music_master''', conn)

def reload():
    global df
    global top_20_results

    top_20_query ='''SELECT * FROM music_master'''

    top_20_results = cursor.execute(top_20_query).fetchall()

    # コミットして変更を保存
    conn.commit()

    # 整形後のデータ出力用
    df = pd.read_sql('''SELECT * FROM music_master''', conn)
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

def updatetitle(Title,Artist,oldUnique):
    params = (Title,Artist,oldUnique)
    cursor.execute("UPDATE music_master SET Title = ? WHERE Artist= ? AND Unique_id = ?;",params)

def updateartist(Title,Artist,oldUnique):
    params = (Artist,Title,oldUnique)
    cursor.execute("UPDATE music_master SET Artist = ? WHERE Title= ? AND Unique_id = ?;",params)

def updateunique(Title,Artist,Unique): #Unique_id更新用（使用していません）
    params = (Unique,Artist,Title)
    query ="SELECT * FROM music_master WHERE Unique_id = ?;",params
    results = cursor.execute(query).fetchall()
    print(results)
    if results != None: #重複があった場合
        params = (Artist,Title,Unique)
        cursor.execute("UPDATE music_master SET Artist = ? WHERE Title= ? AND Unique_id = ?;",params)


# データフレームをPySimpleGUIの表に変換
table_data = df.values.tolist()
header_list = ['楽曲名','アーティスト','得点','前回の順位','前回ランクイン','ランクイン回数','独自ID']
window_size = [20,20,8,8,8,8,8]
# PySimpleGUIのレイアウト
layout = [
    [sg.Text('並び替え'),sg.Combo(['曲名で並び替え', 'アーティスト名で並び替え', 'ランクイン回数順で並び替え','最新回順に並び替え'], default_value="選択して下さい", size=(60,1),key='Combo'),sg.Button('並び替え',key='Select')],
    [sg.Table(values=table_data, headings=header_list, col_widths=window_size,auto_size_columns=False,enable_events=True,key='-TABLE-',
              display_row_numbers=False, justification='left', num_rows=min(25, len(df.head(200))))],

    [sg.Button('曲名修正',size=(10,1),key='曲名修正',button_color=('white','#000080')),
     sg.Button('アーティスト名修正',size=(18,1),key='アーティスト名修正',button_color=('white','#000080')),
     sg.Button('削除',size=(10,1),key='削除',button_color=('white','red')),
     sg.Button('エラーログ出力',size=(15,1),key='エラーログ',button_color=('black','#ff6347')),
     sg.Button('元データ復元',size=(15,1),key='csv',button_color=('white','#4b0082')),
     sg.Button('終了・書き込み',size=(12,1),key='end',button_color=('black', '#00ff00'))
    #  sg.Button('アーティスト名検索',size=(18,3),key='アーティスト名検索')
    ]
]

# ウィンドウを作成
window = sg.Window('管理者画面', layout,resizable=True,icon='FM-BACS.ico')

# イベントループ
while True:
    event, values = window.read()
    if event is None:
      # print('exit')
        break

    elif event == '曲名修正':
        selected_rows = values['-TABLE-']
        if selected_rows:
            # 選択された行を取得
            selected_row_index = values['-TABLE-'][0]
            # 選択された行の楽曲名を取得
            selected_row_Title = table_data[selected_row_index][0]
            # 選択された行のアーティスト名を取得
            selected_row_Artist = table_data[selected_row_index][1]
            selected_row_oldUnique = table_data[selected_row_index][6]
            NewTitle = sg.popup_get_text('新しい楽曲名を入力してください','曲名修正',default_text=str(selected_row_Title),no_titlebar=True)
            if str(NewTitle) == 'None':
                continue

            result2 = sg.popup_ok_cancel(selected_row_Title+'の曲名を\n'+str(NewTitle)+'に更新しますか？',no_titlebar=True)
            if result2 == 'OK':
                # sg.popup('削除しました')
                # 選択された行を削除
                NewID = GetData.generate_unique_id(NewTitle,selected_row_Artist)
                updatetitle(NewTitle,selected_row_Artist,selected_row_oldUnique)
                reload()
                # テーブルのデータを更新
                table_data = df.values.tolist()
                # テーブルを更新
                window['-TABLE-'].update(values=table_data)
            elif result2 == 'Cancel':
                # sg.popup('キャンセルが選択されました')
                continue

    elif event == 'アーティスト名修正':
        selected_rows = values['-TABLE-']
        if selected_rows:
            # 選択された行を取得
            selected_row_index = values['-TABLE-'][0]
            # 選択された行の楽曲名を取得
            selected_row_Title = table_data[selected_row_index][0]
            # 選択された行のアーティスト名を取得
            selected_row_Artist = table_data[selected_row_index][1]
            selected_row_oldUnique = table_data[selected_row_index][6]
            NewArtist = sg.popup_get_text('新しいアーティスト名を入力してください','曲名修正',default_text=str(selected_row_Artist),no_titlebar=True)
            if str(NewArtist) == 'None':
                continue

            result2 = sg.popup_ok_cancel(selected_row_Title+'のアーティスト名を\n'+str(selected_row_Artist)+'から'+str(NewArtist)+'に更新しますか？',no_titlebar=True)
            if result2 == 'OK':
                # sg.popup('削除しました')
                # 選択された行を削除
                NewID = GetData.generate_unique_id(selected_row_Title,NewArtist)
                updateartist(selected_row_Title,NewArtist,selected_row_oldUnique)
                reload()
                # テーブルのデータを更新
                table_data = df.values.tolist()
                # テーブルを更新
                window['-TABLE-'].update(values=table_data)
            elif result2 == 'Cancel':
                # sg.popup('キャンセルが選択されました')
                continue

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
            result = sg.popup_ok_cancel(selected_row_Title+'/'+selected_row_Artist+'を削除しますか？',title='削除確認',no_titlebar=True)
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
    elif event == 'Select':
        value = values['Combo']
        
        if value == 'ランクイン回数順で並び替え':
            sortrankin()
            # reload()
            # テーブルのデータを更新
            table_data = df.values.tolist()
            # テーブルを更新
            window['-TABLE-'].update(values=table_data)

        elif value == '最新回順に並び替え':
            sortlastepisode()
            # reload()
            # テーブルのデータを更新
            table_data = df.values.tolist()
            # テーブルを更新
            window['-TABLE-'].update(values=table_data)
            
        elif value == '曲名で並び替え':
            sorttitle()
            # reload()
            # テーブルのデータを更新
            table_data = df.values.tolist()
            # テーブルを更新
            window['-TABLE-'].update(values=table_data)

        elif value == 'アーティスト名で並び替え':
            sortartist()
            # reload()
            # テーブルのデータを更新
            table_data = df.values.tolist()
            # テーブルを更新
            window['-TABLE-'].update(values=table_data)  

    elif event == 'エラーログ':
         # logファイルをダウンロードフォルダーに保存します

        if os.path.isfile('error.log'):
            import shutil
            user_folder = os.path.expanduser("~")
            folder = os.path.join(user_folder, "Downloads")
            shutil.copy('error.log', folder)
            os.chdir(os.path.dirname(sys.argv[0]))
            os.remove('error.log')
            sg.popup_ok('エラーログをダウンロードフォルダにコピーしました',no_titlebar=True)

        else:
            sg.popup('ログがありませんでした。',no_titlebar=True) 


    elif event == 'csv':
        result = sg.popup_ok_cancel('csvを2120回から復元しますか？\n復元するともとには戻せません',title='csv復元確認',no_titlebar=True)
        if result == 'OK':
            import CreateDB2
            import LearningRank
            sg.popup('処理が終了しました。システムを終了します',no_titlebar=True)
            sys.exit()
        else:
            break

    elif event == 'end':
        break


# ウィンドウを閉じる
window.close()