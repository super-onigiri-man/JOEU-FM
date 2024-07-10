import PySimpleGUI as sg
import pandas as pd
import sqlite3
import sys
import GetData
import datetime

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

def RankingUpdate():

    with open('Reload.txt','r',encoding='UTF-8') as f:
        info = str(f.read())
    print(info)
    info=info.split(',') 
    ID = info[0]
    HaruyaPath = info[1]
    SelectDay = info[2]

    if int(ID) == 1: #識別番号が今週（1）だった場合
        GetData.ResetData()
        GetData.GetThisWeekRank(HaruyaPath)

    elif int(ID) == 2: #識別番号が2（先週）だった場合
        GetData.ResetData()
        GetData.GetLastWeekRank()

    elif int(ID) == 3: #識別番号が3（任意週）の場合
        GetData.ResetData()
        SelectDay = datetime.datetime.strptime(SelectDay, '%Y-%m-%d')
        GetData.GetSelectWeekRank(SelectDay,True)

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

def updatetitle(Title,Artist,oldUnique):
    params = (Title,Artist,oldUnique)
    cursor.execute("UPDATE music_master SET Title = ? WHERE Artist= ? AND Unique_id = ?;",params)

def updateartist(Title,Artist,oldUnique):
    params = (Artist,Title,oldUnique)
    cursor.execute("UPDATE music_master SET Artist = ? WHERE Title= ? AND Unique_id = ?;",params)

def updateunique(Title,Artist,newUnique):
    params = (newUnique,Title,Artist)
    cursor.execute("UPDATE music_master SET Unique_id = ? WHERE Title= ? AND  Artist = ?;",params)
# df = df.head(20)

# データフレームをPySimpleGUIの表に変換
table_data = df.values.tolist()
header_list = ['楽曲名','アーティスト','得点','前回の順位','LastNumber','Onchart']
window_size = [20,20,8,8,8,8]
# PySimpleGUIの,レイアウト
layout = [
    [sg.Table(values=table_data, headings=header_list, col_widths=window_size,auto_size_columns=False,enable_events=True,key='-TABLE-',
              display_row_numbers=False, justification='left', num_rows=min(25, len(df.head(20))))],

    [sg.Button('曲名修正',size=(10,1),key='曲名修正',button_color=('black', 'orange')),
     sg.Button('アーティスト名修正',size=(21,1),key='アーティスト名修正',button_color=('black', 'orange')),
     sg.Button('ランキング再取得',size=(18,1),key='ランキング再取得',button_color=('white', '#4b0082'))],

    [sg.Button('削除',size=(10,1),key='削除',button_color=('white', 'red')),
     sg.Button('アーティスト名で削除',size=(21,1),key='アーティスト名で削除',button_color=('white', 'red')),
     sg.Button('得点修正',size=(12,1),key='得点修正'),
     sg.Button('3点追加',size=(12,1),key = '3点追加'),
     sg.Button('Excel書き込み',size=(12,1),key='書き込み',button_color=('white', '#006400'))
     ]
     
]

window = sg.Window('FM Besthit Automatic Create System', layout,resizable=False,icon='FM-BACS.ico',finalize=True)

# イベントループ
while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED:
        cursor.execute('UPDATE music_master SET Score = 0;') #Scoreの値を全消し
        cursor.execute('''DELETE FROM music_master WHERE Last_Number = '' OR Last_Number = 0;''') #Last_Numberの値が0もしくは空文字の場合消す
        conn.commit()
        sg.popup('ランキングは作成されませんでした\n作るには最初から操作をやり直してください',no_titlebar=True)
        sys.exit()

    if event is None:
        cursor.execute('UPDATE music_master SET Score = 0;') #Scoreの値を全消し
        cursor.execute('''DELETE FROM music_master WHERE Last_Number = '' OR Last_Number = 0;''') #Last_Numberの値が0もしくは空文字の場合消す
        conn.commit()
        sg.popup('ランキングは作成されませんでした\n作るには最初から操作をやり直してください',no_titlebar=True)
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
            result = sg.popup_ok_cancel('【警告！】'+selected_row_Artist+'の曲をすべて削除しますか？',title='削除確認',no_titlebar=True)
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

            NewScore = sg.popup_get_text('新しい得点を入力してください', '得点修正',default_text=str(selected_row_Score),no_titlebar=True)

            if str(NewScore) == 'None':
                continue

            result2 = sg.popup_ok_cancel(selected_row_Title+'/'+selected_row_Artist+'の得点を\n'+str(selected_row_Score)+'点から'+str(NewScore)+'点に更新しますか？',no_titlebar=True)
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

            result2 = sg.popup_ok_cancel(selected_row_Title+'/'+selected_row_Artist+'の得点に＋３点しますか？',no_titlebar=True)
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
                updateunique(NewTitle,selected_row_Artist,NewID)
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
                updateunique(selected_row_Title,NewArtist,NewID)
                reload()
                
                # テーブルのデータを更新
                table_data = df.values.tolist()
                # テーブルを更新
                window['-TABLE-'].update(values=table_data)
            elif result2 == 'Cancel':
                # sg.popup('キャンセルが選択されました')
                continue

    elif event == 'ランキング再取得':
        result2 = sg.popup_ok_cancel('選択すると今まで変更したデータはすべて失われます\n実行しますか？',no_titlebar=True)
        if result2 == 'OK':
            table_data = df[:0] #データフレームを一旦すべて空にする
            RankingUpdate()
            reload()
            # テーブルのデータを更新
            table_data = df.values.tolist()
            # テーブルを更新
            window['-TABLE-'].update(values=table_data)
            continue
        elif result2 == 'Cancel':
            continue


    elif event == '書き込み':
        break
        

# ウィンドウを閉じる
window.close() #ここを実行するとランキング生成に行きます