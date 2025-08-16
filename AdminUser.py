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
    GetData.WriteLog(2,"管理者画面：データベース："+Title+"/"+Artist+"を削除")
        
def updatescore(Title,Artist,Score):
    params = (Score,Title,Artist)
    cursor.execute("UPDATE music_master SET Score = ? WHERE Title = ? AND Artist= ?;",params)
    GetData.WriteLog(2,"管理者画面：データベース："+Title+"/"+Artist+"の得点を"+Score+"へ変更")

def sortrankin():
    global df
    df = pd.read_sql("SELECT * FROM music_master ORDER BY On_Chart DESC;",conn)
    GetData.WriteLog(2,"管理者画面：データベース：ランクイン回数順に並び替えを実行")

def sortlastepisode():
    global df
    df = pd.read_sql("SELECT * FROM music_master ORDER BY Last_Number DESC;",conn)
    GetData.WriteLog(2,"管理者画面：データベース：最終ランクイン順に並び替えを実行")

def addmusic(Title,Artist):
    params = (Title,Artist)
    cursor.execute('''INSERT INTO music_master
     (Title, Artist, Score, Last_Rank, Last_Number, On_chart)
     VALUES (?, ?, 0.0, 0, 0, 0)''',params)
    GetData.WriteLog(2,"管理者画面：データベース："+Title+"/"+Artist+"を追加")

def sorttitle():
    global df 
    df = pd.read_sql("SELECT * FROM music_master ORDER BY Title ASC;",conn)
    GetData.WriteLog(2,"管理者画面：データベース：曲名順に並び替えを実行")

def sortartist():
    global df 
    df = pd.read_sql("SELECT * FROM music_master ORDER BY Artist ASC;",conn)
    GetData.WriteLog(2,"管理者画面：データベース：アーティスト名順に並び替えを実行")

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
    GetData.WriteLog(2,"管理者画面：データベース："+Title+"を修正")

def updateartist(Title,Artist,oldUnique):
    params = (Artist,Title,oldUnique)
    cursor.execute("UPDATE music_master SET Artist = ? WHERE Title= ? AND Unique_id = ?;",params)
    GetData.WriteLog(2,"管理者画面：データベース："+Artist+"を修正")

def updateunique(Title,Artist,Unique): #Unique_id更新用（使用していません）
    params = (Unique,Artist,Title)
    query ="SELECT * FROM music_master WHERE Unique_id = ?;",params
    results = cursor.execute(query).fetchall()
    print(results)
    if results != None: #重複があった場合
        params = (Artist,Title,Unique)
        cursor.execute("UPDATE music_master SET Artist = ? WHERE Title= ? AND Unique_id = ?;",params)

def insert_music_data(Title,Artist,LastRank,LastNumber,Onchart,NewUnique_id):
    params = (Title,Artist,LastRank,LastNumber,Onchart,NewUnique_id)
    cursor.execute("INSERT OR REPLACE INTO music_master VALUES (?, ?, 0, ?, ?, ?, ?);", params)
    GetData.WriteLog(2,"管理者画面：データベース："+Title+"/"+Artist+"を追加")

def FormatCheck(Title,Artist,LastRank,LastNumber,Onchart):

    if Title is None or Title == '' or Artist is None  or Artist == '':
        sg.popup_error('曲名、アーティスト名が入力されていません',no_titlebar=True)
        return False
        
    
    if LastRank == '0'  or LastNumber == '0' or Onchart == '0':
        sg.popup_error('最終順位、最終ランクイン、ランクイン回数に0が含まれています\n0は登録できません',no_titlebar=True)
        return False
    
    if LastRank == ''  or LastNumber == '' or Onchart == '':
        sg.popup_error('最終順位、最終ランクイン、ランクイン回数が入力されていません',no_titlebar=True)
        return False
    
    Unique_id = GetData.generate_unique_id(Title, Artist)
    query = "SELECT EXISTS(SELECT * FROM music_master WHERE Unique_id = ?) AS unique_check;"
    # SQLクエリを実行
    cursor.execute(query, (Unique_id,))
    results = cursor.fetchone() 

    if results[0] == 1:
        popupresult = sg.popup_ok_cancel('既に同じ曲が登録されています\n更新しますか？\n※曲名、アーティスト名の頭3文字で判定しています',no_titlebar=True)
        if popupresult == 'OK':
            return True
        else:
            return False
        
    else:
        return True

    

# データフレームをPySimpleGUIの表に変換
table_data = df.values.tolist()
header_list = ['楽曲名','アーティスト','得点','前回の順位','前回ランクイン','ランクイン回数','独自ID']
window_size = [25,25,8,8,8,8,8]
# PySimpleGUIのレイアウト
layout = [
    [sg.Text('並び替え'),sg.Combo(['曲名で並び替え', 'アーティスト名で並び替え', 'ランクイン回数順で並び替え','最新回順に並び替え'], default_value="選択して下さい", size=(60,1),key='Combo'),sg.Button('並び替え',key='Select'),sg.Text('現在No.'+str(GetData.GetLastNumber())+'まで登録されています',justification='right')],
    [sg.Table(values=table_data, headings=header_list, col_widths=window_size,auto_size_columns=False,enable_events=True,key='-TABLE-',
              display_row_numbers=False, justification='left', num_rows=min(25, len(df.head(200))))],

    [sg.Button('楽曲データ修正',size=(15,1),key='修正',button_color=('white','#000080')),
     sg.Button('削除',size=(10,1),key='削除',button_color=('white','red')),
     sg.Button('ログ出力',size=(15,1),key='エラーログ',button_color=('black','#ff6347')),
     sg.Button('元データ復元',size=(15,1),key='csv',button_color=('white','#4b0082')),
     sg.Button('内部ファイルを開く',size=(18,1),key='setting',button_color=('white','#ffa500')),
     sg.Button('終了・書き込み',size=(16,1),key='end',button_color=('black', '#00ff00')),
    ]
]

# ウィンドウを作成
window = sg.Window('管理者画面', layout,resizable=True,icon='FM-BACS.ico')
GetData.WriteLog(0,"管理者画面：管理者画面起動")

# イベントループ
while True:
    event, values = window.read()
    if event is None:
      # print('exit')
        break

    elif event == '修正':

        GetData.WriteLog(1,"管理者画面：楽曲修正を選択")

        selected_rows = values['-TABLE-']
        if selected_rows:
            # 選択された行を取得
            selected_row_index = values['-TABLE-'][0]
            # 選択された行の楽曲名を取得
            select_Title = table_data[selected_row_index][0]
            # 選択された行のアーティスト名を取得
            select_Artist = table_data[selected_row_index][1]
            select_LastRank = table_data[selected_row_index][3]
            select_LastNumber = table_data[selected_row_index][4]
            select_Onchart = table_data[selected_row_index][5]
            select_oldUnique = table_data[selected_row_index][6]

        else:
            continue

        # レイアウト
        layout = [
        [sg.Text('曲名', size=(15, 1)), sg.InputText(default_text=str(select_Title),key='NewTitle')],
        [sg.Text('アーティスト', size=(15, 1)), sg.InputText(default_text=str(select_Artist),key='NewArtist')],
        [sg.Text('最終順位', size=(15, 1)), sg.InputText(default_text=str(select_LastRank),key='NewLastRank')],
        [sg.Text('最終ランクインNo', size=(15, 1)), sg.InputText(default_text=str(select_LastNumber),key='NewLastNumber')],
        [sg.Text('ランクイン回数', size=(15, 1)), sg.InputText(default_text=str(select_Onchart),key='NewOnchart')],
        [sg.Button('確定'),sg.Button('戻る')]
        ]

        window2 = sg.Window('楽曲データ修正', layout,finalize=True,icon='FM-BACS.ico')
        GetData.WriteLog(0,"管理者画面：楽曲データ修正画面起動")

        while True:
            event, values = window2.read()
            if event is None:
            # print('exit')
                break

            elif event == '確定':

                # 変換
                Title = values['NewTitle']
                Artist = values['NewArtist']
                LastRank = values['NewLastRank']
                LastNumber = values['NewLastNumber']
                Onchart = values['NewOnchart']
                Unique_id = GetData.generate_unique_id(Title,Artist)

                if FormatCheck(Title,Artist,LastRank,LastNumber,Onchart):

                    result = sg.popup_ok_cancel('曲名：'+str(Title)+'\nアーティスト：'+str(Artist)+'\n最終順位：'+str(LastRank)+'位\n最終ランクインNo：'+str(LastNumber)+'回\nランクイン回数：'+str(Onchart)+'回\nに更新しますか？',no_titlebar=True)
                    if result == 'OK':
                        deleterow(select_Title,select_Artist)
                        insert_music_data(Title,Artist,LastRank,LastNumber,Onchart,Unique_id)
                        sg.popup('データベースに書き込みました',no_titlebar=True)
                        reload()
                        # テーブルのデータを更新
                        table_data = df.values.tolist()
                        # テーブルを更新
                        window['-TABLE-'].update(values=table_data)
                        GetData.WriteLog(5,"管理者画面：楽曲データ修正画面終了")
                        window2.close()
                    else:
                        continue

                else:
                    continue

            elif event == '戻る':
                result = sg.popup_ok_cancel('変更した内容は保存されません\n終了しますか？',no_titlebar=True)
                if result == 'OK':
                    GetData.WriteLog(5,"管理者画面：楽曲データ修正画面終了")
                    window2.close()
                else:
                    continue

            

    elif event == '削除':

        GetData.WriteLog(1,"管理者画面：削除を選択")
    
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
                GetData.WriteLog(1,"管理者画面："+selected_row_Title+"/"+selected_row_Artist+"の削除を承認")
                # sg.popup('削除しました')
                # 選択された行を削除
                deleterow(selected_row_Title,selected_row_Artist)
                reload()
                # テーブルのデータを更新
                table_data = df.values.tolist()
                # テーブルを更新
                window['-TABLE-'].update(values=table_data)
            elif result == 'Cancel':
                GetData.WriteLog(1,"管理者画面："+selected_row_Title+"/"+selected_row_Artist+"の削除をキャンセル")
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
            import shutil
            user_folder = os.path.expanduser("~")
            folder = os.path.join(user_folder, "Downloads")
            shutil.copy('MasterLog.log', folder)
            os.chdir(os.path.dirname(sys.argv[0]))
            sg.popup('操作ログを出力します。\nエラーログはありませんでした。',no_titlebar=True) 


    elif event == 'csv':
        GetData.WriteLog(1,"管理者画面：csv復元を選択")
        result = sg.popup_ok_cancel('csvを2120回から復元しますか？\n復元するともとには戻せません',title='csv復元確認',no_titlebar=True)
        if result == 'OK':
            GetData.WriteLog(1,"管理者画面：csv復元を承認")
            import CreateDB2
            import LearningRank
            sg.popup('処理が終了しました。システムを終了します',no_titlebar=True)
            GetData.WriteLog(5,"管理者画面：FM BACS終了\n")
            sys.exit()
        else:
            GetData.WriteLog(1,"管理者画面：csv復元をキャンセル")
            break

    elif event == 'setting':
        os.startfile('C:\FM Besthit Automatic Create System')

    elif event == 'end':
        GetData.WriteLog(5,"管理者画面：管理者画面終了")
        window.close()


# ウィンドウを閉じる
# window.close()