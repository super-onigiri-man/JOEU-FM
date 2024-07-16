import os
import sys
import PySimpleGUI as sg
import GetData
import sqlite3
import datetime
import RevisionRank

dbname = ('test.db')#データベース名.db拡張子で設定
conn = sqlite3.connect(dbname, isolation_level=None)#データベースを作成、自動コミット機能ON
cursor = conn.cursor() #カーソルオブジェクトを作成

Oriconday = GetData.OriconTodays()
path = 'Rank_BackUp/'+str(Oriconday)+'ベストヒットランキング.xlsx'
count = 0
is_file = os.path.isfile(path)
if is_file:
    sg.popup('今週のデータはすでに生成済みです。\n今週のデータを作り直すにはデータベースの変更が必要です',no_titlebar=True)

    result = sg.popup_yes_no("今週のランキングデータを消してランキング生成しますか？", title="確認",no_titlebar=True)

    if result == 'Yes':
        try:
            # import CreateDB
            # 最終回の確認
            # クエリの実行
            query = "SELECT MAX(Last_Number) FROM music_master;"
            cursor.execute(query)
            # 結果の取得
            max_last_number = cursor.fetchone()[0]

            max_last_number = int(max_last_number)

            cursor.execute("DELETE FROM music_master WHERE Last_Number = ?;",(max_last_number,)) 

            while True:

                if count == 0:
                    rollbackday = Oriconday - datetime.timedelta(days=7)
                
                else:
                    rollbackday = rollbackday - datetime.timedelta(days=7)
                filepath = 'Rank_BackUp/'+str(rollbackday)+'ベストヒットランキング.xlsx' 
                if os.path.isfile(filepath): #ファイルがあったら
                    RevisionRank.RevisionRank(filepath)
                    os.remove('Rank_BackUp/'+str(Oriconday)+'ベストヒットランキング.xlsx')
                    break
        
        except Exception as e:
            import traceback
            with open('error.log', 'a') as f:
                traceback.print_exc( file=f)
            sg.popup_error('ロールバック処理に失敗しました',no_titlebar=True)
            

    else:
        sg.popup('システムを終了します',no_titlebar=True)
        sys.exit()
else:
    pass # パスが存在しないかファイルではない