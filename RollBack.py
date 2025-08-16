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

try:
    GetData.WriteLog(3,"ロールバック：ロールバックを実行")
    Oriconday = GetData.OriconTodays()
    print(Oriconday)
    # path = 'Rank_BackUp/'+str(Oriconday)+'ベストヒットランキング.xlsx'
    count = 0
    count2 = 0
    flag = False
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
        if flag :
            break

        if count == 0:
            rollbackday = Oriconday
        elif count == 1:
            rollbackday = Oriconday - datetime.timedelta(days=7)
        else:
            rollbackday = rollbackday - datetime.timedelta(days=1)

        filepath = 'Rank_BackUp/'+str(rollbackday)+'ベストヒットランキング.xlsx' 

        if os.path.isfile(filepath): #ファイルがあったら

            while True:
                if count2 == 0:
                    rollbackday2 = rollbackday - datetime.timedelta(days=1)
                else:
                    rollbackday2 = rollbackday2 - datetime.timedelta(days=1)

                filepath2 = 'Rank_BackUp/'+str(rollbackday2)+'ベストヒットランキング.xlsx'
                print(os.path.isfile(filepath2)) 
                if os.path.isfile(filepath2):
                    flag = True
                    break

                count2 = count2 + 1

        count = count + 1

except Exception as e:
    import traceback
    with open('error.log', 'a') as f:
        traceback.print_exc( file=f)
    GetData.WriteLog(4,"ロールバック：ロールバック処理に失敗")
    sg.popup_error('ロールバック処理に失敗しました',no_titlebar=True)

finally:
    print('DB関連処理')
    print(filepath)
    print(filepath2)
    RevisionRank.RevisionRank(filepath2)
    os.remove(filepath)
    # RevisionRank.RevisionRank()