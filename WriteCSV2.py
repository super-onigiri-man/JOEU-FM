import sqlite3
import csv
import PySimpleGUI as sg
import GetData

# データベースへの接続
conn = sqlite3.connect('test.db')
cursor = conn.cursor()

# データベースから情報を取得
cursor.execute("SELECT * FROM music_master")
data = cursor.fetchall()

# 重複チェックとデータ加工
unique_ids = {}
for row in data:
    unique_id = GetData.generate_unique_id(row[0],row[1])  # Unique_idのインデックス
    if unique_id in unique_ids:
        # 重複があった場合
        existing_data = list(unique_ids[unique_id])
        if row[4] > unique_ids[unique_id][4]:  # Last_Numberが大きい方を選択
            unique_ids[unique_id] = list(row)  # Last_Numberが大きい方を採用
            unique_ids[unique_id][5] = existing_data[5] + row[5] #Onchartを足す
        
        else:
            existing_data[5] = unique_ids[unique_id][5] + row[5]
    else:
        # 重複なし
        unique_ids[unique_id] = row

# データをリストに変換
processed_data = list(unique_ids.values())

# CSVファイルに書き込む
with open('楽曲データ.csv', 'w', newline='',encoding='UTF-8') as csvfile:
    writer = csv.writer(csvfile)
    # ヘッダー行を書き込む
    writer.writerow(['Title', 'Artist', 'Score', 'Last_Rank', 'Last_Number', 'On_Chart', 'Unique_id'])
    # データ行を書き込む
    writer.writerows(processed_data)

# 接続を閉じる
conn.close()

sg.popup("データがCSVファイルに書き込まれました。\nデータの整合性を確実にするため、ソフトを再起動してください",no_titlebar=True)