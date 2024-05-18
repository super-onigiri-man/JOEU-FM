import mojimoji
import openpyxl
import sqlite3
import datetime
import csv
import shutil
import PySimpleGUI as sg

dbname = ('test.db')
conn = sqlite3.connect(dbname, isolation_level=None)#データベースを作成、自動コミット機能ON
cursor = conn.cursor() #カーソルオブジェクトを作成

def WriteCSV(Oriconday):
    excel_file = 'Rank_BackUp/'+str(Oriconday)+'ベストヒットランキング.xlsx'
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active


    # データの処理と挿入
    for row in range(6, 45, 2):
        title = mojimoji.zen_to_han(sheet['E' + str(row)].value, kana=False)
        artist = mojimoji.zen_to_han(sheet['F' + str(row)].value, kana=False)
        rank = sheet['B' + str(row)].value
        on_chart = sheet['D'+str(row)].value
        this_number = mojimoji.zen_to_han(sheet['B3'].value,kana=False)
        this_number = this_number.replace('No.','')
        unique_id = sheet['F' + str(row+1)].value

        if "再" in str(sheet['C' + str(row)].value) or "圏外" in str(sheet['C' + str(row)].value) :
            last_number = int(this_number)
        elif "初" in str(sheet['C' + str(row)].value):
            on_chart = 1
            last_number = this_number
        else:
            last_number = this_number

        # データベースに挿入
        cursor.execute('''INSERT INTO music_master
            (Title, Artist, Score, Last_Rank, Last_Number, On_chart,Unique_id)
            VALUES (?, ?, 0.0, ?, ?, ? ,?)
            ON CONFLICT(Unique_id) DO UPDATE SET Last_Number = ?,Last_Rank = ?,On_Chart = ?''',
                (title, artist, rank, last_number, on_chart,unique_id,last_number,rank,on_chart))

    cursor.execute('UPDATE music_master SET Score = 0;') #Scoreの値を全消し

    cursor.execute('''DELETE FROM music_master WHERE Last_Number = '' OR Last_Number = 0;''') #Last_Numberが空文字もしくは0の場合はその曲を消す（ランクインしなかった曲）

    # コミット
    conn.commit()

    with open('楽曲データ.csv', 'w', newline='',encoding='UTF-8') as f:
        # CSVライターオブジェクトを作成
        csv_writer = csv.writer(f)

        # データを取得
        cursor.execute("SELECT * FROM music_master")
        rows = cursor.fetchall()

        # データをCSVに書き込む
        for row in rows:
            csv_writer.writerow(row)

    sg.popup_ok('CSVファイルを更新しました',no_titlebar=True)

    