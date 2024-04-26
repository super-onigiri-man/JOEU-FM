import csv
import sqlite3
import PySimpleGUI as sg
import requests

dbname = ('test.db')#データベース名.db拡張子で設定
conn = sqlite3.connect(dbname, isolation_level=None)#データベースを作成、自動コミット機能ON
cursor = conn.cursor() #カーソルオブジェクトを作成

try:
    url='https://github.com/super-onigiri-man/JOEU-FM/blob/main/%E6%A5%BD%E6%9B%B2%E3%83%87%E3%83%BC%E3%82%BF.csv'
    filename='楽曲データ.csv'

    urlData = requests.get(url).content

    with open(filename ,mode='wb') as f: # wb でバイト型を書き込める
        f.write(urlData)

except Exception as e:
        sg.popup_error("楽曲データ.csvを取得できませんでした。\n 管理者へお問い合わせください",title="エラー")
try:
    sql = """CREATE TABLE music_master (
        Title TEXT,
        Artist TEXT,
        Score DOUBLE,
        Last_Rank INT,
        Last_Number INT,
        On_Chart INT,
        PRIMARY KEY (Title, Artist)
    );"""
    #↑DBのフォーマット設定（別のところに書いておきます）
    cursor.execute(sql)#executeコマンドでSQL文を実行
    conn.commit()#データベースにコミット(Excelでいう上書き保存。自動コミット設定なので不要だが一応・・)
    with open('楽曲データ.csv',encoding='utf-8') as f:
        # CSVリーダーオブジェクトを作成
        csv_reader = csv.reader(f)
        for row in csv_reader:
            # テーブルにデータを挿入
            cursor.execute("INSERT INTO music_master VALUES (?,?,?,?,?,?)", (row[0], row[1],row[2],row[3],row[4],row[5]))

except Exception as e:
        sg.popup_error("DBを作成できませんでした。\n「楽曲データ.csv」を正しい位置に配置してください",title="エラー")
            