import csv
import sqlite3
import PySimpleGUI as sg
import traceback
import sys

dbname = 'test.db'  # データベース名
csv_file = '楽曲データ.csv'  # CSVファイル名

conn = sqlite3.connect(dbname, isolation_level=None)  # データベースに接続（自動コミット機能ON）
cursor = conn.cursor()  # カーソルオブジェクトを作成

try:
        
        # テーブル作成のSQL文
    sql = """CREATE TABLE music_master (
        Title TEXT,
        Artist TEXT,
        Score DOUBLE,
        Last_Rank INT,
        Last_Number INT,
        On_Chart INT,
        Unique_id TEXT,
        PRIMARY KEY (Unique_id)
        );"""
    #↑DBのフォーマット設定（別のところに書いておきます）
    cursor.execute(sql)#executeコマンドでSQL文を実行
    conn.commit()#データベースにコミット(Excelでいう上書き保存。自動コミット設定なので不要だが一応・・)
    with open('result.csv',encoding='utf-8') as f:
        # CSVリーダーオブジェクトを作成
        csv_reader = csv.reader(f)
        for row in csv_reader:
            # テーブルにデータを挿入
            cursor.execute("INSERT INTO music_master VALUES (?,?,?,?,?,?,?)", (row[0],row[1],row[2],row[3],row[4],row[5],row[6]))


except Exception as e:
    # エラーログを出力
    with open('error.log', 'a') as f:
        traceback.print_exc(file=f)
    sg.popup_error("DBを作成できませんでした。\n システムを終了します", title="エラー",no_titlebar=True)
    sys.exit()