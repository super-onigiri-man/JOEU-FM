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
        create_table_sql = """CREATE TABLE IF NOT EXISTS music_master (
            Title TEXT,
            Artist TEXT,
            Score DOUBLE,
            Last_Rank INT,
            Last_Number INT,
            On_Chart INT,
            PRIMARY KEY (Title, Artist)
        );"""
        cursor.execute(create_table_sql)  # テーブル作成

        # CSVファイルからデータを読み込んでバルクインサート
        with open(csv_file, 'r', encoding='utf-8') as f:
            csv_reader = csv.reader(f)
            data = [tuple(row) for row in csv_reader]  # CSVデータをタプルのリストに変換
            insert_sql = "INSERT OR IGNORE INTO music_master VALUES (?, ?, ?, ?, ?, ?);"  # バルクインサート用のSQL文
            cursor.executemany(insert_sql, data)  # バルクインサート実行


except Exception as e:
    # エラーログを出力
    with open('error.log', 'a') as f:
        traceback.print_exc(file=f)
    sg.popup_error("DBを作成できませんでした。\n システムを終了します", title="エラー",no_titlebar=True)
    sys.exit()