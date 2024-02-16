import csv
import sqlite3

dbname = (r'C:\Users\wiiue\JOEU-FM\test.db')#データベース名.db拡張子で設定
conn = sqlite3.connect(dbname, isolation_level=None)#データベースを作成、自動コミット機能ON
cursor = conn.cursor() #カーソルオブジェクトを作成
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
with open(r'C:\Users\wiiue\JOEU-FM\楽曲データ.csv',encoding='UTF-8') as f:
    # CSVリーダーオブジェクトを作成
    csv_reader = csv.reader(f)
    for row in csv_reader:
        # テーブルにデータを挿入
        cursor.execute("INSERT INTO music_master VALUES (?,?,?,?,?,?)", (row[0], row[1],row[2],row[3],row[4],row[5]))