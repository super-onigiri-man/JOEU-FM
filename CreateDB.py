import csv
import sqlite3
import PySimpleGUI as sg
import traceback
import sys
import pandas as pd
import os

dbname = 'test.db'  # データベース名
csv_file = '楽曲データ.csv'  # CSVファイル名

conn = sqlite3.connect(dbname, isolation_level=None)  # データベースに接続（自動コミット機能ON）
cursor = conn.cursor()  # カーソルオブジェクトを作成

try:

    layout = [
        [sg.Text('読み込み中...', size=(15, 1)), sg.ProgressBar(3, orientation='h', size=(20, 20), key='progressbar')],
        [sg.Text(key = 'progmsg')]
    ]

    window = sg.Window('システム起動中', layout,finalize=True)

    def update_progress_bar(progress_bar,progmsg, value,msg):
        progress_bar.update_bar(value)
        progmsg.update(msg)
        window.refresh()

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

    update_progress_bar(window['progressbar'],window['progmsg'], 1,'主キー重複確認中')

    # CSVファイルを読み込む
    df = pd.read_csv('楽曲データ.csv', names=['曲名', 'アーティスト', '得点', '前回順位', '前回ランクインNo', 'ランクイン回数', '独自ID'],header=0)

    # 重複を処理するための関数
    def resolve_duplicates(group):
        # '前回ランクインNo'（row[4]）が最大の行を残す
        max_row = group.loc[group['前回ランクインNo'].idxmax()]
        # 'ランクイン回数'（row[5]）を合計する
        total_row5 = group['ランクイン回数'].sum()
        
        # 更新する値を設定
        max_row['ランクイン回数'] = total_row5
        
        return max_row

    # '独自ID'（row[6]）を基準にグループ化して重複を解決
    df_resolved = df.groupby('独自ID').apply(resolve_duplicates).reset_index(drop=True)

    # 結果を新しいCSVファイルに保存
    df_resolved.to_csv('楽曲データ.csv',index=False)

    update_progress_bar(window['progressbar'],window['progmsg'], 2,'DB作成中')

    with open('楽曲データ.csv',encoding='utf-8') as f:
        # CSVリーダーオブジェクトを作成
        next(f)
        csv_reader = csv.reader(f)
        for row in csv_reader:
            # テーブルにデータを挿入
            cursor.execute("INSERT INTO music_master VALUES (?,?,?,?,?,?,?)", (row[0],row[1],row[2],row[3],row[4],row[5],row[6]))

    update_progress_bar(window['progressbar'],window['progmsg'], 3,'DB作成中')

    window.close()

except Exception as e:
    # エラーログを出力
    with open('error.log', 'a') as f:
        traceback.print_exc(file=f)
    sg.popup_error("DBを作成できませんでした。\n システムを終了します", title="エラー",no_titlebar=True)
    sys.exit()