# import csv

# def generate_unique_id(song_title, artist_name):
#     # 曲名とアーティスト名の頭3文字を取得
#     song_prefix = song_title[:3]
#     artist_prefix = artist_name[:3]

#     # 頭3文字をUnicodeコードポイントに変換
#     song_code_points = [ord(char) for char in song_prefix]
#     artist_code_points = [ord(char) for char in artist_prefix]

#     # コードポイントを文字列に変換して結合
#     song_code_str = ''.join(f'{cp:04x}' for cp in song_code_points)
#     artist_code_str = ''.join(f'{cp:04x}' for cp in artist_code_points)

#     # 独自IDを生成
#     unique_id = song_code_str + artist_code_str

#     return unique_id

# with open('楽曲データ.csv',encoding='utf-8') as f:
#         # CSVリーダーオブジェクトを作成
#         csv_reader = csv.reader(f)
#         next(f)
#         for row in csv_reader:
#                 id = generate_unique_id(row[0],row[1])
#                 print(row[0]+','+row[1]+','+row[2]+','+row[3]+','+row[4]+','+row[5]+','+id)


import pandas as pd

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
df_resolved.to_csv('楽曲データ.csv', index=False)