import requests
from bs4 import BeautifulSoup #スクレイピング（取得）用
import datetime #日付計算用
import mojimoji
import openpyxl
import sqlite3
import pandas as pd
import PySimpleGUI as sg

dbname = ('test.db')
conn = sqlite3.connect(dbname, isolation_level=None)#データベースを作成、自動コミット機能ON
cursor = conn.cursor() #カーソルオブジェクトを作成
#scoreを元にしたオリコンシングル・オリコンデジタル・ビルボード総合ランキング

# オリコン週間ランキング用
OriconWeekData = []
# オリコンデジタルランキング用
OriconDigitalData = []
# ビルボードランキング用
BillboardData = []
# 明屋書店ランキング用
HaruyaData = []



def OriconTodays():
    # 今日の日付を求める
    d_today = datetime.date.today()
    # 今日の曜日を求める
    todayweek = datetime.date.today().weekday()
    #print(d_today, todayweek)

    #オリコンの発表は毎週水曜日のため、火曜日までは先週のランキングを表示
    #日付は来週の水曜日付となる。
    if (todayweek == 0):  # 今日が月曜日(先週(今週月曜日)のランキング表示)
        Oriconday = d_today
    elif (todayweek == 1):# 今日が火曜日(先週(今週月曜日)のランキング表示)
        Oriconday = d_today - datetime.timedelta(days=1)
    elif (todayweek == 2):# 今日が水曜日
        Oriconday = d_today + datetime.timedelta(days=5)
    elif (todayweek == 3):# 今日が木曜日
        Oriconday = d_today + datetime.timedelta(days=4)
    elif (todayweek == 4):# 今日が金曜日
        Oriconday = d_today + datetime.timedelta(days=3)
    elif (todayweek == 5):# 今日が土曜日
        Oriconday = d_today + datetime.timedelta(days=2)
    elif (todayweek == 6):# 今日が日曜日
        Oriconday = d_today + datetime.timedelta(days=1)

    return Oriconday

def OriconLastWeek():
    # 今日の日付を求める
    d_today = datetime.date.today()
    # 今日の曜日を求める
    todayweek = datetime.date.today().weekday()
    # print(d_today, todayweek)

    # オリコンの発表は毎週水曜日のため、火曜日までは先週のランキングを表示
    # 日付は来週の月曜日付となる。
    if (todayweek == 0):  # 今日が月曜日(先週(今週月曜日)のランキング表示)
        Oriconday = d_today - datetime.timedelta(days=7)
    elif (todayweek == 1):  # 今日が火曜日(先週(今週月曜日)のランキング表示)
        Oriconday = d_today - datetime.timedelta(days=8)
    elif (todayweek == 2):  # 今日が水曜日
        Oriconday = d_today - datetime.timedelta(days=2)
    elif (todayweek == 3):  # 今日が木曜日
        Oriconday = d_today - datetime.timedelta(days=3)
    elif (todayweek == 4):  # 今日が金曜日
        Oriconday = d_today - datetime.timedelta(days=4)
    elif (todayweek == 5):  # 今日が土曜日
        Oriconday = d_today - datetime.timedelta(days=5)
    elif (todayweek == 6):  # 今日が日曜日
        Oriconday = d_today - datetime.timedelta(days=6)

    return Oriconday


def OriconSelectWeek(SelectDay):
    # 日付の入力を促す
    # date = input("2020年8月3日以降の日付を入力してください (YYYY-MM-DD): ")

    # 入力された日付をdatetimeオブジェクトに変換
    dt = datetime.datetime.strptime(SelectDay, "%Y-%m-%d")
    dt = dt.date()
    # 曜日を取得
    weekday = dt.weekday()

    # オリコンの発表は毎週水曜日のため、火曜日までは先週のランキングを表示
    # 日付は来週の月曜日付となる。
    if (weekday == 0):  # 今日が月曜日
        Oriconday = dt
    elif (weekday == 1):  # 今日が火曜日
        Oriconday = dt - datetime.timedelta(days=1)
    elif (weekday == 2):  # 今日が水曜日
        Oriconday = dt - datetime.timedelta(days=2)
    elif (weekday == 3):  # 今日が木曜日
        Oriconday = dt - datetime.timedelta(days=3)
    elif (weekday == 4):  # 今日が金曜日
        Oriconday = dt - datetime.timedelta(days=4)
    elif (weekday == 5):  # 今日が土曜日
        Oriconday = dt - datetime.timedelta(days=5)
    elif (weekday == 6):  # 今日が日曜日
        Oriconday = dt - datetime.timedelta(days=6)

    return Oriconday

def OriconWeekRank(Oriconday):#オリコン週間ランキング
    try:
        #1位から10位
        load_url = "https://www.oricon.co.jp/rank/js/w/" + str(Oriconday) + "/"
        html = requests.get(load_url)
        soup = BeautifulSoup(html.text, "html.parser")
        links = soup.find(class_="content-rank-main").find_all('h2',class_='title') #曲名
        artist = soup.find(class_="content-rank-main").find_all('p',class_='name') #アーティスト名
        score = 6.0 #独自スコア
        rank = 1 #ランキング
        print(str(Oriconday) + "付けオリコン週間シングルランキング")

        index = 0
        for link, artist in zip(links, artist):
            if '/' in link.text:
                # If yes, split the title and create two separate entries in the array
                titles = link.text.split('/')
                for title in titles:
                    mojimoji.zen_to_han(title.strip())  # Strip to remove leading/trailing whitespaces
                    mojimoji.zen_to_han(artist.text)
                    OriconWeekData.append([title.strip(), artist.text, "{:.1f}".format(score)])

                    if title == titles[-1]:
                        rank = rank + 1
                        score = score - 0.3
            else:
                # If no slash, just add the entry to the array
                mojimoji.zen_to_han(link.text)
                mojimoji.zen_to_han(artist.text)
                OriconWeekData.append([link.text, artist.text, "{:.1f}".format(score)])
                rank = rank + 1
                score = score - 0.3

        # 壊れたときの表示用
        # if rank != 10:#10位以下（１ケタの場合）なら（点数の位置を揃えるため）
        #   print(" " + str(rank) + "位 " + "{:.1f}　 ".format(score) + link.text + "/" + artist.text)
        # else: #10位（２ケタの場合）なら
        #   print(str(rank) + "位 " + "{:.1f}　 ".format(score) + link.text + "/" + artist.text)

        # print(OriconWeekData)

        #11位から20位
        load_url = "https://www.oricon.co.jp/rank/js/w/" + str(Oriconday) + "/p/2/"
        html = requests.get(load_url)
        soup = BeautifulSoup(html.text, "html.parser")
        links = soup.find(class_="content-rank-main").find_all('h2', class_='title')  # 曲名
        artist = soup.find(class_="content-rank-main").find_all('p', class_='name')  # アーティスト名
        for link, artist in zip(links, artist):
            if '/' in link.text:
            # 「楽曲A/楽曲B」のような表現方法の場合
                titles = link.text.split('/') #"/"で２曲に分ける
                for title in titles:
                    mojimoji.zen_to_han(title.strip())  # Strip to remove leading/trailing whitespaces
                    mojimoji.zen_to_han(artist.text)
                    OriconWeekData.append([title.strip(), artist.text, "{:.1f}".format(score)])

                    if title == titles[-1]:
                        rank = rank + 1
                        score = score - 0.3
            else:
                # If no slash, just add the entry to the array
                mojimoji.zen_to_han(link.text)
                mojimoji.zen_to_han(artist.text)
                OriconWeekData.append([link.text, artist.text, "{:.1f}".format(score)])
                rank = rank + 1
                score = score - 0.3
                # 壊れたときの表示用
            # print(str(rank) + "位 " + "{:.1f}　 ".format(score) + link.text + "/" + artist.text)

    except Exception as e:
        sg.popup_error("「オリコン週間ランキング」が取得できませんでした",title="エラー")



def OriconDigitalRank(Oriconday):#オリコンデジタルシングルランキング

    try:
        # 1位から10位
        load_url = "https://www.oricon.co.jp/rank/dis/w/" + str(Oriconday) + "/"
        html = requests.get(load_url)
        soup = BeautifulSoup(html.text, "html.parser")
        links = soup.find(class_="content-rank-main").find_all('h2', class_='title')
        artist = soup.find(class_="content-rank-main").find_all('p', class_='name')  # アーティスト名
        rank = 1
        score = float(6.0)
        print(str(Oriconday) + "付けオリコン週間デジタルシングルランキング")
        for link, artist in zip(links, artist):
        # 壊れたときの表示用
            # if rank != 10:#10位以下（１ケタの場合）なら（点数の位置を揃えるため）
            #     print(" " + str(rank) + "位 " + "{:.1f}　 ".format(score) + link.text + "/" + artist.text)
            # else: #10位（２ケタの場合）なら
            #     print(str(rank) + "位 " + "{:.1f}　 ".format(score) + link.text + "/" + artist.text)
            mojimoji.zen_to_han(link.text)
            mojimoji.zen_to_han(artist.text)
            OriconDigitalData.append([link.text,artist.text,"{:.1f}".format(score)])
            rank = rank + 1
            score = score - 0.3

        # 11位から20位
        load_url = "https://www.oricon.co.jp/rank/dis/w/" + str(Oriconday) + "/p/2/"
        html = requests.get(load_url)
        soup = BeautifulSoup(html.text, "html.parser")
        links = soup.find(class_="content-rank-main").find_all('h2', class_='title')
        artist = soup.find(class_="content-rank-main").find_all('p', class_='name')  # アーティスト名
        for link, artist in zip(links, artist):
            OriconDigitalData.append([link.text,artist.text,"{:.1f}".format(score)])
            # print(str(rank) + "位 " + "{:.1f}　 ".format(score) + link.text + "/" + artist.text)
            rank = rank + 1
            score = score - 0.3

    except Exception as e:
        sg.popup_error("「オリコンデジタルランキング」が取得できませんでした",title="エラー")


def BillboadRank(Oriconday):#ビルボードJAPAN HOT100ランキング

    try:
        # オリコンの日付とビルボードの発表日の差を埋めるための計算
        Billday = Oriconday - datetime.timedelta(days=5)

        print(str(Billday) + "付けビルボードJAPAN HOT100ランキング")

        #URL(ここを変更すると読み込まなくなります)
        url = 'https://www.billboard-japan.com/charts/detail?a=hot100&year='+str(Oriconday.year)+'&month='+str(Oriconday.month)+'&day='+str(Oriconday.day)
        #URLを取得してくる
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        songs = soup.find_all('p', class_='musuc_title') #曲名
        artists = soup.find_all('p', class_='artist_name') #アーティスト名
        score = 6.0 #基準点

        for i in range(20):#20回繰り替えす
            song = songs[i].text.strip()
            artist = artists[i].text.strip()
            mojimoji.zen_to_han(song)
            mojimoji.zen_to_han(artist)
            BillboardData.append([song,artist,format(score, '.1f')])
            # if i < 9:
            #   print(f" {i+1}位: {format(score, '.1f')} {song} / {artist}") #1位から9位までのランキング
            # else:
            #   print(f"{i+1}位: {format(score, '.1f')} {song} / {artist}") #10位から20位までのランキング
            score = score - 0.3 #scoreを-0.3する


    except Exception as e:
        sg.popup_error("「ビルボードランキング」が取得できませんでした",title="エラー")


def HaruyaRank(HaruyaPath):
    try:
        # Excelファイルを読み込む
        wb = openpyxl.load_workbook(HaruyaPath)
        ws = wb.active

        # データを2次元配列に挿入する
        for row in range(4, 24):
            song_names = mojimoji.zen_to_han(ws[f"D{row}"].value, kana=False).split('/')
            artist_name = mojimoji.zen_to_han(ws[f"C{row}"].value, kana=False)
            song_name = str(song_name)
            artist_name = str(artist_name)
            point = round(6.0 - ((row - 4) * 0.3), 2)  # 点数を計算する
            for song_name in song_names:
                if "/" in song_name:
                    song_name_a, song_name_b = song_name.split("/")
                    HaruyaData.append([song_name_a.strip(), artist_name, point])
                    HaruyaData.append([song_name_b.strip(), artist_name, point])
                else:
                    HaruyaData.append([song_name.strip(), artist_name, point])

        print('明屋書店データOK')

    except Exception as e:
        sg.popup_error("「明屋書店ランキング」が取得できませんでした",title="エラー")


def insertOriconWeekData():
   for entry in OriconWeekData:
    title, artist, score = entry
    # 既存のデータがあればScoreを足して更新、なければ新規追加
    cursor.execute('''
        INSERT INTO music_master (Title, Artist, Score, Last_Rank, Last_Number, On_Chart)
        VALUES (?, ?, ?, 0, 0, 0)
        ON CONFLICT(Title, Artist) DO UPDATE SET Score = Score + ?
    ''', (title,artist, score, score))

    conn.commit()

def insertOriconDegitalData():
   for entry in OriconDigitalData:
    title, artist, score = entry
    # 既存のデータがあればScoreを足して更新、なければ新規追加
    cursor.execute('''
        INSERT INTO music_master (Title, Artist, Score, Last_Rank, Last_Number, On_Chart)
        VALUES (?, ?, ?, 0, 0, 0)
        ON CONFLICT(Title, Artist) DO UPDATE SET Score = Score + ?
    ''', (title, artist, score, score))

    conn.commit()

def insertBillboardData():
   for entry in BillboardData:
    title, artist, score = entry
    # 既存のデータがあればScoreを足して更新、なければ新規追加
    cursor.execute('''
        INSERT INTO music_master (Title, Artist, Score,Last_Rank, Last_Number, On_Chart)
        VALUES (?, ?, ?, 0, 0, 0)
        ON CONFLICT(Title, Artist) DO UPDATE SET Score = Score + ?
    ''', (title, artist, score, score))

    conn.commit()

def insertHaruyaData():
   for entry in HaruyaData:
    title, artist, score = entry
    # 既存のデータがあればScoreを足して更新、なければ新規追加
    cursor.execute('''
        INSERT INTO music_master (Title, Artist, Score, Last_Rank, Last_Number, On_Chart)
        VALUES (?, ?, ?, 0, 0, 0)
        ON CONFLICT(Title, Artist) DO UPDATE SET Score = Score + ?
    ''', (title, artist, score, score))

    conn.commit()



def GetThisWeekRank(HaruyaPath):
    layout = [
        [sg.Text('読み込み中...', size=(15, 1)), sg.ProgressBar(72, orientation='h', size=(20, 20), key='progressbar')],
        [sg.Button('読み込み中止')]
    ]

    window = sg.Window('今週のランキング取得', layout,finalize=True)

    def update_progress_bar(progress_bar, value):
        progress_bar.update_bar(value)
        window.refresh()

    while True:
        
        update_progress_bar(window['progressbar'], 8)
        OriconTodays()
        update_progress_bar(window['progressbar'], 16)
        OriconWeekRank(OriconTodays())
        update_progress_bar(window['progressbar'], 24)
        OriconDigitalRank(OriconTodays())
        update_progress_bar(window['progressbar'], 32)
        BillboadRank(OriconTodays())
        update_progress_bar(window['progressbar'], 40)
        HaruyaRank(HaruyaPath)
        update_progress_bar(window['progressbar'], 48)
        insertOriconWeekData()
        update_progress_bar(window['progressbar'], 56)
        insertOriconDegitalData()
        update_progress_bar(window['progressbar'], 64)
        insertBillboardData()
        update_progress_bar(window['progressbar'], 72)
        insertHaruyaData()

        window.close()

        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'キャンセル':
                break
            


def GetLastWeekRank():
  
    layout = [
        [sg.Text('読み込み中...', size=(15, 1)), sg.ProgressBar(72, orientation='h', size=(20, 20), key='progressbar')],
        [sg.Button('読み込み中止')]
    ]

    window = sg.Window('先週のランキングを取得', layout,finalize=True)

    def update_progress_bar(progress_bar, value):
        progress_bar.update_bar(value)
        window.refresh()

    while True:
        
        update_progress_bar(window['progressbar'], 9)
        OriconLastWeek()
        update_progress_bar(window['progressbar'], 18)
        OriconWeekRank(OriconLastWeek())
        update_progress_bar(window['progressbar'], 27)
        OriconDigitalRank(OriconLastWeek())
        update_progress_bar(window['progressbar'], 36)
        BillboadRank(OriconLastWeek())
        update_progress_bar(window['progressbar'], 45)
        insertOriconWeekData()
        update_progress_bar(window['progressbar'], 54)
        insertOriconDegitalData()
        update_progress_bar(window['progressbar'], 63)
        insertBillboardData()
        update_progress_bar(window['progressbar'], 72)

        window.close()

        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'キャンセル':
                break
  

def GetSelectWeekRank(SelectDay):
  
    layout = [
        [sg.Text('読み込み中...', size=(15, 1)), sg.ProgressBar(72, orientation='h', size=(20, 20), key='progressbar')],
        [sg.Button('読み込み中止')]
    ]

    window = sg.Window('任意週のランキングを取得', layout,finalize=True)

    def update_progress_bar(progress_bar, value):
        progress_bar.update_bar(value)
        window.refresh()

    while True:
        
        update_progress_bar(window['progressbar'], 9)
        global OSW
        update_progress_bar(window['progressbar'], 18)
        OSW=OriconSelectWeek(SelectDay)
        update_progress_bar(window['progressbar'], 27)
        OriconDigitalRank(OSW)
        update_progress_bar(window['progressbar'], 36)
        BillboadRank(OSW)
        update_progress_bar(window['progressbar'], 45)
        insertOriconWeekData()
        update_progress_bar(window['progressbar'], 54)
        insertOriconDegitalData()
        update_progress_bar(window['progressbar'], 63)
        insertBillboardData()
        update_progress_bar(window['progressbar'], 72)

        window.close()

        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'キャンセル':
                break

def GetThisWeekDate():
   return OriconTodays()

def GetLastWeekDate():
   return OriconLastWeek()

def GetSelectWeekDate():
   return OSW