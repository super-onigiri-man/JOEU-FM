import PySimpleGUI as sg

# テーブルのデータ
data = [['John', 25], ['Doe', 30], ['Smith', 35]]

# レイアウト
layout = [
    [sg.Table(values=data, headings=['Name', 'Age'], key='-TABLE-', enable_events=True)],
    [sg.Button('Get Selected Row')]
]

# ウィンドウの生成
window = sg.Window('Table Example', layout)

# イベントループ
while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break
    elif event == 'Get Selected Row':
        # 選択した行のインデックスを取得
        selected_row_index = values['-TABLE-'][0]
        # テーブルのデータから選択した行を取得
        selected_row_data = data[selected_row_index]
        print("Selected Row Data:", selected_row_data)

# ウィンドウを閉じる
window.close()