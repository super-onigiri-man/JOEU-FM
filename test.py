try:
    #コードをここに書くと発生する例外がファイルに出力される。
    print(あああ) # あああ　は文字列変数を指定していなければ"(ダブルクォート）で囲む必要がある。
except Exception as e:
    import traceback
    with open('error.log', 'a') as f:
        traceback.print_exc( file=f)