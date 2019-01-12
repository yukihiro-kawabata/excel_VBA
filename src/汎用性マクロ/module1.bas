'##########################################################
'
' ディレクトリ内にあるファイル全てを取得する
'
' == 仕様 ===
' フォルダPathをエクセル内に記載して
' 記載したセルを選択した状態で実行すると
' 記載セルから下2行目にファイル一覧が書き出される
'
'##########################################################
Sub getFileListInDir()

    '選択場所を取得してディレクトリPathを取得する
    dirPath = ActiveSheet.Cells(Selection.Row, Selection.Column) & "\"

    '次の行から書き出す
    '実際には空欄を含むので、次の次から書き出す
    cnt = Selection.Row + 1

    'ファイルをすべて書き出す
    filePath = Dir(dirPath & "*")
    Do While filePath <> ""
        cnt = cnt + 1
        Cells(cnt, 1) = filePath
        filePath = Dir()
    Loop
End Sub
