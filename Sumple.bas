Sub SelectFolderAndDisplayPath()
    Dim fd As FileDialog
    Dim selectedFolder As String
    Dim targetCell As Range

    ' フォルダ選択ダイアログの初期化
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    If fd.Show = -1 Then ' ユーザーがOKを押した場合
        selectedFolder = fd.SelectedItems(1) ' 選択されたフォルダのパスを取得
        
        ' フルパスを表示したいセルを指定（例: シート1のA1セル）
        Set targetCell = ThisWorkbook.Sheets("Sheet1").Range("A1")
        
        ' 選択されたフォルダのフルパスをセルに設定
        targetCell.Value = selectedFolder
        
        ' オプション: セルの編集を可能にする
        targetCell.Locked = False
    Else
        MsgBox "No folder selected."
    End If

    Set fd = Nothing
End Sub
