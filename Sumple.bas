Sub SelectFolderAndDisplayPath(cellAddress As String)
    Dim fd As FileDialog
    Dim selectedFolder As String
    Dim targetCell As Range

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    If fd.Show = -1 Then
        selectedFolder = fd.SelectedItems(1)
        
        Set targetCell = ThisWorkbook.Sheets("Sheet1").Range(cellAddress)
        targetCell.Value = selectedFolder
        targetCell.Locked = False ' セルのロックを解除（必要に応じて）
    Else
        MsgBox "No folder selected."
    End If
End Sub


' ボタン1用のラッパーマクロ
Sub Button1_Click()
    SelectFolderAndDisplayPath "A1"
End Sub

' ボタン2用のラッパーマクロ
Sub Button2_Click()
    SelectFolderAndDisplayPath "B1"
End Sub
