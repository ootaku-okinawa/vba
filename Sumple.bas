Sub WriteLog(LogMessage As String, Optional IsError As Boolean = False)
    Dim ws As Worksheet
    Dim lastRow As Long
    Const SheetName As String = "ログ"
    
    ' ログを出力するシートを設定
    On Error Resume Next ' シートが存在しない場合のエラーを無視
    Set ws = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0 ' エラーハンドリングを通常に戻す
    
    ' シートが存在しない場合、新しいシートを作成
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SheetName
    End If
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 最終行が空でなければ、次の行に移動
    If Not IsEmpty(ws.Cells(lastRow, 1).Value) Then lastRow = lastRow + 1
    
    ' ログメッセージと現在の日時を出力
    With ws.Cells(lastRow, 1)
        .Value = Now & " - " & LogMessage
        ' エラーログの場合は文字色を赤にする
        If IsError Then .Font.Color = RGB(255, 0, 0)
    End With
End Sub



Sub SampleProcess()
    ' 正常なログ出力
    WriteLog "処理を開始しました。"
    
    ' 何らかの問題が発生した場合
    WriteLog "エラーが発生しました。", True
    
    ' 処理完了のログを出力
    WriteLog "処理が完了しました。"
End Sub


Sub SampleProcessWithErrorHandling()
    On Error GoTo ErrorHandler ' エラーハンドラを設定
    
    ' ここに通常の処理を記述
    Debug.Print 1 / 0 ' ゼロ除算エラーを意図的に発生させる
    
    Exit Sub ' 正常終了時はエラーハンドラを回避
    
ErrorHandler:
    ' エラー発生時の処理
    WriteLog "エラーが発生しました。エラー番号: " & Err.Number & ", 説明: " & Err.Description, True
    Resume Next ' 次の行で処理を再開（または、必要に応じて適切なエラー処理を行う）
End Sub
