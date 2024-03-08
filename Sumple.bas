Option Explicit

' 定数定義
Const LOG_SHEET_NAME As String = "ログシート"
Const ERROR_LOG_SHEET_NAME As String = "エラーログシート"

' 通常のログを書き込むための関数
Sub WriteLog(Message As String)
    Dim ws As Worksheet
    ' ログシートを取得または作成
    Set ws = GetOrCreateSheet(ThisWorkbook, LOG_SHEET_NAME)
    ' ログデータをシートに書き込み
    WriteToSheet ws, Array(Now, Message)
End Sub

' エラーログを書き込むための関数
Sub WriteErrorLog(LogLevel As String, Message As String)
    Dim ws As Worksheet
    ' エラーログシートを取得または作成
    Set ws = GetOrCreateSheet(ThisWorkbook, ERROR_LOG_SHEET_NAME)
    ' エラーログデータをシートに書き込み
    WriteToSheet ws, Array(Now, LogLevel, Message), LogLevel
End Sub

' 指定された名前のシートを取得または作成する関数
Function GetOrCreateSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    ' シートの存在確認
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    ' シートが存在しない場合は新規作成
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
        ' シートの初期設定実行
        SetupSheet ws, sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

' シートの初期設定を行う関数
Sub SetupSheet(ws As Worksheet, sheetName As String)
    ' シート名に応じてヘッダを設定
    If sheetName = LOG_SHEET_NAME Then
        ws.Range("A1:B1").Value = Array("日時", "ログ内容")
    ElseIf sheetName = ERROR_LOG_SHEET_NAME Then
        ws.Range("A1:C1").Value = Array("日時", "ログレベル", "ログ内容")
    End If
    ' オートフィルターの設定
    ws.Rows("1:1").AutoFilter
End Sub

' 指定されたシートにデータを書き込む関数
Sub WriteToSheet(ws As Worksheet, data As Variant, Optional LogLevel As String = "")
    Dim lastRow As Long
    ' 最終行を取得（空白行を探す）
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ' ヘッダが未設定の場合は設定
    If lastRow = 2 And ws.Cells(1, 1).Value = "" Then SetupSheet ws, ws.Name
    ' データをシートに書き込み
    ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, UBound(data) + 1)).Value = data
    ' エラーレベルが"ERROR"の場合、セルを黄色で塗りつぶす
    If LogLevel = "ERROR" Then
        ws.Cells(lastRow, 2).Interior.Color = RGB(255, 255, 0)
    End If
