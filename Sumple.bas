Attribute VB_Name = "Module1"

Sub AggregateData()
    ' Excelのパフォーマンスを向上させるための初期設定
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' テンプレートシート名の定義
    Const TemplateAllSheetName As String = "個人別(template)"
    Const TemplateBranchSheetName As String = "支社別(template)"
    Const TemplateBuilderSheetName As String = "通建会社別(template)"
    
    ' 必要な変数の宣言
    Dim wsTemplateAll As Worksheet, wsTemplateBranch As Worksheet, wsTemplateBuilder As Worksheet
    Dim dictCompany As Object: Set dictCompany = CreateObject("Scripting.Dictionary")
    Dim dictBranch As Object: Set dictBranch = CreateObject("Scripting.Dictionary")
    Dim dictSection As Object: Set dictSection = CreateObject("Scripting.Dictionary")
    Dim folderPath As String, fileName As String
    Dim lastRow As Long, startRow As Long: startRow = 8
    
    ' テンプレートシートの設定
    Set wsTemplateAll = ThisWorkbook.Sheets(TemplateAllSheetName)
    Set wsTemplateBranch = ThisWorkbook.Sheets(TemplateBranchSheetName)
    Set wsTemplateBuilder = ThisWorkbook.Sheets(TemplateBuilderSheetName)
    
    ' 既存シートの削除
    DeleteSheetIfExists "個人別"
    DeleteSheetIfExists "支社別"
    DeleteSheetIfExists "通建会社別"
    
    ' テンプレートシートをコピーして新しいシートを作成
    wsTemplateBranch.Copy Before:=wsTemplateBranch: ActiveSheet.Name = "支社別"
    wsTemplateBuilder.Copy Before:=wsTemplateBuilder: ActiveSheet.Name = "通建会社別"
    wsTemplateAll.Copy Before:=wsTemplateAll: ActiveSheet.Name = "個人別"
    
    ' データフォルダのパス設定
    folderPath = ThisWorkbook.Path & "\personal_data_develop\"
    fileName = Dir(folderPath & "*.xlsx")
    
    ' ファイルをループして処理
    Do While fileName <> ""
        ProcessFile folderPath & fileName, wsTemplateAll, dictCompany, dictBranch, dictSection, startRow
        fileName = Dir()
    Loop
    
    ' 集計結果の出力（この部分の実装は省略されていますが、必要に応じて実装してください）
    ' OutputAveragesBranch ...
    ' OutputAveragesBuilder ...
    
    ' Excelの設定を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub DeleteSheetIfExists(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then ws.Delete
End Sub

Private Sub ProcessFile(filePath As String, ByRef wsTemplate As Worksheet, ByRef dictCompany As Object, _
                        ByRef dictBranch As Object, ByRef dictSection As Object, ByRef startRow As Long)
    ' ファイルからデータを読み込み、必要な処理を行う（具体的な実装は省略されています）
    ' ここでデータを読み込み、集計し、wsTemplateにデータを出力するロジックを実装します。
    ' filePath: 処理するファイルのパス
    ' wsTemplate: データを出力するテンプレートワークシート
    ' dictCompany, dictBranch, dictSection: 集計データを保持する辞書
    ' startRow: データ出
