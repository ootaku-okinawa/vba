Sub AggregateDataForF1ToF10()
    Dim wsCompanyAvg As Worksheet, wsBranchAvg As Worksheet, wsBranchSectionAvg As Worksheet
    Dim dictCompany As Object, dictBranch As Object, dictSection As Object
    Dim fileName As String, folderPath As String
    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, j As Long, score As Double
    Dim companyName As String, branchName As String, sectionName As String
    Dim arrScores(1 To 10) As Double ' F1からF10のスコアを格納する配列
    
    ' コレクションの初期化
    Set dictCompany = CreateObject("Scripting.Dictionary")
    Set dictBranch = CreateObject("Scripting.Dictionary")
    Set dictSection = CreateObject("Scripting.Dictionary")
    
    ' 結果を出力するシートの設定
    Set wsCompanyAvg = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsCompanyAvg.Name = "Company Avg F1-F10"
    Set wsBranchAvg = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsBranchAvg.Name = "Branch Avg F1-F10"
    Set wsBranchSectionAvg = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsBranchSectionAvg.Name = "Section Avg F1-F10"
    
    ' personディレクトリ内のファイルをループ
    folderPath = ThisWorkbook.Path & "\person\"
    fileName = Dir(folderPath & "*.xlsx")
    
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        Set ws = wb.Sheets(1)
        
        ' データの読み取り
        companyName = ws.Range("A4").Value
        branchName = ws.Range("A2").Value
        sectionName = ws.Range("A3").Value
        
        ' F1からF10までのスコアの読み取り
        For i = 1 To 10
            score = ws.Range("F" & i).Value
            
            ' 担当会社ごとの集計
            If Not dictCompany.exists(companyName & " F" & i) Then
                dictCompany(companyName & " F" & i) = Array(score, 1) ' スコアの合計とカウント
            Else
                dictCompany(companyName & " F" & i) = Array(dictCompany(companyName & " F" & i)(0) + score, dictCompany(companyName & " F" & i)(1) + 1)
            End If
            
            ' 支社ごとの集計
            If Not dictBranch.exists(branchName & " F" & i) Then
                dictBranch(branchName & " F" & i) = Array(score, 1)
            Else
                dictBranch(branchName & " F" & i) = Array(dictBranch(branchName & " F" & i)(0) + score, dictBranch(branchName & " F" & i)(1) + 1)
            End If
            
            ' 係ごとの集計
            If Not dictSection.exists(branchName & " " & sectionName & " F" & i) Then
                dictSection(branchName & " " & sectionName & " F" & i) = Array(score, 1)
            Else
                dictSection(branchName & " " & sectionName & " F" & i) = Array(dictSection(branchName & " " & sectionName & " F" & i)(0) + score, dictSection(branchName & " " & sectionName & " F" & i)(1) + 1)
            End If
        Next i
        
        wb.Close False ' ファイルを保存せずに閉じる
        fileName = Dir() ' 次のファイルへ
    Loop
    
    ' 結果の出力
    OutputAverages wsCompanyAvg, dictCompany
    OutputAverages wsBranchAvg, dictBranch
    OutputAverages wsBranchSectionAvg, dictSection
    
    MsgBox "集計が完了しました。"
End Sub

' 平均値を出力する補助関数
Sub OutputAverages(ws As Worksheet, dict As Object)
    Dim key As Variant, i As Long
    i = 1
    For Each key In dict.keys
        ws.Cells(i, 1).Value = key
        ws.Cells(i, 2).Value = dict(key)(0) / dict(key)(1) ' 合計をカウントで割る
        i = i + 1
    Next key
End Sub
