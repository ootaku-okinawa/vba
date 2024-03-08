Dim originalEnableEvents As Boolean
Dim originalScreenUpdating As Boolean
Dim originalCalculation As XlCalculation

Sub SaveApplicationSettings()
    originalEnableEvents = Application.EnableEvents
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
End Sub

Sub RestoreApplicationSettings()
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
End Sub


Sub MyMainProcedure()
    ' 設定を保存
    SaveApplicationSettings()

    ' 処理を行う
    ' （ここに処理コードを記述）

    ' 設定を元に戻す
    RestoreApplicationSettings()
End Sub
