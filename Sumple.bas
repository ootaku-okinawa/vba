Dim originalEnableEvents As Boolean
Dim originalScreenUpdating As Boolean
Dim originalCalculation As XlCalculation
Dim originalDisplayAlerts As Boolean

Sub SaveApplicationSettings()
    ' 現在の設定を保存
    originalEnableEvents = Application.EnableEvents
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalDisplayAlerts = Application.DisplayAlerts
End Sub

Sub RestoreApplicationSettings()
    ' 保存した設定に戻す
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = originalDisplayAlerts
End Sub



Sub MyMainProcedure()
    ' 設定を保存
    SaveApplicationSettings()

    ' 処理を行う
    ' （ここに処理コードを記述）

    ' 設定を元に戻す
    RestoreApplicationSettings()
End Sub
