Function ProcessValue(ByVal dic As Object, ByVal key As Variant) As Variant
    If Not dic.Exists(key) Then
        ProcessValue = "Key not found"
        Exit Function
    End If
    
    If IsArray(dic(key)) Then
        ' 配列の場合、合計値をカウント値で割る
        Dim total As Double
        Dim count As Double
        total = dic(key)(0)
        count = dic(key)(1)
        If count <> 0 Then
            ProcessValue = total / count
        Else
            ProcessValue = "Error: Count is zero"
        End If
    ElseIf VarType(dic(key)) = vbString Then
        ' 文字列の場合、そのまま使用
        ProcessValue = dic(key)
    Else
        ProcessValue = "Error: Invalid data type"
    End If
End Function
