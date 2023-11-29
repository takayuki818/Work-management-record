Attribute VB_Name = "転記系"
Option Explicit
Public Const A始列 As Long = 2
Public Const B始列 As Long = 7
Public Const 休列 As Long = 11
Function 日計勤務時間検査() As String
    Dim 行 As Long, 列群 As Variant, 列 As Variant
    With Sheets("個別シフト表")
        列群 = Array(5, 10, 12)
        For 行 = 7 To 37
            For Each 列 In 列群
                If .Cells(行, 列) <> "" And .Cells(行, 列) > TimeSerial(7, 45, 0) Then
                    日計勤務時間検査 = "【注意】日計7.75h超過日あり" & vbCrLf & "本ツールは時間外勤務の記録に対応していません"
                    Exit Function
                End If
            Next
        Next
    End With
End Function
Sub 台帳記録()
    Dim 始 As Long, 終 As Long, 添字 As Long, 添数 As Long, 行 As Long, 列 As Long, 最下行 As Long, 記録済 As Long
    Dim 例月区分 As String, 氏名 As String, 文 As String
    Dim 配列(1 To 31, 1 To 9), 貼付(1 To 1, 1 To 9)
    With Sheets("個別シフト表")
        例月区分 = Format(.Range("例月区分"), "ge.m")
        氏名 = .Range("氏名")
        For 行 = 7 To 37
            If .Cells(行, 1) <> "" Then
                If .Cells(行, A始列) <> "" Or .Cells(行, B始列) <> "" Or .Cells(行, 休列) <> "" Then
                    添字 = 添字 + 1
                    配列(添字, 1) = 氏名
                    配列(添字, 2) = .Cells(行, 1)
                    For 列 = 3 To 5
                        配列(添字, 列) = .Cells(行, A始列 + 列 - 3)
                    Next
                    For 列 = 6 To 8
                        配列(添字, 列) = .Cells(行, B始列 + 列 - 6)
                    Next
                    配列(添字, 9) = .Cells(行, 休列)
                End If
            End If
        Next
    End With
    With Sheets("管理台帳")
        Call 記録行探査(例月区分, 氏名, 始, 終)
        添数 = 添字
        Select Case 始
            Case Is > 0
                文 = "台帳を更新しました"
                For 添字 = 1 To 添数
                    For 列 = 1 To 9
                        貼付(1, 列) = 配列(添字, 列)
                    Next
                    記録済 = 0
                    For 行 = 始 To 終
                        If .Cells(行, 2) = 貼付(1, 2) Then
                            .Cells(行, 1).Resize(1, 9) = 貼付
                            記録済 = 1
                            Exit For
                        End If
                    Next
                    If 記録済 = 0 Then
                        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
                        .Cells(最下行 + 1, 1).Resize(1, 9) = 貼付
                    End If
                Next
                For 行 = 始 To 終
                    記録済 = 0
                    For 添字 = 1 To 添数
                        If .Cells(行, 2) = 配列(添字, 2) Then
                            記録済 = 1
                            Exit For
                        End If
                    Next
                    If 記録済 = 0 Then .Rows(行).Delete
                Next
            Case Else
                Select Case 添数
                    Case 0
                        文 = "登録対象データがありません"
                    Case Else
                        文 = "台帳に新規登録しました"
                        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
                        .Cells(最下行 + 1, 1).Resize(添数, 9) = 配列
                End Select
        End Select
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        Call 記録戻し
        MsgBox 文
        文 = 日計勤務時間検査()
        If 文 <> "" Then MsgBox 文
    End With
End Sub
Sub 台帳並替() '集計数式のオートフィル、罫線・条件付き書式設定を含む。
    Dim 最下行 As Long, 最右列 As Long
    Dim 条件 As FormatCondition
    With Sheets("管理台帳")
        最下行 = .Cells(Rows.Count, 2).End(xlUp).Row
        最右列 = .Cells(1, Columns.Count).End(xlToLeft).Column
        Select Case 最下行
            Case Is < 2: Exit Sub
            Case Is >= 2
                .Cells(2, 10).Formula = "=IF(B2="""",""-"",DATE(YEAR(B2),MONTH(B2),1))"
                .Cells(2, 11).Formula = "=D2-C2-E2"
                .Cells(2, 12).Formula = "=MAX(0,D2-TIME(22,0,0))"
                .Cells(2, 13).Formula = "=G2-F2-H2"
                .Cells(2, 14).Formula = "=SUM(K2,M2)"
                If 最下行 >= 3 Then .Cells(2, 10).Resize(1, 5).AutoFill .Cells(2, 10).Resize(最下行 - 1, 5)
        End Select
        .Cells(1, 1).Resize(最下行, 最右列).Characters.PhoneticCharacters = ""
        With .Sort
            With .SortFields
                .Clear
                .Add Key:=Range("A1"), Order:=xlAscending
                .Add Key:=Range("B1"), Order:=xlAscending
            End With
            .SetRange Range(Cells(1, 1), Cells(最下行, 最右列))
            .Header = xlYes
            .Apply
        End With
        .Cells(1, 1).Resize(Rows.Count, Columns.Count).Borders.LineStyle = False
        .Cells(1, 1).Resize(最下行, 最右列).Borders.LineStyle = True
        .Cells.FormatConditions.Delete
        Set 条件 = .Cells(2, 11).Resize(最下行, 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=K2>TIME(7,45,0)")
        条件.Font.Color = RGB(255, 0, 0)
        Set 条件 = .Cells(2, 13).Resize(最下行, 2).FormatConditions.Add(Type:=xlExpression, Formula1:="=M2>TIME(7,45,0)")
        条件.Font.Color = RGB(255, 0, 0)
    End With
End Sub
Sub 記録行探査(例月区分 As String, 氏名 As String, 始 As Long, 終 As Long)
    Dim 行 As Long, 最下行 As Long
    With Sheets("管理台帳")
        Call 台帳並替
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        If 最下行 < 2 Then Exit Sub
        For 行 = 2 To 最下行
            If .Cells(行, 1) = 氏名 And Format(.Cells(行, 2), "ge.m") = 例月区分 Then
                If 始 = 0 Then 始 = 行
                If 始 > 0 Then 終 = 行
                Else: If 終 > 0 Then Exit For
            End If
        Next
    End With
End Sub
Function 編集差分確認() As String
    Dim 氏名 As String, 例月区分 As String
    Dim 始 As Long, 終 As Long, 行 As Long, 列 As Long, 添字 As Long
    Dim 配列(1 To 31, 1 To 9)
    With Sheets("個別シフト表")
        例月区分 = Format(.Range("例月区分"), "ge.m")
        氏名 = .Range("氏名")
        If 例月区分 = "" Or 氏名 = "" Then
            編集差分確認 = ""
            Exit Function
        End If
        For 行 = 7 To 37
            If .Cells(行, 1) <> "" Then
                If .Cells(行, A始列) <> "" Or .Cells(行, B始列) <> "" Or .Cells(行, 休列) <> "" Then
                    添字 = 添字 + 1
                    配列(添字, 1) = 氏名
                    配列(添字, 2) = .Cells(行, 1)
                    For 列 = 3 To 5
                        配列(添字, 列) = .Cells(行, A始列 + 列 - 3)
                    Next
                    For 列 = 6 To 8
                        配列(添字, 列) = .Cells(行, B始列 + 列 - 6)
                    Next
                    配列(添字, 9) = .Cells(行, 休列)
                End If
            End If
        Next
    End With
    With Sheets("管理台帳")
        Call 記録行探査(例月区分, 氏名, 始, 終)
        Select Case 始
            Case Is > 0
                For 行 = 始 To 終
                    For 列 = 1 To 9
                        If .Cells(行, 列) <> 配列(行 - 始 + 1, 列) Then
                            編集差分確認 = "差分あり"
                            Exit Function
                        End If
                    Next
                Next
            Case Else
                Select Case 添字 'フォームへのデータ入力有無で分岐
                    Case Is > 0: 編集差分確認 = "未登録"
                    Case Else: 編集差分確認 = ""
                End Select
        End Select
    End With
End Function
Sub 記録戻し()
    Dim 氏名 As String, 例月区分 As String
    Dim 始 As Long, 終 As Long, 行 As Long, 列 As Long, 添字 As Long
    Dim 貼付(1 To 1, 1 To 3)
    With Sheets("管理台帳")
        例月区分 = Format(Sheets("個別シフト表").Range("例月区分"), "ge.m")
        氏名 = Sheets("個別シフト表").Range("氏名")
        Call 個別シフト表クリア
        Call 記録行探査(例月区分, 氏名, 始, 終)
        If 始 = 0 Then: Exit Sub
        ReDim 配列(1 To 終 - 始 + 1, 1 To 8)
        For 行 = 始 To 終
            For 列 = 2 To 9
                配列(行 - 始 + 1, 列 - 1) = .Cells(行, 列)
            Next
        Next
    End With
    With Sheets("個別シフト表")
        For 行 = 7 To 37
            For 添字 = 1 To 終 - 始 + 1
                If .Cells(行, 1) = 配列(添字, 1) Then
                    Erase 貼付
                    For 列 = 2 To 4
                        貼付(1, 列 - 1) = 配列(添字, 列)
                    Next
                    .Cells(行, A始列).Resize(1, 3) = 貼付
                    Erase 貼付
                    For 列 = 5 To 7
                        貼付(1, 列 - 4) = 配列(添字, 列)
                    Next
                    .Cells(行, B始列).Resize(1, 3) = 貼付
                    .Cells(行, 休列) = 配列(添字, 8)
                End If
            Next
        Next
    End With
End Sub
