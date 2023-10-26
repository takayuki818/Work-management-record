Attribute VB_Name = "Module1"
Option Explicit
Sub 台帳記録()
    Dim A始列 As Long, B始列 As Long, 休列 As Long, 始 As Long, 終 As Long
    Dim 添字 As Long, 行 As Long, 列 As Long, 最下行 As Long
    Dim 氏名 As String, 文 As String
    Dim 配列(1 To 31, 1 To 9)
    A始列 = 2: B始列 = 5: 休列 = 8
    With Sheets("個別シフト表")
        氏名 = .Range("氏名")
        For 行 = 7 To 37
            If .Cells(行, 1) <> "" Then
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
        Next
    End With
    With Sheets("管理台帳")
        Call 台帳並替
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        Call 記録行探査(氏名, 始, 終)
        If 始 > 0 Then
            If 添字 <> 終 - 始 + 1 Then
                MsgBox "管理台帳の記録行数と一致しません" & vbCrLf & vbCrLf & "台帳記録の更新を中止します"
                Exit Sub
            End If
        End If
        .Cells(最下行 + 1, 1).Resize(添字, 9) = 配列
        .Cells(1, 1).Resize(最下行 + 添字, 9).Borders.LineStyle = True
        Select Case 始
            Case Is > 0
                .Rows(始 & ":" & 終).Delete
                Call 台帳並替
                文 = "台帳記録を更新しました"
            Case Else
                文 = "台帳に新規登録しました"
        End Select
        Call 入力規則再設定
        MsgBox 文
    End With
End Sub
Sub 台帳並替()
    Dim 最下行 As Long
    With Sheets("管理台帳")
        最下行 = .Cells(Rows.Count, 2).End(xlUp).Row
        .Cells(1, 1).Resize(最下行, 9).Characters.PhoneticCharacters = ""
        With .Sort
            With .SortFields
                .Clear
                .Add Key:=Range("A1"), Order:=xlAscending
                .Add Key:=Range("B1"), Order:=xlAscending
            End With
            .SetRange Range(Cells(1, 1), Cells(最下行, 9))
            .Header = xlYes
            .Apply
        End With
    End With
End Sub
Sub 記録行探査(氏名 As String, 始 As Long, 終 As Long)
    Dim 行 As Long, 最下行 As Long
    With Sheets("管理台帳")
        Call 台帳並替
        最下行 = .Cells(Rows.Count, 2).End(xlUp).Row
        For 行 = 2 To 最下行
            If .Cells(行, 1) = 氏名 Then
                If 始 = 0 Then 始 = 行
                If 始 > 0 Then 終 = 行
                Else: If 終 > 0 Then Exit For
            End If
        Next
    End With
End Sub
Sub 記録戻し()
    Dim 氏名 As String
    Dim 始 As Long, 終 As Long, 行 As Long, 列 As Long
    With Sheets("管理台帳")
        氏名 = Sheets("個別シフト表").Range("氏名")
        Call 個別シフト表クリア
        Call 記録行探査(氏名, 始, 終)
        If 始 = 0 Then: Exit Sub
        ReDim A区(始 To 終, 3 To 5)
        ReDim B区(始 To 終, 6 To 8)
        ReDim 休区(始 To 終, 9 To 9)
        For 行 = 始 To 終
            For 列 = 3 To 9
                Select Case 列
                    Case 3, 4, 5: A区(行, 列) = .Cells(行, 列)
                    Case 6, 7, 8: B区(行, 列) = .Cells(行, 列)
                    Case 9: 休区(行, 列) = .Cells(行, 列)
                End Select
            Next
        Next
    End With
    With Sheets("個別シフト表")
        .Cells(7, 2).Resize(終 - 始 + 1, 3) = A区
        .Cells(7, 5).Resize(終 - 始 + 1, 3) = B区
        .Cells(7, 8).Resize(終 - 始 + 1, 1) = 休区
    End With
End Sub
Sub 入力規則再設定()
    Dim 最下行 As Long, 行 As Long
    Dim 氏名 As String, リスト As String
    With Sheets("管理台帳")
        Call 台帳並替
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        氏名 = .Cells(2, 1)
        リスト = 氏名
        For 行 = 3 To 最下行
            If .Cells(行, 1) <> 氏名 Then
                 氏名 = .Cells(行, 1)
                リスト = リスト & "," & 氏名
            End If
        Next
    End With
    With Sheets("個別シフト表").Range("氏名").Validation
        .Delete
        .Add Type:=xlValidateList, Operator:=xlEqual, Formula1:=リスト
        .ShowError = False
    End With
End Sub
Sub 個別シフト表クリア()
    With Sheets("個別シフト表")
        .Cells(7, 2).Resize(31, 3).ClearContents
        .Cells(7, 5).Resize(31, 3).ClearContents
        .Cells(7, 8).Resize(31, 1).ClearContents
    End With
End Sub
Sub 保護切替()
    With ActiveSheet
        Select Case .ProtectContents
            Case True: .Unprotect: MsgBox "シート保護を解除しました"
            Case False: .Protect: MsgBox "シートを保護しました"
        End Select
    End With
End Sub
