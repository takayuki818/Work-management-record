Attribute VB_Name = "その他"
Option Explicit
Sub 集計シート構築()
    Dim 行数 As Long, 最下行 As Long, 添字 As Long
    Dim 給与(), 自家用車(), 特日昼(), 特日夜()
    With Sheets("氏名マスタ")
        行数 = .Cells(Rows.Count, 4).End(xlUp).Row - 1
        If 行数 < 1 Then
            MsgBox "氏名マスタが作成されていません"
            Exit Sub
        End If
        給与 = .Cells(2, 1).Resize(行数, 5).Value
        ReDim 謝礼氏名(1 To 行数, 1 To 2)
        For 添字 = 1 To 行数
            謝礼氏名(添字, 1) = .Cells(添字 + 1, 2)
            謝礼氏名(添字, 2) = .Cells(添字 + 1, 4)
        Next
        自家用車 = .Cells(2, 6).Resize(行数, 1).Value
        特日昼 = .Cells(2, 7).Resize(行数, 1).Value
        特日夜 = .Cells(2, 8).Resize(行数, 1).Value
    End With
    With Sheets("給与集計")
        最下行 = .Cells(Rows.Count, 4).End(xlUp).Row
        .Cells(7, 1).Resize(最下行 - 6, 31).Clear
        .Cells(6, 1).Resize(1, 31).AutoFill .Cells(6, 1).Resize(行数, 31)
        .Cells(6, 1).Resize(行数, 5) = 給与
    End With
    With Sheets("謝礼集計")
        最下行 = .Cells(Rows.Count, 2).End(xlUp).Row
        .Cells(6, 1).Resize(最下行 - 5, 21).Clear
        .Cells(5, 1).Resize(1, 21).AutoFill .Cells(5, 1).Resize(行数, 21)
        .Cells(5, 1).Resize(行数, 2) = 謝礼氏名
        .Cells(5, 4).Resize(行数, 1) = 自家用車
        .Cells(5, 9).Resize(行数, 2) = 謝礼氏名
        .Cells(5, 12).Resize(行数, 1) = 特日昼
        .Cells(5, 15).Resize(行数, 1) = 特日夜
        Call 全勤務日設定
        Call 特別日勤務設定
        Call 集計管理モード
    End With
    MsgBox "集計シートの構築が完了しました"
End Sub
Sub プルダウン設定(モード As String) '256字以上の文字列形式でプルダウン設定するとファイルが破損することに注意！
    Dim 最下行 As Long, 行 As Long, 添字 As Long
    Dim 範囲式 As String
    With Sheets("給与集計")
        最下行 = .Cells(Rows.Count, 4).End(xlUp).Row
        If 最下行 < 6 Then 最下行 = 6
        ReDim リスト(1 To 最下行 - 5, 1 To 1)
        For 行 = 6 To 最下行
            Select Case モード
                Case "氏名全件"
                    リスト(行 - 5, 1) = .Cells(行, 4)
                Case "要支払氏名のみ"
                    If .Cells(行, 30) > 0 Then
                        添字 = 添字 + 1
                        リスト(添字, 1) = .Cells(行, 4)
                    End If
            End Select
        Next
    End With
    With Sheets("プルダウン設定")
        .Cells(1, 1).Resize(Rows.Count, 1).ClearContents
        .Cells(1, 1).Resize(最下行 - 5, 1) = リスト
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        範囲式 = "=プルダウン設定!$A$1:$A$" & 最下行
    End With
    With Sheets("個別シフト表")
        .Unprotect
        With .Range("氏名").Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=範囲式
        End With
        .Protect
    End With
End Sub
Sub 夜定型シフト代入()
    Dim 行 As Long, 配列()
    With Sheets("個別シフト表")
        行 = ActiveCell.Row
        If 行 > 6 And 行 < 38 Then
            配列 = .Range("夜標準シフト").Value
            .Cells(行, 2).Resize(1, 3) = 配列
        End If
    End With
End Sub
Sub 昼定型シフト代入()
    Dim 行 As Long, 配列()
    With Sheets("個別シフト表")
        行 = ActiveCell.Row
        If 行 > 6 And 行 < 38 Then
            配列 = .Range("昼標準シフト").Value
            .Cells(行, 7).Resize(1, 3) = 配列
        End If
    End With
End Sub
Sub 個別シフト表クリア()
    With Sheets("個別シフト表")
        .Cells(7, A始列).Resize(31, 3).ClearContents
        .Cells(7, B始列).Resize(31, 3).ClearContents
        .Cells(7, 休列).Resize(31, 1).ClearContents
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
Sub 特別日勤務設定()
    Dim 氏名 As String, 月区 As String
    Dim 行 As Long, 最下行 As Long
    Dim 戻り値()
    With Sheets("給与集計")
        月区 = Format(.Range("台帳_集計月"), "ge.m")
    End With
    With Sheets("特別日マスタ")
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 特別日リスト(1 To 最下行 - 1)
        Select Case 最下行
            Case Is > 1
                For 行 = 2 To 最下行
                    特別日リスト(行 - 1) = .Cells(行, 1)
                Next
            Case Else
                MsgBox "特別日リストが設定されていません"
                Exit Sub
        End Select
    End With
    With Sheets("謝礼集計")
        最下行 = .Cells(Rows.Count, 10).End(xlUp).Row
        ReDim 昼(1 To 最下行 - 4, 1 To 1)
        ReDim 夜(1 To 最下行 - 4, 1 To 1)
        ReDim 備考(1 To 最下行 - 4, 1 To 1)
        For 行 = 5 To 最下行
            氏名 = .Cells(行, 10)
            戻り値 = 特別日勤務取得(氏名, 月区, 特別日リスト)
            昼(行 - 4, 1) = 戻り値(1)
            夜(行 - 4, 1) = 戻り値(2)
            備考(行 - 4, 1) = 戻り値(3)
        Next
        .Cells(5, 11).Resize(最下行 - 4, 1) = 昼
        .Cells(5, 14).Resize(最下行 - 4, 1) = 夜
        .Cells(5, 21).Resize(最下行 - 4, 1) = 備考
    End With
End Sub
Function 特別日勤務取得(氏名 As String, 月区 As String, 特別日リスト As Variant)
    Dim 行 As Long, 最下行 As Long
    Dim 特別日 As Variant
    Dim 戻り値(1 To 3)
    With Sheets("管理台帳")
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        For 行 = 2 To 最下行
            If .Cells(行, 1) = 氏名 And Format(.Cells(行, 10), "ge.m") = 月区 Then
                For Each 特別日 In 特別日リスト
                    If .Cells(行, 2) = 特別日 Then
                        If .Cells(行, 13) > 0 Then 戻り値(1) = 戻り値(1) + 1
                        If .Cells(行, 11) > 0 Then 戻り値(2) = 戻り値(2) + 1
                        Select Case 戻り値(3)
                            Case "": 戻り値(3) = "勤務日：" & Format(.Cells(行, 2), "d")
                            Case Else: 戻り値(3) = 戻り値(3) & "," & Format(.Cells(行, 2), "d")
                        End Select
                        Exit For
                    End If
                Next
            End If
        Next
    End With
    特別日勤務取得 = 戻り値
End Function
Sub 全勤務日設定()
    Dim 氏名 As String, 月区 As String
    Dim 行 As Long, 最下行 As Long
    With Sheets("給与集計")
        月区 = Format(.Range("台帳_集計月"), "ge.m")
    End With
    With Sheets("謝礼集計")
        最下行 = .Cells(Rows.Count, 2).End(xlUp).Row
        ReDim 配列(1 To 最下行 - 4, 1 To 1)
        For 行 = 5 To 最下行
            氏名 = .Cells(行, 2)
            配列(行 - 4, 1) = 勤務日取得(氏名, 月区)
        Next
        .Cells(5, 7).Resize(最下行 - 4, 1) = 配列
    End With
End Sub
Function 勤務日取得(氏名 As String, 月区 As String) As String
    Dim 行 As Long, 最下行 As Long
    With Sheets("管理台帳")
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        For 行 = 2 To 最下行
            If .Cells(行, 1) = 氏名 And Format(.Cells(行, 10), "ge.m") = 月区 Then
                Select Case 勤務日取得
                    Case "": 勤務日取得 = "勤務日：" & Format(.Cells(行, 2), "d")
                    Case Else: 勤務日取得 = 勤務日取得 & "," & Format(.Cells(行, 2), "d")
                End Select
            End If
        Next
    End With
End Function
