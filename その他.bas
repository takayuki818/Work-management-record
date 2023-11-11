Attribute VB_Name = "その他"
Option Explicit
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
    Dim 氏名 As String, 月区 As Date
    Dim 行 As Long, 最下行 As Long
    Dim 戻り値()
    With Sheets("給与集計")
        月区 = .Range("台帳_集計月")
    End With
    With Sheets("謝礼集計")
        最下行 = .Cells(Rows.Count, 9).End(xlUp).Row
        ReDim 特別日リスト(1 To 最下行 - 3)
        Select Case 最下行
            Case Is > 3
                For 行 = 4 To 最下行
                    特別日リスト(行 - 3) = .Cells(行, 9)
                Next
            Case Else
                MsgBox "特別日リストが設定されていません"
                Exit Sub
        End Select
        
        最下行 = .Cells(Rows.Count, 12).End(xlUp).Row
        ReDim 昼(1 To 最下行 - 3, 1 To 1)
        ReDim 夜(1 To 最下行 - 3, 1 To 1)
        ReDim 備考(1 To 最下行 - 3, 1 To 1)
        For 行 = 4 To 最下行
            氏名 = .Cells(行, 12)
            戻り値 = 特別日勤務取得(氏名, 月区, 特別日リスト)
            昼(行 - 3, 1) = 戻り値(1)
            夜(行 - 3, 1) = 戻り値(2)
            備考(行 - 3, 1) = 戻り値(3)
        Next
        .Cells(4, 13).Resize(最下行 - 3, 1) = 昼
        .Cells(4, 16).Resize(最下行 - 3, 1) = 夜
        .Cells(4, 23).Resize(最下行 - 3, 1) = 備考
    End With
End Sub
Function 特別日勤務取得(氏名 As String, 月区 As Date, 特別日リスト As Variant)
    Dim 行 As Long, 最下行 As Long
    Dim 特別日 As Variant
    Dim 戻り値(1 To 3)
    With Sheets("管理台帳")
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        For 行 = 2 To 最下行
            If .Cells(行, 1) = 氏名 And .Cells(行, 10) = 月区 Then
                For Each 特別日 In 特別日リスト
                    If .Cells(行, 2) = 特別日 Then
                        If .Cells(行, 11) > 0 Then 戻り値(1) = 戻り値(1) + 1
                        If .Cells(行, 13) > 0 Then 戻り値(2) = 戻り値(2) + 1
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
    Dim 氏名 As String, 月区 As Date
    Dim 行 As Long, 最下行 As Long
    With Sheets("給与集計")
        月区 = .Range("台帳_集計月")
    End With
    With Sheets("謝礼集計")
        最下行 = .Cells(Rows.Count, 2).End(xlUp).Row
        ReDim 配列(1 To 最下行 - 3, 1 To 1)
        For 行 = 4 To 最下行
            氏名 = .Cells(行, 2)
            配列(行 - 3, 1) = 勤務日取得(氏名, 月区)
        Next
        .Cells(4, 7).Resize(最下行 - 3, 1) = 配列
    End With
End Sub
Function 勤務日取得(氏名 As String, 月区 As Date) As String
    Dim 行 As Long, 最下行 As Long
    With Sheets("管理台帳")
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        For 行 = 2 To 最下行
            If .Cells(行, 1) = 氏名 And .Cells(行, 10) = 月区 Then
                Select Case 勤務日取得
                    Case "": 勤務日取得 = "勤務日：" & Format(.Cells(行, 2), "d")
                    Case Else: 勤務日取得 = 勤務日取得 & "," & Format(.Cells(行, 2), "d")
                End Select
            End If
        Next
    End With
End Function
