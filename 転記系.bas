Attribute VB_Name = "�]�L�n"
Option Explicit
Public Const A�n�� As Long = 2
Public Const B�n�� As Long = 7
Public Const �x�� As Long = 11
Function ���v�Ζ����Ԍ���() As String
    Dim �s As Long, ��Q As Variant, �� As Variant
    With Sheets("�ʃV�t�g�\")
        ��Q = Array(5, 10, 12)
        For �s = 7 To 37
            For Each �� In ��Q
                If .Cells(�s, ��) <> "" And .Cells(�s, ��) > TimeSerial(7, 45, 0) Then
                    ���v�Ζ����Ԍ��� = "�y���Ӂz���v7.75h���ߓ�����" & vbCrLf & "�{�c�[���͎��ԊO�Ζ��̋L�^�ɑΉ����Ă��܂���"
                    Exit Function
                End If
            Next
        Next
    End With
End Function
Sub �䒠�L�^()
    Dim �n As Long, �I As Long, �Y�� As Long, �Y�� As Long, �s As Long, �� As Long, �ŉ��s As Long, �L�^�� As Long
    Dim �ጎ�敪 As String, ���� As String, �� As String
    Dim �z��(1 To 31, 1 To 9), �\�t(1 To 1, 1 To 9)
    With Sheets("�ʃV�t�g�\")
        �ጎ�敪 = Format(.Range("�ጎ�敪"), "ge.m")
        ���� = .Range("����")
        For �s = 7 To 37
            If .Cells(�s, 1) <> "" Then
                If .Cells(�s, A�n��) <> "" Or .Cells(�s, B�n��) <> "" Or .Cells(�s, �x��) <> "" Then
                    �Y�� = �Y�� + 1
                    �z��(�Y��, 1) = ����
                    �z��(�Y��, 2) = .Cells(�s, 1)
                    For �� = 3 To 5
                        �z��(�Y��, ��) = .Cells(�s, A�n�� + �� - 3)
                    Next
                    For �� = 6 To 8
                        �z��(�Y��, ��) = .Cells(�s, B�n�� + �� - 6)
                    Next
                    �z��(�Y��, 9) = .Cells(�s, �x��)
                End If
            End If
        Next
    End With
    With Sheets("�Ǘ��䒠")
        Call �L�^�s�T��(�ጎ�敪, ����, �n, �I)
        �Y�� = �Y��
        Select Case �n
            Case Is > 0
                �� = "�䒠���X�V���܂���"
                For �Y�� = 1 To �Y��
                    For �� = 1 To 9
                        �\�t(1, ��) = �z��(�Y��, ��)
                    Next
                    �L�^�� = 0
                    For �s = �n To �I
                        If .Cells(�s, 2) = �\�t(1, 2) Then
                            .Cells(�s, 1).Resize(1, 9) = �\�t
                            �L�^�� = 1
                            Exit For
                        End If
                    Next
                    If �L�^�� = 0 Then
                        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
                        .Cells(�ŉ��s + 1, 1).Resize(1, 9) = �\�t
                    End If
                Next
                For �s = �n To �I
                    �L�^�� = 0
                    For �Y�� = 1 To �Y��
                        If .Cells(�s, 2) = �z��(�Y��, 2) Then
                            �L�^�� = 1
                            Exit For
                        End If
                    Next
                    If �L�^�� = 0 Then .Rows(�s).Delete
                Next
            Case Else
                Select Case �Y��
                    Case 0
                        �� = "�o�^�Ώۃf�[�^������܂���"
                    Case Else
                        �� = "�䒠�ɐV�K�o�^���܂���"
                        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
                        .Cells(�ŉ��s + 1, 1).Resize(�Y��, 9) = �z��
                End Select
        End Select
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        Call �L�^�߂�
        MsgBox ��
        �� = ���v�Ζ����Ԍ���()
        If �� <> "" Then MsgBox ��
    End With
End Sub
Sub �䒠����() '�W�v�����̃I�[�g�t�B���A�r���E�����t�������ݒ���܂ށB
    Dim �ŉ��s As Long, �ŉE�� As Long
    Dim ���� As FormatCondition
    With Sheets("�Ǘ��䒠")
        �ŉ��s = .Cells(Rows.Count, 2).End(xlUp).Row
        �ŉE�� = .Cells(1, Columns.Count).End(xlToLeft).Column
        Select Case �ŉ��s
            Case Is < 2: Exit Sub
            Case Is >= 2
                .Cells(2, 10).Formula = "=IF(B2="""",""-"",DATE(YEAR(B2),MONTH(B2),1))"
                .Cells(2, 11).Formula = "=D2-C2-E2"
                .Cells(2, 12).Formula = "=MAX(0,D2-TIME(22,0,0))"
                .Cells(2, 13).Formula = "=G2-F2-H2"
                .Cells(2, 14).Formula = "=SUM(K2,M2)"
                If �ŉ��s >= 3 Then .Cells(2, 10).Resize(1, 5).AutoFill .Cells(2, 10).Resize(�ŉ��s - 1, 5)
        End Select
        .Cells(1, 1).Resize(�ŉ��s, �ŉE��).Characters.PhoneticCharacters = ""
        With .Sort
            With .SortFields
                .Clear
                .Add Key:=Range("A1"), Order:=xlAscending
                .Add Key:=Range("B1"), Order:=xlAscending
            End With
            .SetRange Range(Cells(1, 1), Cells(�ŉ��s, �ŉE��))
            .Header = xlYes
            .Apply
        End With
        .Cells(1, 1).Resize(Rows.Count, Columns.Count).Borders.LineStyle = False
        .Cells(1, 1).Resize(�ŉ��s, �ŉE��).Borders.LineStyle = True
        .Cells.FormatConditions.Delete
        Set ���� = .Cells(2, 11).Resize(�ŉ��s, 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=K2>TIME(7,45,0)")
        ����.Font.Color = RGB(255, 0, 0)
        Set ���� = .Cells(2, 13).Resize(�ŉ��s, 2).FormatConditions.Add(Type:=xlExpression, Formula1:="=M2>TIME(7,45,0)")
        ����.Font.Color = RGB(255, 0, 0)
    End With
End Sub
Sub �L�^�s�T��(�ጎ�敪 As String, ���� As String, �n As Long, �I As Long)
    Dim �s As Long, �ŉ��s As Long
    With Sheets("�Ǘ��䒠")
        Call �䒠����
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        If �ŉ��s < 2 Then Exit Sub
        For �s = 2 To �ŉ��s
            If .Cells(�s, 1) = ���� And Format(.Cells(�s, 2), "ge.m") = �ጎ�敪 Then
                If �n = 0 Then �n = �s
                If �n > 0 Then �I = �s
                Else: If �I > 0 Then Exit For
            End If
        Next
    End With
End Sub
Function �ҏW�����m�F() As String
    Dim ���� As String, �ጎ�敪 As String
    Dim �n As Long, �I As Long, �s As Long, �� As Long, �Y�� As Long
    Dim �z��(1 To 31, 1 To 9)
    With Sheets("�ʃV�t�g�\")
        �ጎ�敪 = Format(.Range("�ጎ�敪"), "ge.m")
        ���� = .Range("����")
        If �ጎ�敪 = "" Or ���� = "" Then
            �ҏW�����m�F = ""
            Exit Function
        End If
        For �s = 7 To 37
            If .Cells(�s, 1) <> "" Then
                If .Cells(�s, A�n��) <> "" Or .Cells(�s, B�n��) <> "" Or .Cells(�s, �x��) <> "" Then
                    �Y�� = �Y�� + 1
                    �z��(�Y��, 1) = ����
                    �z��(�Y��, 2) = .Cells(�s, 1)
                    For �� = 3 To 5
                        �z��(�Y��, ��) = .Cells(�s, A�n�� + �� - 3)
                    Next
                    For �� = 6 To 8
                        �z��(�Y��, ��) = .Cells(�s, B�n�� + �� - 6)
                    Next
                    �z��(�Y��, 9) = .Cells(�s, �x��)
                End If
            End If
        Next
    End With
    With Sheets("�Ǘ��䒠")
        Call �L�^�s�T��(�ጎ�敪, ����, �n, �I)
        Select Case �n
            Case Is > 0
                For �s = �n To �I
                    For �� = 1 To 9
                        If .Cells(�s, ��) <> �z��(�s - �n + 1, ��) Then
                            �ҏW�����m�F = "��������"
                            Exit Function
                        End If
                    Next
                Next
            Case Else
                Select Case �Y�� '�t�H�[���ւ̃f�[�^���͗L���ŕ���
                    Case Is > 0: �ҏW�����m�F = "���o�^"
                    Case Else: �ҏW�����m�F = ""
                End Select
        End Select
    End With
End Function
Sub �L�^�߂�()
    Dim ���� As String, �ጎ�敪 As String
    Dim �n As Long, �I As Long, �s As Long, �� As Long, �Y�� As Long
    Dim �\�t(1 To 1, 1 To 3)
    With Sheets("�Ǘ��䒠")
        �ጎ�敪 = Format(Sheets("�ʃV�t�g�\").Range("�ጎ�敪"), "ge.m")
        ���� = Sheets("�ʃV�t�g�\").Range("����")
        Call �ʃV�t�g�\�N���A
        Call �L�^�s�T��(�ጎ�敪, ����, �n, �I)
        If �n = 0 Then: Exit Sub
        ReDim �z��(1 To �I - �n + 1, 1 To 8)
        For �s = �n To �I
            For �� = 2 To 9
                �z��(�s - �n + 1, �� - 1) = .Cells(�s, ��)
            Next
        Next
    End With
    With Sheets("�ʃV�t�g�\")
        For �s = 7 To 37
            For �Y�� = 1 To �I - �n + 1
                If .Cells(�s, 1) = �z��(�Y��, 1) Then
                    Erase �\�t
                    For �� = 2 To 4
                        �\�t(1, �� - 1) = �z��(�Y��, ��)
                    Next
                    .Cells(�s, A�n��).Resize(1, 3) = �\�t
                    Erase �\�t
                    For �� = 5 To 7
                        �\�t(1, �� - 4) = �z��(�Y��, ��)
                    Next
                    .Cells(�s, B�n��).Resize(1, 3) = �\�t
                    .Cells(�s, �x��) = �z��(�Y��, 8)
                End If
            Next
        Next
    End With
End Sub
