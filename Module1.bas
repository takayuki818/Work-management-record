Attribute VB_Name = "Module1"
Option Explicit
Sub �䒠�L�^()
    Dim A�n�� As Long, B�n�� As Long, �x�� As Long, �n As Long, �I As Long
    Dim �Y�� As Long, �s As Long, �� As Long, �ŉ��s As Long
    Dim ���� As String, �� As String
    Dim �z��(1 To 31, 1 To 9)
    A�n�� = 2: B�n�� = 5: �x�� = 8
    With Sheets("�ʃV�t�g�\")
        ���� = .Range("����")
        For �s = 7 To 37
            If .Cells(�s, 1) <> "" Then
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
        Next
    End With
    With Sheets("�Ǘ��䒠")
        Call �䒠����
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        Call �L�^�s�T��(����, �n, �I)
        If �n > 0 Then
            If �Y�� <> �I - �n + 1 Then
                MsgBox "�Ǘ��䒠�̋L�^�s���ƈ�v���܂���" & vbCrLf & vbCrLf & "�䒠�L�^�̍X�V�𒆎~���܂�"
                Exit Sub
            End If
        End If
        .Cells(�ŉ��s + 1, 1).Resize(�Y��, 9) = �z��
        .Cells(1, 1).Resize(�ŉ��s + �Y��, 9).Borders.LineStyle = True
        Select Case �n
            Case Is > 0
                .Rows(�n & ":" & �I).Delete
                Call �䒠����
                �� = "�䒠�L�^���X�V���܂���"
            Case Else
                �� = "�䒠�ɐV�K�o�^���܂���"
        End Select
        Call ���͋K���Đݒ�
        MsgBox ��
    End With
End Sub
Sub �䒠����()
    Dim �ŉ��s As Long
    With Sheets("�Ǘ��䒠")
        �ŉ��s = .Cells(Rows.Count, 2).End(xlUp).Row
        .Cells(1, 1).Resize(�ŉ��s, 9).Characters.PhoneticCharacters = ""
        With .Sort
            With .SortFields
                .Clear
                .Add Key:=Range("A1"), Order:=xlAscending
                .Add Key:=Range("B1"), Order:=xlAscending
            End With
            .SetRange Range(Cells(1, 1), Cells(�ŉ��s, 9))
            .Header = xlYes
            .Apply
        End With
    End With
End Sub
Sub �L�^�s�T��(���� As String, �n As Long, �I As Long)
    Dim �s As Long, �ŉ��s As Long
    With Sheets("�Ǘ��䒠")
        Call �䒠����
        �ŉ��s = .Cells(Rows.Count, 2).End(xlUp).Row
        For �s = 2 To �ŉ��s
            If .Cells(�s, 1) = ���� Then
                If �n = 0 Then �n = �s
                If �n > 0 Then �I = �s
                Else: If �I > 0 Then Exit For
            End If
        Next
    End With
End Sub
Sub �L�^�߂�()
    Dim ���� As String
    Dim �n As Long, �I As Long, �s As Long, �� As Long
    With Sheets("�Ǘ��䒠")
        ���� = Sheets("�ʃV�t�g�\").Range("����")
        Call �ʃV�t�g�\�N���A
        Call �L�^�s�T��(����, �n, �I)
        If �n = 0 Then: Exit Sub
        ReDim A��(�n To �I, 3 To 5)
        ReDim B��(�n To �I, 6 To 8)
        ReDim �x��(�n To �I, 9 To 9)
        For �s = �n To �I
            For �� = 3 To 9
                Select Case ��
                    Case 3, 4, 5: A��(�s, ��) = .Cells(�s, ��)
                    Case 6, 7, 8: B��(�s, ��) = .Cells(�s, ��)
                    Case 9: �x��(�s, ��) = .Cells(�s, ��)
                End Select
            Next
        Next
    End With
    With Sheets("�ʃV�t�g�\")
        .Cells(7, 2).Resize(�I - �n + 1, 3) = A��
        .Cells(7, 5).Resize(�I - �n + 1, 3) = B��
        .Cells(7, 8).Resize(�I - �n + 1, 1) = �x��
    End With
End Sub
Sub ���͋K���Đݒ�()
    Dim �ŉ��s As Long, �s As Long
    Dim ���� As String, ���X�g As String
    With Sheets("�Ǘ��䒠")
        Call �䒠����
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        ���� = .Cells(2, 1)
        ���X�g = ����
        For �s = 3 To �ŉ��s
            If .Cells(�s, 1) <> ���� Then
                 ���� = .Cells(�s, 1)
                ���X�g = ���X�g & "," & ����
            End If
        Next
    End With
    With Sheets("�ʃV�t�g�\").Range("����").Validation
        .Delete
        .Add Type:=xlValidateList, Operator:=xlEqual, Formula1:=���X�g
        .ShowError = False
    End With
End Sub
Sub �ʃV�t�g�\�N���A()
    With Sheets("�ʃV�t�g�\")
        .Cells(7, 2).Resize(31, 3).ClearContents
        .Cells(7, 5).Resize(31, 3).ClearContents
        .Cells(7, 8).Resize(31, 1).ClearContents
    End With
End Sub
Sub �ی�ؑ�()
    With ActiveSheet
        Select Case .ProtectContents
            Case True: .Unprotect: MsgBox "�V�[�g�ی���������܂���"
            Case False: .Protect: MsgBox "�V�[�g��ی삵�܂���"
        End Select
    End With
End Sub
