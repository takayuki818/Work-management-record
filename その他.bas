Attribute VB_Name = "���̑�"
Option Explicit
Sub �W�v�V�[�g�\�z()
    Dim �s�� As Long, �ŉ��s As Long, �Y�� As Long
    Dim ���^(), ���Ɨp��(), ������(), ������()
    With Sheets("�����}�X�^")
        �s�� = .Cells(Rows.Count, 4).End(xlUp).Row - 1
        If �s�� < 1 Then
            MsgBox "�����}�X�^���쐬����Ă��܂���"
            Exit Sub
        End If
        ���^ = .Cells(2, 1).Resize(�s��, 5).Value
        ReDim �ӗ玁��(1 To �s��, 1 To 2)
        For �Y�� = 1 To �s��
            �ӗ玁��(�Y��, 1) = .Cells(�Y�� + 1, 2)
            �ӗ玁��(�Y��, 2) = .Cells(�Y�� + 1, 4)
        Next
        ���Ɨp�� = .Cells(2, 6).Resize(�s��, 1).Value
        ������ = .Cells(2, 7).Resize(�s��, 1).Value
        ������ = .Cells(2, 8).Resize(�s��, 1).Value
    End With
    With Sheets("���^�W�v")
        �ŉ��s = .Cells(Rows.Count, 4).End(xlUp).Row
        .Cells(7, 1).Resize(�ŉ��s - 6, 31).Clear
        .Cells(6, 1).Resize(1, 31).AutoFill .Cells(6, 1).Resize(�s��, 31)
        .Cells(6, 1).Resize(�s��, 5) = ���^
    End With
    With Sheets("�ӗ�W�v")
        �ŉ��s = .Cells(Rows.Count, 2).End(xlUp).Row
        .Cells(6, 1).Resize(�ŉ��s - 5, 21).Clear
        .Cells(5, 1).Resize(1, 21).AutoFill .Cells(5, 1).Resize(�s��, 21)
        .Cells(5, 1).Resize(�s��, 2) = �ӗ玁��
        .Cells(5, 4).Resize(�s��, 1) = ���Ɨp��
        .Cells(5, 9).Resize(�s��, 2) = �ӗ玁��
        .Cells(5, 12).Resize(�s��, 1) = ������
        .Cells(5, 15).Resize(�s��, 1) = ������
        Call �S�Ζ����ݒ�
        Call ���ʓ��Ζ��ݒ�
        Call �W�v�Ǘ����[�h
    End With
    MsgBox "�W�v�V�[�g�̍\�z���������܂���"
End Sub
Sub �v���_�E���ݒ�(���[�h As String) '256���ȏ�̕�����`���Ńv���_�E���ݒ肷��ƃt�@�C�����j�����邱�Ƃɒ��ӁI
    Dim �ŉ��s As Long, �s As Long, �Y�� As Long
    Dim �͈͎� As String
    With Sheets("���^�W�v")
        �ŉ��s = .Cells(Rows.Count, 4).End(xlUp).Row
        If �ŉ��s < 6 Then �ŉ��s = 6
        ReDim ���X�g(1 To �ŉ��s - 5, 1 To 1)
        For �s = 6 To �ŉ��s
            Select Case ���[�h
                Case "�����S��"
                    ���X�g(�s - 5, 1) = .Cells(�s, 4)
                Case "�v�x�������̂�"
                    If .Cells(�s, 30) > 0 Then
                        �Y�� = �Y�� + 1
                        ���X�g(�Y��, 1) = .Cells(�s, 4)
                    End If
            End Select
        Next
    End With
    With Sheets("�v���_�E���ݒ�")
        .Cells(1, 1).Resize(Rows.Count, 1).ClearContents
        .Cells(1, 1).Resize(�ŉ��s - 5, 1) = ���X�g
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        �͈͎� = "=�v���_�E���ݒ�!$A$1:$A$" & �ŉ��s
    End With
    With Sheets("�ʃV�t�g�\")
        .Unprotect
        With .Range("����").Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=�͈͎�
        End With
        .Protect
    End With
End Sub
Sub ���^�V�t�g���()
    Dim �s As Long, �z��()
    With Sheets("�ʃV�t�g�\")
        �s = ActiveCell.Row
        If �s > 6 And �s < 38 Then
            �z�� = .Range("��W���V�t�g").Value
            .Cells(�s, 2).Resize(1, 3) = �z��
        End If
    End With
End Sub
Sub ����^�V�t�g���()
    Dim �s As Long, �z��()
    With Sheets("�ʃV�t�g�\")
        �s = ActiveCell.Row
        If �s > 6 And �s < 38 Then
            �z�� = .Range("���W���V�t�g").Value
            .Cells(�s, 7).Resize(1, 3) = �z��
        End If
    End With
End Sub
Sub �ʃV�t�g�\�N���A()
    With Sheets("�ʃV�t�g�\")
        .Cells(7, A�n��).Resize(31, 3).ClearContents
        .Cells(7, B�n��).Resize(31, 3).ClearContents
        .Cells(7, �x��).Resize(31, 1).ClearContents
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
Sub ���ʓ��Ζ��ݒ�()
    Dim ���� As String, ���� As String
    Dim �s As Long, �ŉ��s As Long
    Dim �߂�l()
    With Sheets("���^�W�v")
        ���� = Format(.Range("�䒠_�W�v��"), "ge.m")
    End With
    With Sheets("���ʓ��}�X�^")
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim ���ʓ����X�g(1 To �ŉ��s - 1)
        Select Case �ŉ��s
            Case Is > 1
                For �s = 2 To �ŉ��s
                    ���ʓ����X�g(�s - 1) = .Cells(�s, 1)
                Next
            Case Else
                MsgBox "���ʓ����X�g���ݒ肳��Ă��܂���"
                Exit Sub
        End Select
    End With
    With Sheets("�ӗ�W�v")
        �ŉ��s = .Cells(Rows.Count, 10).End(xlUp).Row
        ReDim ��(1 To �ŉ��s - 4, 1 To 1)
        ReDim ��(1 To �ŉ��s - 4, 1 To 1)
        ReDim ���l(1 To �ŉ��s - 4, 1 To 1)
        For �s = 5 To �ŉ��s
            ���� = .Cells(�s, 10)
            �߂�l = ���ʓ��Ζ��擾(����, ����, ���ʓ����X�g)
            ��(�s - 4, 1) = �߂�l(1)
            ��(�s - 4, 1) = �߂�l(2)
            ���l(�s - 4, 1) = �߂�l(3)
        Next
        .Cells(5, 11).Resize(�ŉ��s - 4, 1) = ��
        .Cells(5, 14).Resize(�ŉ��s - 4, 1) = ��
        .Cells(5, 21).Resize(�ŉ��s - 4, 1) = ���l
    End With
End Sub
Function ���ʓ��Ζ��擾(���� As String, ���� As String, ���ʓ����X�g As Variant)
    Dim �s As Long, �ŉ��s As Long
    Dim ���ʓ� As Variant
    Dim �߂�l(1 To 3)
    With Sheets("�Ǘ��䒠")
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        For �s = 2 To �ŉ��s
            If .Cells(�s, 1) = ���� And Format(.Cells(�s, 10), "ge.m") = ���� Then
                For Each ���ʓ� In ���ʓ����X�g
                    If .Cells(�s, 2) = ���ʓ� Then
                        If .Cells(�s, 13) > 0 Then �߂�l(1) = �߂�l(1) + 1
                        If .Cells(�s, 11) > 0 Then �߂�l(2) = �߂�l(2) + 1
                        Select Case �߂�l(3)
                            Case "": �߂�l(3) = "�Ζ����F" & Format(.Cells(�s, 2), "d")
                            Case Else: �߂�l(3) = �߂�l(3) & "," & Format(.Cells(�s, 2), "d")
                        End Select
                        Exit For
                    End If
                Next
            End If
        Next
    End With
    ���ʓ��Ζ��擾 = �߂�l
End Function
Sub �S�Ζ����ݒ�()
    Dim ���� As String, ���� As String
    Dim �s As Long, �ŉ��s As Long
    With Sheets("���^�W�v")
        ���� = Format(.Range("�䒠_�W�v��"), "ge.m")
    End With
    With Sheets("�ӗ�W�v")
        �ŉ��s = .Cells(Rows.Count, 2).End(xlUp).Row
        ReDim �z��(1 To �ŉ��s - 4, 1 To 1)
        For �s = 5 To �ŉ��s
            ���� = .Cells(�s, 2)
            �z��(�s - 4, 1) = �Ζ����擾(����, ����)
        Next
        .Cells(5, 7).Resize(�ŉ��s - 4, 1) = �z��
    End With
End Sub
Function �Ζ����擾(���� As String, ���� As String) As String
    Dim �s As Long, �ŉ��s As Long
    With Sheets("�Ǘ��䒠")
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        For �s = 2 To �ŉ��s
            If .Cells(�s, 1) = ���� And Format(.Cells(�s, 10), "ge.m") = ���� Then
                Select Case �Ζ����擾
                    Case "": �Ζ����擾 = "�Ζ����F" & Format(.Cells(�s, 2), "d")
                    Case Else: �Ζ����擾 = �Ζ����擾 & "," & Format(.Cells(�s, 2), "d")
                End Select
            End If
        Next
    End With
End Function
