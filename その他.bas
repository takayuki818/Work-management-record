Attribute VB_Name = "���̑�"
Option Explicit
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
    Dim ���� As String, ���� As Date
    Dim �s As Long, �ŉ��s As Long
    Dim �߂�l()
    With Sheets("���^�W�v")
        ���� = .Range("�䒠_�W�v��")
    End With
    With Sheets("�ӗ�W�v")
        �ŉ��s = .Cells(Rows.Count, 9).End(xlUp).Row
        ReDim ���ʓ����X�g(1 To �ŉ��s - 3)
        Select Case �ŉ��s
            Case Is > 3
                For �s = 4 To �ŉ��s
                    ���ʓ����X�g(�s - 3) = .Cells(�s, 9)
                Next
            Case Else
                MsgBox "���ʓ����X�g���ݒ肳��Ă��܂���"
                Exit Sub
        End Select
        
        �ŉ��s = .Cells(Rows.Count, 12).End(xlUp).Row
        ReDim ��(1 To �ŉ��s - 3, 1 To 1)
        ReDim ��(1 To �ŉ��s - 3, 1 To 1)
        ReDim ���l(1 To �ŉ��s - 3, 1 To 1)
        For �s = 4 To �ŉ��s
            ���� = .Cells(�s, 12)
            �߂�l = ���ʓ��Ζ��擾(����, ����, ���ʓ����X�g)
            ��(�s - 3, 1) = �߂�l(1)
            ��(�s - 3, 1) = �߂�l(2)
            ���l(�s - 3, 1) = �߂�l(3)
        Next
        .Cells(4, 13).Resize(�ŉ��s - 3, 1) = ��
        .Cells(4, 16).Resize(�ŉ��s - 3, 1) = ��
        .Cells(4, 23).Resize(�ŉ��s - 3, 1) = ���l
    End With
End Sub
Function ���ʓ��Ζ��擾(���� As String, ���� As Date, ���ʓ����X�g As Variant)
    Dim �s As Long, �ŉ��s As Long
    Dim ���ʓ� As Variant
    Dim �߂�l(1 To 3)
    With Sheets("�Ǘ��䒠")
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        For �s = 2 To �ŉ��s
            If .Cells(�s, 1) = ���� And .Cells(�s, 10) = ���� Then
                For Each ���ʓ� In ���ʓ����X�g
                    If .Cells(�s, 2) = ���ʓ� Then
                        If .Cells(�s, 11) > 0 Then �߂�l(1) = �߂�l(1) + 1
                        If .Cells(�s, 13) > 0 Then �߂�l(2) = �߂�l(2) + 1
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
    Dim ���� As String, ���� As Date
    Dim �s As Long, �ŉ��s As Long
    With Sheets("���^�W�v")
        ���� = .Range("�䒠_�W�v��")
    End With
    With Sheets("�ӗ�W�v")
        �ŉ��s = .Cells(Rows.Count, 2).End(xlUp).Row
        ReDim �z��(1 To �ŉ��s - 3, 1 To 1)
        For �s = 4 To �ŉ��s
            ���� = .Cells(�s, 2)
            �z��(�s - 3, 1) = �Ζ����擾(����, ����)
        Next
        .Cells(4, 7).Resize(�ŉ��s - 3, 1) = �z��
    End With
End Sub
Function �Ζ����擾(���� As String, ���� As Date) As String
    Dim �s As Long, �ŉ��s As Long
    With Sheets("�Ǘ��䒠")
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        For �s = 2 To �ŉ��s
            If .Cells(�s, 1) = ���� And .Cells(�s, 10) = ���� Then
                Select Case �Ζ����擾
                    Case "": �Ζ����擾 = "�Ζ����F" & Format(.Cells(�s, 2), "d")
                    Case Else: �Ζ����擾 = �Ζ����擾 & "," & Format(.Cells(�s, 2), "d")
                End Select
            End If
        Next
    End With
End Function
