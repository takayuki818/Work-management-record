Attribute VB_Name = "���[�h�ؑ�"
Option Explicit
Sub �S�W�J()
    Dim �V�[�g
    Application.ScreenUpdating = False
    For Each �V�[�g In Sheets
        �V�[�g.Visible = True
    Next
    Sheets("MENU").Activate
    Application.ScreenUpdating = True
End Sub
Sub MENU�߂�()
    Dim ��\��(), �V�[�g��
    Application.ScreenUpdating = False
    Sheets("MENU").Visible = True
    ��\�� = Array("�ʃV�t�g�\", "�Ǘ��䒠", "���^�W�v", "�ӗ�W�v", "���ʓ��}�X�^", "�����}�X�^", "�v���_�E���ݒ�")
    For Each �V�[�g�� In ��\��
        Sheets(�V�[�g��).Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
Sub �V�t�g���̓��[�h()
    Dim ��\��(), �V�[�g��
    Application.ScreenUpdating = False
    Call �S�W�J
    ��\�� = Array("MENU", "���^�W�v", "�ӗ�W�v", "���ʓ��}�X�^", "�����}�X�^", "�v���_�E���ݒ�")
    For Each �V�[�g�� In ��\��
        Sheets(�V�[�g��).Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
Sub �W�v�Ǘ����[�h()
    Dim ��\��(), �V�[�g��
    Application.ScreenUpdating = False
    Call �S�W�J
    ��\�� = Array("MENU", "�ʃV�t�g�\", "���ʓ��}�X�^", "�����}�X�^", "�v���_�E���ݒ�")
    For Each �V�[�g�� In ��\��
        Sheets(�V�[�g��).Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
Sub �}�X�^�ݒ胂�[�h()
    Dim ��\��(), �V�[�g��
    Application.ScreenUpdating = False
    Call �S�W�J
    ��\�� = Array("MENU", "�ʃV�t�g�\", "�Ǘ��䒠", "���^�W�v", "�ӗ�W�v", "�v���_�E���ݒ�")
    For Each �V�[�g�� In ��\��
        Sheets(�V�[�g��).Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
