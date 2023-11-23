Attribute VB_Name = "モード切替"
Option Explicit
Sub 全展開()
    Dim シート
    Application.ScreenUpdating = False
    For Each シート In Sheets
        シート.Visible = True
    Next
    Sheets("MENU").Activate
    Application.ScreenUpdating = True
End Sub
Sub MENU戻り()
    Dim 非表示(), シート名
    Application.ScreenUpdating = False
    Sheets("MENU").Visible = True
    非表示 = Array("個別シフト表", "管理台帳", "給与集計", "謝礼集計", "特別日マスタ", "氏名マスタ", "プルダウン設定")
    For Each シート名 In 非表示
        Sheets(シート名).Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
Sub シフト入力モード()
    Dim 非表示(), シート名
    Application.ScreenUpdating = False
    Call 全展開
    非表示 = Array("MENU", "給与集計", "謝礼集計", "特別日マスタ", "氏名マスタ", "プルダウン設定")
    For Each シート名 In 非表示
        Sheets(シート名).Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
Sub 集計管理モード()
    Dim 非表示(), シート名
    Application.ScreenUpdating = False
    Call 全展開
    非表示 = Array("MENU", "個別シフト表", "特別日マスタ", "氏名マスタ", "プルダウン設定")
    For Each シート名 In 非表示
        Sheets(シート名).Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
Sub マスタ設定モード()
    Dim 非表示(), シート名
    Application.ScreenUpdating = False
    Call 全展開
    非表示 = Array("MENU", "個別シフト表", "管理台帳", "給与集計", "謝礼集計", "プルダウン設定")
    For Each シート名 In 非表示
        Sheets(シート名).Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
