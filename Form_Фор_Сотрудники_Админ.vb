VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_���_����������_�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    ������73.UseTheme = False
    ������73.Enabled = False
End Sub

Private Sub ������16_Click()

    ������73.Enabled = True
    ������73.UseTheme = True
    Dim St As String
    St = InputBox("������� ���", "����� �� �����", "���")
    If St = "" Then Exit Sub
    
    Me.Filter = "��� LIKE ""*" & St & "*"""
    Me.FilterOn = True

End Sub

Private Sub ������17_Click()

    ������73.Enabled = True
    ������73.UseTheme = True
    Dim St As String
    St = InputBox("������� �������", "����� �� �������", "�������")
    If St = "" Then Exit Sub
    
    Me.Filter = "������� LIKE ""*" & St & "*"""
    Me.FilterOn = True

End Sub

Private Sub ������18_Click()

    ������73.Enabled = True
    ������73.UseTheme = True
    Dim St As String
    St = InputBox("������� ��������", "����� �� ��������", "��������")
    If St = "" Then Exit Sub
    
    Me.Filter = "�������� LIKE ""*" & St & "*"""
    Me.FilterOn = True

End Sub

Private Sub ������73_Click()
    Me.FilterOn = False
    ������73.Enabled = False
    ������73.UseTheme = False
End Sub
