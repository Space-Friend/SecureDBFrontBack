VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_���_�����_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    
    If Forms![���_������].EveryoneCanSeeMe = 2 Then
        ������51.Enabled = True
    End If
    If Forms![���_������].EveryoneCanSeeMe = 3 Then
        ������51.Enabled = True
    End If
    
    If Forms![���_������].Tip = "" Then
        lbl1.Caption = Forms![���_������].Oboz
        Me.Filter = "����������� LIKE ""*" & Forms![���_������].Oboz & "*"" "
        ������61.UseTheme = True
        ������61.Enabled = True
        Me.FilterOn = True
    Else
        lbl1.Caption = Forms![���_������].Tip
        Me.Filter = "[��� ��������] LIKE ""*" & Forms![���_������].Tip & "*"" "
        ������61.UseTheme = True
        ������61.Enabled = True
        Me.FilterOn = True
    End If
End Sub

Private Sub ������61_Click()
    Me.FilterOn = False
    ������61.Caption = "������: ���"
    ������61.UseTheme = False
    ������61.Enabled = False
    lbl1.Caption = "���� �����."
End Sub

Private Sub ������51_Click()
    
    If MsgBox("�� ����������� ��������� ������� ����� '" & Me![�����] & "' �� ������� '" & Forms![���_������].Num & "', �� �����?", vbYesNo, "�������������") = vbYes Then
        Forms![���_������].Kod = Me![���]
        Dim db As Database
        Dim rs As Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset(OpenArgs)
        rs.FindFirst "���=" & Forms![���_������].Num & ""
        
        If rs![�����] = "0" Then
            rs.Edit
            rs![�����] = Me![���]
            rs![����� ���������] = Forms![���_������].Namesys
            rs.Update
            DoCmd.Close acForm, "���_�����_����"
        Else
            MsgBox "�������, �� ���� ����� ��� ����� �������. ������� ������� ���.", , "������!"
        End If
        
    End If
  
End Sub

Private Sub ������80_Click()
    Forms![���_������].Tip = ""
    Forms![���_������].Oboz = ""
    DoCmd.Close acForm, Me.Name
End Sub
