VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()

'
'DoCmd.ShowToolbar "Ribbon", acToolbarNo
'
    Forms![���_������].EveryoneCanSeeMe = Forms![���_������].lll.Caption
    If Forms![���_������].EveryoneCanSeeMe = "1" Then
        lbltest.Caption = "������� �����"
        ������694.Enabled = True
        Kart.Enabled = False
        sotr.Enabled = False
        log.Enabled = False
        logkart.Enabled = False
        shft.Enabled = False
        
    End If
    If Forms![���_������].EveryoneCanSeeMe = "2" Then
        lbltest.Caption = "����������� �����"
        ������694.Enabled = True
        Kart.Enabled = True
        sotr.Enabled = False
        log.Enabled = False
        logkart.Enabled = False
        shft.Enabled = False
        
    End If
    If Forms![���_������].EveryoneCanSeeMe = "3" Then
        lbltest.Caption = "����������������� �����"
        ������694.Enabled = True
        Kart.Enabled = True
        sotr.Enabled = True
        log.Enabled = True
        logkart.Enabled = True
        shft.Enabled = True
        
    End If
    
    On Error Resume Next
    ������694.SetFocus
    
End Sub

Private Sub log_Click()
    
    DoCmd.OpenForm "���_���_����"
    DoCmd.Close acForm, Me.Name
    
End Sub

Private Sub shft_Click()
    
    Dim prop As Property
    On Error GoTo Setproperty
    Set prop = CurrentDb.CreateProperty("allowbypasskey", dbBoolean, False)
    
    CurrentDb.Properties.Append prop
    
Setproperty:
    If MsgBox("�� ������ �������� ������� ������� (Shift)?  � ����� ������ �� ����� �������� ���� ������ ��� ����������� ���� ������� �� ��������� � ������ �������������� �� ������������ �����������.", vbYesNo, "�������� ������� �������") = vbYes Then
        CurrentDb.Properties("allowbypasskey") = True
    Else
        CurrentDb.Properties("allowbypasskey") = False
    End If
    
End Sub

Private Sub logkart_Click()
    
    DoCmd.OpenForm "���_����_��������"
    DoCmd.Close acForm, Me.Name
    
End Sub

Private Sub sotr_Click()
    
        If Forms![���_������].EveryoneCanSeeMe = "3" Then
        DoCmd.OpenForm "���_����������_�����", , , , , acDialog
    Else: DoCmd.OpenForm "���_����������", , , , , acDialog
    End If
    
End Sub


Private Sub ������100_Click()
    MsgBox "���������� ������ 774-5. ������ ��������� �� 30.11.2022 To space with Space!", vbInformation, "������"
End Sub

Private Sub ������33_Click()
    DoCmd.OpenForm "���_�����"
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub ������663_Click()
    DoCmd.Close acForm, Me.Name
    DoCmd.CloseDatabase
End Sub

Private Sub ������77_Click()

End Sub
