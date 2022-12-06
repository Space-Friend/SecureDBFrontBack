VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_���_��������_��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bWasNewRecord As Boolean

Private Sub Form_AfterDelConfirm(Status As Integer)
    Call AuditDelEnd("audTmp_���_��������", "aud_���_��������", Status)
End Sub

Private Sub Form_AfterUpdate()
    Call AuditEditEnd("���_��������", "audTmp_���_��������", "aud_���_��������", "���", Nz(Me!���, 0), bWasNewRecord)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    bWasNewRecord = Me.NewRecord
    Call AuditEditBegin("���_��������", "audTmp_���_��������", "���", Nz(Me.���, 0), bWasNewRecord)
End Sub

Private Sub Form_Delete(Cancel As Integer)
    Call AuditDelBegin("���_��������", "audTmp_���_��������", "���", Nz(Me.���, 0))
End Sub

Private Sub ������2743_Click()
    DoCmd.Close
End Sub

Private Sub Form_Load()

    If Forms![���_������].EveryoneCanSeeMe = "1" Then
        lbl1.Caption = "������� �����"
        ������260.Enabled = False
        kno1.Enabled = False
        kno26.Enabled = False
        ������256.Enabled = False
        ������200.Enabled = False
    End If
    If Forms![���_������].EveryoneCanSeeMe = "2" Then
        lbl1.Caption = "����������� �����"
        ������260.Enabled = True
        kno1.Enabled = True
        kno26.Enabled = True
        ������256.Enabled = True
        ������200.Enabled = False
    End If
    If Forms![���_������].EveryoneCanSeeMe = "3" Then
        lbl1.Caption = "����������������� �����"
        ������260.Enabled = True
        kno1.Enabled = True
        kno26.Enabled = True
        ������256.Enabled = True
        ������200.Enabled = True
    End If
    
End Sub

Private Sub kno1_Click()
    Form.AllowEdits = True
End Sub

Private Sub ������239_Click()
    
    If Forms![���_������].EveryoneCanSeeMe = "3" Then
        DoCmd.OpenForm "���_����������_�����", , , , , acDialog
    Else: DoCmd.OpenForm "���_����������", , , , , acDialog
    End If
    
End Sub

Private Sub ������308_Click()
    
Setproperty:
    If MsgBox("����� �������������� �� ������? (����������� ��������)", vbYesNo, "����� ��������������") = vbYes Then
        ���_������.Value = ""
        DoCmd.RefreshRecord
    End If
    
End Sub

Private Sub ������309_Click()
    
Setproperty:
    If MsgBox("����� �������������� �� ��������? (����������� ��������)", vbYesNo, "����� ��������������") = vbYes Then
        ���_��������.Value = ""
        DoCmd.RefreshRecord
    End If
    
End Sub

Private Sub ������324_Click()
    
    On Error GoTo sht
    DoCmd.OpenForm "���_����_��������_��������", acNormal, , "[���] =" & Me![���], , acDialog
End

sht:
    MsgBox "������ ����������� �������. ��������� ����������� ��������."

End Sub
