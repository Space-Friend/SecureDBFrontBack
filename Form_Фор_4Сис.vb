VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_���_4���"
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
    Call AuditEditEnd("���_��������", "audTmp_���_��������", "aud_���_��������", "���", Nz(Me![���_��������.���], 0), bWasNewRecord)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    bWasNewRecord = Me.NewRecord
    Call AuditEditBegin("���_��������", "audTmp_���_��������", "���", Nz(Me![���_��������.���], 0), bWasNewRecord)
End Sub

Private Sub Form_Delete(Cancel As Integer)
    Call AuditDelBegin("���_��������", "audTmp_���_��������", "���", Nz(Me.���, 0))
End Sub

Private Sub Form_Load()
        
    Form.AllowEdits = False
    ������80.SetFocus
    If Forms![���_������].lll.Caption = "1" Then
        lbl1.Caption = "� ������ - ���������"
        lbl2.Caption = "� ��������� - ���������"
        lbl3.Caption = "� �������� - ���������"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(186, 20, 25)
        Me.lbl3.ForeColor = RGB(186, 20, 25)
        kno1.Enabled = False
        ������91.Enabled = False
        ������3073.Enabled = False
        ������102.Enabled = False
    End If
    If Forms![���_������].lll.Caption = "2" Then
        lbl1.Caption = "� ������ - ���������"
        lbl2.Caption = "� ��������� - ���������"
        lbl3.Caption = "� �������� - ���������"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(63, 118, 42)
        Me.lbl3.ForeColor = RGB(186, 20, 25)
        kno1.Enabled = True
        ������91.Enabled = True
        ������3073.Enabled = True
        ������102.Enabled = True
    End If
    If Forms![���_������].lll.Caption = "3" Then
        lbl1.Caption = "� ������ - ���������"
        lbl2.Caption = "� ��������� - ���������"
        lbl3.Caption = "� �������� - ���������"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(63, 118, 42)
        Me.lbl3.ForeColor = RGB(63, 118, 42)
        kno1.Enabled = True
        ������91.Enabled = True
        ������3073.Enabled = True
        ������102.Enabled = True
    End If

End Sub

Private Sub kno1_Click()
    Form.AllowEdits = True
End Sub

Private Sub ������102_Click()
    
        Dim db As Database
        Dim rs As Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("���_4���")
        rs.FindFirst "���=" & Me![���] & ""
        If rs![�����] = 0 Then End
    If MsgBox("�� ����������� ��������� ������� �� ����� � ������� '" & rs![���] & "', �� �����?", vbYesNo, "�������������") = vbYes Then
        rs.Edit
        rs![����� ���������] = "�� ������"
        rs.Update
        rs.Edit
        rs![�����] = "0"
        rs.Update
        MsgBox ("���������! �� �������� �������� ���� ��������� ��������.")
        DoCmd.RefreshRecord
        
    End If
  
End Sub

Public Sub ������65_Click()
    
    Forms![���_������].EveryoneCanSeeMe = Forms![���_������].lll.Caption
    Forms![���_������].Namesys = "����� ������ ����� ��������� ���� 2 1.04.0000.141 �2"
    Forms![���_������].Num = DLookup("[���]", "���_4���", "[���] = " & Forms!���_4���!���)
    Forms![���_������].Tip = ""
    On Error GoTo ohsht
    Forms![���_������].Oboz = DLookup("[�������]", "���_4���", "[���] = " & Forms!���_4���!���)
    Forms![���_������].EveryoneCanSeeMe = Forms![���_������].lll.Caption
    DoCmd.OpenForm "���_�����_����", , , , acFormReadOnly, acDialog, "���_4���"
    DoCmd.RefreshRecord
    
End

ohsht:
    If MsgBox("� �������� ����������� �����������, ����� ��������� ����� ��������. ��������� ����� �� ����?", vbYesNo) = vbYes Then
        Forms![���_������].Tip = DLookup("[�����]", "���_4���", "[���] = " & Forms!���_4���!���)
        Forms![���_������].EveryoneCanSeeMe = Forms![���_������].lll.Caption
        DoCmd.OpenForm "���_�����_����", , , , acFormReadOnly, acDialog, "���_4���"
        DoCmd.RefreshRecord
    End If
End Sub

Private Sub ������61_Click()

    Me.FilterOn = False
    ������61.Caption = "������: ���"
    ������61.UseTheme = False
    ������61.Enabled = False

End Sub

Private Sub ������67_Click()
'
    Dim St As String
    St = InputBox("����������, ������� ����� ��� ������ �������� ���� ��������", "��� ��������", "������")
    If St = "" Then Exit Sub
    
    Me.Filter = "[�����] LIKE ""*" & St & "*"" "
    Me.FilterOn = True
    ������61.Caption = "������ �����"
    ������61.UseTheme = True
    ������61.Enabled = True

End Sub

Private Sub ������68_Click()
'
    Dim D As String
    
    D = InputBox("����������, ������� ����� ��� ������ �������� ����������� ��������", "�����������", "� 100")
    If D = "" Then Exit Sub
    
    Me.Filter = "[�������] LIKE ""*" & D & "*"" "
    Me.FilterOn = True
    ������61.Caption = "������ �����"
    ������61.UseTheme = True
    ������61.Enabled = True
    
End Sub

Private Sub ������73_Click()

    Me.FilterOn = False
    ������61.Caption = "������: ���"
    ������61.UseTheme = False
    ������61.Enabled = False
    
End Sub

Private Sub ������74_Click()
    
On Error GoTo ohsht
    Forms![���_������].EveryoneCanSeeMe = Forms![���_������].lll.Caption
    DoCmd.OpenForm "���_��������_��������", acNormal, , "[���] =" & Me![���_��������.���]
End

ohsht:
    If MsgBox("������ ������� �������� ��������������� ��������. ������� ��������� ����� ��������?", vbYesNo) = vbYes Then
        Call ������65_Click
    End If
End

End Sub

Private Sub ������80_Click()
    
    Forms![���_������].EveryoneCanSeeMe = Forms![���_������].lll.Caption
    DoCmd.Close acForm, Me.Name, acSavePrompt
    DoCmd.OpenForm "�����"
    
End Sub

Private Sub ������96_Click()
    If ������96.Caption = "��� /\" Then
        Me.OrderByOn = False
        ������96.Caption = "���"
        ������97.Caption = "���"
        ������98.Caption = "���"
        ������_����_���.Caption = "����� ���������"
        ������_�����.Caption = "��� ��������"
        ������_����_������.Caption = "�����������"
        ������_���_�����.Caption = "��������� �����"
        ������49.Caption = "�������"
        ������_���_�����.Caption = "���� �������"
        ������108.Caption = "����."
        ������21.Caption = "��"
        ������39.Caption = "���� ���������"
    Else
        Me.OrderBy = "���_1����.��� DESC"
        Me.OrderByOn = True
        ������98.Caption = "���"
        ������96.Caption = "��� /\"
        ������97.Caption = "���"
        ������_����_���.Caption = "����� ���������"
        ������_�����.Caption = "��� ��������"
        ������_����_������.Caption = "�����������"
        ������_���_�����.Caption = "��������� �����"
        ������49.Caption = "�������"
        ������_���_�����.Caption = "���� �������"
        ������108.Caption = "����."
        ������21.Caption = "��"
        ������39.Caption = "���� ���������"
    End If
End Sub

Private Sub ������97_Click()
    If ������97.Caption = "��� /\" Then
        Me.OrderByOn = False
        ������96.Caption = "���"
        ������97.Caption = "���"
        ������98.Caption = "���"
        ������_����_���.Caption = "����� ���������"
        ������_�����.Caption = "��� ��������"
        ������_����_������.Caption = "�����������"
        ������_���_�����.Caption = "��������� �����"
        ������49.Caption = "�������"
        ������_���_�����.Caption = "���� �������"
        ������108.Caption = "����."
        ������21.Caption = "��"
        ������39.Caption = "���� ���������"
    Else
        Me.OrderBy = "���_1����.��� DESC"
        Me.OrderByOn = True
        ������98.Caption = "���"
        ������96.Caption = "���"
        ������97.Caption = "��� /\"
        ������_����_���.Caption = "����� ���������"
        ������_�����.Caption = "��� ��������"
        ������_����_������.Caption = "�����������"
        ������_���_�����.Caption = "��������� �����"
        ������49.Caption = "�������"
        ������_���_�����.Caption = "���� �������"
        ������108.Caption = "����."
        ������21.Caption = "��"
        ������39.Caption = "���� ���������"
    End If
End Sub

Private Sub ������98_Click()
    If ������98.Caption = "��� /\" Then
        Me.OrderByOn = False
        ������96.Caption = "���"
        ������97.Caption = "���"
        ������98.Caption = "���"
        ������_����_���.Caption = "����� ���������"
        ������_�����.Caption = "��� ��������"
        ������_����_������.Caption = "�����������"
        ������_���_�����.Caption = "��������� �����"
        ������49.Caption = "�������"
        ������_���_�����.Caption = "���� �������"
        ������108.Caption = "����."
        ������21.Caption = "��"
        ������39.Caption = "���� ���������"
    Else
        Me.OrderBy = "���_1����.��� DESC"
        Me.OrderByOn = True
        ������98.Caption = "��� /\"
        ������96.Caption = "���"
        ������97.Caption = "���"
        ������_����_���.Caption = "����� ���������"
        ������_�����.Caption = "��� ��������"
        ������_����_������.Caption = "�����������"
        ������_���_�����.Caption = "��������� �����"
        ������49.Caption = "�������"
        ������_���_�����.Caption = "���� �������"
        ������108.Caption = "����."
        ������21.Caption = "��"
        ������39.Caption = "���� ���������"
    End If
End Sub

