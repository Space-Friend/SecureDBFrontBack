VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_���_������_���������"
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

Private Sub Form_Load()

    If Forms![���_������].EveryoneCanSeeMe = "1" Then
        Me.AllowDeletions = False
        lbl1.Caption = "� ������ - ���������"
        lbl2.Caption = "� ��������� - ���������"
        lbl3.Caption = "� �������� - ���������"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(186, 20, 25)
        Me.lbl3.ForeColor = RGB(186, 20, 25)
        kno1.Enabled = False
        ������91.Enabled = False
        ������3073.Enabled = False
        ������94.Enabled = False
    End If
    If Forms![���_������].EveryoneCanSeeMe = "2" Then
        Me.AllowDeletions = False
        lbl1.Caption = "� ������ - ���������"
        lbl2.Caption = "� ��������� - ���������"
        lbl3.Caption = "� �������� - ���������"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(63, 118, 42)
        Me.lbl3.ForeColor = RGB(186, 20, 25)
        kno1.Enabled = True
        ������91.Enabled = True
        ������3073.Enabled = True
        ������94.Enabled = False
    End If
    If Forms![���_������].EveryoneCanSeeMe = "3" Then
        Me.AllowDeletions = True
        lbl1.Caption = "� ������ - ���������"
        lbl2.Caption = "� ��������� - ���������"
        lbl3.Caption = "� �������� - ���������"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(63, 118, 42)
        Me.lbl3.ForeColor = RGB(63, 118, 42)
        kno1.Enabled = True
        ������91.Enabled = True
        ������3073.Enabled = True
        ������94.Enabled = True
    End If
    
End Sub

Private Sub kno1_Click()
    Form.AllowEdits = True
End Sub

Private Sub ������60_Click()

    Dim s As String
    
    s = InputBox("����������, ������� ����� ��� ������ �������� ����� ��������� ��������", "����� ���������", "������� ������ ������������� �����")
    If s = "" Then Exit Sub

    Me.Filter = "[����� ���������] LIKE ""*" & s & "*"" "
    Me.FilterOn = True
    ������61.Caption = "������ �����"
    ������61.UseTheme = True
    ������61.Enabled = True
    
End Sub

Private Sub ������61_Click()

    Me.FilterOn = False
    ������61.Caption = "������: ���"
    ������61.UseTheme = False
    ������61.Enabled = False

End Sub

Private Sub ������67_Click()

    Dim St As String
    St = InputBox("����������, ������� ����� ��� ������ �������� ���� ��������", "��� ��������", "������")
    If St = "" Then Exit Sub
    
    Me.Filter = "[��� ��������] LIKE ""*" & St & "*"" "
    Me.FilterOn = True
    ������61.Caption = "������ �����"
    ������61.UseTheme = True
    ������61.Enabled = True

End Sub

Private Sub ������68_Click()

    Dim D As String
    
    D = InputBox("����������, ������� ����� ��� ������ �������� ����������� ��������", "�����������", "� 100")
    If D = "" Then Exit Sub
    
    Me.Filter = "[�����������] LIKE ""*" & D & "*"" "
    Me.FilterOn = True
    ������61.Caption = "������ �����"
    ������61.UseTheme = True
    ������61.Enabled = True
    
End Sub

Private Sub ������69_Click()

    Dim D As String
    
    D = InputBox("����������, ������� ����� ��� ������ �������� ������ ��������", "�����")
    If D = "" Then Exit Sub
    
    Me.Filter = "[�����] LIKE ""*" & D & "*"" "
    Me.FilterOn = True
    ������61.Caption = "������ �����"
    ������61.UseTheme = True
    ������61.Enabled = True
    
End Sub

Private Sub ������70_Click()

    Dim D As String
    
    D = InputBox("����������, ������� ��� ��������", "���", "001")
    If D = "" Then Exit Sub
    
    Me.Filter = "[���] LIKE ""*" & D & "*"" "
    Me.FilterOn = True
    ������61.Caption = "������ �����"
    ������61.UseTheme = True
    ������61.Enabled = True
    
End Sub

Private Sub ������71_Click()

    Dim D As String
    
    D = InputBox("����������, ������� ��� ��������", "���", "10")
    If D = "" Then Exit Sub
    
    Me.Filter = "[���] LIKE " & D & " "
    Me.FilterOn = True
    ������61.Caption = "������ �����"
    ������61.UseTheme = True
    ������61.Enabled = True
    
End Sub

Private Sub ������72_Click()

    Dim D As String
    
    D = InputBox("����������, ������� ��� ��������", "���", "0.9")
    If D = "" Then Exit Sub
    
    Me.Filter = "[���] LIKE ""*" & D & "*"" "
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
    
    Forms![���_������].EveryoneCanSeeMe = Forms![���_������].lll.Caption
    DoCmd.OpenForm "���_��������_��������", , , "[���] =" & Me![���], acFormReadOnly

End Sub
