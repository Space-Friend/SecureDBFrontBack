VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_���_�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
    DoCmd.OpenForm "���_������", , , , acFormEdit, acHidden
End Sub

Private Sub ������11_Click()

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("���_����������", dbOpenSnapshot, dbReadOnly)
    rs.FindFirst "�����='" & Me.log & "'"
    
    If rs.NoMatch = True Then
        Me.�������7.Visible = True
        Me.�������8.Visible = False
        Me.log.SetFocus
        Exit Sub
    End If
    Me.�������7.Visible = False

    If rs!������ = Me.par Then
        
        Dim a
        Dim StrSQL As String
        a = rs![��� ����������]
        StrSQL = "INSERT INTO [���_���_����] ([��� ����������]) VALUES ('" & a & "' );"
        DoCmd.SetWarnings False
        DoCmd.RunSQL StrSQL
        
        If rs!������� = 1 Then
            Forms![���_������].EveryoneCanSeeMe = "1"
        End If
        If rs!������� = 2 Then
            Forms![���_������].EveryoneCanSeeMe = "2"
        End If
        If rs!������� = 3 Then
            Forms![���_������].EveryoneCanSeeMe = "3"
        End If
        
        Forms![���_������].lll.Caption = Forms![���_������].EveryoneCanSeeMe
        DoCmd.OpenForm "�����"
        DoCmd.Close acForm, "���_�����"
    Else
        Me.�������8.Visible = True
        Me.par.SetFocus
    End If
    
End Sub
