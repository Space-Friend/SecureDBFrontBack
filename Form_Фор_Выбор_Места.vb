VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_���_�����_�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    ������80.SetFocus
End Sub

Private Sub sys1_Click()
    DoCmd.Close acForm, "���_�����_�����"
    DoCmd.Close acForm, "�����"
    DoCmd.OpenForm "���_1���"
End Sub

Private Sub sys2_Click()
    DoCmd.Close acForm, "���_�����_�����"
    DoCmd.Close acForm, "�����"
    DoCmd.OpenForm "���_2���"
End Sub
Private Sub sys3_Click()
    DoCmd.Close acForm, "���_�����_�����"
    DoCmd.Close acForm, "�����"
    DoCmd.OpenForm "���_3���"
End Sub
Private Sub sys4_Click()
    DoCmd.Close acForm, "���_�����_�����"
    DoCmd.Close acForm, "�����"
    DoCmd.OpenForm "���_4���"
End Sub
Private Sub sys5_Click()
    DoCmd.Close acForm, "���_�����_�����"
    DoCmd.Close acForm, "�����"
    DoCmd.OpenForm "���_5���"
End Sub
Private Sub sys6_Click()
    DoCmd.Close acForm, "���_�����_�����"
    DoCmd.Close acForm, "�����"
    DoCmd.OpenForm "���_6���"
End Sub
Private Sub sys7_Click()
    DoCmd.Close acForm, "���_�����_�����"
    DoCmd.Close acForm, "�����"
    DoCmd.OpenForm "���_7���"
End Sub
Private Sub sys8_Click()
    DoCmd.Close acForm, "���_�����_�����"
    DoCmd.Close acForm, "�����"
    DoCmd.OpenForm "���_8���"
End Sub

Private Sub ������16_Click()
    DoCmd.Close acForm, "���_�����_�����"
    DoCmd.Close acForm, "�����"
    DoCmd.OpenForm "���_�����"
End Sub
