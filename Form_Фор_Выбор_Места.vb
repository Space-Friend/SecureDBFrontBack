VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Фор_Выбор_Места"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    Кнопка80.SetFocus
End Sub

Private Sub sys1_Click()
    DoCmd.Close acForm, "Фор_Выбор_Места"
    DoCmd.Close acForm, "Старт"
    DoCmd.OpenForm "Фор_1Сис"
End Sub

Private Sub sys2_Click()
    DoCmd.Close acForm, "Фор_Выбор_Места"
    DoCmd.Close acForm, "Старт"
    DoCmd.OpenForm "Фор_2Сис"
End Sub
Private Sub sys3_Click()
    DoCmd.Close acForm, "Фор_Выбор_Места"
    DoCmd.Close acForm, "Старт"
    DoCmd.OpenForm "Фор_3Сис"
End Sub
Private Sub sys4_Click()
    DoCmd.Close acForm, "Фор_Выбор_Места"
    DoCmd.Close acForm, "Старт"
    DoCmd.OpenForm "Фор_4Сис"
End Sub
Private Sub sys5_Click()
    DoCmd.Close acForm, "Фор_Выбор_Места"
    DoCmd.Close acForm, "Старт"
    DoCmd.OpenForm "Фор_5Сис"
End Sub
Private Sub sys6_Click()
    DoCmd.Close acForm, "Фор_Выбор_Места"
    DoCmd.Close acForm, "Старт"
    DoCmd.OpenForm "Фор_6Сис"
End Sub
Private Sub sys7_Click()
    DoCmd.Close acForm, "Фор_Выбор_Места"
    DoCmd.Close acForm, "Старт"
    DoCmd.OpenForm "Фор_7Сис"
End Sub
Private Sub sys8_Click()
    DoCmd.Close acForm, "Фор_Выбор_Места"
    DoCmd.Close acForm, "Старт"
    DoCmd.OpenForm "Фор_8Сис"
End Sub

Private Sub Кнопка16_Click()
    DoCmd.Close acForm, "Фор_Выбор_Места"
    DoCmd.Close acForm, "Старт"
    DoCmd.OpenForm "Фор_склад"
End Sub
