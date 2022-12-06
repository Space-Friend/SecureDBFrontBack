VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Старт"
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
    Forms![Фор_Статус].EveryoneCanSeeMe = Forms![Фор_Статус].lll.Caption
    If Forms![Фор_Статус].EveryoneCanSeeMe = "1" Then
        lbltest.Caption = "Базовые права"
        Кнопка694.Enabled = True
        Kart.Enabled = False
        sotr.Enabled = False
        log.Enabled = False
        logkart.Enabled = False
        shft.Enabled = False
        
    End If
    If Forms![Фор_Статус].EveryoneCanSeeMe = "2" Then
        lbltest.Caption = "Расширенные права"
        Кнопка694.Enabled = True
        Kart.Enabled = True
        sotr.Enabled = False
        log.Enabled = False
        logkart.Enabled = False
        shft.Enabled = False
        
    End If
    If Forms![Фор_Статус].EveryoneCanSeeMe = "3" Then
        lbltest.Caption = "Администраторские права"
        Кнопка694.Enabled = True
        Kart.Enabled = True
        sotr.Enabled = True
        log.Enabled = True
        logkart.Enabled = True
        shft.Enabled = True
        
    End If
    
    On Error Resume Next
    Кнопка694.SetFocus
    
End Sub

Private Sub log_Click()
    
    DoCmd.OpenForm "Фор_Лог_Вход"
    DoCmd.Close acForm, Me.Name
    
End Sub

Private Sub shft_Click()
    
    Dim prop As Property
    On Error GoTo Setproperty
    Set prop = CurrentDb.CreateProperty("allowbypasskey", dbBoolean, False)
    
    CurrentDb.Properties.Append prop
    
Setproperty:
    If MsgBox("Вы хотите включить клавишу допуска (Shift)?  В таком случае во время загрузки базы данных при удерживании этой клавиши БД откроется в режиме редактирования со стандартными настройками.", vbYesNo, "Включить клавишу допуска") = vbYes Then
        CurrentDb.Properties("allowbypasskey") = True
    Else
        CurrentDb.Properties("allowbypasskey") = False
    End If
    
End Sub

Private Sub logkart_Click()
    
    DoCmd.OpenForm "Фор_Логи_Карточек"
    DoCmd.Close acForm, Me.Name
    
End Sub

Private Sub sotr_Click()
    
        If Forms![Фор_Статус].EveryoneCanSeeMe = "3" Then
        DoCmd.OpenForm "Фор_Сотрудники_Админ", , , , , acDialog
    Else: DoCmd.OpenForm "Фор_Сотрудники", , , , , acDialog
    End If
    
End Sub


Private Sub Кнопка100_Click()
    MsgBox "Разработка отдела 774-5. Версия программы от 30.11.2022 To space with Space!", vbInformation, "Версия"
End Sub

Private Sub Кнопка33_Click()
    DoCmd.OpenForm "Фор_Логин"
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Кнопка663_Click()
    DoCmd.Close acForm, Me.Name
    DoCmd.CloseDatabase
End Sub

Private Sub Кнопка77_Click()

End Sub
