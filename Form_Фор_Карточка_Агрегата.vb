VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Фор_Карточка_Агрегата"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bWasNewRecord As Boolean

Private Sub Form_AfterDelConfirm(Status As Integer)
    Call AuditDelEnd("audTmp_Таб_Агрегаты", "aud_Таб_Агрегаты", Status)
End Sub

Private Sub Form_AfterUpdate()
    Call AuditEditEnd("Таб_Агрегаты", "audTmp_Таб_Агрегаты", "aud_Таб_Агрегаты", "Код", Nz(Me!Код, 0), bWasNewRecord)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    bWasNewRecord = Me.NewRecord
    Call AuditEditBegin("Таб_Агрегаты", "audTmp_Таб_Агрегаты", "Код", Nz(Me.Код, 0), bWasNewRecord)
End Sub

Private Sub Form_Delete(Cancel As Integer)
    Call AuditDelBegin("Таб_Агрегаты", "audTmp_Таб_Агрегаты", "Код", Nz(Me.Код, 0))
End Sub

Private Sub Кнопка2743_Click()
    DoCmd.Close
End Sub

Private Sub Form_Load()

    If Forms![Фор_Статус].EveryoneCanSeeMe = "1" Then
        lbl1.Caption = "Базовые права"
        Кнопка260.Enabled = False
        kno1.Enabled = False
        kno26.Enabled = False
        Кнопка256.Enabled = False
        Кнопка200.Enabled = False
    End If
    If Forms![Фор_Статус].EveryoneCanSeeMe = "2" Then
        lbl1.Caption = "Расширенные права"
        Кнопка260.Enabled = True
        kno1.Enabled = True
        kno26.Enabled = True
        Кнопка256.Enabled = True
        Кнопка200.Enabled = False
    End If
    If Forms![Фор_Статус].EveryoneCanSeeMe = "3" Then
        lbl1.Caption = "Администраторские права"
        Кнопка260.Enabled = True
        kno1.Enabled = True
        kno26.Enabled = True
        Кнопка256.Enabled = True
        Кнопка200.Enabled = True
    End If
    
End Sub

Private Sub kno1_Click()
    Form.AllowEdits = True
End Sub

Private Sub Кнопка239_Click()
    
    If Forms![Фор_Статус].EveryoneCanSeeMe = "3" Then
        DoCmd.OpenForm "Фор_Сотрудники_Админ", , , , , acDialog
    Else: DoCmd.OpenForm "Фор_Сотрудники", , , , , acDialog
    End If
    
End Sub

Private Sub Кнопка308_Click()
    
Setproperty:
    If MsgBox("Снять ответственного за монтаж? (необратимое действие)", vbYesNo, "Снять ответственного") = vbYes Then
        Отв_монтаж.Value = ""
        DoCmd.RefreshRecord
    End If
    
End Sub

Private Sub Кнопка309_Click()
    
Setproperty:
    If MsgBox("Снять ответственного за ДЕмонтаж? (необратимое действие)", vbYesNo, "Снять ответственного") = vbYes Then
        Отв_Демонтаж.Value = ""
        DoCmd.RefreshRecord
    End If
    
End Sub

Private Sub Кнопка324_Click()
    
    On Error GoTo sht
    DoCmd.OpenForm "Фор_Логи_Карточек_Контекст", acNormal, , "[Код] =" & Me![Код], , acDialog
End

sht:
    MsgBox "Ошибка отображения истории. Проверьте целостность карточки."

End Sub
