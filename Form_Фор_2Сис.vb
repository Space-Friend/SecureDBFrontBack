VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Фор_2Сис"
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
    Call AuditEditEnd("Таб_Агрегаты", "audTmp_Таб_Агрегаты", "aud_Таб_Агрегаты", "Код", Nz(Me![Таб_Агрегаты.Код], 0), bWasNewRecord)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    bWasNewRecord = Me.NewRecord
    Call AuditEditBegin("Таб_Агрегаты", "audTmp_Таб_Агрегаты", "Код", Nz(Me![Таб_Агрегаты.Код], 0), bWasNewRecord)
End Sub

Private Sub Form_Delete(Cancel As Integer)
    Call AuditDelBegin("Таб_Агрегаты", "audTmp_Таб_Агрегаты", "Код", Nz(Me.Код, 0))
End Sub

Private Sub Form_Load()
        
    Form.AllowEdits = False
    Кнопка80.SetFocus
    If Forms![Фор_Статус].lll.Caption = "1" Then
        lbl1.Caption = "• Чтение - разрешено"
        lbl2.Caption = "• Изменение - запрещено"
        lbl3.Caption = "• Удаление - запрещено"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(186, 20, 25)
        Me.lbl3.ForeColor = RGB(186, 20, 25)
        kno1.Enabled = False
        Кнопка91.Enabled = False
        Кнопка3073.Enabled = False
        Кнопка102.Enabled = False
    End If
    If Forms![Фор_Статус].lll.Caption = "2" Then
        lbl1.Caption = "• Чтение - разрешено"
        lbl2.Caption = "• Изменение - разрешено"
        lbl3.Caption = "• Удаление - запрещено"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(63, 118, 42)
        Me.lbl3.ForeColor = RGB(186, 20, 25)
        kno1.Enabled = True
        Кнопка91.Enabled = True
        Кнопка3073.Enabled = True
        Кнопка102.Enabled = True
    End If
    If Forms![Фор_Статус].lll.Caption = "3" Then
        lbl1.Caption = "• Чтение - разрешено"
        lbl2.Caption = "• Изменение - разрешено"
        lbl3.Caption = "• Удаление - разрешено"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(63, 118, 42)
        Me.lbl3.ForeColor = RGB(63, 118, 42)
        kno1.Enabled = True
        Кнопка91.Enabled = True
        Кнопка3073.Enabled = True
        Кнопка102.Enabled = True
    End If

End Sub

Private Sub kno1_Click()
    Form.AllowEdits = True
End Sub

Private Sub Кнопка102_Click()
    
        Dim db As Database
        Dim rs As Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("Зап_2сис")
        rs.FindFirst "Ном=" & Me![Ном] & ""
        If rs![Связь] = 0 Then End
    If MsgBox("Вы собираетесь отправить агрегат на склад с позиции '" & rs![Ном] & "', всё верно?", vbYesNo, "Подтверждение") = vbYes Then
        rs.Edit
        rs![Место установки] = "На складе"
        rs.Update
        rs.Edit
        rs![Связь] = "0"
        rs.Update
        MsgBox ("Выполнено! Не забудьте поменять дату демонтажа агрегата.")
        DoCmd.RefreshRecord
        
    End If
  
End Sub

Public Sub Кнопка65_Click()
    
    Forms![Фор_Статус].EveryoneCanSeeMe = Forms![Фор_Статус].lll.Caption
    Forms![Фор_Статус].Namesys = "Система подачи дистиллированной воды 1.04.0360.085. Г2"
    Forms![Фор_Статус].Num = DLookup("[Ном]", "Зап_2сис", "[Ном] = " & Forms!Фор_2Сис!Ном)
    Forms![Фор_Статус].Tip = ""
    On Error GoTo ohsht
    Forms![Фор_Статус].Oboz = DLookup("[Обознач]", "Зап_2сис", "[Ном] = " & Forms!Фор_2Сис!Ном)
    Forms![Фор_Статус].EveryoneCanSeeMe = Forms![Фор_Статус].lll.Caption
    DoCmd.OpenForm "Фор_Склад_Окно", , , , acFormReadOnly, acDialog, "Зап_2сис"
    DoCmd.RefreshRecord
    
End

ohsht:
    If MsgBox("У агрегата отсутствует обозначение, чтобы выполнить поиск аналогов. Выполнить поиск по типу?", vbYesNo) = vbYes Then
        Forms![Фор_Статус].Tip = DLookup("[ТипАг]", "Зап_1сис", "[Ном] = " & Forms!Фор_2Сис!Ном)
        Forms![Фор_Статус].EveryoneCanSeeMe = Forms![Фор_Статус].lll.Caption
        DoCmd.OpenForm "Фор_Склад_Окно", , , , acFormReadOnly, acDialog, "Зап_2сис"
        DoCmd.RefreshRecord
    End If
End Sub

Private Sub Кнопка61_Click()

    Me.FilterOn = False
    Кнопка61.Caption = "Фильтр: НЕТ"
    Кнопка61.UseTheme = False
    Кнопка61.Enabled = False

End Sub

Private Sub Кнопка67_Click()
'
    Dim St As String
    St = InputBox("Пожалуйста, введите ЧАСТЬ или ПОЛНОЕ НАЗВАНИЕ ТИПА агрегата", "Тип агрегата", "Клапан")
    If St = "" Then Exit Sub
    
    Me.Filter = "[ТипАг] LIKE ""*" & St & "*"" "
    Me.FilterOn = True
    Кнопка61.Caption = "Фильтр СБРОС"
    Кнопка61.UseTheme = True
    Кнопка61.Enabled = True

End Sub

Private Sub Кнопка68_Click()
'
    Dim D As String
    
    D = InputBox("Пожалуйста, введите ЧАСТЬ или ПОЛНОЕ НАЗВАНИЕ ОБОЗНАЧЕНИЯ агрегата", "Обозначение", "Т 100")
    If D = "" Then Exit Sub
    
    Me.Filter = "[Обознач] LIKE ""*" & D & "*"" "
    Me.FilterOn = True
    Кнопка61.Caption = "Фильтр СБРОС"
    Кнопка61.UseTheme = True
    Кнопка61.Enabled = True
    
End Sub

Private Sub Кнопка73_Click()

    Me.FilterOn = False
    Кнопка61.Caption = "Фильтр: НЕТ"
    Кнопка61.UseTheme = False
    Кнопка61.Enabled = False
    
End Sub

Private Sub Кнопка74_Click()
    
On Error GoTo ohsht
    Forms![Фор_Статус].EveryoneCanSeeMe = Forms![Фор_Статус].lll.Caption
    DoCmd.OpenForm "Фор_Карточка_Агрегата", acNormal, , "[Код] =" & Me![Таб_Агрегаты.Код]
End

ohsht:
    If MsgBox("Нельзя открыть карточку неприкреплённого агрегата. Желаете выполнить поиск аналогов?", vbYesNo) = vbYes Then
        Call Кнопка65_Click
    End If
End

End Sub

Private Sub Кнопка80_Click()
    
    Forms![Фор_Статус].EveryoneCanSeeMe = Forms![Фор_Статус].lll.Caption
    DoCmd.Close acForm, Me.Name, acSavePrompt
    DoCmd.OpenForm "Старт"
    
End Sub

Private Sub Кнопка96_Click()
    If Кнопка96.Caption = "КРС /\" Then
        Me.OrderByOn = False
        Кнопка96.Caption = "КРС"
        Кнопка97.Caption = "ГРС"
        Кнопка98.Caption = "УДУ"
        кнопка_эксп_РАЗ.Caption = "Место установки"
        кнопка_ТипОБ.Caption = "Тип агрегата"
        кнопка_наим_оборуд.Caption = "Обозначение"
        кнопка_зав_номер.Caption = "Заводской номер"
        Кнопка49.Caption = "Подпись"
        кнопка_инв_номер.Caption = "Дата монтажа"
        Кнопка108.Caption = "Сраб."
        Кнопка21.Caption = "ПД"
        Кнопка39.Caption = "Дата демонтажа"
    Else
        Me.OrderBy = "Таб_1Сист.КРС DESC"
        Me.OrderByOn = True
        Кнопка98.Caption = "УДУ"
        Кнопка96.Caption = "КРС /\"
        Кнопка97.Caption = "ГРС"
        кнопка_эксп_РАЗ.Caption = "Место установки"
        кнопка_ТипОБ.Caption = "Тип агрегата"
        кнопка_наим_оборуд.Caption = "Обозначение"
        кнопка_зав_номер.Caption = "Заводской номер"
        Кнопка49.Caption = "Подпись"
        кнопка_инв_номер.Caption = "Дата монтажа"
        Кнопка108.Caption = "Сраб."
        Кнопка21.Caption = "ПД"
        Кнопка39.Caption = "Дата демонтажа"
    End If
End Sub

Private Sub Кнопка97_Click()
    If Кнопка97.Caption = "ГРС /\" Then
        Me.OrderByOn = False
        Кнопка96.Caption = "КРС"
        Кнопка97.Caption = "ГРС"
        Кнопка98.Caption = "УДУ"
        кнопка_эксп_РАЗ.Caption = "Место установки"
        кнопка_ТипОБ.Caption = "Тип агрегата"
        кнопка_наим_оборуд.Caption = "Обозначение"
        кнопка_зав_номер.Caption = "Заводской номер"
        Кнопка49.Caption = "Подпись"
        кнопка_инв_номер.Caption = "Дата монтажа"
        Кнопка108.Caption = "Сраб."
        Кнопка21.Caption = "ПД"
        Кнопка39.Caption = "Дата демонтажа"
    Else
        Me.OrderBy = "Таб_1Сист.ГРС DESC"
        Me.OrderByOn = True
        Кнопка98.Caption = "УДУ"
        Кнопка96.Caption = "КРС"
        Кнопка97.Caption = "ГРС /\"
        кнопка_эксп_РАЗ.Caption = "Место установки"
        кнопка_ТипОБ.Caption = "Тип агрегата"
        кнопка_наим_оборуд.Caption = "Обозначение"
        кнопка_зав_номер.Caption = "Заводской номер"
        Кнопка49.Caption = "Подпись"
        кнопка_инв_номер.Caption = "Дата монтажа"
        Кнопка108.Caption = "Сраб."
        Кнопка21.Caption = "ПД"
        Кнопка39.Caption = "Дата демонтажа"
    End If
End Sub

Private Sub Кнопка98_Click()
    If Кнопка98.Caption = "УДУ /\" Then
        Me.OrderByOn = False
        Кнопка96.Caption = "КРС"
        Кнопка97.Caption = "ГРС"
        Кнопка98.Caption = "УДУ"
        кнопка_эксп_РАЗ.Caption = "Место установки"
        кнопка_ТипОБ.Caption = "Тип агрегата"
        кнопка_наим_оборуд.Caption = "Обозначение"
        кнопка_зав_номер.Caption = "Заводской номер"
        Кнопка49.Caption = "Подпись"
        кнопка_инв_номер.Caption = "Дата монтажа"
        Кнопка108.Caption = "Сраб."
        Кнопка21.Caption = "ПД"
        Кнопка39.Caption = "Дата демонтажа"
    Else
        Me.OrderBy = "Таб_1Сист.УДУ DESC"
        Me.OrderByOn = True
        Кнопка98.Caption = "УДУ /\"
        Кнопка96.Caption = "КРС"
        Кнопка97.Caption = "ГРС"
        кнопка_эксп_РАЗ.Caption = "Место установки"
        кнопка_ТипОБ.Caption = "Тип агрегата"
        кнопка_наим_оборуд.Caption = "Обозначение"
        кнопка_зав_номер.Caption = "Заводской номер"
        Кнопка49.Caption = "Подпись"
        кнопка_инв_номер.Caption = "Дата монтажа"
        Кнопка108.Caption = "Сраб."
        Кнопка21.Caption = "ПД"
        Кнопка39.Caption = "Дата демонтажа"
    End If
End Sub

