VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Фор_Список_Агрегатов"
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

Private Sub Form_Load()

    If Forms![Фор_Статус].EveryoneCanSeeMe = "1" Then
        Me.AllowDeletions = False
        lbl1.Caption = "• Чтение - разрешено"
        lbl2.Caption = "• Изменение - запрещено"
        lbl3.Caption = "• Удаление - запрещено"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(186, 20, 25)
        Me.lbl3.ForeColor = RGB(186, 20, 25)
        kno1.Enabled = False
        Кнопка91.Enabled = False
        Кнопка3073.Enabled = False
        Кнопка94.Enabled = False
    End If
    If Forms![Фор_Статус].EveryoneCanSeeMe = "2" Then
        Me.AllowDeletions = False
        lbl1.Caption = "• Чтение - разрешено"
        lbl2.Caption = "• Изменение - разрешено"
        lbl3.Caption = "• Удаление - запрещено"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(63, 118, 42)
        Me.lbl3.ForeColor = RGB(186, 20, 25)
        kno1.Enabled = True
        Кнопка91.Enabled = True
        Кнопка3073.Enabled = True
        Кнопка94.Enabled = False
    End If
    If Forms![Фор_Статус].EveryoneCanSeeMe = "3" Then
        Me.AllowDeletions = True
        lbl1.Caption = "• Чтение - разрешено"
        lbl2.Caption = "• Изменение - разрешено"
        lbl3.Caption = "• Удаление - разрешено"
        Me.lbl1.ForeColor = RGB(63, 118, 42)
        Me.lbl2.ForeColor = RGB(63, 118, 42)
        Me.lbl3.ForeColor = RGB(63, 118, 42)
        kno1.Enabled = True
        Кнопка91.Enabled = True
        Кнопка3073.Enabled = True
        Кнопка94.Enabled = True
    End If
    
End Sub

Private Sub kno1_Click()
    Form.AllowEdits = True
End Sub

Private Sub Кнопка60_Click()

    Dim s As String
    
    s = InputBox("Пожалуйста, введите ЧАСТЬ или ПОЛНОЕ НАЗВАНИЕ МЕСТА УСТАНОВКИ агрегата", "Место установки", "Система подачи газообразного азота")
    If s = "" Then Exit Sub

    Me.Filter = "[Место установки] LIKE ""*" & s & "*"" "
    Me.FilterOn = True
    Кнопка61.Caption = "Фильтр СБРОС"
    Кнопка61.UseTheme = True
    Кнопка61.Enabled = True
    
End Sub

Private Sub Кнопка61_Click()

    Me.FilterOn = False
    Кнопка61.Caption = "Фильтр: НЕТ"
    Кнопка61.UseTheme = False
    Кнопка61.Enabled = False

End Sub

Private Sub Кнопка67_Click()

    Dim St As String
    St = InputBox("Пожалуйста, введите ЧАСТЬ или ПОЛНОЕ НАЗВАНИЕ ТИПА агрегата", "Тип агрегата", "Клапан")
    If St = "" Then Exit Sub
    
    Me.Filter = "[Тип агрегата] LIKE ""*" & St & "*"" "
    Me.FilterOn = True
    Кнопка61.Caption = "Фильтр СБРОС"
    Кнопка61.UseTheme = True
    Кнопка61.Enabled = True

End Sub

Private Sub Кнопка68_Click()

    Dim D As String
    
    D = InputBox("Пожалуйста, введите ЧАСТЬ или ПОЛНОЕ НАЗВАНИЕ ОБОЗНАЧЕНИЯ агрегата", "Обозначение", "Т 100")
    If D = "" Then Exit Sub
    
    Me.Filter = "[Обозначение] LIKE ""*" & D & "*"" "
    Me.FilterOn = True
    Кнопка61.Caption = "Фильтр СБРОС"
    Кнопка61.UseTheme = True
    Кнопка61.Enabled = True
    
End Sub

Private Sub Кнопка69_Click()

    Dim D As String
    
    D = InputBox("Пожалуйста, введите ЧАСТЬ или ПОЛНОЕ НАЗВАНИЕ НОМЕРА агрегата", "Номер")
    If D = "" Then Exit Sub
    
    Me.Filter = "[Номер] LIKE ""*" & D & "*"" "
    Me.FilterOn = True
    Кнопка61.Caption = "Фильтр СБРОС"
    Кнопка61.UseTheme = True
    Кнопка61.Enabled = True
    
End Sub

Private Sub Кнопка70_Click()

    Dim D As String
    
    D = InputBox("Пожалуйста, введите ГРС агрегата", "ГРС", "001")
    If D = "" Then Exit Sub
    
    Me.Filter = "[ГРС] LIKE ""*" & D & "*"" "
    Me.FilterOn = True
    Кнопка61.Caption = "Фильтр СБРОС"
    Кнопка61.UseTheme = True
    Кнопка61.Enabled = True
    
End Sub

Private Sub Кнопка71_Click()

    Dim D As String
    
    D = InputBox("Пожалуйста, введите УДУ агрегата", "УДУ", "10")
    If D = "" Then Exit Sub
    
    Me.Filter = "[УДУ] LIKE " & D & " "
    Me.FilterOn = True
    Кнопка61.Caption = "Фильтр СБРОС"
    Кнопка61.UseTheme = True
    Кнопка61.Enabled = True
    
End Sub

Private Sub Кнопка72_Click()

    Dim D As String
    
    D = InputBox("Пожалуйста, введите КРС агрегата", "КРС", "0.9")
    If D = "" Then Exit Sub
    
    Me.Filter = "[КРС] LIKE ""*" & D & "*"" "
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
    
    Forms![Фор_Статус].EveryoneCanSeeMe = Forms![Фор_Статус].lll.Caption
    DoCmd.OpenForm "Фор_Карточка_Агрегата", , , "[Код] =" & Me![Код], acFormReadOnly

End Sub
