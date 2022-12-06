VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Фор_Склад_Окно"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    
    If Forms![Фор_Статус].EveryoneCanSeeMe = 2 Then
        Кнопка51.Enabled = True
    End If
    If Forms![Фор_Статус].EveryoneCanSeeMe = 3 Then
        Кнопка51.Enabled = True
    End If
    
    If Forms![Фор_Статус].Tip = "" Then
        lbl1.Caption = Forms![Фор_Статус].Oboz
        Me.Filter = "Обозначение LIKE ""*" & Forms![Фор_Статус].Oboz & "*"" "
        Кнопка61.UseTheme = True
        Кнопка61.Enabled = True
        Me.FilterOn = True
    Else
        lbl1.Caption = Forms![Фор_Статус].Tip
        Me.Filter = "[Тип агрегата] LIKE ""*" & Forms![Фор_Статус].Tip & "*"" "
        Кнопка61.UseTheme = True
        Кнопка61.Enabled = True
        Me.FilterOn = True
    End If
End Sub

Private Sub Кнопка61_Click()
    Me.FilterOn = False
    Кнопка61.Caption = "Фильтр: НЕТ"
    Кнопка61.UseTheme = False
    Кнопка61.Enabled = False
    lbl1.Caption = "Весь склад."
End Sub

Private Sub Кнопка51_Click()
    
    If MsgBox("Вы собираетесь назначить агрегат номер '" & Me![Номер] & "' на позицию '" & Forms![Фор_Статус].Num & "', всё верно?", vbYesNo, "Подтверждение") = vbYes Then
        Forms![Фор_Статус].Kod = Me![Код]
        Dim db As Database
        Dim rs As Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset(OpenArgs)
        rs.FindFirst "Ном=" & Forms![Фор_Статус].Num & ""
        
        If rs![Связь] = "0" Then
            rs.Edit
            rs![Связь] = Me![Код]
            rs![Место установки] = Forms![Фор_Статус].Namesys
            rs.Update
            DoCmd.Close acForm, "Фор_склад_окно"
        Else
            MsgBox "Кажется, на этом месте уже стоит агрегат. Сначала снимите тот.", , "Занято!"
        End If
        
    End If
  
End Sub

Private Sub Кнопка80_Click()
    Forms![Фор_Статус].Tip = ""
    Forms![Фор_Статус].Oboz = ""
    DoCmd.Close acForm, Me.Name
End Sub
