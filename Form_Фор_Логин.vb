VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Фор_Логин"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
    DoCmd.OpenForm "Фор_Статус", , , , acFormEdit, acHidden
End Sub

Private Sub Кнопка11_Click()

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("Таб_Сотрудники", dbOpenSnapshot, dbReadOnly)
    rs.FindFirst "Логин='" & Me.log & "'"
    
    If rs.NoMatch = True Then
        Me.Надпись7.Visible = True
        Me.Надпись8.Visible = False
        Me.log.SetFocus
        Exit Sub
    End If
    Me.Надпись7.Visible = False

    If rs!Пароль = Me.par Then
        
        Dim a
        Dim StrSQL As String
        a = rs![Код сотрудника]
        StrSQL = "INSERT INTO [Таб_Лог_Вход] ([Код сотрудника]) VALUES ('" & a & "' );"
        DoCmd.SetWarnings False
        DoCmd.RunSQL StrSQL
        
        If rs!Уровень = 1 Then
            Forms![Фор_Статус].EveryoneCanSeeMe = "1"
        End If
        If rs!Уровень = 2 Then
            Forms![Фор_Статус].EveryoneCanSeeMe = "2"
        End If
        If rs!Уровень = 3 Then
            Forms![Фор_Статус].EveryoneCanSeeMe = "3"
        End If
        
        Forms![Фор_Статус].lll.Caption = Forms![Фор_Статус].EveryoneCanSeeMe
        DoCmd.OpenForm "Старт"
        DoCmd.Close acForm, "Фор_Логин"
    Else
        Me.Надпись8.Visible = True
        Me.par.SetFocus
    End If
    
End Sub
