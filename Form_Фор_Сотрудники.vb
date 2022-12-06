VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Фор_Сотрудники"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    Кнопка73.UseTheme = False
    Кнопка73.Enabled = False
End Sub

Private Sub Кнопка16_Click()

    Кнопка73.Enabled = True
    Кнопка73.UseTheme = True
    Dim St As String
    St = InputBox("Введите имя", "Поиск по имени", "Имя")
    If St = "" Then Exit Sub
    
    Me.Filter = "Имя LIKE ""*" & St & "*"""
    Me.FilterOn = True

End Sub

Private Sub Кнопка17_Click()

    Кнопка73.Enabled = True
    Кнопка73.UseTheme = True
    Dim St As String
    St = InputBox("Введите Фамилию", "Поиск по фамилии", "Фамилия")
    If St = "" Then Exit Sub
    
    Me.Filter = "Фамилия LIKE ""*" & St & "*"""
    Me.FilterOn = True

End Sub

Private Sub Кнопка18_Click()

    Кнопка73.Enabled = True
    Кнопка73.UseTheme = True
    Dim St As String
    St = InputBox("Введите отчество", "Поиск по отчеству", "отчество")
    If St = "" Then Exit Sub
    
    Me.Filter = "Отчество LIKE ""*" & St & "*"""
    Me.FilterOn = True

End Sub

Private Sub Кнопка73_Click()

    Me.FilterOn = False
    Кнопка73.Enabled = False
    Кнопка73.UseTheme = False
    
End Sub
