VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Фор_Статус"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public EveryoneCanSeeMe As String
Public Oboz As String
Public Tip As String
Public Num As Byte
Public Kod As String
Public Namesys As String

Private Sub Form_Load()

    DoCmd.MoveSize 25000, 25000
     
End Sub

' https://dzen.ru/media/denchik_notes/novaia-ficha-ot-maikrosofta-risk-bezopasnosti-korporaciia-maikrosoft-zablokirovala-zapusk-makrosov-6310800a15d5545b90beb7ff
