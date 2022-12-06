Option Compare Database

' Author: Allen Browne, allen@allenbrowne.com, 2006.
' Implemented by: Space Friend, Spacefrnd@mail.ru, 2022.

' Purpose: Audit trail, to track Deletes, Edits, and Inserts.
' Does not audit any Cascading Updates/Deletes.

' Requirements: The table to be audited must have an AutoNumber primary key.
' Data entry must be through a form.

' Method: Makes a copy of the record in a temp table, and logs the
' change when it is guaranteed. The temp table copes with:
' - multiple deletes at once (continuous/datasheet view)
' - cancelled deletes or failed updates.
' - requirement for sequential numbering in the audit table.
' On a multi-user split (front-end/back-end) database, the
' temp table may reside in the front end, and the audit log
' in the back-end.

' Result: The audit table will contain one record for each deletion or
' insertion, and two records for each edit (before and after).
' Delete Copy of the deleted record, marked "Delete".
' Insert Copy of the new record, marked "Insert".
' Change: Copy of the record before change, marked "EditFrom".
' Copy of the record after change, marked "EditTo".
' This approach, together with the sequential numbering of the
' AutoNumber in the audit table makes tampering with the audit
' log more detectable.

'Note: Record confirmations need to be on. When opening the database:
' If Not Application.GetOption("Confirm Record Changes") Then
' Application.SetOption ("Confirm Record Changes"), True
' End If

Option Explicit

Private Const conMod As String = "ajbAudit"
Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias _
 "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Function NetworkUserName() As String
On Error GoTo Err_Handler
 'Purpose: Returns the network login name
 Dim lngLen As Long
 Dim lngX As Long
 Dim strUserName As String

 NetworkUserName = "Unknown"

 strUserName = String$(254, 0)
 lngLen = 255&
 lngX = apiGetUserName(strUserName, lngLen)
 If (lngX > 0&) Then
 NetworkUserName = Left$(strUserName, lngLen - 1&)
 End If

Exit_Handler:
 Exit Function

Err_Handler:
 Call LogError(Err.Number, Err.Description, conMod & ".NetworkUserName", , False)
 Resume Exit_Handler
End Function


Function AuditDelBegin(sTable As String, sAudTmpTable As String, sKeyField As String, lngKeyValue As Long) As Boolean
On Error GoTo Err_AuditDelBegin
 'Purpose: Write a copy of the record to a tmp audit table.
 ' Copy to be written to real audit table in AfterDelConfirm.
 'Arguments: sTable = name of table to be audited.
 ' sAudTmpTable = the name of the temp audit table.
 ' sKeyField = name of AutoNumber field in table.
 ' lngKeyValue = number in the AutoNumber field.
 'Return: True if successful.
 'Usage: Call from a form's Delete event. Example:
 ' Call AuditDelBegin("tblInvoice", "audTmpInvoice", "InvoiceID", Me.InvoiceID)
 'Note: Must also call AuditDelEnd in the form's AfterDelConfirm event.
 Dim db As DAO.Database ' Current database
 Dim sSQL As String ' Append query.

 ' Append record to the temp audit table.
 Set db = DBEngine(0)(0)
 sSQL = "INSERT INTO " & sAudTmpTable & " ( audType, audDate, audUser ) " & _
 "SELECT 'Delete' AS Expr1, Now() AS Expr2, NetworkUserName() AS Expr3, " & sTable & ".* " & _
 "FROM " & sTable & " WHERE (" & sTable & "." & sKeyField & " = " & lngKeyValue & ");"
 db.Execute sSQL, dbFailOnError

Exit_AuditDelBegin:
 Set db = Nothing
 Exit Function

Err_AuditDelBegin:
 Call LogError(Err.Number, Err.Description, conMod & ".AuditDelBegin()", , False)
 Resume Exit_AuditDelBegin
End Function


Function AuditDelEnd(sAudTmpTable As String, sAudTable As String, Status As Integer) As Boolean
On Error GoTo Err_AuditDelEnd
 'Purpose: If the deletion was completed, copy the data from the
 ' temp table to the autit table. Empty temp table.
 'Arguments: sAudTmpTable = name of temp audit table
 ' sAudTable = name of audit table
 ' Status = Status from the form's AfterDelConfirm event.
 'Return: True if successful.
 'Usage: Call from form's AfterDelConfirm event. Example:
 ' Call AuditDelEnd("audTmpInvoice", "audInvoice", Status)
 Dim db As DAO.Database ' Currrent database
 Dim sSQL As String ' Append query.

 ' If the Delete proceeded, copy the record(s) from temp table to delete table.
 ' Note: Only "Delete" types are copied: cancelled Edits may be there as well
  Set db = DBEngine(0)(0)
 If Status = acDeleteOK Then
 sSQL = "INSERT INTO " & sAudTable & " SELECT " & sAudTmpTable & ".* FROM " & sAudTmpTable & _
 " WHERE (" & sAudTmpTable & ".audType = 'Delete');"
 db.Execute sSQL, dbFailOnError
 End If

 'Remove the temp record(s)
  sSQL = "DELETE FROM " & sAudTmpTable & ";"
 db.Execute sSQL, dbFailOnError
 AuditDelEnd = True

Exit_AuditDelEnd:
 Set db = Nothing
 Exit Function

Err_AuditDelEnd:
 Call LogError(Err.Number, Err.Description, conMod & ".AuditDelEnd()", False)
 Resume Exit_AuditDelEnd
End Function


Function AuditEditBegin(sTable As String, sAudTmpTable As String, sKeyField As String, _
 lngKeyValue As Long, bWasNewRecord As Boolean) As Boolean
On Error GoTo Err_AuditEditBegin
 'Purpose: Write a copy of the old values to temp table.
 ' It is then copied to the true audit table in AuditEditEnd.
 'Arugments: sTable = name of table being audited.
 ' sAudTmpTable = name of the temp audit table.
 ' sKeyField = name of the AutoNumber field.
 ' lngKeyValue = Value of the AutoNumber field.
 ' bWasNewRecord = True if this was a new insert.
 'Return: True if successful
 'Usage: Called in form's BeforeUpdate event. Example:
 ' bWasNewRecord = Me.NewRecord
 ' Call AuditEditBegin("tblInvoice", "audTmpInvoice", "InvoiceID", Me.InvoiceID, bWasNewRecord)
 Dim db As DAO.Database ' Current database
 Dim sSQL As String

 'Remove any cancelled update still in the tmp table.
 Set db = DBEngine(0)(0)
 sSQL = "DELETE FROM " & sAudTmpTable & ";"
 db.Execute sSQL

 ' If this was not a new record, save the old values.
 If Not bWasNewRecord Then
 sSQL = "INSERT INTO " & sAudTmpTable & " ( audType, audDate, audUser ) " & _
 "SELECT 'EditFrom' AS Expr1, Now() AS Expr2, NetworkUserName() AS Expr3, " & sTable & ".* " & _
 "FROM " & sTable & " WHERE (" & sTable & "." & sKeyField & " = " & lngKeyValue & ");"
 db.Execute sSQL, dbFailOnError
 End If
 AuditEditBegin = True

Exit_AuditEditBegin:
 Set db = Nothing
 Exit Function

Err_AuditEditBegin:
 Call LogError(Err.Number, Err.Description, conMod & ".AuditEditBegin()", , False)
 Resume Exit_AuditEditBegin
End Function


Function AuditEditEnd(sTable As String, sAudTmpTable As String, sAudTable As String, _
 sKeyField As String, lngKeyValue As Long, bWasNewRecord As Boolean) As Boolean
On Error GoTo Err_AuditEditEnd
 'Purpose: Write the audit trail to the audit table.
 'Arguments: sTable = name of table being audited.
 ' sAudTmpTable = name of the temp audit table.
 ' sAudTable = name of the audit table.
 ' sKeyField = name of the AutoNumber field.
 ' lngKeyValue = Value of the AutoNumber field.
 ' bWasNewRecord = True if this was a new insert.
 'Return: True if successful
 'Usage: Called in form's AfterUpdate event. Example:
 ' Call AuditEditEnd("tblInvoice", "audTmpInvoice", "audInvoice", "InvoiceID", Me.InvoiceID, bWasNewRecord)
 Dim db As DAO.Database
 Dim sSQL As String
 Set db = DBEngine(0)(0)

 If bWasNewRecord Then
 ' Copy the new values as "Insert".
 sSQL = "INSERT INTO " & sAudTable & " ( audType, audDate, audUser ) " & _
 "SELECT 'Insert' AS Expr1, Now() AS Expr2, NetworkUserName() AS Expr3, " & sTable & ".* " & _
 "FROM " & sTable & " WHERE (" & sTable & "." & sKeyField & " = " & lngKeyValue & ");"
 db.Execute sSQL, dbFailOnError
 Else
 ' Copy the latest edit from temp table as "EditFrom".
 sSQL = "INSERT INTO " & sAudTable & " SELECT TOP 1 " & sAudTmpTable & ".* FROM " & sAudTmpTable & _
 " WHERE (" & sAudTmpTable & ".audType = 'EditFrom') ORDER BY " & sAudTmpTable & ".audDate DESC;"
 db.Execute sSQL
 ' Copy the new values as "EditTo"
 sSQL = "INSERT INTO " & sAudTable & " ( audType, audDate, audUser ) " & _
 "SELECT 'EditTo' AS Expr1, Now() AS Expr2, NetworkUserName() AS Expr3, " & sTable & ".* " & _
 "FROM " & sTable & " WHERE (" & sTable & "." & sKeyField & " = " & lngKeyValue & ");"
 db.Execute sSQL
 ' Empty the temp table.
 sSQL = "DELETE FROM " & sAudTmpTable & ";"
 db.Execute sSQL, dbFailOnError
 End If
 AuditEditEnd = True

Exit_AuditEditEnd:
 Set db = Nothing
 Exit Function

Err_AuditEditEnd:
 Call LogError(Err.Number, Err.Description, conMod & ".AuditEditEnd()", , False)
 Resume Exit_AuditEditEnd
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function LogError(ByVal lngErrNumber As Long, ByVal strErrDescription As String, _
 strCallingProc As String, Optional vParameters, Optional bShowUser As Boolean = True) As Boolean
On Error GoTo Err_LogError
 ' Purpose: Generic error handler.
 ' Logs errors to table "tLogError".
 ' Arguments: lngErrNumber - value of Err.Number
 ' strErrDescription - value of Err.Description
 ' strCallingProc - name of sub|function that generated the error.
 ' vParameters - optional string: List of parameters to record.
 ' bShowUser - optional boolean: If False, suppresses display.
 ' Author: Allen Browne, allen@allenbrowne.com

 Dim strMsg As String ' String for display in MsgBox
 Dim rst As DAO.Recordset ' The tLogError table

 Select Case lngErrNumber
 Case 0
 Debug.Print strCallingProc & " called error 0."
 Case 2501 ' Cancelled
 'Do nothing.
 Case 3314, 2101, 2115 ' Can't save.
 If bShowUser Then
 strMsg = "Record cannot be saved at this time." & vbCrLf & _
 "Complete the entry, or press <Esc> to undo."
 MsgBox strMsg, vbExclamation, strCallingProc
 End If
 Case Else
 If bShowUser Then
 strMsg = "Error " & lngErrNumber & ": " & strErrDescription
 MsgBox strMsg, vbExclamation, strCallingProc
 End If
 Set rst = CurrentDb.OpenRecordset("tLogError", , dbAppendOnly)
 rst.AddNew
 rst![ErrNumber] = lngErrNumber
 rst![ErrDescription] = Left$(strErrDescription, 255)
 rst![ErrDate] = Now()
 rst![CallingProc] = strCallingProc
 rst![UserName] = CurrentUser()
 rst![ShowUser] = bShowUser
 If Not IsMissing(vParameters) Then
 rst![Parameters] = Left(vParameters, 255)
 End If
 rst.Update
 rst.Close
 LogError = True
 End Select

Exit_LogError:
 Set rst = Nothing
 Exit Function

Err_LogError:
 strMsg = "An unexpected situation arose in your program." & vbCrLf & _
 "Please write down the following details:" & vbCrLf & vbCrLf & _
 "Calling Proc: " & strCallingProc & vbCrLf & _
 "Error Number " & lngErrNumber & vbCrLf & strErrDescription & vbCrLf & vbCrLf & _
 "Unable to record because Error " & Err.Number & vbCrLf & Err.Description
 MsgBox strMsg, vbCritical, "LogError()"
 Resume Exit_LogError
End Function
