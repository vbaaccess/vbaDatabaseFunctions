Option Compare Database
Option Explicit

Private Const CurrentModeName = "modInitDB"

Public db As New clsDBFunctions          ' object need db connection

' Functions:
'
' db.DoesTheFieldAlreadyExist(TableName As String, FieldName As String) As Boolean
' db.GetValue(strSQL As String, Optional FieldIndex As Single = 0) As Variant
' db.SetBlankCurrentConnection() As Variant
