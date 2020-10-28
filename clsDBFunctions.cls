'---------------------------------------------------------------------------------------
' Module    : clsDBFunctions
' Author    : paweld
' Date      : 2019-12-02
' Purpose   : functions that require a database connection
' Licence   : Free
'
' Changes:
'----------------------------------------------------------------------------------------------------------------------------------
' 2020-10-21     , PD : beta wersion
'----------------------------------------------------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const CurrentModName = "clsDBFunctions"

Public Function DoesTheFieldAlreadyExist(TableName As String, FieldName As String) As Boolean


    Dim FieldExists As Byte
    Dim SqlS As String
    Dim Rst As Recordset

    SqlS = "SELECT CASE WHEN COL_LENGTH('" & NazwaTabeli & "', '" & NazwaPola & "') IS NOT NULL THEN 1 ELSE 0 END AS FieldExists"
    If OpenRst(Rst, SqlS) Then
        FieldExists = Rst!FieldExists
        Rst.Close
        Set Rst = Nothing
    End If
    
    If FieldExists = 1 Then
        DoesTheFieldAlreadyExist = True
    ElseIf FieldExists = 0 Then
        DoesTheFieldAlreadyExist = False
    End If
    
End Function
    
Public Function SetBlankCurrentConnection()
    CurrentProject.OpenConnection ("")
End Function
