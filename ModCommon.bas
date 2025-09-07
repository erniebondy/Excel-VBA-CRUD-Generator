Attribute VB_Name = "ModCommon"
Option Explicit

Dim ADOConnection_ As ADODB.Connection

' It would be nice to automate the modification the VB_PredeclaredId of a class module

'Public Property Get ADOConnection() As ADODB.Connection
'  Set ADOConnection = ADOConnection_
'End Property

Function NextID(TableName) As Integer
  Dim cmd As ADODB.Command: Set cmd = ModCommon.CommandNew
  cmd.CommandText = "SELECT MAX(id) FROM [" & TableName & "$]"
  Dim RS As ADODB.Recordset: Set RS = cmd.Execute
  
  NextID = 1
  
  On Error Resume Next
  NextID = RS(0) + 1
  On Error GoTo 0
  
End Function

Sub CommandDeleteParameters(cmd As ADODB.Command)
  Dim I As Integer
  For I = 0 To cmd.Parameters.Count - 1
    cmd.Parameters.Delete 0
  Next
End Sub

Function CommandNew() As ADODB.Command
  Set CommandNew = New ADODB.Command
  CommandNew.ActiveConnection = ADOConnection_
End Function

Sub ConnectionOpen()
  ADOConnection_.Open
End Sub

Sub ConnectionClose()
  ADOConnection_.Close
End Sub

Sub ConnectionNew()
  Set ADOConnection_ = New ADODB.Connection
  'ADOConnection_.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0;HDR=Yes;READONLY=FALSE"";"
  'ADOConnection_.ConnectionString = "DRIVER=SQLite3 ODBC Driver;Database=c:\users\ernie\onedrive\documents\programming\TestDB.db;LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;"
  ADOConnection_.ConnectionString = "DRIVER=SQLite3 ODBC Driver;Database=c:\users\ernie\onedrive\documents\programming\TestDB.db;"
End Sub

Function StrIntr(str As String, ParamArray params()) As String
  
  StrIntr = str
  Dim strBegin As Integer: strBegin = InStr(1, str, "#{")
  
  Do While strBegin > 0
      
    strBegin = strBegin + 2
    Dim strEnd As Integer: strEnd = InStr(strBegin, str, "}")
    Dim idxsStr As String: idxsStr = Mid(str, strBegin, strEnd - strBegin)
    
    Dim idx As Integer
    
    If IsNumeric(idxsStr) Then
      idx = idxsStr
    Else
      ' Fail
    End If
    
    StrIntr = Replace(StrIntr, "#{" & idxsStr & "}", params(idx - 1))
    
    strBegin = InStr(strBegin, str, "#{")
  Loop

End Function
