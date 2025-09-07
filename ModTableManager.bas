Attribute VB_Name = "ModTableManager"
Option Explicit

Sub TT()
  
  Call ModCommon.ConnectionNew
  Call TableCreate("test226")
  
End Sub

Sub TableCreate(TableName As String)

  Call ModCommon.ConnectionOpen
  
  Dim cmd As ADODB.Command: Set cmd = ModCommon.CommandNew

  cmd.Parameters.Refresh
  cmd.Parameters.Append cmd.CreateParameter("id", adInteger, , , Null)
  cmd.Parameters.Append cmd.CreateParameter("name", adVarWChar, , Len(TableName), TableName)
  cmd.Parameters.Append cmd.CreateParameter("created_at", adDate, , , Now)
  cmd.Parameters.Append cmd.CreateParameter("updated_at", adDate, , , Null)

  'cmd.CommandText = "INSERT INTO TABLES VALUES(@id, @name, @created_at, @update_at)"
  cmd.CommandText = "INSERT INTO TABLES VALUES(?, ?, ?, ?)"
  cmd.Execute
  
  Call ModCommon.ConnectionClose
  
'  Call ModCommon.ConnectionOpen
'
'  Dim cmd As ADODB.Command: Set cmd = ModCommon.CommandNew
'
'  shTables.Range("A1:A2").NumberFormat = 0
'
'  Dim NextID As Long: NextID = ModCommon.NextID(shTables.Name)
'
'  cmd.Parameters.Refresh
'  cmd.Parameters.Append cmd.CreateParameter("id", adInteger, , , NextID)
'  cmd.Parameters.Append cmd.CreateParameter("name", adVarWChar, , Len(TableName), TableName)
'  cmd.Parameters.Append cmd.CreateParameter("created_at", adDate, , , Now)
'  cmd.Parameters.Append cmd.CreateParameter("updated_at", adDate, , , Null)
'
'  cmd.CommandText = "INSERT INTO [" & shTables.Name & "$] VALUES(@id, @name, @created_at, @updated_at)"
'  cmd.Execute
'
'  If shTables.Cells(2, 1).Value = vbNullString Then shTables.Rows(2).Delete
'
'  Call ModCommon.ConnectionClose
  
End Sub

'ShTABLES.Range("A1:A2").NumberFormat = 0

Sub TablesLoad()
  
  Call ModCommon.ConnectionOpen
    
  Dim cmd As ADODB.Command: Set cmd = ModCommon.CommandNew
  'Cmd.ActiveConnection = ModCommon.ADOConnection
  'cmd.CommandText = "SELECT * FROM [" & ShTABLES.Name & "$]"
  'cmd.CommandText = StrIntr("SELECT * FROM [#{1}$]", shTables)
  cmd.CommandText = "SELECT * FROM TABLES"
  
  Dim RS As ADODB.Recordset: Set RS = cmd.Execute
  
  FrmMain.LboTables.Clear
  
  Do While Not RS.EOF
    'Debug.Print Rs("id"), Rs("name"), Rs("created_at"), Rs("updated_at")
    FrmMain.LboTables.AddItem RS("name")
    FrmMain.LboTables.List(FrmMain.LboTables.ListCount - 1, 1) = RS("id")
    RS.MoveNext
  Loop
  
  Call ModCommon.ConnectionClose
  
End Sub

