Attribute VB_Name = "ModExportComponents"
Option Explicit
''' Add reference Microsoft Visual Basic for Applications Extensibility 5.3

Sub ExportComponentsMain()
  
  Dim ThisProject As VBProject
  Set ThisProject = ThisWorkbook.VBProject
  
  Dim Component As VBComponent
  For Each Component In ThisProject.VBComponents
    
    Dim Ext As String
    Select Case Component.Type
    Case VBIDE.vbext_ct_StdModule
      'Debug.Print "[STDMODULE] " & Component.Name
      Ext = ".bas"
    Case VBIDE.vbext_ct_ClassModule
      'Debug.Print "[CLASSMODULE] " & Component.Name
      Ext = ".cls"
    Case VBIDE.vbext_ct_Document
      'Debug.Print "[DOCUMENT] " & Component.Name
      Ext = ".cls"
    Case VBIDE.vbext_ct_MSForm
      'Debug.Print "[MSFORM] " & Component.Name
      Ext = ".cls"
    Case Else
      Debug.Print "[INFO] Could not determine extension for component " & Component.Name
    End Select
        
    On Error Resume Next
    Component.Export "crud_generator_components/" & Component.Name & Ext
    
    If Err.Number <> 0 Then
      Debug.Print "[ERROR] Could not export component " & Component.Name
    End If
    On Error GoTo 0
    
  Next
  
End Sub
 
