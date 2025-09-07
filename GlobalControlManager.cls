VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GlobalControlManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' Add reference Microsoft Scripting Runtime


Dim GlFormObj As Object
Dim WithEvents GlForm As MSForms.UserForm
Attribute GlForm.VB_VarHelpID = -1
Dim WithEvents GlLabel As MSForms.Label
Attribute GlLabel.VB_VarHelpID = -1

Dim GlControls As Scripting.Dictionary
Public WithEvents GlActiveControl As GlobalControl

Public Event GlobalMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal Src As String)

''' Explicitly using type 'MSForms.UserForm' results in the UserForm
''' variable to hold some data - like the forms' controls.
''' But using type 'object' gives the proper form type.
''' For instance FrmTest will result in the UserForm variable
''' to be of type Object/FrmTest as opposed to just UserForm/UserForm
''' This may have something to do with the Attribute VB_PredeclaredId = True
''' but that's pure speculation
Sub GlManagerInit(UserForm As Object)

  Set GlForm = UserForm
  Set GlFormObj = UserForm
  
  Set GlControls = New Scripting.Dictionary
  GlControls.Add UserForm.Name, GlForm
  
  Dim I As Integer
  For I = 0 To GlForm.Controls.Count - 1
  
    Dim GlCtrl As GlobalControl
    Set GlCtrl = New GlobalControl
    GlCtrl.Init Me, GlForm.Controls(I)
    GlControls.Add GlCtrl.InnerControl.Name, GlCtrl

  Next
  
End Sub

Private Sub GlActiveControl_GlControlMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal Src As String)
  RaiseEvent GlobalMouseMove(Button, Shift, X, Y, GlActiveControl.InnerControl.Name)
End Sub

Private Sub GlForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 RaiseEvent GlobalMouseMove(Button, Shift, X, Y, GlFormObj.Name)
End Sub
