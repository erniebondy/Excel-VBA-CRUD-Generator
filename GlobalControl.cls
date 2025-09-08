VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GlobalControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GlobalOffsetReturnType
  XOffset As Single
  YOffset As Single
End Type

Dim WithEvents pGlLabel As MSForms.Label
Attribute pGlLabel.VB_VarHelpID = -1
Dim WithEvents pGlTextBox As MSForms.TextBox
Attribute pGlTextBox.VB_VarHelpID = -1
Dim WithEvents pGlButton As MSForms.CommandButton
Attribute pGlButton.VB_VarHelpID = -1
Dim WithEvents pGlComboBox As MSForms.ComboBox
Attribute pGlComboBox.VB_VarHelpID = -1
Dim WithEvents pGlListBox As MSForms.ListBox
Attribute pGlListBox.VB_VarHelpID = -1
Dim WithEvents pGlFrame As MSForms.Frame
Attribute pGlFrame.VB_VarHelpID = -1
Dim WithEvents pGlMultiPage As MSForms.MultiPage
Attribute pGlMultiPage.VB_VarHelpID = -1

Public GlManager As GlobalControlManager
Public InnerControl As MSForms.Control

Public Event GlControlMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, ByVal Src As String)

Sub Init(GlManager As GlobalControlManager, InnerControl As MSForms.Control)
  Set Me.GlManager = GlManager
  Set Me.InnerControl = InnerControl
  
  If TypeOf InnerControl Is MSForms.Label Then
    Set pGlLabel = InnerControl
  ElseIf TypeOf InnerControl Is MSForms.Frame Then
    Set pGlFrame = InnerControl
  ElseIf TypeOf InnerControl Is MSForms.CommandButton Then
    Set pGlButton = InnerControl
  ElseIf TypeOf InnerControl Is MSForms.ComboBox Then
    Set pGlComboBox = InnerControl
  ElseIf TypeOf InnerControl Is MSForms.TextBox Then
    Set pGlTextBox = InnerControl
  ElseIf TypeOf InnerControl Is MSForms.ListBox Then
    Set pGlListBox = InnerControl
  ElseIf TypeOf InnerControl Is MSForms.MultiPage Then
    Set pGlMultiPage = InnerControl
  Else
    Debug.Print "[INFO] Control not handled! " & TypeName(InnerControl)
  End If
  
End Sub

Private Function GetGlobalOffset() As GlobalOffsetReturnType
  
  Dim P As Object
  Set P = InnerControl.Parent
  
  ''' For some reason when testing if P was TypeOf UserForm
  ''' it evaluated to true, no matter if P was a Control or a UserForm
  ''' Testing if P is TypeOf Control seems to work
  
  Dim XSum As Single
  Dim YSum As Single
  
  Do While TypeOf P Is MSForms.Control
    ''' Page control has no left/top
    ''' FIX: Clean this ?
    On Error Resume Next
    XSum = XSum + P.Left
    YSum = YSum + P.Top
    If Err.Number <> 0 Then
      Debug.Print "[ERROR] No left/top properties for control type " & TypeName(P)
    End If
    On Error GoTo 0
    Set P = P.Parent
  Loop

  ' Return (BTW Return is a keyword in VBA... But it doesn't do what you think it does :3)
  GetGlobalOffset.XOffset = XSum
  GetGlobalOffset.YOffset = YSum
  
End Function

''' Not too sure how to reduce the repetition here, if it's even possible
''' FIX: Controls in a Frame of MultiPage zero out to their 'owner' or 'parent'
Private Sub pGlButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Call GlMouseMoveHost(Button, Shift, x, Y)
End Sub

Private Sub pGlComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Call GlMouseMoveHost(Button, Shift, x, Y)
End Sub

Private Sub pGlFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Call GlMouseMoveHost(Button, Shift, x, Y)
End Sub

Private Sub pGlLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Call GlMouseMoveHost(Button, Shift, x, Y)
End Sub

Private Sub pGlListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Call GlMouseMoveHost(Button, Shift, x, Y)
End Sub

Private Sub pGlMultiPage_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Call GlMouseMoveHost(Button, Shift, x, Y)
End Sub

Private Sub pGlTextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Call GlMouseMoveHost(Button, Shift, x, Y)
End Sub

''' Because someone who 'raises' an event is the host?
Private Sub GlMouseMoveHost(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Set GlManager.GlActiveControl = Me
  Dim Offsets As GlobalOffsetReturnType
  Offsets = GetGlobalOffset
  RaiseEvent GlControlMove(Button, Shift, InnerControl.Left + Offsets.XOffset + x, InnerControl.Top + Offsets.YOffset + Y, InnerControl.Name)
End Sub
