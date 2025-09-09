VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTestMovables 
   Caption         =   "Test Movables"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7035
   OleObjectBlob   =   "FrmTestMovables.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmTestMovables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' TODO: Fix Z-Ordering
Dim MouseDown As Boolean

Dim Movables(2) As MSForms.Label
Dim SelectedLbl As MSForms.Label
Dim DX As Single
Dim DY As Single
Dim PX As Single
Dim PY As Single

Private Sub Label2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  MouseDown = True
  Dim I As Integer
  For I = LBound(Movables) To UBound(Movables)
    If Label2.Left + X >= Movables(I).Left And Label2.Left + X <= Movables(I).Left + Movables(I).Width And _
       Label2.Top + Y >= Movables(I).Top And Label2.Top + Y <= Movables(I).Top + Movables(I).Height Then
      Set SelectedLbl = Movables(I)
      DX = (Label2.Left + X) - SelectedLbl.Left
      DY = (Label2.Top + Y) - SelectedLbl.Top
      PX = X
      PY = Y
      Exit For
    End If
  Next
  
End Sub

Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Label3.Caption = Label3.Left
  Label4.Caption = Label4.Left
  Label5.Caption = Label5.Left
  
  
  Label6.Caption = "LX " & X & vbNewLine & "FX " & Label2.Left + X
  
  If SelectedLbl Is Nothing Then Exit Sub
  
  Label6.Caption = Label6.Caption & vbNewLine & "SL.Left " & SelectedLbl.Left
  
  Dim NX As Single
  NX = Label2.Left + X - DX
  
  If SelectedLbl.Left - (PX - X) < Label2.Left Then
    NX = Label2.Left
  ElseIf SelectedLbl.Left + SelectedLbl.Width - (PX - X) > Label2.Left + Label2.Width Then
    NX = Label2.Left + Label2.Width - SelectedLbl.Width
  Else
    PX = X
  End If
  
  
  Dim NY As Single
  NY = Label2.Top + Y - DY
  
  If SelectedLbl.Top - (PY - Y) < Label2.Top Then
    NY = Label2.Top
  ElseIf SelectedLbl.Top + SelectedLbl.Height - (PY - Y) > Label2.Top + Label2.Height Then
    NY = Label2.Top + Label2.Height - SelectedLbl.Height
  Else
    PY = Y
  End If
  
  SelectedLbl.Left = NX
  SelectedLbl.Top = NY
  
End Sub

Private Sub Label2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  MouseDown = False
  Set SelectedLbl = Nothing
End Sub

Private Sub UserForm_Initialize()
  
  Label2.Caption = Label2.Left
  Label3.Caption = Label3.Left
  Label4.Caption = Label4.Left
  Label5.Caption = Label5.Left
  
  MouseDown = False
  Set Movables(0) = Me.Label3
  Set Movables(1) = Me.Label4
  Set Movables(2) = Me.Label5
  
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  MouseDown = False
End Sub


