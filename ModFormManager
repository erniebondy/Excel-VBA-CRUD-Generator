Attribute VB_Name = "ModFormManager"
Option Explicit

Sub ApplyPalette(UForm As MSForms.UserForm)

  UForm.BackColor = ColorBackground
  UForm.ForeColor = ColorForeground
  
  Dim Ctrl As MSForms.Control
  For Each Ctrl In UForm.Controls
    'Debug.Print TypeName(Ctrl)
    Select Case TypeName(Ctrl)
      Case "CommandButton"
        Dim TempBtn As MSForms.CommandButton
        Set TempBtn = Ctrl
        TempBtn.BackColor = ColorTextBackground
        TempBtn.ForeColor = ColorTextForeground
      Case "ListBox"
        Dim TempLbo As MSForms.ListBox
        Set TempLbo = Ctrl
        TempLbo.BackColor = ColorTextBackground
        TempLbo.ForeColor = ColorTextForeground
      Case Else
        'Debug.Print "No palette applied to control: " & TypeName(Ctrl)
    End Select
  Next
  
End Sub
