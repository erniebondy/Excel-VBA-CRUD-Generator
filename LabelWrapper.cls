VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabelWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ThisLabel As MSForms.Label
Attribute ThisLabel.VB_VarHelpID = -1

Private Sub ThisLabel_Click()
  ''' Yes, VBA has semi-colons, no they don't do what you think they do.
  ''' A semi-colon allows the next Debug.Print statement to remain on
  ''' on the same line
  Debug.Print ThisLabel.Name & " SELECTED"
End Sub
