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

Public LabelWrappers As LabelCollectionWrapper
Public WithEvents ThisLabel As MSForms.Label
Attribute ThisLabel.VB_VarHelpID = -1

Dim pSelected As Boolean

Private Sub Class_Initialize()
  pSelected = False
  Set LabelWrappers = New LabelCollectionWrapper
End Sub

Private Sub ThisLabel_Click()
  ''' Yes, VBA has semi-colons, no they don't do what you think they do.
  ''' A semi-colon allows the next Debug.Print statement to remain on
  ''' on the same line
  'Debug.Print ThisLabel.Name; " SELECTED"
  
  ''' Could create a label of the same size as the swatch + padding
  ''' and center it to the swatch to create the effect of a "thick" border
  
  pSelected = Not pSelected
  
  ThisLabel.BorderStyle = fmBorderStyleSingle
  ThisLabel.BorderColor = IIf(pSelected, vbRed, vbBlack)
  ThisLabel.Parent.Repaint ''' Updates visuals
  
  If pSelected Then Set FrmPaletteViewer.SelectedSwatch = ThisLabel
  
  Call ClearSelection
  
End Sub

''' Potentially look into EventTracing using WinAPI to determine
''' the callee of the sub
Sub Unselect()
  pSelected = False
End Sub

Private Sub ClearSelection()
  
  Dim L As LabelWrapper
  Dim I As Integer
  
  For I = 1 To LabelWrappers.ThisCollection.Count
    Set L = LabelWrappers.ThisCollection(I)
    If Not L Is Me Then
      L.ThisLabel.BorderStyle = fmBorderStyleSingle
      L.ThisLabel.BorderColor = vbBlack
      L.ThisLabel.Parent.Repaint ''' Updates visuals
      L.Unselect
    End If
  Next
  
End Sub
