VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPaletteViewer 
   Caption         =   "Palette Viewer"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7605
   OleObjectBlob   =   "FrmPaletteViewer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmPaletteViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''' Add reference Microsoft Scripting Runtime

Const WSName As String = "__SWATCHES__"
Dim SaveWS As Worksheet
Dim Swatches As Scripting.Dictionary
Dim Labels As VBA.Collection

Private Enum SaveWSCols
  enHex = 1
  enName
End Enum

Private Sub BtnAddSwatch_Click()
  ''' FIX: Add error checking
  Call SwatchAdd("&H" & TxtSwatchValue.Value, TxtName.Value)
End Sub


Private Sub BtnSave_Click()
  Call SwatchSave
End Sub

Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  CommandButton2.BackColor = SystemColorConstants.vbHighlight
End Sub

'Dim Entered As Boolean
'
'Private Sub TxtSwatchValue_Enter()
'  If Entered Then Exit Sub
'  Entered = True
'  TxtSwatchValue.Text = vbNullString
'  TxtSwatchValue.Font.Italic = False
'  TxtSwatchValue.ForeColor = SystemColorConstants.vbWindowText
'End Sub

Private Sub TxtSwatchValue_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  If KeyCode.Value = KeyCodeConstants.vbKeyReturn Then
    '
  End If
End Sub

Private Sub UserForm_Initialize()

  Set Swatches = New Scripting.Dictionary
  Set Labels = New VBA.Collection
  
  Call SwatchesLoad
  
'  Entered = False
  ListBox1.AddItem "I"
  ListBox1.AddItem "am"
  ListBox1.AddItem "a"
  ListBox1.AddItem "listbox"
  
'  Dim Palettes(5 - 1, 1)
'  Palettes(0, 0) = PaletteColorAsparagus
'  Palettes(0, 1) = "Asparagus"
'  Palettes(1, 0) = PaletteColorAlmond
'  Palettes(1, 1) = "Almond"
'  Palettes(2, 0) = PaletteColorSpaceCadet
'  Palettes(2, 1) = "Space Cadet"
'  Palettes(3, 0) = PaletteColorTaupeGray
'  Palettes(3, 1) = "Taupe Gray"
'  Palettes(4, 0) = PaletteColorRoseQuartz
'  Palettes(4, 1) = "Rose Quartz"
'
'  Dim I As Integer
'  Dim Padding As Single: Padding = 20
'  Dim LblWidth As Single: LblWidth = 50
'  Dim LblHeight As Single: LblHeight = 50
'
'  For I = LBound(Palettes) To UBound(Palettes)
'    Dim Lbl As MSForms.Label
'    Set Lbl = Me.Controls.Add("Forms.Label.1", , True)
'    Lbl.Left = Padding + (I * (LblWidth + Padding))
'    Lbl.Top = Padding
'    Lbl.Width = LblWidth
'    Lbl.Height = LblHeight
'    Lbl.BorderColor = vbBlack
'    Lbl.BorderStyle = fmBorderStyleSingle
'    Lbl.BackColor = Palettes(I, 0)
'    Lbl.Caption = Palettes(I, 1)
'    Lbl.TextAlign = fmTextAlignCenter
'
'    ' Possibly add a 'slider' to form to adjust threshold
'    If Palettes(I, 0) < 6000000 Then Lbl.ForeColor = vbWhite
'    'Debug.Print Palettes(I, 0)
'  Next
'
'  Call ApplyPalette(Me)

  
  
End Sub

Private Sub SwatchAdd(Hex As Long, Name As String)

  If Not Swatches.Exists(Hex) Then Swatches.Add Hex, Name

  Dim Padding As Single: Padding = 20
  Dim LblWidth As Single: LblWidth = 50
  Dim LblHeight As Single: LblHeight = 50
  
  Dim Lbl As MSForms.Label
  Set Lbl = FrSwatches.Controls.Add("Forms.Label.1", "Lbl" & Name, True)
  Lbl.Left = Padding + ((Swatches.Count - 1) * (LblWidth + Padding))
  Lbl.Top = Padding
  Lbl.Width = LblWidth
  Lbl.Height = LblHeight
  Lbl.BorderColor = vbBlack
  Lbl.BorderStyle = fmBorderStyleSingle
  Lbl.BackColor = Hex
  Lbl.Caption = Name
  Lbl.TextAlign = fmTextAlignCenter
  
  Dim LblWrapper As LabelWrapper
  Set LblWrapper = New LabelWrapper
  Set LblWrapper.ThisLabel = Lbl
  
  Labels.Add LblWrapper
  
  
  ' Possibly add a 'slider' to form to adjust threshold
  'If Palettes(I, 0) < 6000000 Then Lbl.ForeColor = vbWhite

End Sub

Private Sub SwatchSave()

  Set SaveWS = Nothing
  
  ''' Get or create save worksheet
  On Error Resume Next
  Set SaveWS = ThisWorkbook.Worksheets(WSName)
  On Error GoTo 0
  
  If SaveWS Is Nothing Then
    Set SaveWS = ThisWorkbook.Worksheets.Add
    SaveWS.Name = WSName
    'SaveWS.Visible = xlSheetVeryHidden '...very
  End If
  
  ''' Iterating on a Dictionary iterates the keys by default
  Dim Key
  Dim Row As Long: Row = 1
  For Each Key In Swatches
    SaveWS.Cells(Row, SaveWSCols.enHex) = Key
    SaveWS.Cells(Row, SaveWSCols.enName) = Swatches(Key)
    Inc Row
  Next
  
  ''' Warn user that saving swatches saves the workbook
  ''' before continuing!
  ThisWorkbook.Save
  
End Sub

Private Sub SwatchesLoad()

  Set SaveWS = Nothing
  
  On Error Resume Next
  Set SaveWS = ThisWorkbook.Worksheets(WSName)
  On Error GoTo 0
  
  If SaveWS Is Nothing Then Exit Sub
  
  Dim Row As Long: Row = 1
  Do While SaveWS.Cells(Row, SaveWSCols.enHex) <> vbNullString
    Swatches.Add SaveWS.Cells(Row, SaveWSCols.enHex).Value, _
                 SaveWS.Cells(Row, SaveWSCols.enName).Value
    Call SwatchAdd(CLng(Swatches.Keys(Swatches.Count - 1)), _
                   CStr(Swatches.Items(Swatches.Count - 1)))
    Inc Row
  Loop
  
End Sub



























