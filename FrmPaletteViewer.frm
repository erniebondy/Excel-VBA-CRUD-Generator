VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPaletteViewer 
   Caption         =   "Palette Viewer"
   ClientHeight    =   10665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11370
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

''' TODO: Reverse Hex input. User should enter RRGGBB, program reads BBGGRR

Const WSNAME As String = "__SWATCHES__"
Const PADDING As Single = 10
Const LBL_WIDTH As Single = 44
Const LBL_HEIGHT As Single = 44
  
Dim SaveWS As Worksheet
Dim Swatches As Scripting.Dictionary
Dim LabelWrappers As LabelCollectionWrapper
Dim SelectedControl As MSForms.Control
Dim Entered As Boolean

''' Not very 'safe'?
Public SelectedSwatch As MSForms.Label

Private Enum SaveWSCols
  enHex = 1
  enName
End Enum

Private Sub BtnAddSwatch_Click()
  ''' FIX: Add error checking
  Dim HexVal As String
  HexVal = TxtSwatchValue.Value
  Call HexReverse(HexVal)
  HexVal = "&H" & HexVal
  
  Swatches.Add CStr(CLng(HexVal)), TxtName.Value
  Call SwatchesAdd
End Sub

Private Sub BtnApplySwatch_Click()
  If SelectedSwatch Is Nothing Then Exit Sub

  Call MessageSet("Select control to apply swatch")
    
  Set SelectedControl = Nothing
  Do While SelectedControl Is Nothing: DoEvents: Loop
  
  Call SwatchApply(SelectedControl)
  Call MessageClear
  
End Sub

Private Sub BtnDeleteSwatch_Click()
  Call SwatchDelete
End Sub

Private Sub BtnSave_Click()
  Call SwatchSave
End Sub

Private Sub CheckBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Set SelectedControl = CheckBox1
End Sub

Private Sub ComboBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Set SelectedControl = ComboBox1
End Sub

Private Sub CommandButton1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Set SelectedControl = CommandButton1
End Sub

Private Sub CommandButton2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Set SelectedControl = CommandButton2
End Sub

Private Sub FrBackground_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Set SelectedControl = FrBackground
End Sub

Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Set SelectedControl = Label1
End Sub

Private Sub ListBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Set SelectedControl = ListBox1
End Sub

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Set SelectedControl = TextBox1
End Sub

Private Sub UserForm_Initialize()

  Set Swatches = New Scripting.Dictionary
  Set LabelWrappers = New LabelCollectionWrapper
  
  Call SwatchesLoad
  
'  Entered = False
  ListBox1.AddItem "I"
  ListBox1.AddItem "am"
  ListBox1.AddItem "a"
  ListBox1.AddItem "listbox"
  
  
  
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////'

Private Sub SwatchesAdd()

  Dim Key
  Dim Col As Integer: Col = 1
  Dim Row As Integer: Row = 1
  
  For Each Key In Swatches
    
    If PADDING + (Col * (LBL_WIDTH + PADDING)) >= FrSwatches.InsideWidth Then
      Col = 1
      Inc Row
    End If
    
    Call SwatchAdd(CStr(Key), Swatches(Key), Row, Col)
    Inc Col
    
  Next
  
End Sub

Private Sub SwatchAdd(Hex As Long, Name As String, Row As Integer, Col As Integer)
  
  ''' Determine Left/Top (X/Y), notice Col and Row are 1-based index
  Dim Left As Single
  Left = PADDING + ((Col - 1) * (LBL_WIDTH + PADDING))
  
  Dim Top As Single
  Top = PADDING + ((Row - 1) * (LBL_HEIGHT + PADDING))
    
  Dim Lbl As MSForms.Label
  Set Lbl = FrSwatches.Controls.Add("Forms.Label.1", , True)
  Lbl.Left = Left
  Lbl.Top = Top
  Lbl.Width = LBL_WIDTH
  Lbl.Height = LBL_HEIGHT
  Lbl.BorderColor = vbBlack
  Lbl.BorderStyle = fmBorderStyleSingle
  Lbl.BackColor = Hex
  Lbl.Caption = Name
  Lbl.TextAlign = fmTextAlignCenter
  
  Dim LblWrapper As LabelWrapper
  Set LblWrapper = New LabelWrapper
  Set LblWrapper.ThisLabel = Lbl
  Set LblWrapper.LabelWrappers = LabelWrappers
  
  LabelWrappers.ThisCollection.Add LblWrapper, LblWrapper.ThisLabel.Name
  
  
  ' Possibly add a 'slider' to form to adjust threshold
  'If Palettes(I, 0) < 6000000 Then Lbl.ForeColor = vbWhite

End Sub

Private Sub SwatchDelete()
  
  If SelectedSwatch Is Nothing Then Exit Sub
  
  Dim HexVal As Long: HexVal = SelectedSwatch.BackColor
  Dim LblName As String: LblName = SelectedSwatch.Name
  
  ''' Remove the label
  Me.Controls.Remove LblName
  
  ''' Remove the label wrapper
  LabelWrappers.ThisCollection.Remove LblName
  
  ''' Remove the swatch
  Swatches.Remove CStr(HexVal)
  
  ''' Clear selected swatch
  Set SelectedSwatch = Nothing
  
End Sub

Private Sub SwatchSave()

  Set SaveWS = Nothing
  
  ''' Get or create save worksheet
  On Error Resume Next
  Set SaveWS = ThisWorkbook.Worksheets(WSNAME)
  On Error GoTo 0
  
  If SaveWS Is Nothing Then
    Set SaveWS = ThisWorkbook.Worksheets.Add
    SaveWS.Name = WSNAME
    'SaveWS.Visible = xlSheetVeryHidden '...very
  End If
  
  SaveWS.Cells.Clear
  
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
  Set SaveWS = ThisWorkbook.Worksheets(WSNAME)
  On Error GoTo 0
  
  If SaveWS Is Nothing Then Exit Sub
  
  Dim Row As Long: Row = 1
  Do While SaveWS.Cells(Row, SaveWSCols.enHex) <> vbNullString
    Swatches.Add CStr(SaveWS.Cells(Row, SaveWSCols.enHex).Value), _
                 CStr(SaveWS.Cells(Row, SaveWSCols.enName).Value)
    Inc Row
  Loop
  
  Call SwatchesAdd
  
End Sub

Private Sub SwatchApply(Ctrl As MSForms.Control)
  
  If OBtnForeground.Value Then
    Ctrl.ForeColor = SelectedSwatch.BackColor
  ElseIf OBtnBackground.Value Then
    Ctrl.BackColor = SelectedSwatch.BackColor
  End If
  
End Sub

Private Sub MessageSet(Message As String)
  LblInfo.Caption = Message
End Sub

Private Sub MessageClear()
  LblInfo.Caption = vbNullString
End Sub

''' This could probably be smarter and needs error handling
Private Sub HexReverse(ByRef HexVal As String)

  Dim Comp1 As String, Comp2 As String, Comp3 As String
  
  If Len(HexVal) = 2 Then Exit Sub
  
  If Len(HexVal) = 4 Then
    Comp1 = Mid(HexVal, 1, 2)
    Comp2 = Mid(HexVal, 3, 2)
    HexVal = Comp2 & Comp1
    Exit Sub
  End If
  
  If Len(HexVal) = 6 Then
    Comp1 = Mid(HexVal, 1, 2)
    Comp2 = Mid(HexVal, 3, 2)
    Comp3 = Mid(HexVal, 5, 2)
    HexVal = Comp3 & Comp2 & Comp1
    Exit Sub
  End If
  
End Sub























