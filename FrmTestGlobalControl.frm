VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTestGlobalControl 
   Caption         =   "Test Global Control"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4740
   OleObjectBlob   =   "FrmTestGlobalControl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmTestGlobalControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents GlCtrlMgr As GlobalControlManager
Attribute GlCtrlMgr.VB_VarHelpID = -1

Private Sub GlCtrlMgr_GlobalMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, ByVal Src As String)
  LblDebug.Caption = "[" & Src & "] " & x & ", " & Y
End Sub

Private Sub UserForm_Initialize()
  
  Set GlCtrlMgr = New GlobalControlManager
  GlCtrlMgr.GlManagerInit Me
  
End Sub
