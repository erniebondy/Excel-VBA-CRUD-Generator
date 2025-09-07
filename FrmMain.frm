VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain 
   Caption         =   "CRUD Generator"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "FrmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnTableCreate_Click()
  Call ModTableManager.TableCreate("my_table")
  Call ModTableManager.TablesLoad
End Sub

Private Sub UserForm_Initialize()
  Call ModFormManager.ApplyPalette(Me)
  Call ModCommon.ConnectionNew
  Call ModTableManager.TablesLoad
End Sub
