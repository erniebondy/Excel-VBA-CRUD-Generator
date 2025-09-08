VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmChat 
   Caption         =   "Chat"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "FrmChat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub BtnSend_Click()
  Call SendMessage(TxtName.Value, TxtMessage.Value)
End Sub

Private Sub TxtMessage_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  If KeyCode.Value = KeyCodeConstants.vbKeyReturn Then
    KeyCode = 0 ''' Prevents default action of moving focus to next control (ChatGPT)
    Call SendMessage(TxtName.Value, TxtMessage.Value)
  End If
End Sub

Private Sub UserForm_Initialize()
'
End Sub

Private Sub SendMessage(Name As String, Message As String)

  Dim ChatForm
  For Each ChatForm In VBA.UserForms
    If Not TypeOf ChatForm Is FrmChat Then GoTo Continue ''' Holy statement Batman
    If ChatForm Is Me Then GoTo Continue
    Dim Messages As MSForms.ListBox
    Set Messages = ChatForm.Controls("LboMessages")
    Messages.AddItem Name
    Messages.List(Messages.ListCount - 1, 1) = Message
    
    'ChatForm.Controls("TxtMessage").Value = vbNullString
    'ChatForm.Controls("TxtMessage").SetFocus
    
'    LboMessages.AddItem Name
'    LboMessages.List(LboMessages.ListCount - 1, 1) = Message
'    TxtMessage.Value = vbNullString
'    TxtMessage.SetFocus
Continue:
  Next
  
  LboMessages.AddItem Name
  LboMessages.List(LboMessages.ListCount - 1, 1) = Message
  TxtMessage.Value = vbNullString
  TxtMessage.SetFocus
  
End Sub
