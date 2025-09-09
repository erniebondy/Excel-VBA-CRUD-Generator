VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabelCollectionWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ThisCollection As VBA.Collection

Private Sub Class_Initialize()
  Set ThisCollection = New VBA.Collection
End Sub


