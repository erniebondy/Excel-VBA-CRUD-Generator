Attribute VB_Name = "Module1"
Option Explicit

Sub AAA()


  Dim Dic As Scripting.Dictionary
  Set Dic = New Scripting.Dictionary
  
  Dic.Add "a", "apple"
  Dic.Add "b", "banana"
  Dic.Add "c", "carrot"
  
  Debug.Print Dic.Items(Dic.Count - 1)
  
End Sub

Sub BBB()

  Dim F1 As FrmChat
  Set F1 = New FrmChat
  F1.Show vbModeless
  
  Dim F2 As FrmChat
  Set F2 = New FrmChat
  F2.Show vbModeless
  
  FrmInput.Show vbModeless
  
  
  Debug.Print VBA.UserForms.Count

  
End Sub
