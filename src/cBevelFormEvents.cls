VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBevelFormEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Host As Object
Public WithEvents GenerateButton As MSForms.CommandButton
Public WithEvents CancelButton As MSForms.CommandButton
Public WithEvents ModeCombo As MSForms.ComboBox
Public WithEvents LockCombo As MSForms.ComboBox

Private Sub GenerateButton_Click()
    Host.GenerateClicked
End Sub

Private Sub CancelButton_Click()
    Host.CancelClicked
End Sub

Private Sub ModeCombo_Change()
    Host.UpdateUI
End Sub

Private Sub LockCombo_Change()
    Host.UpdateUI
End Sub
