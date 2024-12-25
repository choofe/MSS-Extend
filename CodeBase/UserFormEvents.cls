VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserFormEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents commandButtonNumber As MSForms.CommandButton
Attribute commandButtonNumber.VB_VarHelpID = -1
Public tagNumber As Integer

Private Sub commandButtonNumber_Click()
tagNumber = commandButtonNumber.Tag
removeItem (tagNumber)
End Sub

