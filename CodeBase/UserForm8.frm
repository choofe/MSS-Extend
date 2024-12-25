VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm8 
   Caption         =   "ËÈÊ ÊÚãíÑÇÊ"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10590
   OleObjectBlob   =   "UserForm8.frx":0000
End
Attribute VB_Name = "UserForm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
'itemCounter = itemCounter + 1
'MsgBox itemCounter

addItem

End Sub

Private Sub CommandButton2_Click()
Dim i As Integer
'.Controls.Remove ("cost" & itemCounter)
'.Controls.Remove ("description" & itemCounter)
'.Controls.Remove ("itemnumber" & itemCounter)
'.Controls.Remove ("deleterow" & itemCounter)

For i = 1 To itemCounter
With Frame3
If .Controls("description" & i).text = "" Then
    MsgBox ("ÑÏíÝ " & i & " ÝÇÞÏ ÊæÖíÍÇÊ ÇÓÊ!")
    Exit Sub
End If
If .Controls("cost" & i).text = "" Then
    MsgBox ("ÑÏíÝ " & i & " ÝÇÞÏ åÒíäå!")
    Exit Sub
End If
If Not IsNumeric(.Controls("cost" & i).text) Then
    MsgBox ("ÑÏíÝ " & i & " ãÞÇÏíÑ ÚÏÏí!")
    Exit Sub
End If

End With
Next
repairRegister (Label14.Caption)
End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub



Private Sub UserForm_Initialize()
Dim sheetIndex
sheetIndex = indexFind(UserForm1.Frame2.ListBox1.Value, originalAssetList) + 1
Label14.Caption = Service.Sheets(sheetIndex).Name

Set collectionEvent = New Collection
itemCounter = 0
UserForm8.Left = ((getScreenX / 2) / 1.333) - (UserForm8.Width / 2)
UserForm8.Top = 25
'Set Service = ThisWorkbook
If Len(Dir(ThisWorkbook.Path & "\history.xlsm")) = 0 Then
    MsgBox "ÝÇíá ÊÇÑíÎå ãæÌæÏ äíÓÊ!"
Else
    Set History = Workbooks.Open(ThisWorkbook.Path & "\history.xlsm")
End If

Dim jalaliToday As Variant: jalaliToday = toJalaaliFromDateObject(Date)
Dim deployDate As Variant: deployDate = toJalaaliFromDateObject(Date - 5)
TextBox55.Value = jalaliToday(2)
TextBox56.Value = jalaliToday(1)
TextBox57.Value = jalaliToday(0)

TextBox58.Value = deployDate(2)
TextBox59.Value = deployDate(1)
TextBox60.Value = deployDate(0)

addItem
'-----
End Sub

'Sub addItem()
'Dim itemNumLabel As MSForms.label
'Set itemNumLabel = Me.Frame3.Controls.Add("forms.label.1", "itemNumber" & itemCounter, True)
'With itemNumLabel
'    .Top = itemCounter * 32
'    .Height = 20
'    .Width = 36
'    .Left = 444
'    .Caption = itemCounter
'    .Font.Size = 12
'    .Font.Bold = msoTrue
'    .TextAlign = 2 'center align
'    .BorderStyle = 1
'End With
'
'Dim descriptionTextBox As MSForms.TextBox
'Set descriptionTextBox = Me.Frame3.Controls.Add("forms.textbox.1", "Description" & itemCounter, True)
'With descriptionTextBox
'    .Top = itemCounter * 32
'    .Height = 20
'    .Width = 300
'    .Left = 132
'    .Font.Size = 12
'    .Font.Bold = msoTrue
'    .TextAlign = 3 'right to left
'    .BorderStyle = 1
'End With
'
'Dim costTextBox As MSForms.TextBox
'Set costTextBox = Me.Frame3.Controls.Add("forms.textbox.1", "Cost" & itemCounter, True)
'With costTextBox
'    .Top = itemCounter * 32
'    .Height = 20
'    .Width = 75
'    .Left = 48
'    .Font.Size = 12
'    .Font.Bold = msoTrue
'    .TextAlign = 1 'left to right
'    .BorderStyle = 1
'
'End With
'Dim clickCMDEvents As UserFormEvents
'
'Dim deleteRowButton As MSForms.CommandButton
'Set deleteRowButton = Me.Frame3.Controls.Add("Forms.commandbutton.1", "deleteRow" & itemCounter, True)
'With deleteRowButton
'    .Top = itemCounter * 32
'    .Height = 20
'    .Width = 25
'    .Left = 16
'    .Caption = "-"
'    .Font.Size = 12
'    .Font.Bold = msoTrue
'    .ForeColor = RGB(250, 0, 0)
'    .Tag = itemCounter
'End With
'Set clickCMDEvents = New UserFormEvents
'Set clickCMDEvents.commandButtonNumber = deleteRowButton
'collectionEvent.Add clickCMDEvents
'
'Dim bigFrame As Boolean
'If UserForm8.Height > 620 Then
'    bigFrame = True
'    Frame3.ScrollBars = fmScrollBarsVertical
'    Frame3.ScrollHeight = Frame3.ScrollHeight + 32
'    CommandButton1.Top = CommandButton1.Top + 32
'    Frame3.ScrollTop = Frame3.ScrollHeight - 70
'Else
'    UserForm8.Height = Frame3.Top + itemCounter * 32 + 150
'    Frame3.ScrollBars = fmScrollBarsNone
'    Frame3.Height = (itemCounter + 1) * 32 + 40
'    CommandButton1.Top = CommandButton1.Top + 32
'    CommandButton2.Top = UserForm8.Height - 70
'    CommandButton3.Top = UserForm8.Height - 70
'
'End If
'itemCounter = itemCounter + 1
'End Sub



'Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'if controls("deleterow"&
'End Sub
