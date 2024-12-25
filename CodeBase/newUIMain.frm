VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newUIMain 
   Caption         =   "NewUI"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19695
   OleObjectBlob   =   "newUIMain.frx":0000
   RightToLeft     =   -1  'True
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "newUIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub UserForm_Activate()
'With ListView1
'   .CheckBoxes = True
'   .Gridlines = True
'   With .ColumnHeaders
'      .Clear
'      .Add , , "Item", 70
'      .Add , , "Subitem-1", 70
'      .Add , , "Subitem-2", 70
'   End With
'
'   Dim li As ListItem
'
'   Set li = .ListItems.Add()
'   li.ListSubItems.Add , , "Subitem 1.1"
'   li.ListSubItems.Add , , "Subitem 1.2"
'
'   Set li = .ListItems.Add()
'   li.ListSubItems.Add , , "Subitem 2.1"
'   li.ListSubItems.Add , , "Subitem 2.2"
'
'   Set li = .ListItems.Add()
'   li.ListSubItems.Add , , "Subitem 3.1"
'   li.ListSubItems.Add , , "Subitem 3.2"
'
''   .ColumnHeaders(1).Position = 2
'End With

With ListView1
   .View = lvwReport
   .FullRowSelect = True
    .ColumnHeaders.Add , , "ÂíÊã ÓÑæíÓ", 70
   
    .ColumnHeaders.Add , , "˜íáæãÊÑ ÊÚæíÖ", 100
    .ColumnHeaders(2).Alignment = lvwColumnCenter

'With ListView1
Dim i
For i = 1 To 5
    With .ColumnHeaders
        .Add , , i, 50
        If i > 1 Then ListView1.ColumnHeaders(i).Alignment = lvwColumnCenter
    End With

Next
For i = 1 To 5
     .ListItems.Add.ListSubItems.Add , , i
     
Next
'.ListItems.Add.ListSubItems.Add , , "hello"
'.ListItems.Add.ListSubItems.Add , , "hello"
'.ListItems.Add , , "hello"
.Gridlines = True

End With
End Sub

Private Sub UserForm_Click()

With ListView1
'    .ColumnWidths = "4 cm ; 2 cm ; 3 cm ; 3 cm ; 3 cm"
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "ÂíÊã ÓÑæíÓ", 40
    .ColumnHeaders.Add , , "˜íáæãÊÑ ÊÚæíÖ", 50
'    .addItem ""
'    .List(0, 1) = "˜ÇÑ˜ÑÏ ÂíÊã"
'    .List(0, 2) = "˜íáæãÊÑ ÊÚæíÖ"
'    .List(0, 3) = "˜íáæãÊÑ ÊÇ ÊÚæíÖ"
'    .List(0, 4) = "ÊÚæíÖ"

End With
Set Service = ThisWorkbook

'If Len(Dir(ThisWorkbook.Path & "\history.xlsm")) = 0 Then
'    MsgBox "İÇíá ÊÇÑíÎå ãæÌæÏ äíÓÊ!"
'Else
'    Set History = Workbooks.Open(ThisWorkbook.Path & "\history.xlsm")
'End If

Dim saveCopyAnswer As Long
'saveCopyAnswer = MsgBox(" ÂíÇ ÇÒ æÖÚíÊ İÚáí í˜ ˜í ÇíÌÇÏ ÔæÏ¿", vbYesNo)
'If saveCopyAnswer = vbYes Then ThisWorkbook.SaveCopyAs (Application.ThisWorkbook.Path & "\temp" & Replace(Date, "/", "", 1) & " " & Replace(Time, ":", "", 1) & ".xlsm")
Success = False
Service.Sheets("Kilometrage").Activate
Dim lastRow As Integer

lastRow = Service.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row 'count of unique values
If lastRow <> 1 Then
    'ComboBox1.List = Service.Sheets(1).Range("A1:A" & lastRow).Value
    'ComboBox1.ListIndex = 0
    'ListView1.List = Service.Sheets(1).Range("A1:A" & lastRow).Value
'    ListBox1.ListIndex = 0
    'ReDim originalAssetList(1 To lastRow)
    ReDim templist(1 To lastRow)
    For i = 1 To lastRow
        templist(i) = Service.Sheets(1).Range("A" & i).Value
    Next
    originalAssetList = templist
    'Service.Sheets(1).Range("A1:A" & lastRow).Value
End If


End Sub
