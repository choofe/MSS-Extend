VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Motorpool Service Scheduler"
   ClientHeight    =   11805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21165
   OleObjectBlob   =   "UserForm1.frx":0000
   RightToLeft     =   -1  'True
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    equipmentWithItems (UserForm1.Frame2.ListBox1)
    If UserForm1.Frame2.ListBox1.ListCount > 1 Then
        'UserForm1.Frame2.ListBox1.ListIndex = 1
    Else
        UserForm1.Frame2.ListBox1.Clear
    End If
Else
    ListBox1.Clear
    ListBox1.List = originalAssetList
    With Frame2.ListBox2
    .Clear
    .ColumnWidths = "4 cm ; 2 cm ; 3 cm ; 3 cm ; 3 cm"
    .addItem "¬Ì „ ”—ÊÌ”"
    .List(0, 1) = "ò«—ò—œ ¬Ì „"
    .List(0, 2) = "òÌ·Ê„ —  ⁄ÊÌ÷"
    .List(0, 3) = "òÌ·Ê„ —  «  ⁄ÊÌ÷"
    .List(0, 4) = " ⁄ÊÌ÷"
End With
    
End If

End Sub

Private Sub cmdContinue_Click()
normalViewMainform
hideAll
Service.Sheets(1).Activate
End Sub

Private Sub ComboBox1_Change()
If ComboBox1.ListIndex < 1 Then
   CommandButton1.Enabled = False
   CommandButton2.Enabled = False
   CommandButton4.Enabled = False
   CommandButton9.Enabled = False
   CommandButton10.Enabled = False
   CommandButton11.Enabled = False
   CommandButton12.Enabled = False
   CommandButton13.Enabled = False
Else
   CommandButton1.Enabled = True
   CommandButton2.Enabled = True
   CommandButton4.Enabled = True
   CommandButton9.Enabled = True
   CommandButton10.Enabled = True
   CommandButton11.Enabled = True
   CommandButton12.Enabled = True
   CommandButton13.Enabled = True
End If
End Sub

Private Sub CommandButton1_Click()

'Service.Sheets(ComboBox1.Value).Activate
UserForm5.Show

'History.Save
UserForm1.Show
UserForm1.Frame2.ListBox1.ListIndex = UserForm1.Frame2.ListBox1.ListIndex
End Sub

Private Sub CommandButton10_Click()
UserForm6.Show
End Sub

Private Sub CommandButton11_Click()
'Service.Sheets(ComboBox1.Value).Activate
UserForm8.Show
End Sub

Private Sub CommandButton12_Click()
hideAll
Service.Sheets(ComboBox1.ListIndex + 2).Visible = True
minimizeMainForm
MotorPoolPlanning.ThisWorkbook.Sheets(ComboBox1.ListIndex + 2).Activate
End Sub


Private Sub CommandButton13_Click()
equipmentChecklist (indexFind(ListBox1.Value, originalAssetList) + 1)
End Sub

Private Sub CommandButton14_Click()

End Sub

Private Sub CommandButton2_Click()
Call checklistMaker


End Sub

Private Sub CommandButton3_Click()
UserForm2.Show


End Sub

Private Sub CommandButton4_Click()
'Service.Sheets (Frame2.ListBox1.Value)
UserForm3.Show
'Unload Me
End Sub

Private Sub CommandButton5_Click()
'MotorPoolPlanning.ThisWorkbook.Sheets(ComboBox1.ListIndex + 2).Activate
UserForm7.Show
End Sub

Private Sub CommandButton6_Click()
Service.Sheets("kilometrage").Activate
'History.Close savechanges:=True
hideAll
Unload Me
End Sub

Private Sub CommandButton9_Click()
Service.Sheets(ComboBox1.Value).Activate
Dim deleteSheet As VbMsgBoxResult, sheetIndex As Integer
deleteSheet = MsgBox("¬Ì« «“ Õ–› ’›ÕÂ Ê”Ì·Â „ÿ„∆‰ Â” Ìœø «Ì‰ ⁄„·Ì«  €Ì—ﬁ«»· »«“ê‘  „Ì »«‘œ", vbYesNo, "Õ–›")
Dim createCopyAnswer As VbMsgBoxResult
If deleteSheet = vbYes Then
    sheetIndex = Service.Sheets(ComboBox1.Value).index
    Service.Sheets("Kilometrage").Activate
    ActiveSheet.Unprotect
    Rows(sheetIndex - 1).Delete
    ActiveSheet.Protect
    Service.Sheets(ComboBox1.Value).Activate
    ActiveSheet.Unprotect
    Application.DisplayAlerts = False
    Service.Sheets(ComboBox1.Value).Delete
    
'--------------------------------------
'creating history backup on demand
createCopyAnswer = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «“ ’›ÕÂ  «—ÌŒçÂ Ê”Ì·Â ‰ﬁ·ÌÂ Ìò òÅÌ œ— Ìò ›«Ì· „Ã“«  ÂÌÂ ‘Êœø", vbYesNo, "«ÌÃ«œ òÅÌ")
        If createCopyAnswer = vbYes Then saveSingleHistoryToNewWorkbook (sheetIndex)
    History.Sheets(ComboBox1.Value).Delete
    Application.DisplayAlerts = True
    
End If
History.Save
ComboBox1.Clear
Service.Sheets("Kilometrage").Activate
Dim lastRow As Integer
lastRow = Service.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row 'count of unique values
If lastRow <> 1 Then
    ComboBox1.List = Service.Sheets(1).Range("A1:A" & lastRow).Value
    ComboBox1.ListIndex = 0
End If
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub hideButton_Click()
'    If hideButton.Caption = "hide all" Then
'        hideAll
'        hideButton.Caption = "unhide all"
'    Else
'        unHideAll
'        hideButton.Caption = "hide all"
'    End If
UserForm9.Show
End Sub
Private Sub Label17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label16.ForeColor = vbBlack
End Sub

Private Sub Label24_Click()

End Sub

Private Sub Label25_Click()

End Sub

Private Sub ListBox1_Click()
hideAll
Dim sheetIndex
sheetIndex = indexFind(ListBox1.Value, originalAssetList) + 1


'Service.Sheets(ListBox1.ListIndex + 2).Visible = True
'Dim equipmentSheet As Sheets
If ListBox1.ListIndex <> "‘„«—Â Å·«ò «‰ Ÿ«„Ì" Then
With Service.Sheets(sheetIndex)
Dim rowCount As Integer
Frame2.ListBox2.Clear
With Frame2.ListBox2
    .ColumnWidths = "4 cm ; 2 cm ; 3 cm ; 3 cm ; 3 cm"
    .addItem "¬Ì „ ”—ÊÌ”"
    .List(0, 1) = "ò«—ò—œ ¬Ì „"
    .List(0, 2) = "òÌ·Ê„ —  ⁄ÊÌ÷"
    .List(0, 3) = "òÌ·Ê„ —  «  ⁄ÊÌ÷"
    .List(0, 4) = " ⁄ÊÌ÷"
End With
lblEquipmentCode.Caption = .Name
lblModel.Caption = .Range("A4").Value
lblColor.Caption = .Range("E6").Value
lblDepartment.Caption = .Range("G6").Value
lblMileage.Caption = .Range("E" & STARTING_ROW - 2).Value
lblDriver.Caption = .Range("H4").Value
lblPhone.Caption = .Range("I4").Value
rowCount = 1
For i = STARTING_ROW To findLastRow(Service.Sheets(sheetIndex), "B")
    Frame2.ListBox2.addItem .Range("B" & i).Value
    Frame2.ListBox2.List(rowCount, 1) = .Range("G" & i).Value
    Frame2.ListBox2.List(rowCount, 2) = .Range("F" & i).Value + .Range("D" & i).Value
    If .Range("H" & i).Value >= 0 Then Frame2.ListBox2.List(rowCount, 3) = .Range("H" & i).Value Else Frame2.ListBox2.List(rowCount, 3) = "OverRun"
    If .Range("H" & i).Value < 100 And .Range("H" & i).Value >= 0 Then Frame2.ListBox2.List(rowCount, 4) = "[C]"
    If .Range("H" & i).Value < 0 Then Frame2.ListBox2.List(rowCount, 4) = "[O]"
    If .Range("H" & i).Value >= 100 And .Range("H" & i).Value < 300 Then Frame2.ListBox2.List(rowCount, 4) = "[A]"
    If .Range("H" & i).Value >= 200 Then Frame2.ListBox2.List(rowCount, 4) = "[ ]"
    
    rowCount = rowCount + 1
Next

End With
End If
If ListBox1.ListIndex < 1 Then
   CommandButton1.Enabled = False
   CommandButton2.Enabled = False
   CommandButton4.Enabled = False
   CommandButton9.Enabled = False
   CommandButton10.Enabled = False
   CommandButton11.Enabled = False
   CommandButton12.Enabled = False
   CommandButton13.Enabled = False
Else
   CommandButton1.Enabled = True
   CommandButton2.Enabled = True
   CommandButton4.Enabled = True
   CommandButton9.Enabled = True
   CommandButton10.Enabled = True
   CommandButton11.Enabled = True
   CommandButton12.Enabled = True
   CommandButton13.Enabled = True
End If


End Sub

Private Sub TextBox1_Change()
    Call search(TextBox1.text, originalAssetList, UserForm1.ListBox1)
End Sub

Private Sub UserForm_Initialize()

With Frame2.ListBox2
    .ColumnWidths = "4 cm ; 2 cm ; 3 cm ; 3 cm ; 3 cm"
    .addItem "¬Ì „ ”—ÊÌ”"
    .List(0, 1) = "ò«—ò—œ ¬Ì „"
    .List(0, 2) = "òÌ·Ê„ —  ⁄ÊÌ÷"
    .List(0, 3) = "òÌ·Ê„ —  «  ⁄ÊÌ÷"
    .List(0, 4) = " ⁄ÊÌ÷"
End With
Set Service = ThisWorkbook

If Len(Dir(ThisWorkbook.Path & "\history.xlsm")) = 0 Then
    MsgBox "›«Ì·  «—ÌŒçÂ „ÊÃÊœ ‰Ì” !"
Else
    'Set History = Workbooks.Open(ThisWorkbook.Path & "\history.xlsm")
End If

Dim saveCopyAnswer As Long
'saveCopyAnswer = MsgBox(" ¬Ì« «“ Ê÷⁄Ì  ›⁄·Ì Ìò òÅÌ «ÌÃ«œ ‘Êœø", vbYesNo)
'If saveCopyAnswer = vbYes Then ThisWorkbook.SaveCopyAs (Application.ThisWorkbook.Path & "\temp" & Replace(Date, "/", "", 1) & " " & Replace(Time, ":", "", 1) & ".xlsm")
Success = False
Service.Sheets("Kilometrage").Activate
Dim lastRow As Integer

lastRow = Service.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row 'count of unique values
If lastRow <> 1 Then
    ComboBox1.List = Service.Sheets(1).Range("A1:A" & lastRow).Value
    ComboBox1.ListIndex = 0
    Frame2.ListBox1.List = Service.Sheets(1).Range("A1:B" & lastRow).Value
    For i = 0 To Frame2.ListBox1.ListCount - 1
        Frame2.ListBox1.List(i, 1) = " "
    Next
    'ListBox1.ListIndex = 0
    'ReDim originalAssetList(1 To lastRow)
    ReDim templist(1 To lastRow)
    For i = 1 To lastRow
        templist(i) = Service.Sheets(1).Range("A" & i).Value
    Next
    originalAssetList = templist
    'Service.Sheets(1).Range("A1:A" & lastRow).Value
End If
'ListBox1.ColumnCount = 2
    'ListBox1.addItem ""
    Frame2.ListBox1.List(0, 1) = "«Œÿ«—"
Call warningPages
End Sub
