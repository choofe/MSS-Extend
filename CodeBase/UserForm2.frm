VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Ê—Êœ «ÿ·«⁄«  Ê”Ì·Â ‰ﬁ·ÌÂ ÃœÌœ"
   ClientHeight    =   11985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15195
   OleObjectBlob   =   "UserForm2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Planning and organizing service routines for Car and Motorcycles
'Written and developed BY M.Amin Shafeie
'Started on 1400/04/10 (2021/07/01)
'PCSP (Pishgaman Saze Tejarat Pars)
'IMCC group IranMall 2021
Option Explicit
Const ITEM_NUMBER As Integer = 16
Dim Success As Boolean
Private Function dateCheck(Year As Integer, Month As Integer, Day As Integer) As Boolean
'checks the date entered is correct

If Year > 1420 Then
    MsgBox "”«· ò„ — «“ 1420 »«‘œ!"
    dateCheck = False
    Exit Function
End If
If Month > 6 And Month < 12 And Day > 30 Then
    MsgBox "—Ê“ —« ò‰ —· ‰„«ÌÌœ! ‘‘ „«Â œÊ„ ”«· ò„ — «“ 31 —Ê“ Â” ‰œ"
    dateCheck = False
    Exit Function
End If
If Month = 12 And Day > 29 Then
    MsgBox " «—ÌŒ —« ò‰ —· ò‰Ìœ! ¬Ì« Ê«ﬁ⁄« œ—  ⁄ÿÌ·«  ¬Œ— ”«·  ÕÊÌ· ’Ê—  ê—› Â «” ø"
    dateCheck = False
    Exit Function
End If
dateCheck = True
End Function

Private Sub CommandButton10_Click()
MsgBox (dateCheck(TextBox57.Value, TextBox56.Value, TextBox55.Value))

Dim gregorianDate As Date
Dim T As Date
gregorianDate = toGregorianDateObject(1400, 4, 21)
T = toGregorianDateObject(TextBox57.Value, TextBox56.Value, TextBox55.Value)
MsgBox "Date Object From Jalali= " & T

Dim jalaliDateArray
jalaliDateArray = toJalaali(2016, 3, 9)

'MsgBox jalaliDateArray(0) & "/" & jalaliDateArray(1) & "/" & jalaliDateArray(2)

Dim gregorialDateArray
gregorialDateArray = toGregorian(1394, 12, 19)

'MsgBox gregorialDateArray(0) & "/" & gregorialDateArray(1) & "/" & gregorialDateArray(2)

End Sub

'------------------------------------------------------------------------------------
Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    TextBox1.Enabled = True
    TextBox39.Enabled = True
    MotorOilType.Enabled = True
    Else
    TextBox1.Enabled = False
    TextBox39.Enabled = False
    MotorOilType.Enabled = False
End If
End Sub

Private Sub CheckBox10_Click()
If CheckBox10.Value = True Then
    TextBox10.Enabled = True
    TextBox48.Enabled = True
    Else
    TextBox10.Enabled = False
    TextBox48.Enabled = False
End If
End Sub

Private Sub CheckBox11_Click()
If CheckBox11.Value = True Then
    TextBox11.Enabled = True
    TextBox49.Enabled = True
    Else
    TextBox11.Enabled = False
    TextBox49.Enabled = False
End If
End Sub

Private Sub CheckBox12_Click()
If CheckBox12.Value = True Then
    TextBox12.Enabled = True
    TextBox50.Enabled = True
    Else
    TextBox12.Enabled = False
    TextBox50.Enabled = False
End If
End Sub

Private Sub CheckBox13_Click()
If CheckBox13.Value = True Then
    TextBox13.Enabled = True
    TextBox51.Enabled = True
    Else
    TextBox13.Enabled = False
    TextBox51.Enabled = False
End If
End Sub

Private Sub CheckBox14_Click()
If CheckBox14.Value = True Then
    TextBox14.Enabled = True
    TextBox52.Enabled = True
    GearBoxOilType.Enabled = True
    Else
    TextBox14.Enabled = False
    TextBox52.Enabled = False
    GearBoxOilType.Enabled = False
End If
End Sub

Private Sub CheckBox15_Click()
If CheckBox15.Value = True Then
    TextBox15.Enabled = True
    TextBox53.Enabled = True
    tyreType.Enabled = True
    Else
    TextBox15.Enabled = False
    TextBox53.Enabled = False
    tyreType.Enabled = False
End If
End Sub

Private Sub CheckBox16_Click()
If CheckBox16.Value = True Then
    TextBox16.Enabled = True
    TextBox54.Enabled = True
    batteryCap.Enabled = True
    Else
    TextBox16.Enabled = False
    TextBox54.Enabled = False
    batteryCap.Enabled = False
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
    TextBox2.Enabled = True
    TextBox40.Enabled = True
    Else
    TextBox2.Enabled = False
    TextBox40.Enabled = False
End If
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then
    TextBox3.Enabled = True
    TextBox41.Enabled = True
    Else
    TextBox3.Enabled = False
    TextBox41.Enabled = False
End If
End Sub

Private Sub CheckBox4_Click()
If CheckBox4.Value = True Then
    TextBox4.Enabled = True
    TextBox42.Enabled = True
    Else
    TextBox4.Enabled = False
    TextBox42.Enabled = False
End If
End Sub

Private Sub CheckBox5_Click()
If CheckBox5.Value = True Then
    TextBox5.Enabled = True
    TextBox43.Enabled = True
    Else
    TextBox5.Enabled = False
    TextBox43.Enabled = False
End If
End Sub

Private Sub CheckBox6_Click()
If CheckBox6.Value = True Then
    TextBox6.Enabled = True
    TextBox44.Enabled = True
    BrakeFluidType.Enabled = True
    Else
    TextBox6.Enabled = False
    TextBox44.Enabled = False
    BrakeFluidType.Enabled = False
End If
End Sub

Private Sub CheckBox7_Click()
If CheckBox7.Value = True Then
    TextBox7.Enabled = True
    TextBox45.Enabled = True
    HydraulicOilType.Enabled = True
    Else
    TextBox7.Enabled = False
    TextBox45.Enabled = False
    HydraulicOilType.Enabled = False
End If
End Sub

Private Sub CheckBox8_Click()
If CheckBox8.Value = True Then
    TextBox8.Enabled = True
    TextBox46.Enabled = True
    Else
    TextBox8.Enabled = False
    TextBox46.Enabled = False
End If
End Sub

Private Sub CheckBox9_Click()
If CheckBox9.Value = True Then
    TextBox9.Enabled = True
    TextBox47.Enabled = True
    Else
    TextBox9.Enabled = False
    TextBox47.Enabled = False
End If
End Sub
Private Sub CommandButton1_Click()
Dim i As Long
For i = 1 To ITEM_NUMBER
   If OptionButton1.Value = True Then Controls("CheckBox" & i).Locked = False
   If OptionButton1.Value = True Or Controls("CheckBox" & i).Locked = False Then Controls("CheckBox" & i).Value = True
Next

End Sub

Private Sub CommandButton2_Click()
Dim i As Long
For i = 1 To ITEM_NUMBER
   If OptionButton1.Value = True Then Controls("CheckBox" & i).Locked = False
   If OptionButton1.Value = True Or Controls("CheckBox" & i).Locked = False Then Controls("CheckBox" & i).Value = False
Next

End Sub

Private Sub CommandButton3_Click()
TextBox1.text = STD_MOTOR_OIL
TextBox2.text = STD_MOTOR_OIL_FILTER
TextBox3.text = STD_AIR_FILTER
TextBox4.text = STD_CABIN_AIR_FILTER
TextBox5.text = STD_COOLANT_FLUID
TextBox6.text = STD_BRAKE_FLUID
TextBox7.text = STD_HYDRAULIC_OIL
TextBox8.text = STD_SPARK_PLUG
TextBox9.text = STD_SPARK_WIRE
TextBox10.text = STD_CLUTCH
TextBox11.text = STD_FRONT_BRAKE_PAD
TextBox12.text = STD_REAR_BRAKE_PAD
TextBox13.text = STD_TIMING_BELT
TextBox14.text = STD_GEARBOX_OIL
TextBox15.text = STD_TYRES
TextBox16.text = STD_BATTERY

End Sub

Private Sub CommandButton4_Click()
Call PageExist(True)
End Sub
Private Function PageExist(controlButton As Boolean) As Boolean
Dim NewName As String
Dim exist As Boolean
exist = False
Dim i As Long
'create standard name based on wether motorcycle or car option is selected
If OptionButton1.Value = True Then
NewName = "ŒÊœ—Ê" & TextBox21.Value & TextBox22.Value & TextBox23.Value
Else
NewName = "„Ê Ê—" & TextBox22.Value & "-" & TextBox23.Value
End If


Dim answer As VbMsgBoxResult
'check all previously cereated pages to determine if the NewName is match with page name
For i = 1 To Sheets.Count
    'there will be a msgbox informing user that page alredy exist and if he/she would like to show content
    If Sheets(i).Name = NewName Then
        exist = True
        PageExist = True
        answer = MsgBox("Ìò »—êÂ »« ‘„«—Â Å·«ò " + NewName + " ÊÃÊœ œ«—œ. ¬Ì« „Ì ŒÊ«ÂÌœ »Â »—êÂ „Ê—œ ‰Ÿ— „‰ ﬁ· ‘ÊÌœø ", vbYesNo)
        If answer = vbYes Then
            Sheets(i).Activate
            Unload UserForm2
            'unload form!!!!!!unload app
        End If
        Exit For
    End If
Next
'at this point "exist" remained false and the page has not been created before
'controlButton reveals if user has clicked the check button (controlButton=true) or
'create page button (controlButton=false)
'if user had clicked control button he/she will recieve a "not found" magbox
'otherwise there won't be any massage and proccess will go on creating page

'user clicked check button:
If Not exist And controlButton Then
        MsgBox ("«Ì‰ Å·«ò ﬁ»·« À»  ‰‘œÂ")
        PageExist = False
End If
'user clicked create page button
If Not exist Then
    PageExist = False
End If

End Function
Private Sub CommandButton5_Click()

'----Plate Input Check
If OptionButton1.Value = True Then
    If TextBox20.text = "" Or _
       TextBox21.text = "" Or _
       TextBox22.text = "" Or _
       TextBox23.text = "" Then
       
        MsgBox "Å·«ò «‰ Ÿ«„Ì »Â ÿÊ— ò«„· Ê«—œ ‰‘œÂ «” "
    Exit Sub
    End If
    If Val(TextBox20.text) / 100 >= 10 Or _
       Val(TextBox21.text) / 100 >= 10 Or _
       Val(TextBox23.text) / 1000 >= 10 Or _
       IsNumeric(TextBox22.text) Then
       
        MsgBox "«ÿ·«⁄«  Ê—ÊœÌ Å·«ò —« ò‰ —· ò‰Ìœ"
        Exit Sub
    End If
    
Else
    If TextBox22.text = "" Or _
       TextBox23.text = "" Then
            MsgBox "Å·«ò «‰ Ÿ«„Ì »Â ÿÊ— ò«„· Ê«—œ ‰‘œÂ «” "
    Exit Sub
    End If
    If Val(TextBox22.text) / 100 >= 10 Or _
       Val(TextBox23.text) / 100000 >= 10 Then

        MsgBox "«ÿ·«⁄«  Ê—ÊœÌ Å·«ò —« ò‰ —· ò‰Ìœ"
        Exit Sub
    End If
End If


If TextBox57.Value <= 99 And TextBox57.Value > 95 Then TextBox57.Value = 1300 + TextBox57.Value
If TextBox57.Value < 95 And TextBox57.Value >= 0 Then TextBox57.Value = 1400 + TextBox57.Value
If TextBox57.Value < 1395 Then MsgBox "⁄œœ ”«· òÊçò «” "
If TextBox57.Value > 1495 Then MsgBox "⁄œœ ”«· »“—ê «” "

'checking if the plate no. is already exist
'if the page already exist, it'll ask user whether they want to go to corresponding sheet or not

If PageExist(False) Then Exit Sub

Dim i

'-------------------
'check if deliver kilometrage is entered
If TextBox38.Value = "" Then
    MsgBox "òÌ·Ê„ —«é  ÕÊÌ·Ì Ê«—œ ‰‘œÂ «” "
    Exit Sub
End If
For i = 30 To 37
    If Controls("Textbox" & i).text = "" Then
        MsgBox "ÌòÌ «“ „‘Œ’«  ŒÊœ—Ê Ê«—œ ‰‘œÂ «” !"
        Exit Sub
    End If
Next

'--------------------
'check if there is at least one item checked
Dim ChkBoxArray(1 To ITEM_NUMBER) As Control
Dim LeastChecked As Boolean
LeastChecked = False

For i = 1 To ITEM_NUMBER
    Set ChkBoxArray(i) = Controls("CheckBox" & i)
    If ChkBoxArray(i).Value Then LeastChecked = True
Next
If LeastChecked = False Then
    MsgBox ("ÂÌç ¬Ì „ ”—ÊÌ”Ì «‰ Œ«» ‰‘œÂ «” !")
    Exit Sub
End If
'-------------------
'cheking value ranges
Dim StandardValues(1 To ITEM_NUMBER) As Control
Dim OutOfRange As Boolean
OutOfRange = False

For i = 1 To ITEM_NUMBER
    Set StandardValues(i) = Controls("TextBox" & i)
  If StandardValues(i).Enabled = True Then
    If StandardValues(i).Value < 3500 Then
        With StandardValues(i)
            .SetFocus
            .ForeColor = RGB(200, 0, 0)
            .ControlTipText = "Too Low"
        End With
        OutOfRange = True
    End If
    If StandardValues(i).Value > 95000 Then
        With StandardValues(i)
            .SetFocus
            .ForeColor = RGB(200, 0, 0)
            .ControlTipText = "Too High"
        End With
        OutOfRange = True
    End If
    If StandardValues(i).Value >= 3500 And StandardValues(i).Value <= 95000 Then
        With StandardValues(i)
            .SetFocus
            .ForeColor = RGB(0, 0, 0)
            .ControlTipText = ""
        End With
    End If
  End If
Next
If OutOfRange Then
    MsgBox "Ìò Ì« ç‰œ „ﬁœ«— «” «‰œ«—œ œ— „ÕœÊœÂ „ﬁ«œÌ— ﬁ—«— ‰œ«—œ!"
    Exit Sub
End If
'------------------------------
'date input check
Dim transferDate As Date
Dim checkDate As Boolean
If TextBox55.text = "" And _
   TextBox56.text = "" And _
   TextBox57.text = "" Then
   transferDate = Date
Else
    If TextBox55.text = "" Or _
        TextBox56.text = "" Or _
        TextBox57.text = "" Then
            MsgBox " «—ÌŒ »Â ÿÊ— ò«„· Ê«—œ ‰‘œÂ «” "
            Exit Sub
    End If
End If
If IsNumeric(TextBox55.text) And _
   IsNumeric(TextBox56.text) And _
   IsNumeric(TextBox57.text) Then
    If dateCheck(TextBox57.Value, TextBox56.Value, TextBox55.Value) Then
        checkDate = True
    Else
    Exit Sub
End If
End If





If OptionButton1.Value = True Then
    CarPageCreation
Else
    MotorPageCreation
End If
If Success Then
'History.Sheets("RAW").Activate
'
'History.Sheets("RAW").Copy after:=History.Sheets(Sheets.Count)
'History.ActiveSheet.Name = Service.Sheets(Service.Sheets.Count).Name
'History.Sheets(Sheets.Count).Range("A9:A24").EntireRow.Delete
'    With History.ActiveSheet
'        .Range("A5").Value = Service.Sheets(Service.Sheets.Count).Range("A4").Value
'        .Range("D5").Value = Service.Sheets(Service.Sheets.Count).Range("D4").Value
'        .Range("E5").Value = Service.Sheets(Service.Sheets.Count).Range("E4").Value
'        .Range("F5").Value = Service.Sheets(Service.Sheets.Count).Range("F4").Value
'        .Range("G5").Value = Service.Sheets(Service.Sheets.Count).Range("G4").Value
'        .Range("H5").Value = Service.Sheets(Service.Sheets.Count).Range("H4").Value
'        .Range("I5").Value = Service.Sheets(Service.Sheets.Count).Range("I4").Value
'        .Range("B6").Value = Service.Sheets(Service.Sheets.Count).Range("B5").Value
'        .Range("G6").Value = Service.Sheets(Service.Sheets.Count).Range("G5").Value
'        .Range("C7").Value = Service.Sheets(Service.Sheets.Count).Range("C6").Value
'        .Range("E7").Value = Service.Sheets(Service.Sheets.Count).Range("E6").Value
'        .Range("G7").Value = Service.Sheets(Service.Sheets.Count).Range("G6").Value
'        .Range("I7").Value = Service.Sheets(Service.Sheets.Count).Range("I6").Value
'    End With
Dim lastRow As Integer
Dim sheetIndex
sheetIndex = Service.Sheets.Count
lastRow = findLastRow(Service.Sheets(sheetIndex), "B")
'creating a named range with "Status[sheetNo.]" format
'makes a break under the infos and puts a "history" header below the info page
    With Service.Sheets(sheetIndex)
        .Unprotect
        Service.Names.Add Name:="Status" & sheetIndex, RefersTo:=Service.Sheets(sheetIndex).Range("A1:" & "I" & lastRow)
        .Cells.PageBreak = xlNone
        .Rows(lastRow + 1).RowHeight = Application.CentimetersToPoints(0.3)
        .Rows(lastRow + 2).RowHeight = Application.CentimetersToPoints(0.3)
        .Rows(lastRow + 2).PageBreak = xlPageBreakManual
        .Range("A" & lastRow + 3 & ":I" & lastRow + 3).Merge
        .Range("A" & lastRow + 3 & ":I" & lastRow + 3).VerticalAlignment = xlCenter
        .Range("A" & lastRow + 3 & ":I" & lastRow + 3).HorizontalAlignment = xlCenter
        .Range("A" & lastRow + 3 & ":I" & lastRow + 3).Borders.LineStyle = xlContinuous
        .Range("A" & lastRow + 3 & ":I" & lastRow + 3).Value = " «—ÌŒçÂ"
        .Range("A" & lastRow + 3 & ":I" & lastRow + 3).Font.Bold = True
        .Range("A" & lastRow + 3 & ":I" & lastRow + 3).Font.Size = 15
        .Range("A" & lastRow + 3).Font.Name = "B Nazanin"
        .Range("A" & lastRow + 3).Interior.ColorIndex = 42
        .Protect
    End With

Unload Me
End If
Service.Sheets("Kilometrage").Activate
lastRow = Service.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row 'count of unique values
UserForm1.Frame2.ListBox1.List = Service.Sheets(1).Range("A1:A" & lastRow).Value
'UserForm1.ComboBox1.ListIndex = 0

End Sub
Private Sub CarPageCreation()
Dim NewName As String
NewName = "ŒÊœ—Ê" & TextBox21.Value & TextBox22.Value & TextBox23.Value
'---------------
'Copying from raw sheet
Service.Activate
Service.Sheets("raw").Copy after:=Sheets(Sheets.Count)
Dim sheetIndex
sheetIndex = Service.Sheets.Count
ActiveWindow.View = xlPageLayoutView
Sheets(sheetIndex).Visible = False
Sheets(1).Activate
With Service.Sheets(sheetIndex)
    .Unprotect
    .Visible = False
    .Name = NewName
    .DisplayRightToLeft = True
Dim counter

For counter = 65 To 73
    .Range(Chr(counter) & "1").EntireColumn.ColumnWidth = Sheets("RAW").Columns(Chr(counter)).ColumnWidth
Next

With .PageSetup
 .PaperSize = 9
 .LeftMargin = Application.CentimetersToPoints(0.7)
 .RightMargin = Application.CentimetersToPoints(0.75)
 .TopMargin = Application.CentimetersToPoints(1.4)
 .BottomMargin = Application.CentimetersToPoints(1.9)
 .HeaderMargin = Application.CentimetersToPoints(0.8)
 .FooterMargin = Application.CentimetersToPoints(0.8)
 .CenterHorizontally = True
End With

.Range("G4").Value = TextBox20.text 'Iran Number of plate
.Range("C4").Value = TextBox21.text '2 digit Number
.Range("D4").Value = TextBox22.text 'single character of plate
.Range("E4").Value = TextBox23.text '3 digit number
.Range("A4").Value = TextBox30.text 'Car name (model)
.Range("C6").Value = TextBox31.text 'Manufacture year
.Range("B5").Value = TextBox32.text 'Motor serial
.Range("G5").Value = TextBox33.text 'Chasis number
.Range("E6").Value = TextBox34.text 'Color
.Range("H4").Value = TextBox35.text 'Driver name
.Range("I4").Value = TextBox36.text 'driver phone
.Range("G6").Value = TextBox37.text 'Department
.Range("A8").Value = MotorOilType.text
.Range("E8").Value = BrakeFluidType.text
.Range("F8").Value = HydraulicOilType.text
.Range("C8").Value = GearBoxOilType.text
.Range("H8").Value = tyreType.text
.Range("G8").Value = batteryCap.text

Dim transferDate As Date
If TextBox55.text = "" And _
   TextBox56.text = "" And _
   TextBox57.text = "" Then
   transferDate = Date
   Else
transferDate = toGregorianDateObject(TextBox57.Value, TextBox56.Value, TextBox55.Value)
End If
'setting the delivery date
.Range("I6").Value = transferDate

Dim T, i
For T = STARTING_ROW To ITEM_NUMBER + STARTING_ROW - 1
    .Range("D" & T).Value = Controls("TextBox" & T - STARTING_ROW + 1).Value
Next

Dim InitialValues As Control

For i = ITEM_NUMBER + STARTING_ROW - 1 To STARTING_ROW Step -1
    If Not (Controls("CheckBox" & (i - STARTING_ROW + 1))) Then
        .Rows(i).Delete
    Else
        Set InitialValues = Controls("textbox" & i + 27) 'textbox names start at 34 and end in 54
        .Range("E" & i).Value = TextBox38.Value - InitialValues.Value
    End If
Next

    .Range("E" & STARTING_ROW).Value = "=Kilometrage!B" & Sheets.Count - 1
    .Protect
End With

With Service.Sheets("Kilometrage")
    .Unprotect
    .Range("A" & Sheets.Count - 1).Value = NewName
    .Range("B" & Sheets.Count - 1).Value = TextBox38.Value
    .Range("D" & Sheets.Count - 1).Value = transferDate
    .Range("E" & Sheets.Count - 1).Value = TextBox38.Value
    .Protect
End With

eraseRawExtraRows (Service.Sheets.Count)

MsgBox "’›ÕÂ „Ê—œ ‰Ÿ— »« „Ê›ﬁÌ  «ÌÃ«œ ‘œ"

Success = True

End Sub
Private Sub MotorPageCreation()
Dim NewName As String
NewName = "„Ê Ê—" & TextBox22.Value & "-" & TextBox23.Value
Service.Activate
Service.Sheets("raw").Copy after:=Sheets(Sheets.Count)
Dim sheetIndex
sheetIndex = Service.Sheets.Count
ActiveWindow.View = xlPageLayoutView
Sheets(1).Activate
With Service.Sheets(sheetIndex)
    .Unprotect
    .Visible = False
    .Name = NewName
    .DisplayRightToLeft = True
Dim counter

For counter = 65 To 73
    .Range(Chr(counter) & "1").EntireColumn.ColumnWidth = Sheets("RAW").Columns(Chr(counter)).ColumnWidth
Next

With .PageSetup
 .PaperSize = 9
 .LeftMargin = Application.CentimetersToPoints(0.7)
 .RightMargin = Application.CentimetersToPoints(0.75)
 .TopMargin = Application.CentimetersToPoints(1.4)
 .BottomMargin = Application.CentimetersToPoints(1.9)
 .HeaderMargin = Application.CentimetersToPoints(0.8)
 .FooterMargin = Application.CentimetersToPoints(0.8)
 .CenterHorizontally = True
End With



.Range("C4:D4").ClearContents
.Range("C4:D4").Merge
.Range("C4").Value = TextBox23.text '5 digit number
.Range("F4:G4").ClearContents
.Range("F4:G4").Merge
.Range("E4").ClearContents
.Range("E4").Value = "«Ì—«‰"
.Range("F4").Value = TextBox22.text '2 digit Number
.Range("A4").Value = TextBox30.text
.Range("C6").Value = TextBox31.text
.Range("B5").Value = TextBox32.text
.Range("G5").Value = TextBox33.text
.Range("E6").Value = TextBox34.text
.Range("H4").Value = TextBox35.text
.Range("I4").Value = TextBox36.text
.Range("I6").Value = TextBox37.text
.Range("A8").Value = MotorOilType.text
.Range("E8").Value = BrakeFluidType.text
.Range("H8").Value = tyreType.text
.Range("G8").Value = batteryCap.text
.Range("F8:F9").Interior.ColorIndex = 15
.Range("F8:F9").Font.ColorIndex = 2
.Range("C8:C9").Interior.ColorIndex = 15
.Range("C8:C9").Font.ColorIndex = 2

Dim transferDate As Date
If TextBox55.text = "" And _
   TextBox56.text = "" And _
   TextBox57.text = "" Then
   transferDate = Date
   Else
transferDate = toGregorianDateObject(TextBox57.Value, TextBox56.Value, TextBox55.Value)
End If
'setting the delivery date
.Range("I6").Value = transferDate
Dim T, i
For T = STARTING_ROW To ITEM_NUMBER + STARTING_ROW - 1
    .Range("D" & T).Value = Controls("TextBox" & T - STARTING_ROW + 1).Value
Next

Dim InitialValues As Control

For i = ITEM_NUMBER + STARTING_ROW - 1 To STARTING_ROW Step -1
    If Not (Controls("CheckBox" & (i - STARTING_ROW + 1))) Or (Controls("CheckBox" & (i - STARTING_ROW + 1)).Locked) Then
        .Rows(i).Delete
    Else
        Set InitialValues = Controls("textbox" & i + 27) 'textbox names start at 34 and end in 54
        Range("E" & i).Value = TextBox38.Value - InitialValues.Value
    End If
Next

    .Range("E" & STARTING_ROW).Value = "=Kilometrage!B" & Sheets.Count - 1
    .Protect
End With

With Service.Sheets("Kilometrage")
    .Unprotect
    .Range("A" & Sheets.Count - 1).Value = NewName
    .Range("B" & Sheets.Count - 1).Value = TextBox38.Value
    .Range("D" & Sheets.Count - 1).Value = transferDate
    .Range("E" & Sheets.Count - 1).Value = TextBox38.Value
    .Protect
End With
eraseRawExtraRows (Service.Sheets.Count)

MsgBox "’›ÕÂ „Ê—œ ‰Ÿ— »« „Ê›ﬁÌ  «ÌÃ«œ ‘œ"

Success = True

End Sub
Private Sub CommandButton6_Click()
Unload Me
End Sub


Private Sub CommandButton7_Click()
If TextBox57.Value <= 99 And TextBox57.Value > 95 Then TextBox57.Value = 1300 + TextBox57.Value
If TextBox57.Value < 95 And TextBox57.Value >= 0 Then TextBox57.Value = 1400 + TextBox57.Value
If TextBox57.Value < 1395 Then MsgBox "⁄œœ ”«· òÊçò «” "
If TextBox57.Value > 1495 Then MsgBox "⁄œœ ”«· »“—ê «” "


Dim transferDate As Date
Dim checkDate As Boolean
If TextBox55.text = "" And _
   TextBox56.text = "" And _
   TextBox57.text = "" Then
   transferDate = Date
Else
    If TextBox55.text = "" Or _
        TextBox56.text = "" Or _
        TextBox57.text = "" Then
            MsgBox " «—ÌŒ »Â ÿÊ— ò«„· Ê«—œ ‰‘œÂ «” "
            Exit Sub
    End If
End If
If IsNumeric(TextBox55.text) And _
   IsNumeric(TextBox56.text) And _
   IsNumeric(TextBox57.text) Then
    If dateCheck(TextBox57.Value, TextBox56.Value, TextBox55.Value) Then
        checkDate = True
    Else
    Exit Sub
End If
End If





If OptionButton1.Value = True Then
    CarPageCreation
Else
    MotorPageCreation
End If
End Sub

Private Sub CommandButton8_Click()
Dim datePicked() As Integer
datePicked = datePicking()
TextBox57.text = datePicked(3)
TextBox56.text = datePicked(2)
TextBox55.text = datePicked(1)
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub OptionButton1_Click()
TextBox20.Visible = True
TextBox21.Visible = True
Label6.Visible = True
Label8.Visible = True

With Label6
    .Top = 84
    .Left = 240
End With

Dim i As Integer
For i = 1 To 16
    Controls("TextBox" & i).Locked = False
    Controls("CheckBox" & i).Locked = False
    Controls("TextBox" & i + 38).Locked = False
    Controls("CheckBox" & i).Enabled = True
    If Controls("CheckBox" & i).Value Then
        Controls("TextBox" & i).Enabled = True
        Controls("TextBox" & i + 38).Enabled = True
    End If
Next
End Sub

Private Sub OptionButton2_Click()
TextBox20.Visible = False
TextBox21.Visible = False
'Label6.Visible = False
Label8.Visible = False
TextBox22.SetFocus
'------------

With Label6
    .Left = 170
    .Top = 96
End With

With TextBox3
    .Enabled = False
    .Locked = True
End With
With CheckBox3
    .Enabled = False
    .Locked = True
End With
With TextBox41
    .Enabled = False
    .Locked = True
End With
'-------------
With TextBox4
    .Enabled = False
    .Locked = True
End With
With CheckBox4
    .Enabled = False
    .Locked = True
End With
With TextBox42
    .Enabled = False
    .Locked = True
End With
'-------------
With TextBox5
    .Enabled = False
    .Locked = True
End With
With CheckBox5
    .Enabled = False
    .Locked = True
End With
With TextBox43
    .Enabled = False
    .Locked = True
End With
'------------
'With TextBox6
'    .Enabled = False
'    .Locked = True
'End With
'With CheckBox6
'    .Value = False
'    .Enabled = False
'    .Locked = True
'End With
'With TextBox44
'    .Enabled = False
'    .Locked = True
'End With
'With BrakeFluidType
'    .Enabled = False
'    .Locked = True
'End With

'-------------
With TextBox7
    .Enabled = False
    .Locked = True
End With
With CheckBox7
    .Enabled = False
    .Locked = True
End With
With TextBox45
    .Enabled = False
    .Locked = True
End With
'-------------
With TextBox13
    .Enabled = False
    .Locked = True
End With
With CheckBox13
    .Enabled = False
    .Locked = True
End With
With TextBox51
    .Enabled = False
    .Locked = True
End With
'-------------
With TextBox14
    .Enabled = False
    .Locked = True
End With
With CheckBox14
    .Enabled = False
    .Locked = True
End With
With TextBox52
    .Enabled = False
    .Locked = True
End With
'-------------
With TextBox10
    .Enabled = False
    .Locked = True
End With
With CheckBox10
    .Enabled = False
    .Locked = True
End With
With TextBox48
    .Enabled = False
    .Locked = True
End With

End Sub

Private Sub TextBox1_Change()

End Sub
Private Sub CheckValues(ByRef TxtBox As TextBox)
If TxtBox.Value < 3500 Then
    With TxtBox
        .SetFocus
        .ForeColor = RGB(200, 0, 0)
        .ControlTipText = "Too Low"
    End With
    Exit Sub
End If
If TextBox1.Value > 95000 Then
    MsgBox "òÌ·Ê„ — ò«—ò—œ »”Ì«— “Ì«œ «” "
    With TextBox1
        .SetFocus
        .ForeColor = RGB(200, 0, 0)
        .ControlTipText = "Too High"
    End With
End If
If TextBox1.Value >= 3500 And TextBox1.Value <= 95000 Then
    With TextBox1
        .SetFocus
        .ForeColor = RGB(0, 0, 0)
    End With
End If

End Sub
Private Sub TextBox20_Change()
If OptionButton1.Value = True And Len(TextBox20.Value) = 2 Then TextBox21.SetFocus
End Sub

Private Sub TextBox21_Change()
If OptionButton1.Value = True And Len(TextBox21.Value) = 2 Then TextBox22.SetFocus

End Sub

Private Sub TextBox22_Change()
If OptionButton1.Value = True And Len(TextBox22.Value) = 1 Then TextBox23.SetFocus Else If OptionButton2.Value = True And Len(TextBox22.Value) = 3 Then TextBox23.SetFocus
End Sub
Private Sub TextBox55_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox55.Value > 31 Then TextBox55.text = 31
End Sub

Private Sub TextBox56_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If TextBox56.Value > 12 Then TextBox56.text = 12
End Sub

Private Sub TextBox57_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If TextBox57.Value <= 99 And TextBox57.Value > 95 Then TextBox57.Value = 1300 + TextBox57.Value
If TextBox57.Value < 95 And TextBox57.Value >= 0 Then TextBox57.Value = 1400 + TextBox57.Value
If TextBox57.Value < 1395 Then MsgBox "⁄œœ ”«· òÊçò «” "
If TextBox57.Value > 1495 Then MsgBox "⁄œœ ”«· »“—ê «” "

End Sub
Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
ActiveSheet.DisplayRightToLeft = True

Dim todayDate As Date: todayDate = Date
Dim jalaliToday As Variant
jalaliToday = toJalaaliFromDateObject(todayDate)
TextBox55.Value = jalaliToday(2)
TextBox56.Value = jalaliToday(1)
TextBox57.Value = jalaliToday(0)

End Sub
