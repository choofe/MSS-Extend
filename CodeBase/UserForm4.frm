VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "ÇäÊÎÇÈ ÊÇÑíÎ"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBox1_Click()
'On Error Resume Next

globalJDate(0) = ComboBox1.Value
Dim gYear, gMonth, gDay
gYear = globalJDate(0)
gMonth = globalJDate(1)
gDay = globalJDate(2)
Call updateMonthView(gYear, gMonth, gDay)
errHandling:

End Sub

Private Sub ComboBox2_Change()
globalJDate(1) = ComboBox2.ListIndex + 1
Dim gYear, gMonth, gDay
gYear = globalJDate(0)
gMonth = globalJDate(1)
gDay = globalJDate(2)
Call updateMonthView(gYear, gMonth, gDay)

End Sub

Private Sub ComboBox2_Click()
End Sub

Private Sub CommandButton1_Click()
globalJDate = toJalaaliFromDateObject(Date)
Dim gYear, gMonth, gDay
gYear = globalJDate(0)
gMonth = globalJDate(1)
gDay = globalJDate(2)
ComboBox1.ListIndex = gYear - 1399
ComboBox2.ListIndex = gMonth - 1
Call updateMonthView(gYear, gMonth, gDay)
End Sub

Private Sub CommandButton2_Click()
If ComboBox2.ListIndex = 11 Then
    ComboBox2.ListIndex = 0
    ComboBox1.ListIndex = ComboBox1.ListIndex + 1
Else
    ComboBox2.ListIndex = ComboBox2.ListIndex + 1
End If
End Sub
Private Sub CommandButton3_Click()
If ComboBox2.ListIndex = 0 Then
    ComboBox2.ListIndex = 11
    ComboBox1.ListIndex = ComboBox1.ListIndex - 1
Else
    ComboBox2.ListIndex = ComboBox2.ListIndex - 1
End If
End Sub

Private Sub CommandButton4_Click()
ComboBox1.ListIndex = ComboBox1.ListIndex + 1

End Sub
Private Sub CommandButton5_Click()
ComboBox1.ListIndex = ComboBox1.ListIndex - 1
End Sub

Private Sub updateMonthView(ByVal gYear As Long, ByVal gMonth As Long, ByVal gDay As Long)
For i = 1 To 35
    Controls("label" & i).Caption = ""
    Controls("label" & i).BorderColor = RGB(0, 0, 0)
Next
Dim todayJDate As Variant
todayJDate = toJalaaliFromDateObject(Date)
Select Case gMonth
    Case 1: Label43.Caption = "ÝÑæÑÏíä"
    Case 2: Label43.Caption = "ÇÑÏíÈåÔÊ"
    Case 3: Label43.Caption = "ÎÑÏÇÏ"
    Case 4: Label43.Caption = "ÊíÑ"
    Case 5: Label43.Caption = "ãÑÏÇÏ"
    Case 6: Label43.Caption = "ÔåÑíæÑ"
    Case 7: Label43.Caption = "ãåÑ"
    Case 8: Label43.Caption = "ÂÈÇä"
    Case 9: Label43.Caption = "ÂÐÑ"
    Case 10: Label43.Caption = "Ïí"
    Case 11: Label43.Caption = "Èåãä"
    Case 12: Label43.Caption = "ÇÓÝäÏ"
End Select
Label44.Caption = gYear
Dim monthDayCount As Long
Dim tempDate As Date
'------------
'findig month days count
tempDate = toGregorianDateObject(gYear, gMonth, 1)
'with the current month first day georgian date is calculated
'by adding 30 days to this date the resulted date will be calculated to jalali date
'with day result of the calculated date month day number is determined.
'if the day is 31 then month has 31 days , if the day is 1 it is the first day of next month and
'current month has 30 days. this way id the day is 2 the month is 29 days
'with this trick there is no worry about leap year and month Esfand 30 days

If (toJalaaliFromDateObject(tempDate + 30)(2)) = 31 Then monthDayCount = 31 Else
If (toJalaaliFromDateObject(tempDate + 30)(2)) = 1 Then monthDayCount = 30 Else
If (toJalaaliFromDateObject(tempDate + 30)(2)) = 2 Then monthDayCount = 29
'---------------------
Dim weekDayShift As Integer
Dim newGdate As Date
newGdate = toGregorianDateObject(gYear, gMonth, gDay)

weekDayShift = Weekday(newGdate - toJalaaliFromDateObject(newGdate)(2)) + 1
If weekDayShift >= 7 Then weekDayShift = weekDayShift - 7

For i = 1 To monthDayCount
    If (i + weekDayShift) <= 35 Then
    Controls("label" & i + weekDayShift).Caption = i
    If todayJDate(0) = gYear And todayJDate(1) = gMonth And (i - weekDayShift) = todayJDate(2) _
        Then Controls("label" & i).BorderColor = RGB(250, 0, 0)
    Else
    Controls("label" & i - 35 + weekDayShift).Caption = i
    If todayJDate(0) = gYear And todayJDate(1) = gMonth And (i - weekDayShift) = todayJDate(2) _
        Then Controls("label" & i).BorderColor = RGB(250, 0, 0)
    End If
Next

Dim JalaaliDate
Dim GeorgianDate As Date
GeorgianDate = toGregorianDateObject(gYear, gMonth, gDay)
JalaaliDate = toJalaaliFromDateObject(GeorgianDate)
TextBox1.text = JalaaliDate(0) & "/" & JalaaliDate(1) & "/" & JalaaliDate(2)

End Sub

Private Sub CommandButton6_Click()

globalSelectedDate = TextBox1.text
Unload Me
End Sub

Private Sub TextBox1_Change()
globalSelectedDate = TextBox1.text

End Sub

Private Sub UserForm_Deactivate()
End Sub

Private Sub UserForm_Initialize()
globalJDate = toJalaaliFromDateObject(Date)
gYear = globalJDate(0)
gMonth = globalJDate(1)
gDay = globalJDate(2)

For i = 1399 To 1415
    ComboBox1.addItem (i)
Next
ComboBox1.ListIndex = gYear - 1399

ComboBox2.ListRows = 12
ComboBox2.addItem ("ÝÑæÑÏíä")
ComboBox2.addItem ("ÇÑÏíÈåÔÊ")
ComboBox2.addItem ("ÎÑÏÇÏ")
ComboBox2.addItem ("ÊíÑ")
ComboBox2.addItem ("ãÑÏÇÏ")
ComboBox2.addItem ("ÔåÑíæÑ")
ComboBox2.addItem ("ãåÑ")
ComboBox2.addItem ("ÂÈÇä")
ComboBox2.addItem ("ÂÐÑ")
ComboBox2.addItem ("Ïí")
ComboBox2.addItem ("Èåãä")
ComboBox2.addItem ("ÇÓÝäÏ")
ComboBox2.ListIndex = gMonth - 1

Call updateMonthView(gYear, gMonth, gDay)
globalSelectedDate = TextBox1.text

End Sub
Private Sub Label1_Click()
If Label1.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label1.Caption))
End If
End Sub
Private Sub Label2_Click()
If Label2.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label2.Caption))
End If
End Sub
Private Sub Label3_Click()
If Label3.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label3.Caption))
End If
End Sub
Private Sub Label4_Click()
If Label4.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label4.Caption))
End If
End Sub
Private Sub Label5_Click()
If Label5.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label5.Caption))
End If
End Sub
Private Sub Label6_Click()
If Label6.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label6.Caption))
End If
End Sub
Private Sub Label7_Click()
If Label8.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label7.Caption))
End If
End Sub
Private Sub Label8_Click()
If Label8.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label8.Caption))
End If
End Sub
Private Sub Label9_Click()
If Label9.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label9.Caption))
End If
End Sub
Private Sub Label10_Click()
If Label10.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label10.Caption))
End If
End Sub
Private Sub Label11_Click()
If Label11.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label11.Caption))
End If
End Sub
Private Sub Label12_Click()
If Label12.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label12.Caption))
End If
End Sub
Private Sub Label13_Click()
If Label13.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label13.Caption))
End If
End Sub
Private Sub Label14_Click()
If Label14.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label14.Caption))
End If
End Sub
Private Sub Label15_Click()
If Label15.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label15.Caption))
End If
End Sub
Private Sub Label16_Click()
If Label16.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label16.Caption))
End If
End Sub
Private Sub Label17_Click()
If Label17.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label17.Caption))
End If
End Sub
Private Sub Label18_Click()
If Label18.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label18.Caption))
End If
End Sub
Private Sub Label19_Click()
If Label19.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label19.Caption))
End If
End Sub
Private Sub Label20_Click()
If Label20.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label20.Caption))
End If
End Sub
Private Sub Label21_Click()
If Label21.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label21.Caption))
End If
End Sub
Private Sub Label22_Click()
If Label22.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label22.Caption))
End If
End Sub
Private Sub Label23_Click()
If Label23.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label23.Caption))
End If
End Sub
Private Sub Label24_Click()
If Label24.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label24.Caption))
End If
End Sub
Private Sub Label25_Click()
If Label25.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label25.Caption))
End If
End Sub
Private Sub Label26_Click()
If Label26.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label26.Caption))
End If
End Sub
Private Sub Label27_Click()
If Label27.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label27.Caption))
End If
End Sub
Private Sub Label28_Click()
If Label28.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label28.Caption))
End If
End Sub
Private Sub Label29_Click()
If Label29.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label29.Caption))
End If
End Sub
Private Sub Label30_Click()
If Label30.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label30.Caption))
End If
End Sub
Private Sub Label31_Click()
If Label31.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label31.Caption))
End If
End Sub
Private Sub Label32_Click()
If Label32.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label32.Caption))
End If
End Sub
Private Sub Label33_Click()
If Label33.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label33.Caption))
End If
End Sub
Private Sub Label34_Click()
If Label34.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label34.Caption))
End If
End Sub
Private Sub Label35_Click()
If Label35.Caption <> "" Then
Call updateMonthView(globalJDate(0), globalJDate(1), Int(Label35.Caption))
End If
End Sub


