Attribute VB_Name = "Module1"
Option Explicit
Private Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1
Public Const ITEM_NUMBER = 16
Public globalJDate As Variant
Public globalSelectedDate As Variant
Public History As Workbook
Public Service As Workbook
Public itemCounter As Integer
Global collectionEvent As Collection
Public originalAssetList()
Global Const STARTING_ROW As Integer = 12

Global Const STD_MOTOR_OIL As String = "6000"
Global Const STD_MOTOR_OIL_FILTER As String = "12000"
Global Const STD_AIR_FILTER As String = "6000"
Global Const STD_CABIN_AIR_FILTER As String = "30000"
Global Const STD_COOLANT_FLUID As String = "60000"
Global Const STD_BRAKE_FLUID As String = "40000"
Global Const STD_HYDRAULIC_OIL As String = "45000"
Global Const STD_SPARK_PLUG As String = "75000"
Global Const STD_SPARK_WIRE As String = "75000"
Global Const STD_CLUTCH As String = "75000"
Global Const STD_FRONT_BRAKE_PAD As String = "65000"
Global Const STD_REAR_BRAKE_PAD As String = "80000"
Global Const STD_TIMING_BELT As String = "60000"
Global Const STD_GEARBOX_OIL As String = "65000"
Global Const STD_TYRES As String = "80000"
Global Const STD_BATTERY As String = "60000"

'tabdil as date object be arraye jalali
 Function toJalaaliFromDateObject(gDate As Date)
   toJalaaliFromDateObject = toJalaali(Year(gDate), Month(gDate), Day(gDate))
End Function
' tabdild to object date gregorian
 Function toGregorianDateObject(jy As Long, jm As Long, jd As Long)
    Dim result
    result = toGregorian(jy, jm, jd)
    toGregorianDateObject = DateValue(result(0) & "-" & result(1) & "-" & result(2))
End Function
'tabdil jalali be miladi ba daryaf sale , mahe , rooz jalali
' yek arraye barmigardanad ke index(0)= sal, index(1)=mah, index(2)=rooz
 Function toGregorian(jy As Long, jm As Long, jd As Long)
      toGregorian = d2g(j2d(jy, jm, jd))
End Function
Function getScreenX() As Long
getScreenX = GetSystemMetrics(SM_CXSCREEN)
End Function
Function getScreenY() As Long
getScreenY = GetSystemMetrics(SM_CYSCREEN)
End Function


' tabdile tarikh miladi be jalali ba daryaft sale, mah , rooz miladi
' yek arraye barmigardanad ke index(0)= sal, index(1)=mah, index(2)=rooz
 Function toJalaali(gy As Long, gm As Long, gd As Long)
    toJalaali = d2j(g2d(gy, gm, gd))
End Function
  
' check valid bodan jalali date
 Function isValidJalaaliDate(jy As Long, jm As Long, jd As Long)
    isValidJalaaliDate = jy >= -61 And jy <= 3177 And jm >= 1 And jm <= 12 And jd >= 1 And jd <= jalaaliMonthLength(jy, jm)
End Function
  

' tedade rooz haye mah ra baraye sale , mahe jalali bar migardanad
 Function jalaaliMonthLength(jy As Long, jm As Long)
    If (jm <= 6) Then
        jalaaliMonthLength = 31
        Exit Function
    End If
    If (jm <= 11) Then
     jalaaliMonthLength = 30
     Exit Function
    End If
    If (isLeapJalaaliYear(jy)) Then
        jalaaliMonthLength = 30
        Exit Function
        
    End If
  jalaaliMonthLength = 29
End Function
' check in ke sale jalali kabise as ya na
 Function isLeapJalaaliYear(jy As Long)
    Dim leap As Long
    leap = jalCal(jy)(0)
    
    If (leap = 0) Then
        isLeapJalaaliYear = True
    Else
        isLeapJalaaliYear = False
    End If
    
End Function

' function haye paeeen baraye amaliate dakheli ast va nabayad estefade shavad
 Function j2d(jy As Long, jm As Long, jd As Long)
    Dim r As Long
    Dim rgy As Long
    Dim rmarch As Long
    
    rgy = jalCal(jy)(1)
    rmarch = jalCal(jy)(2)
    j2d = g2d(rgy, 3, rmarch) + ((jm - 1) * 31) - ((jm \ 7) * (jm - 7)) + jd - 1
    
End Function

 Function d2j(jdn As Long)
    Dim gy As Long
    gy = d2g(jdn)(0) ' Calculate Gregorian year (gy)
    Dim jy As Long
    jy = gy - 621
    Dim rmarch  As Long
    jalCal (jy)
    rmarch = jalCal(jy)(2)
    Dim rleap  As Long
    rleap = jalCal(jy)(0)
    
    Dim jdn1f As Long
    jdn1f = g2d(gy, 3, rmarch)  'r.march
    Dim jd As Long
    Dim jm As Long
    Dim k As Long

    ' Find number of days that passed since 1 Farvardin.
    k = jdn - jdn1f
    
    Dim result(3)
     
    If (k >= 0) Then
        If (k <= 185) Then
          ' The first 6 months.
          jm = 1 + (k \ 31)
          jd = (k Mod 31) + 1
          result(0) = jy
          result(1) = jm
          result(2) = jd
             
          d2j = result
          Exit Function
          
          
        Else
          ' The remaining months.
          k = k - 186
        End If
    Else
        ' Previous Jalaali year.
        jy = jy - 1
        k = k + 179
        If (rleap = 1) Then 'r.leap
          k = k + 1
        End If
    End If
    
    
    jm = 7 + (k \ 30)
    jd = (k Mod 30) + 1
    
    result(0) = jy
    result(1) = jm
    result(2) = jd
             
    d2j = result
    
End Function


 Function d2g(jdn As Long)
    Dim j As Long
    Dim i As Long
    Dim gd As Long
    Dim gm As Long
    Dim gy As Long
    j = 4 * jdn + 139361631
    j = j + (((((4 * jdn + 183187720) \ 146097) * 3) \ 4) * 4) - 3908
    
    i = (((j Mod 1461) \ 4) * 5) + 308
    gd = ((i Mod 153) \ 5) + 1
    gm = ((i \ 153) Mod 12) + 1
    gy = (j \ 1461) - 100100 + ((8 - gm) \ 6)
    
    Dim result(3)

    result(0) = gy
    result(1) = gm
    result(2) = gd
  
    d2g = result

End Function



 Function g2d(gy As Long, gm As Long, gd As Long)

    Dim d As Long
    d = (((gy + ((gm - 8) \ 6) + 100100) * 1461) \ 4) + ((153 * ((gm + 9) Mod 12) + 2) \ 5) + gd - 34840408
    d = d - ((((gy + 100100 + ((gm - 8) \ 6)) \ 100) * 3) \ 4) + 752
    g2d = d
End Function
 Function jalCal(jy As Long)

    Dim breaks
    breaks = Array(-61, 9, 38, 199, 426, 686, 756, 818, 1111, 1181, 1210, 1635, 2060, 2097, 2192, 2262, 2324, 2394, 2456, 3178)

    Dim bl As Long
    bl = 20
    Dim gy As Long
    
    gy = jy + 621
    Dim leapJ  As Long
    leapJ = -14
    Dim jp As Long
    jp = breaks(0)
    Dim jm As Long
    Dim jump As Long
    Dim leap As Long
    Dim leapG As Long
    Dim march As Long
    Dim n As Long
    Dim i As Long
    

    If (jy < jp Or jy >= breaks(bl - 1)) Then
        MsgBox "Invalid Jalaali year " & jy
    End If
   

   'Find the limiting years for the Jalaali year jy.
   For i = 1 To (bl - 1) Step 1
        jm = breaks(i)
        jump = jm - jp
        If (jy < jm) Then Exit For
        
        leapJ = leapJ + (jump \ 33) * 8 + ((jump Mod 33) \ 4)
        jp = jm
   Next
   
  
   n = jy - jp

  ' Find the number of leap years from AD 621 to the beginning
  ' of the current Jalaali year in the Persian calendar.
  
  leapJ = leapJ + (n \ 33) * 8 + (((n Mod 33) + 3) \ 4)
  If ((jump Mod 33) = 4 And (jump - n) = 4) Then
    leapJ = leapJ + 1
  End If

  ' And the same in the Gregorian calendar (until the year gy).
  leapG = (gy \ 4) - ((((gy \ 100) + 1) * 3) \ 4) - 150

  ' Determine the Gregorian date of Farvardin the 1st.
  march = 20 + leapJ - leapG

  ' Find how many years have passed since the last leap year.
  If ((jump - n) < 6) Then
    n = n - jump + ((jump + 4) \ 33) * 33
  End If
  
  leap = ((((n + 1) Mod 33) - 1) Mod 4)
  If (leap = -1) Then
    leap = 4
  End If
  
  Dim result(3)

  result(0) = leap
  result(1) = gy
  result(2) = march
  
  jalCal = result


End Function
Sub checklistMaker()

Dim tName2 As String
Dim sheetCount As Long
sheetCount = ThisWorkbook.Sheets.Count
tName2 = "Temp" & Trim("checklist" & Replace(Time, ":", "", 1))

'ThisWorkbook.ActiveSheet.Copy after:=Sheets(Sheets.Count)
Sheets.Add(after:=Sheets(Sheets.Count)).Name = tName2
ActiveSheet.Name = tName2
ActiveSheet.Unprotect
ActiveSheet.DisplayRightToLeft = True
Dim counter As Integer
For counter = 65 To 73
  'iColumnWidth =
    ActiveSheet.Range(Chr(counter) & "1").EntireColumn.ColumnWidth = Sheets(2).Columns(Chr(counter)).ColumnWidth
Next
'Setting the margin
With ActiveSheet.PageSetup
 .LeftMargin = Application.CentimetersToPoints(0.7)
 .RightMargin = Application.CentimetersToPoints(0.75)
 .TopMargin = Application.CentimetersToPoints(1.4)
 .BottomMargin = Application.CentimetersToPoints(1.9)
 .HeaderMargin = Application.CentimetersToPoints(0.8)
 .FooterMargin = Application.CentimetersToPoints(0.8)
 .CenterHorizontally = True
End With
'-------------------

ActiveWindow.View = xlPageLayoutView
Dim rowCount As Long
Dim prevRowCount As Long
prevRowCount = 1

'Dim alarmTrig(1) As Integer
ReDim alarmTrig(sheetCount)
Dim i As Integer
Dim j As Integer
Dim lastRow
For i = 3 To sheetCount
    ThisWorkbook.Sheets(i).Activate
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    alarmTrig(i) = 1
    For j = STARTING_ROW To lastRow 'rows
        If Range("H" & j).Value <= 100 Then
            alarmTrig(i) = j
            'prevRowCount = RowCount
            Exit For
        End If
    Next j
Next

Dim rowCounter As Long: rowCounter = 1
For i = 3 To sheetCount
    If alarmTrig(i) <> 1 Then
        'creating a checklist for an equipment with an item to be cared of
        ThisWorkbook.Sheets(i).Activate
        Range("A1:I9").Copy
        'header creatiton
        With ThisWorkbook.Sheets(tName2)
            .Range("A" & rowCounter & ":I" & rowCounter + 8).PasteSpecial xlPasteValues
            .Range("A" & rowCounter & ":I" & rowCounter + 8).PasteSpecial xlPasteFormats
            .Rows(rowCounter + 7).Delete
            .Range("H" & rowCounter + 5).Value = " «—ÌŒ  Ê·Ìœ çò ·Ì” "
            .Range("I" & rowCounter + 5).Value = Date
            .Range("A" & rowCounter & ":I" & rowCounter).Merge
            .Range("A" & rowCounter & ":I" & rowCounter).VerticalAlignment = xlCenter
            .Range("A" & rowCounter & ":I" & rowCounter).HorizontalAlignment = xlCenter
            .Range("A" & rowCounter & ":I" & rowCounter).Borders.LineStyle = xlContinuous
            .Range("A" & rowCounter & ":I" & rowCounter).Font.Name = "B Nazanin"
            .Range("A" & rowCounter & ":I" & rowCounter).Font.Bold = True
            .Range("A" & rowCounter & ":I" & rowCounter).Font.Size = 15
            .Range("A" & rowCounter & ":I" & rowCounter).Value = "çò ·Ì”   ⁄ÊÌ÷ ¬Ì „ Â«Ì ”—ÊÌ” œÊ—Â «Ì "
        End With
        
        rowCounter = rowCounter + 8 'row for pasting the info
        For j = alarmTrig(i) To Cells(Rows.Count, 2).End(xlUp).Row
            If Range("H" & j).Value <= 100 Then
               Range(j & ":" & j).Copy
               With ThisWorkbook.Sheets(tName2).Range("A" & rowCounter)
                .PasteSpecial xlPasteValues
                .PasteSpecial xlPasteFormats
               End With
               rowCounter = rowCounter + 1
            End If
        Next j
    ThisWorkbook.Sheets(tName2).Rows(rowCounter).PageBreak = xlPageBreakManual
    End If
Next i


ThisWorkbook.Sheets(tName2).Activate
'ActiveSheet.PrintPreview

Application.Dialogs(xlDialogPrint).Show
'ActiveSheet.PrintOut
Dim answer As Integer
answer = MsgBox("¬Ì« ’›ÕÂ „Êﬁ Ì çò ·Ì”  Â« «“ ·Ì”  Õ–› ‘Êœø œ— ’Ê—  «‰ Œ«» ŒÌ— ’›ÕÂ „Ê—œ ‰Ÿ— Õ „« »«Ìœ »Â ’Ê—  œ” Ì Õ–› ‘Êœ", vbYesNo)
If answer = vbYes Then ActiveSheet.Delete
End Sub
Sub equipmentChecklist(index As Integer)
Dim tName2 As String
'Dim sheetCount As Long
'sheetCount = ThisWorkbook.Sheets.Count
tName2 = "Tmp" & ThisWorkbook.Sheets(index).Name & ("checklist" & Replace(Format(Date, "Short Date"), "/", "", 1))

'ThisWorkbook.ActiveSheet.Copy after:=Sheets(Sheets.Count)
Sheets.Add(after:=Sheets(Sheets.Count)).Name = tName2
With Service.Sheets(tName2)
    .Unprotect
    .DisplayRightToLeft = True
Dim counter As Integer
For counter = 65 To 73
    .Range(Chr(counter) & "1").EntireColumn.ColumnWidth = Service.Sheets(2).Columns(Chr(counter)).ColumnWidth
Next
'Setting the margin
With .PageSetup
 .LeftMargin = Application.CentimetersToPoints(0.7)
 .RightMargin = Application.CentimetersToPoints(0.75)
 .TopMargin = Application.CentimetersToPoints(1.4)
 .BottomMargin = Application.CentimetersToPoints(1.9)
 .HeaderMargin = Application.CentimetersToPoints(0.8)
 .FooterMargin = Application.CentimetersToPoints(0.8)
 .CenterHorizontally = True
End With
'-------------------

ActiveWindow.View = xlPageLayoutView
Dim rowCount As Long
Dim prevRowCount As Long
prevRowCount = 1

'Dim alarmTrig(1) As Integer
'ReDim alarmTrig(sheetCount)
Dim i As Integer
'Dim j As Integer
'Dim lastRow
'For i = 3 To sheetCount
'    ThisWorkbook.Sheets(i).Activate
'    alarmTrig(i) = 1
'    For j = 10 To lastRow 'rows
'        If Range("H" & j).Value <= 100 Then
'            alarmTrig(i) = j
'            'prevRowCount = RowCount
'            Exit For
'        End If
'    Next j
'Next
End With 'Tname2
With Service.Sheets(index)
Dim lastRow As Integer
lastRow = lastRowInRange(Range("status" & index))  ' (Cells(Rows.Count, 2).End(xlUp).Row
Dim firstItem As Integer
For i = STARTING_ROW To lastRow 'rows
        If Range("H" & i).Value <= 100 Then
            firstItem = i
            Exit For
        End If
Next i
Dim rowCounter As Long: rowCounter = 1
'For i = 3 To sheetCount
'    If alarmTrig(i) <> 1 Then
        'creating a checklist for an equipment with an item to be cared of
        'header creatiton
        
        With Service.Sheets(tName2)
            Service.Sheets(index).Range("A1:I11").Copy
            .Range("A1:I11").PasteSpecial xlPasteValues
            .Range("A1:I11").PasteSpecial xlPasteFormats
            .Range("B10").Value = "òÌ·Ê„ — ”—ÊÌ”"
            .Range("E10").Value = ""
            .Range("A6").Value = " «—ÌŒ  Ê·Ìœ çò ·Ì” "
            .Range("C6:E6").ClearContents
            .Range("I6").Copy
            .Range("C6").PasteSpecial xlPasteFormats
            .Range("C6:E6").Merge
            .Range("C6").Value = Date
            .Range("A2").Value = "çò ·Ì”   ⁄ÊÌ÷ ¬Ì „ Â«Ì ”—ÊÌ” œÊ—Â «Ì "
            '----------------
            .Rows(5).Delete '|
            '----------------
            .Range("D10:H10").ClearContents
            .Range("D10").Value = "Ê÷⁄Ì  «‰Ã«„"
            .Range("E10:F10").Merge
            .Range("E10").Value = "»—‰œ Ê ‰Ê⁄"
            .Range("G10").Value = "«” «‰œ«—œ ò«—ò—œ"
            .Range("H10").Value = "„ﬁœ«—/ ⁄œ«œ"
        End With
        
        rowCounter = STARTING_ROW - 1 'row for pasting the info (-1 because one row has been deleted)
        For i = firstItem To lastRow
            If .Range("H" & i).Value <= 100 Then
               'Range(i & ":" & i).Copy Destination:=ThisWorkbook.Sheets(tName2).Range(rowCounter & ":" & rowCounter)
               .Range("B" & i & ":C" & i).Copy Destination:=Service.Sheets(tName2).Range("B" & rowCounter & ":C" & rowCounter)
               With ThisWorkbook.Sheets(tName2)
                .Range("E" & rowCounter & ":F" & rowCounter).Merge
                .Range("B" & rowCounter & ":H" & rowCounter).Borders.LineStyle = xlContinuous
                If rowCounter Mod 2 <> 0 Then .Range("B" & rowCounter & ":H" & rowCounter).Interior.color = RGB(217, 217, 217)
               End With
               rowCounter = rowCounter + 1
            End If
        Next i
    'ThisWorkbook.Sheets(tName2).Rows(rowCounter).PageBreak = xlPageBreakManual
    'End If
'Next i
End With
'minimizeMainForm

ThisWorkbook.Sheets(tName2).Activate

'ActiveSheet.PrintPreview
'Application.Dialogs(xlDialogPrint).Show

'ActiveSheet.PrintOut
'normalViewMainform
Dim answer As Integer
answer = MsgBox("¬Ì« ’›ÕÂ „Êﬁ Ì çò ·Ì”  Â« «“ ·Ì”  Õ–› ‘Êœø œ— ’Ê—  «‰ Œ«» ŒÌ— ’›ÕÂ „Ê—œ ‰Ÿ— Õ „« »«Ìœ »Â ’Ê—  œ” Ì Õ–› ‘Êœ", vbYesNo)
Application.DisplayAlerts = False
If answer = vbYes Then ActiveSheet.Delete
Application.DisplayAlerts = True
End Sub
Sub minimizeMainForm()
With UserForm1
    .Width = 150
    .Height = 150
    .Left = 500
    .Top = 170
    .Frame1.Visible = False
    .Frame2.Visible = False
    .cmdContinue.Visible = True
    .cmdContinue.Height = 50
    .cmdContinue.Width = 100
    .cmdContinue.Top = 35
    .cmdContinue.Left = 20
End With
End Sub
Sub normalViewMainform()
With UserForm1
    .Width = 485
    .Height = 445
    .Left = 500
    .Frame1.Visible = True
    .Frame2.Visible = True
    .cmdContinue.Visible = False
End With
End Sub
Sub changenames()
Dim i As Integer
For i = 3 To Sheets.Count
If InStr(1, Sheets(i).Name, "„Ê Ê—") = 0 Then
    Sheets(i).Name = "ŒÊœ—Ê" & Sheets(i).Name
End If
Next
End Sub
Sub hideAll()
Dim i As Integer
For i = 3 To Sheets.Count
    Service.Sheets(i).Visible = False
Next
End Sub
Sub unHideAll()
Dim i As Integer
Set Service = ThisWorkbook
For i = 3 To Sheets.Count
    Service.Sheets(i).Visible = True
Next
End Sub

Sub Historysub()
Dim i As Integer
Set Service = ThisWorkbook
If Len(Dir(ThisWorkbook.Path & "\history.xlsm")) = 0 Then
    MsgBox "›«Ì·  «—ÌŒçÂ „ÊÃÊœ ‰Ì” !"
Else
    Set History = Workbooks.Open(ThisWorkbook.Path & "\history.xlsm")
End If

For i = 3 To Service.Sheets.Count
    History.Sheets("RAW").Copy after:=History.Sheets(Sheets.Count)
    ActiveSheet.Name = Service.Sheets(i).Name
    History.Sheets(Sheets.Count).Range("A9:A15").EntireRow.Delete
    With History.ActiveSheet
        .Range("A5").Value = Service.Sheets(i).Range("A4").Value
        .Range("D5").Value = Service.Sheets(i).Range("D4").Value
        .Range("E5").Value = Service.Sheets(i).Range("E4").Value
        .Range("F5").Value = Service.Sheets(i).Range("F4").Value
        .Range("G5").Value = Service.Sheets(i).Range("G4").Value
        .Range("H5").Value = Service.Sheets(i).Range("H4").Value
        .Range("I5").Value = Service.Sheets(i).Range("I4").Value
        .Range("B6").Value = Service.Sheets(i).Range("B5").Value
        .Range("G6").Value = Service.Sheets(i).Range("G5").Value
        .Range("C7").Value = Service.Sheets(i).Range("C6").Value
        .Range("E7").Value = Service.Sheets(i).Range("E6").Value
        .Range("G7").Value = Service.Sheets(i).Range("G6").Value
        .Range("I7").Value = Service.Sheets(i).Range("I6").Value
    End With
Next

End Sub
Sub update_kilometrage_Names()
Dim i As Integer
Set Service = ThisWorkbook
If Len(Dir(ThisWorkbook.Path & "\history.xlsm")) = 0 Then
    MsgBox "›«Ì·  «—ÌŒçÂ „ÊÃÊœ ‰Ì” !"
Else
    Set History = Workbooks.Open(ThisWorkbook.Path & "\history.xlsm")
End If
 Service.Sheets("kilometrage").Unprotect
For i = 2 To Service.Sheets.Count - 1
    Service.Sheets("Kilometrage").Range("A" & i).Value = Service.Sheets(i + 1).Name
Next
End Sub
Sub saveSingleHistoryToNewWorkbook(index As Integer)
Dim Path As String
Dim FileName As String
Dim newWorkbook As Workbook
History.Sheets(index - 1).Copy
Set newWorkbook = ActiveWorkbook
Path = History.Path
FileName = History.Sheets(index - 1).Name
Application.DisplayAlerts = False


Dim fPth As Object
Set fPth = Application.FileDialog(msoFileDialogSaveAs)

With fPth
.InitialFileName = Path & "\" & FileName
.Title = "Save your File"
.FilterIndex = 1
.InitialView = msoFileDialogViewList
If .Show <> 0 Then

newWorkbook.SaveAs FileName:=.SelectedItems(1), FileFormat:=xlOpenXMLWorkbook
End If
End With
newWorkbook.Close savechanges:=True
MsgBox (fPth.SelectedItems(1) & "›«Ì· Å‘ Ì»«‰ »« „Ê›ﬁÌ  «ÌÃ«œ ‘œ: ")
End Sub
Sub registerService(sheetIndex As Integer)
Dim form5 As Object
Set form5 = UserForm5
Dim transferDate As Date
Dim serviceItems As String
Dim i As Integer
With UserForm5
Dim serviceRowHeight
serviceRowHeight = 0
For i = 1 To ITEM_NUMBER
    If .Controls("checkbox" & i).Value Then
    serviceItems = serviceItems & vbNewLine & .Controls("checkbox" & i).Caption
    serviceRowHeight = serviceRowHeight + 0.75
    End If
Next
End With
'removing first blank line
serviceItems = Replace(serviceItems, vbCrLf, "", 1, 1)

transferDate = toGregorianDateObject(form5.TextBox57.Value, form5.TextBox56.Value, form5.TextBox55.Value)

With Service.Sheets(sheetIndex)

    Dim endRow As Integer
    .Unprotect
    .Visible = True
    .Activate
    
    Dim tmp1, tmp2
    tmp1 = findLastRow(Service.Sheets(sheetIndex), "A")
    tmp2 = findLastRow(Service.Sheets(sheetIndex), "E")
    If tmp1 > tmp2 Then endRow = tmp1 Else endRow = tmp2
    
    
    endRow = endRow + 2
    Service.Sheets("RAW").Range("A40:I42").Copy

    .Range("A" & endRow & ":I" & endRow + 2).Insert Shift:=xlDown
    endRow = endRow + 2
    .Range("A" & endRow).Value = transferDate
    .Range("C" & endRow).Value = form5.TextBox1.Value
    .Range("D" & endRow).Value = form5.TextBox2.Value
    .Range("E" & endRow).Value = serviceItems
    .Range("E" & endRow).Value = Replace(.Range("E" & endRow).Value, Chr(13), " ", 1)
    .Range("M" & endRow).Value = .Range("E" & endRow).Value
    .Range("M" & endRow).WrapText = True
    '.Range("A" & endRow & ":I" & endRow).EntireRow.AutoFit
    .Rows(endRow).RowHeight = Application.CentimetersToPoints(serviceRowHeight) '(Rows(endRow).RowHeight)
    .Range("M" & endRow).ClearContents
    .Protect
End With
MsgBox "”—ÊÌ” »« „Ê›ﬁÌ  À»  ê—œÌœ", , "À»  „Ê›ﬁ"
'Unload Me

End Sub
Function datePicking() As Integer()
UserForm4.Show
Dim jYear As Long
Dim jMonth As Long
Dim jDay As Long
Dim i As Long
Dim tempDate As String
jYear = Int(Mid(globalSelectedDate, 1, 4))
tempDate = Replace(globalSelectedDate, jYear, "", 1)

For i = Len(globalSelectedDate) To 1 Step -1
    Dim TempChar As String
    TempChar = Mid(globalSelectedDate, i, 1)
    If TempChar = "/" Then
        jDay = Int(Mid(globalSelectedDate, i + 1, Len(globalSelectedDate) - i))
      Dim s As String: s = Str(jDay)
        tempDate = Left(tempDate, Len(tempDate) - Len(Str(jDay)))
        tempDate = Replace(tempDate, "/", "", 1)
        jMonth = tempDate
        Exit For
    End If
Next
Dim dateSeparated(1 To 3) As Integer
dateSeparated(1) = jDay
dateSeparated(2) = jMonth
dateSeparated(3) = jYear
datePicking = dateSeparated


End Function
Sub addItem()
itemCounter = itemCounter + 1

Dim itemNumLabel As MSForms.label
Set itemNumLabel = UserForm8.Frame3.Controls.Add("forms.label.1", "itemNumber" & itemCounter, True)
With itemNumLabel
    .Top = itemCounter * 32
    .Height = 20
    .Width = 36
    .Left = 444
    .Caption = itemCounter
    .Font.Size = 12
    .Font.Bold = msoTrue
    .TextAlign = 2 'center align
    .BorderStyle = 1
End With

Dim descriptionTextBox As MSForms.TextBox
Set descriptionTextBox = UserForm8.Frame3.Controls.Add("forms.textbox.1", "Description" & itemCounter, True)
With descriptionTextBox
    .Top = itemCounter * 32
    .Height = 20
    .Width = 300
    .Left = 132
    .Font.Size = 12
    .Font.Bold = msoTrue
    .TextAlign = 3 'right to left
    .BorderStyle = 1
End With

Dim costTextBox As MSForms.TextBox
Set costTextBox = UserForm8.Frame3.Controls.Add("forms.textbox.1", "Cost" & itemCounter, True)
With costTextBox
    .Top = itemCounter * 32
    .Height = 20
    .Width = 75
    .Left = 48
    .Font.Size = 12
    .Font.Bold = msoTrue
    .TextAlign = 1 'left to right
    .BorderStyle = 1
    
End With
Dim clickCMDEvents As UserFormEvents

Dim deleteRowButton As MSForms.CommandButton
Set deleteRowButton = UserForm8.Frame3.Controls.Add("Forms.commandbutton.1", "deleteRow" & itemCounter, True)
With deleteRowButton
    .Top = itemCounter * 32
    .Height = 20
    .Width = 25
    .Left = 16
    .Caption = "-"
    .Font.Size = 12
    .Font.Bold = msoTrue
    .ForeColor = RGB(250, 0, 0)
    .Tag = itemCounter
End With
Set clickCMDEvents = New UserFormEvents
Set clickCMDEvents.commandButtonNumber = deleteRowButton
collectionEvent.Add clickCMDEvents

With UserForm8
Dim bigFrame As Boolean
If UserForm8.Height > 620 Then
    bigFrame = True
    .Frame3.ScrollBars = fmScrollBarsVertical
    .Frame3.ScrollHeight = .Frame3.ScrollHeight + 32
    .CommandButton1.Top = .CommandButton1.Top + 32
    .Frame3.ScrollTop = .Frame3.ScrollHeight - 70
Else
    UserForm8.Height = .Frame3.Top + itemCounter * 32 + 150
    .Frame3.ScrollBars = fmScrollBarsNone
    .Frame3.Height = (itemCounter + 1) * 32 + 40
    .CommandButton1.Top = .CommandButton1.Top + 32
    .CommandButton2.Top = UserForm8.Height - 70
    .CommandButton3.Top = UserForm8.Height - 70

End If
End With

End Sub


Sub removeItem(ByVal tagNum As Integer)
Dim i As Integer
With UserForm8
For i = tagNum + 1 To itemCounter

.Controls("Description" & i - 1).text = .Controls("Description" & i).text
.Controls("cost" & i - 1).text = .Controls("cost" & i).text

Next


If UserForm8.Height > 620 Then
   ' bigFrame = True
    .Frame3.ScrollBars = fmScrollBarsVertical
    .Frame3.ScrollHeight = .Frame3.ScrollHeight - 32
    .CommandButton1.Top = .CommandButton1.Top - 32
    '.Frame3.ScrollTop = .Frame3.ScrollHeight - 70
Else
    UserForm8.Height = UserForm8.Height - 32
    .Frame3.ScrollBars = fmScrollBarsNone
    .Frame3.Height = .Frame3.Height - 32
    .CommandButton1.Top = .CommandButton1.Top - 32
    .CommandButton2.Top = UserForm8.Height - 70
    .CommandButton3.Top = UserForm8.Height - 70
    

End If

.Controls.Remove ("cost" & itemCounter)
.Controls.Remove ("description" & itemCounter)
.Controls.Remove ("itemnumber" & itemCounter)
.Controls.Remove ("deleterow" & itemCounter)

.Frame3.Visible = False
.Frame3.Visible = True
If tagNum = itemCounter Then tagNum = tagNum - 1
If tagNum = 0 Then
    itemCounter = itemCounter - 1
    Exit Sub
End If
.Frame3.Controls("deleterow" & tagNum).SetFocus
End With

itemCounter = itemCounter - 1

End Sub

Sub repairRegister(sheetName As String)
Dim form8 As Object
Set form8 = UserForm8
Dim deployDate As Date
deployDate = toGregorianDateObject(form8.TextBox60.Value, form8.TextBox59.Value, form8.TextBox58.Value)
Dim repairDate As Date
repairDate = toGregorianDateObject(form8.TextBox57.Value, form8.TextBox56.Value, form8.TextBox55.Value)
Dim sheetIndex
sheetIndex = indexFind(sheetName, originalAssetList) + 1
'History.Sheets(sheetName).Activate
With Service.Sheets(sheetIndex)
Service.Sheets(1).Activate
    .Unprotect
Dim endRow As Integer
Dim tmp1, tmp2
    tmp1 = findLastRow(Service.Sheets(sheetIndex), "A")
    tmp2 = findLastRow(Service.Sheets(sheetIndex), "E")
    If tmp1 > tmp2 Then endRow = tmp1 Else endRow = tmp2
endRow = endRow + 2

Dim repairShopData As String

repairShopData = form8.TextBox3.Value & vbCrLf & form8.TextBox4.Value & vbCrLf & form8.TextBox5.Value

Dim repairManName As String
Dim repairManContact As String
repairManName = form8.TextBox1.Value
repairManContact = form8.Frame1.TextBox2.text

Service.Sheets("RAW").Range("A44:I49").Copy
'With History.Sheets(sheetName)
    .Range("A" & endRow & ":I" & endRow + 5).Insert Shift:=xlDown
     endRow = endRow + 2
.Visible = True
.Activate
    .Range("A" & endRow - 1 & ":I" & endRow - 1).EntireRow.AutoFit
    .Range("A" & endRow).Value = deployDate
    .Range("C" & endRow).Value = form8.TextBox8.Value
    .Range("D" & endRow).Value = form8.TextBox7.Value
    .Range("E" & endRow).Value = repairShopData
    .Range("E" & endRow).Value = Replace(.Range("E" & endRow).Value, Chr(13), " ", 1)
    .Range("N" & endRow).Value = .Range("E" & endRow).Value
    .Range("N" & endRow).WrapText = True
   
    .Range("H" & endRow).Value = repairManName
    .Range("I" & endRow).Value = repairManContact
    .Range("O" & endRow).Value = .Range("H" & endRow).Value
    .Range("O" & endRow).WrapText = True
    .Range("A" & endRow & ":I" & endRow).EntireRow.AutoFit
    
    .Rows(endRow).RowHeight = (Rows(endRow).RowHeight)
    .Range("N" & endRow).ClearContents
    .Range("O" & endRow).ClearContents
   
endRow = endRow + 1
    .Range("C" & endRow).Value = form8.TextBox6.Value
endRow = endRow + 1
    .Range("C" & endRow).Value = repairDate

    '.Range("E" & endRow + 2).Value = Replace(.Range("E" & endRow + 2).Value, Chr(13), " ", 1)
endRow = endRow + 2
Dim i As Integer
For i = 1 To itemCounter
    Service.Sheets("RAW").Range("A50:I50").Copy
    .Range("A" & endRow & ":I" & endRow).Insert Shift:=xlDown
    .Range("A" & endRow).Value = i
    .Range("C" & endRow).Value = form8.Frame3.Controls("description" & i).text
    .Range("I" & endRow).Value = form8.Frame3.Controls("cost" & i).Value
    endRow = endRow + 1
Next
Dim sumRange As Range
Set sumRange = .Range("I" & endRow - itemCounter & ":I" & endRow - 1)
    Service.Sheets("RAW").Range("A51:I51").Copy
    .Range("A" & endRow & ":I" & endRow).Insert Shift:=xlDown
    .Range("G" & endRow).Value = Application.WorksheetFunction.Sum(sumRange)
    .Protect
End With

End Sub

Sub search(ByRef text As String, ByRef assets As Variant, ByRef listToUpdate As Variant)
With listToUpdate
'UserForm1.ListBox1

Dim i
'loading = True
.Clear
If text <> "" Then
    .addItem Service.Sheets(1).Range("A1").Value
End If
For Each i In assets
    If InStr(1, i, text, vbTextCompare) <> 0 Then .addItem i
Next
End With
'loading = False
End Sub
Function findLastRow(sheet, columnIndex As String)
 findLastRow = sheet.Range(columnIndex & Rows.Count).End(xlUp).Row 'count of unique values

End Function

Function indexFind(ByVal Value As Variant, arr As Variant) As Integer
    indexFind = Application.Match(Value, arr, False)
End Function
Sub equipmentWithItems(ByRef listToUpdate As Variant)
Dim sheetCount
sheetCount = ThisWorkbook.Sheets.Count
ReDim alarmTrig(sheetCount)
Dim i As Integer
Dim j As Integer
Dim lastRow

For i = 3 To sheetCount
    With ThisWorkbook.Sheets(i)
    lastRow = .Cells(Rows.Count, 2).End(xlUp).Row
    alarmTrig(i) = 1
    For j = STARTING_ROW To lastRow 'rows
        If .Range("H" & j).Value <= 100 Then
            alarmTrig(i) = j
            'prevRowCount = RowCount
            Exit For
        End If
    Next j
    End With
Next

With listToUpdate

.Clear
'If text <> "" Then
    .addItem Service.Sheets(1).Range("A1").Value
'End If

Dim rowCounter As Long: rowCounter = 1
For i = 3 To sheetCount
    If alarmTrig(i) <> 1 Then
        .addItem Service.Sheets(i).Name
    End If
Next

End With

End Sub
Sub updateHistory()
Dim i As Integer
Dim lastRow
Set Service = ThisWorkbook
For i = 3 To Service.Sheets.Count
unHideAll

    lastRow = findLastRow(Service.Sheets(i), "B")
    
    With Service.Sheets(i)
         .Unprotect
       Service.Names.Add Name:="Status" & i, RefersTo:=Service.Sheets(i).Range("A1:" & "I" & lastRow)
        .Cells.PageBreak = xlNone
        .Rows(lastRow + 2).PageBreak = xlPageBreakManual
        .Rows(lastRow + 1).RowHeight = Application.CentimetersToPoints(0.3)
        .Rows(lastRow + 2).RowHeight = Application.CentimetersToPoints(0.3)
        .Range("A" & lastRow + 2 & ":I" & lastRow + 2).UnMerge
        .Range("A" & lastRow + 2 & ":I" & lastRow + 2).ClearContents
        .Range("A" & lastRow + 2 & ":I" & lastRow + 2).ClearFormats
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

Next
End Sub
Function findHistoryRow(sheetIndex As Integer) As Integer
    If sheetIndex > Service.Sheets.Count Then
        MsgBox "out of range"
        Exit Function
    End If
    findHistoryRow = Service.Sheets(sheetIndex).Range("Status" & sheetIndex).Rows.Count + 2
End Function
Sub eraseRawExtraRows(sheetIndex)
Dim i
For i = STARTING_ROW To STARTING_ROW + ITEM_NUMBER + 2
        If Service.Sheets(sheetIndex).Range("b" & i).Value = "" Then
            Service.Sheets(sheetIndex).Unprotect
            MsgBox findLastRow(Service.Sheets(sheetIndex), "E")
            Service.Sheets(sheetIndex).Rows(i & ":" & findLastRow(Service.Sheets(sheetIndex), "E")).Delete
            Service.Sheets(sheetIndex).Protect
            Exit Sub
        End If
    Next
End Sub
Sub updateInfoPage() 'sheetIndex
Dim i
Set Service = ThisWorkbook
For i = 3 To Service.Sheets.Count
    With Service.Sheets(i)
     If .Range("A7").Value = "" Then
        .Unprotect
        Service.Sheets("RAW").Range("A7:I8").Copy
        .Range("A7:I8").Insert Shift:=xlDown
        Service.Sheets("RAW").Range("1:12").Copy
        Dim j
        For j = 1 To 12
        .Range(j & ":" & j).RowHeight = Service.Sheets("RAW").Range(j & ":" & j).RowHeight
        Next
        '.Range("A7:I8").PasteSpecial
        'MsgBox .Name
        
     End If
    End With
Next
Application.CutCopyMode = False
End Sub
Sub protectAll()
Dim i
Set Service = ThisWorkbook
For i = 1 To Service.Sheets.Count
    Service.Sheets(i).Protect
Next
End Sub
Function lastRowInRange(rng As Range)
lastRowInRange = rng.Cells(rng.Rows.Count, 1).Row
End Function
Sub integrityCheck()
Dim i
Set Service = ThisWorkbook
For i = 3 To Service.Sheets.Count
    If Service.Sheets(i).Name <> Service.Sheets(1).Range("A" & i - 1) Then MsgBox ("Assets list name conflict at no " & i)
Next
End Sub

Sub warningPages()
Dim i
Dim j
For i = 3 To Service.Sheets.Count
    With Service.Sheets(i)
        For j = STARTING_ROW To lastRowInRange(Range("status" & i))
        If .Range("H" & j).Value < 300 Then
            UserForm1.Frame2.ListBox1.List(i - 2, 1) = "*"
            Exit For
        End If
        Next
    End With
Next
            
End Sub
