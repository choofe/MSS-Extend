VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "À»   €ÌÌ—/ ÕÊ·"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub

Private Sub CommandButton1_Click()

'Dim Service As Workbook, History As Workbook
'Set Service = Workbooks.Item("Project MSS V2.5 BETA.xlsm")
'Set History = Workbooks.Open(ThisWorkbook.Path & "\history.xlsm")
Dim transferDate As Date
transferDate = toGregorianDateObject(TextBox57.Value, TextBox56.Value, TextBox55.Value)
Dim sheetIndex
sheetIndex = indexFind(Label14.Caption, originalAssetList) + 1
With Service.Sheets(sheetIndex)
.Unprotect
Dim endRow As Integer

Dim tmp1, tmp2
    tmp1 = findLastRow(Service.Sheets(sheetIndex), "A")
    tmp2 = findLastRow(Service.Sheets(sheetIndex), "E")
    If tmp1 > tmp2 Then endRow = tmp1 Else endRow = tmp2
endRow = endRow + 2

Service.Sheets("RAW").Range("A36:I38").Copy
    .Range("G6").Value = TextBox7.Value
    .Range("H4").Value = TextBox8.Value
    .Range("I4").Value = TextBox9.Value
    .Range("A" & endRow & ":I" & endRow + 2).Insert Shift:=xlDown
    .Range("A" & endRow + 2).Value = transferDate
    .Range("C" & endRow + 2).Value = TextBox5.Value
    .Range("D" & endRow + 2).Value = TextBox1.Value
    .Range("E" & endRow + 2).Value = TextBox2.Value
    .Range("F" & endRow + 2).Value = TextBox7.Value
    .Range("G" & endRow + 2).Value = TextBox8.Value
    .Range("H" & endRow + 2).Value = TextBox6.Value
    Dim rowHght
    rowHght = 0.75
    For i = 1 To Len(TextBox6.text)
        If Mid(TextBox6.text, i, 1) = Chr(13) Then
            rowHght = rowHght + 0.75
        End If
    Next
    .Range("H" & endRow + 2).Value = Replace(.Range("H" & endRow + 2).Value, Chr(13), " ", 1)
    .Range("K" & endRow + 2).Value = .Range("H" & endRow + 2).Value
    .Range("K" & endRow + 2).WrapText = True
    .Rows(endRow + 2).RowHeight = Application.CentimetersToPoints(rowHght)
    '.Range("A" & endRow + 2 & ":I" & endRow + 2).EntireRow.AutoFit
'    .Rows(endRow + 2).RowHeight = (Rows(endRow + 2).RowHeight)
    .Range("K" & endRow + 2).ClearContents
.Protect
End With
With Service.Sheets(1)
    .Unprotect
    .Range("B" & sheetIndex - 1).Value = TextBox5.Value
    .Range("C" & sheetIndex - 1).Value = transferDate
    .Protect
End With
MsgBox " €ÌÌ—«  »« „Ê›ﬁÌ  À»  ê—œÌœ", , "À»  „Ê›ﬁ"
Unload Me

End Sub

Private Sub CommandButton8_Click()
Dim datePicked() As Integer
datePicked = datePicking()
TextBox57.text = datePicked(3)
TextBox56.text = datePicked(2)
TextBox55.text = datePicked(1)

End Sub

Private Sub CommandButton9_Click()
Unload Me
End Sub

Private Sub Label4_Click()

End Sub

Private Sub ToggleButton1_Click()
If ToggleButton1.Value = True Then
    TextBox1.Enabled = True
    TextBox2.Enabled = True
Else
    TextBox1.Enabled = False
    TextBox2.Enabled = False
End If

End Sub

Private Sub UserForm_Activate()
'History.Sheets(Service.ActiveSheet.Name).Activate
Dim sheetIndex
sheetIndex = indexFind(UserForm1.Frame2.ListBox1.Value, originalAssetList) + 1
With Service.Sheets(sheetIndex)
    Label14.Caption = .Name
    Dim todayDate As Date: todayDate = Date
    Dim jalaliToday As Variant
    jalaliToday = toJalaaliFromDateObject(todayDate)
    TextBox55.Value = jalaliToday(2)
    TextBox56.Value = jalaliToday(1)
    TextBox57.Value = jalaliToday(0)
TextBox5.text = Service.Sheets(1).Range("B" & sheetIndex - 1).Value

TextBox1.text = .Range("G6").Value
TextBox2.text = .Range("H4").Value
End With
End Sub

