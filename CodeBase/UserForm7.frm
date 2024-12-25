VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "UserForm7"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7170
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton2_Click()
Dim lastRow As Integer, i As Integer
lastRow = Service.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row 'count of unique values
Dim KLMUpdateAnswer As VbMsgBoxResult
Dim readDate As Date
readDate = toGregorianDateObject(TextBox57.Value, TextBox56.Value, TextBox55.Value)
With Service.Sheets(1)
Service.Sheets("Kilometrage").Unprotect
For i = 1 To lastRow - 1
    If Controls("kilometrupdate" & i).Value <> "" And _
        Int(Val(.Range("B" & i + 1).Value)) > Int(Val(Controls("kilometrupdate" & i).Value)) Then
       
       KLMUpdateAnswer = MsgBox("ﬁ—«∆  òÌ·Ê„ — »—«Ì " & .Range("A" & i + 1).Value & _
        " «“ ﬁ—«∆  ﬁ»·Ì ò„ — «” . ¬Ì« ﬁ—«∆  òÌ·Ê„ — »« „ﬁœ«— Ê«—œ ‘œÂ »Â —Ê“ —”«‰Ì ‘Êœø", vbYesNo + vbMsgBoxRight + vbExclamation, "ò«Â‘ òÌ·Ê„ — ò«—ò—œ")
       
       If KLMUpdateAnswer = vbYes Then
        .Range("B" & i + 1).Value = Controls("kilometrupdate" & i).Value
        .Range("C" & i + 1).Value = readDate
       End If
    
    Else
        
        If Controls("kilometrupdate" & i).Value <> "" And _
        Int(Val(.Range("B" & i + 1).Value)) <= Int(Val(Controls("kilometrupdate" & i).Value)) Then
         .Range("B" & i + 1).Value = Controls("kilometrupdate" & i).Value
         .Range("C" & i + 1).Value = readDate
        End If
    
    End If
Next
.Protect
End With
End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub CommandButton8_Click()
Dim datePicked() As Integer
datePicked = datePicking()
TextBox57.text = datePicked(3)
TextBox56.text = datePicked(2)
TextBox55.text = datePicked(1)

End Sub

Private Sub UserForm_Initialize()
Set Service = ThisWorkbook
If Len(Dir(ThisWorkbook.Path & "\history.xlsm")) = 0 Then
    MsgBox "›«Ì·  «—ÌŒçÂ „ÊÃÊœ ‰Ì” !"
Else
    Set History = Workbooks.Open(ThisWorkbook.Path & "\history.xlsm")
End If

Dim todayDate As Date: todayDate = Date
Dim jalaliToday As Variant
jalaliToday = toJalaaliFromDateObject(todayDate)
TextBox55.Value = jalaliToday(2)
TextBox56.Value = jalaliToday(1)
TextBox57.Value = jalaliToday(0)

Dim i As Integer
Dim updKLM As Control
Dim nameLabel As Control
Dim lastRow As Integer
lastRow = Service.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row 'count of unique values
Dim FrameHeight As Integer
FrameHeight = lastRow * 25 + 10
Dim bigFrame As Boolean: bigFrame = False

If FrameHeight > 510 Then
    FrameHeight = 510
    Frame1.ScrollBars = fmScrollBarsVertical
    bigFrame = True
End If

UserForm7.Height = FrameHeight + 140
Frame1.Height = FrameHeight

For i = 0 To lastRow - 1
    Set nameLabel = Me.Frame1.Controls.Add("forms.label.1", "nameLabel" & i, True)
    With nameLabel
    .Left = Frame1.Width - 170
    .TextAlign = 3
    .Width = 150
    .Height = 18
    .Top = (25 * i)
    .Caption = Service.Sheets(1).Range("A" & i + 1).Value
    .Font.Size = 14
    End With
    If i <> 0 Then
    Set updKLM = Me.Frame1.Controls.Add("forms.textbox.1", "KilometrUpdate" & i, True)
    With updKLM
    
    .Height = 18
    .Font.Size = 12
    .BorderStyle = 1
    .SpecialEffect = 0
    .Top = (25 * i)
    .Left = Frame1.Width - 310
    .Width = 150
    End With
    End If
    If bigFrame Then Frame1.ScrollHeight = Frame1.ScrollHeight + 25
Next
CommandButton2.Top = UserForm7.Height - 70
CommandButton3.Top = UserForm7.Height - 70
End Sub
