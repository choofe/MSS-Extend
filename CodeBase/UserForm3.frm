VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "ÊÌ—«Ì‘ «ÿ·«⁄«  Ê”Ì·Â ‰ﬁ·ÌÂ"
   ClientHeight    =   13050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim checkBoxSelection(1 To 16) As Boolean

Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    TextBox1.Enabled = True
    TextBox39.Enabled = True
    Else
    TextBox1.Enabled = False
    TextBox39.Enabled = False
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
    Else
    TextBox14.Enabled = False
    TextBox52.Enabled = False
End If
End Sub

Private Sub CheckBox15_Click()
If CheckBox15.Value = True Then
    TextBox15.Enabled = True
    TextBox53.Enabled = True
    Else
    TextBox15.Enabled = False
    TextBox53.Enabled = False
End If
End Sub

Private Sub CheckBox16_Click()
If CheckBox16.Value = True Then
    TextBox16.Enabled = True
    TextBox54.Enabled = True
    Else
    TextBox16.Enabled = False
    TextBox54.Enabled = False
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
    Else
    TextBox6.Enabled = False
    TextBox44.Enabled = False
End If
End Sub

Private Sub CheckBox7_Click()
If CheckBox7.Value = True Then
    TextBox7.Enabled = True
    TextBox45.Enabled = True
    Else
    TextBox7.Enabled = False
    TextBox45.Enabled = False
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

Private Sub CommandButton10_Click()
UserForm6.Show
End Sub

Private Sub CommandButton2_Click()

End Sub



Private Sub CommandButton4_Click()

End Sub

Private Sub CommandButton11_Click()
Unload Me
End Sub

Private Sub CommandButton5_Click()
'Updating data

ActiveSheet.Unprotect
ActiveSheet.Range("H4").Value = TextBox35.Value
ActiveSheet.Range("I4").Value = TextBox36.Value
ActiveSheet.Range("G6").Value = TextBox37.Value
'ActiveSheet.Range("I6").Value = globalSelectedDate
ActiveSheet.Protect
'---------------
'next phase for editing Items's rows
'MsgBox checkBoxSelection(1)
'For i = 1 To 16
'    If Controls("checkbox" & i).Value Then checkBoxSelection(i) = True
'Next
'For i = 10 To 25
'    If checkBoxSelection(i - 9) Then
'        If Range("B" & i) = Controls("checkbox" & i - 9).Caption Then
'            MsgBox "OK-true and exist"
'        Else
'            MsgBox "new row should be added"
'        End If
'    Else
'        If Range("B" & i) = Controls("checkbox" & i - 9).Caption Then
'            MsgBox "row should be removed"
'        Else
'            MsgBox "OK-False and not exist"
'        End If
'    End If
'
'Next
Service.ActiveSheet.Protect
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
Dim deleteSheet As VbMsgBoxResult, sheetIndex As Integer
deleteSheet = MsgBox("¬Ì« «“ Õ–› ’›ÕÂ Ê”Ì·Â „ÿ„∆‰ Â” Ìœø «Ì‰ ⁄„·Ì«  €Ì—ﬁ«»· »«“ê‘  „Ì »«‘œ", vbYesNo, "Õ–›")
Dim createCopyAnswer As VbMsgBoxResult
If deleteSheet = vbYes Then
    sheetIndex = ActiveSheet.index
    Service.Sheets("Kilometrage").Activate
    ActiveSheet.Unprotect
    Rows(sheetIndex - 1).Delete
    ActiveSheet.Protect
    Service.Sheets(Label14.Caption).Activate
    ActiveSheet.Unprotect
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    
'--------------------------------------
'creating history backup on demand
createCopyAnswer = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «“ ’›ÕÂ  «—ÌŒçÂ Ê”Ì·Â ‰ﬁ·ÌÂ Ìò òÅÌ œ— Ìò ›«Ì· „Ã“«  ÂÌÂ ‘Êœø", vbYesNo, "«ÌÃ«œ òÅÌ")
        If createCopyAnswer = vbYes Then saveSingleHistoryToNewWorkbook (sheetIndex)
    History.Sheets(sheetIndex - 1).Delete
    Application.DisplayAlerts = True
    Unload Me
End If
History.Save
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub UserForm_Initialize()
Dim sheetIndex
sheetIndex = indexFind(UserForm1.Frame2.ListBox1.Value, originalAssetList) + 1
With Service.Sheets(sheetIndex)
Label14.Caption = .Name
Dim rngLastRow
    rngLastRow = lastRowInRange(Range("status" & sheetIndex))
For i = 16 To 1 Step -1
    Dim checkCaption As String
    checkCaption = Controls("checkbox" & i).Caption

    For j = rngLastRow To STARTING_ROW Step -1
        If .Range("B" & j).text = checkCaption Then
            Controls("checkbox" & i).Value = True
            Controls("textbox" & i + 38).Value = .Range("G" & j).Value
            Controls("textbox" & i).Value = .Range("D" & j).Value
            'checkBoxSelection(i - 9) = True
            Exit For
            End If
            
    Next j
Next i
End With
End Sub
