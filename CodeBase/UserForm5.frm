VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "À»  ”—ÊÌ”"
   ClientHeight    =   10815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11325
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Dim currSheetIndex As Integer

'ActiveSheet.Unprotect
'currSheetIndex = ActiveSheet.index
Dim sheetIndex
sheetIndex = indexFind(UserForm1.Frame2.ListBox1.Value, originalAssetList) + 1

Service.Sheets("Kilometrage").Unprotect
Service.Sheets("Kilometrage").Range("B" & sheetIndex - 1).Value = TextBox1.Value
Service.Sheets("Kilometrage").Protect
Dim i As Integer
Dim j As Integer

With Service.Sheets(sheetIndex)
.Unprotect
For i = 1 To 16
   Dim checkCaption As String
   checkCaption = Controls("checkbox" & i).Caption
   If Controls("checkbox" & i).Value Then
        For j = i + 9 To 25
            If .Range("B" & j).text = checkCaption Then
                .Range("E" & j).Value = TextBox1.Value
                If Controls("standard" & i).text <> "" Then
                    .Range("D" & j).Value = Controls("standard" & i).Value
                  Else
                    
                End If
                Exit For
            End If
        Next j
   End If
Next i
.Protect
registerService (sheetIndex)
'History.Save
Unload Me
End With
End Sub

Private Sub CommandButton2_Click()
Dim i As Integer
For i = 1 To ITEM_NUMBER
    If Controls("checkbox" & i).ForeColor = RGB(250, 0, 0) Then
        Controls("checkbox" & i).Value = True
    Else
        Controls("checkbox" & i).Value = False
    End If
Next
End Sub
Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    standard1.Enabled = True
    Else
    standard1.Enabled = False
End If
End Sub
Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
    standard2.Enabled = True
    Else
    standard2.Enabled = False
End If
End Sub
Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then
    standard3.Enabled = True
    Else
    standard3.Enabled = False
End If
End Sub
Private Sub CheckBox4_Click()
If CheckBox4.Value = True Then
    standard4.Enabled = True
    Else
    standard4.Enabled = False
End If
End Sub
Private Sub CheckBox5_Click()
If CheckBox5.Value = True Then
    standard5.Enabled = True
    Else
    standard5.Enabled = False
End If
End Sub
Private Sub CheckBox6_Click()
If CheckBox6.Value = True Then
    standard6.Enabled = True
    Else
    standard6.Enabled = False
End If
End Sub
Private Sub CheckBox7_Click()
If CheckBox7.Value = True Then
    standard7.Enabled = True
    Else
    standard7.Enabled = False
End If
End Sub
Private Sub CheckBox8_Click()
If CheckBox8.Value = True Then
    standard8.Enabled = True
    Else
    standard8.Enabled = False
End If
End Sub
Private Sub CheckBox9_Click()
If CheckBox9.Value = True Then
    standard9.Enabled = True
    Else
    standard9.Enabled = False
End If
End Sub
Private Sub CheckBox10_Click()
If CheckBox10.Value = True Then
    standard10.Enabled = True
    Else
    standard10.Enabled = False
End If
End Sub

Private Sub CheckBox11_Click()
If CheckBox11.Value = True Then
    standard11.Enabled = True
    Else
    standard11.Enabled = False
End If
End Sub
Private Sub CheckBox12_Click()
If CheckBox12.Value = True Then
    standard12.Enabled = True
    Else
    standard12.Enabled = False
End If
End Sub
Private Sub CheckBox13_Click()
If CheckBox13.Value = True Then
    standard13.Enabled = True
    Else
    standard13.Enabled = False
End If
End Sub
Private Sub CheckBox14_Click()
If CheckBox14.Value = True Then
    standard14.Enabled = True
    Else
    standard14.Enabled = False
End If
End Sub
Private Sub CheckBox15_Click()
If CheckBox15.Value = True Then
    standard15.Enabled = True
    Else
    standard15.Enabled = False
End If
End Sub
Private Sub CheckBox16_Click()
If CheckBox16.Value = True Then
    standard16.Enabled = True
    Else
    standard16.Enabled = False
End If
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
Private Sub UserForm_Initialize()

Dim sheetIndex
sheetIndex = indexFind(UserForm1.Frame2.ListBox1.Value, originalAssetList) + 1
Label14.Caption = Service.Sheets(sheetIndex).Name
Dim todayDate As Date: todayDate = Date
Dim jalaliToday As Variant
jalaliToday = toJalaaliFromDateObject(todayDate)
TextBox55.Value = jalaliToday(2)
TextBox56.Value = jalaliToday(1)
TextBox57.Value = jalaliToday(0)
With Service.Sheets(sheetIndex)
Dim i As Integer
Dim j As Integer
    Dim rngLastRow
    rngLastRow = lastRowInRange(Range("status" & sheetIndex))
For i = 16 To 1 Step -1
    Dim checkCaption As String
    checkCaption = Controls("checkbox" & i).Caption
    'Dim rng As Range
    'Set rng = Range("status" & sheetIndex)
    For j = rngLastRow To STARTING_ROW Step -1
        If .Range("B" & j).text = checkCaption Then
            Controls("checkbox" & i).Value = True
            If .Range("H" & j).Value <= 100 Then
                Controls("checkbox" & i).ForeColor = RGB(250, 0, 0)
            End If
            Exit For
            End If
            
    Next j
Next i
For i = 1 To 16
    If Not (Controls("checkbox" & i).Value) Then
        Controls("checkbox" & i).Enabled = False
        Controls("standard" & i).Enabled = False
    End If
Next
End With
End Sub
