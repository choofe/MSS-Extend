VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm9 
   Caption         =   "UserForm3"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3600
   OleObjectBlob   =   "UserForm9.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim i
For i = 0 To ListBox1.ListCount - 1
    ListBox1.Selected(i) = True
Next
End Sub

Private Sub ListBox1_Click()
Me.Hide

End Sub

Private Sub UserForm_Activate()
'Dim sheetslist() As String
ReDim sheetslist(1 To ThisWorkbook.Sheets.Count)
Dim i As Integer
Dim j As Integer
Dim lastRow
Dim counter
Service.Activate
counter = 1
For i = 3 To Service.Sheets.Count
    'Service.Sheets(i).
    'Service.Sheets(i).Activate
    With Service.Sheets(i)
    lastRow = Service.Sheets(i).Cells(Rows.Count, 2).End(xlUp).Row
    'alarmTrig(i) = 1
    For j = STARTING_ROW To lastRow 'rows
        If .Range("H" & j).Value <= 100 Then
            sheetslist(counter) = Service.Sheets(i).Name
            ListBox1.addItem Service.Sheets(i).Name
            'counter = counter + 1
            Exit For
        End If
    Next j
    End With
Next

'For i = ThisWorkbook.Sheets.count To 5 Step -1
'    sheetslist(i) = ThisWorkbook.Sheets(i).Name
'Next
'ListBox1.List = sheetslist
If ListBox1.ListCount * 12.5 < 312 Then
    UserForm9.Height = ListBox1.ListCount * 12.5 + 130
    ListBox1.Height = ListBox1.ListCount * 12.5 + 10
    CommandButton1.Top = ListBox1.Height + 30
    CommandButton2.Top = CommandButton1.Top + 30
Else
    UserForm9.Height = 400
    ListBox1.Height = 315.5
    CommandButton1.Top = 330.5
    CommandButton2.Top = CommandButton1.Top + 30
End If
'UserForm9.Height = ListBox1.Height + 200
End Sub

