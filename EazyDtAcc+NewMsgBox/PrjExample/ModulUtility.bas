Attribute VB_Name = "ModUtil"
Public Buttons As Long
'All of this Code, i got from lecturer, Mr. Edy
'----------------------------------------------
Public Sub aktif(blanko As Form, kondisi As Boolean)
Dim ctrl As Control
For Each ctrl In blanko
 If TypeName(ctrl) = "TextBox" Then ctrl.Enabled = kondisi
 If TypeName(ctrl) = "ComboBox" Then ctrl.Enabled = kondisi
 If TypeName(ctrl) = "MaskEdBox" Then ctrl.Enabled = kondisi
 If TypeName(ctrl) = "DTPicker" Then ctrl.Enabled = kondisi
Next
End Sub
Public Sub kosong(blanko As Form)
Dim ctrl As Control
For Each ctrl In blanko
 If TypeName(ctrl) = "TextBox" Then ctrl.Text = Empty
 If TypeName(ctrl) = "ComboBox" Then ctrl.Text = Empty
 If TypeName(ctrl) = "MaskEdBox" Then ctrl.Text = Empty
Next
End Sub
'------------------------------------------------------

'ex-lecturer-------------------------------------------
Public Sub WarnaNORMAL(forme As Form)
Dim ctrl As Control
For Each ctrl In forme

    '[ case sensitive ]
    If TypeName(ctrl) = "TextBox" Then ctrl.BackColor = &H80000010
    If TypeName(ctrl) = "ComboBox" Then ctrl.BackColor = &H80000010
    If TypeName(ctrl) = "MaskEdBox" Then ctrl.BackColor = &H80000010

Next
End Sub
'Public Sub Button(blanko As Form, kondisi As Boolean)
'Dim ctrl As Control
'For Each ctrl In blanko
'If TypeName(ctrl) = "CommandButton" Then
' If ctrl.Caption = "..." Then
'  ctrl.Enabled = kondisi
' End If
'End If
'Next
'End Sub

Public Function cekKosong(forme As Form)
Dim ctrl As Control
For Each ctrl In forme
 If TypeName(ctrl) = "TextBox" Then
  If ctrl.Text = "" Then
   MsgBox "Data belum diinput..", vbCritical, "Pesan"
   ctrl.BackColor = vbBlack
   cekKosong = True
  Else
   ctrl.BackColor = &H80000010
  End If
 End If

 If TypeName(ctrl) = "MaskEdBox" Then
  If ctrl.Text = "" Then
   MsgBox "Data belum diinput..", vbCritical, "Pesan"
   ctrl.BackColor = vbRed
   cekKosong = True
  Else
   ctrl.BackColor = &H80000010
  End If
 End If

 If TypeName(ctrl) = "ComboBox" Then
  If ctrl.Text = "" Then
   MsgBox "Data belum dipilih..", vbCritical, "Pesan"
   ctrl.BackColor = vbRed
   cekKosong = True
  Else
   ctrl.BackColor = &H80000010
  End If
 End If
Next
End Function
'-----------------------------------------------------









