VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'%#########################################%'
'%Author   : Tendri S (20)                 %'
'%Date     : October 08, 2007              %'
'%Location : Bekasi, Indonesia             %'
'%Email    : mizz_daeng@plasa.com          %'
'%Please Do Not Removes Any Copyrights and %'
'%#########################################%'

Option Explicit

Private sMsg As String

Private WithEvents fMessage As frmMessage
Attribute fMessage.VB_VarHelpID = -1
'Fungsi kyk msgboxnya vb
'Kelar jg nih.. ^^
Public Function TendriMsg(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String)
Dim TmblStr() As String
Dim TmblClose() As Boolean
Dim TipeTmbl As Long
Dim TipeIcon As Long

Const PlusButtons  As Long = &H7
Const PlusIcons As Long = &H71

Set fMessage = New frmMessage

With fMessage
 .JudulMsg = Title
End With
    
'Teks yang akan ditampilkan pd tombol dan tombol u/ ngeclose msgbox(nandain doank-->true)
TipeTmbl = (Buttons And PlusButtons) 'vbRetryCancel, vbYesNo, vbYesNoCancel, vbAbortRetryIgnore, vbOKCancel
Select Case TipeTmbl
 Case vbYesNo
  ReDim TmblStr(1)
  ReDim TmblClose(1)
             
  TmblStr(0) = "Yes"
  TmblStr(1) = "No"
  TmblClose(0) = True
  TmblClose(1) = True
  
 Case vbOKCancel
  ReDim TmblStr(1)
  ReDim TmblClose(1)

  TmblStr(0) = "OK"
  TmblStr(1) = "Cancel"
  TmblClose(0) = True
  TmblClose(1) = True
  
 Case vbRetryCancel
  ReDim TmblStr(1)
  ReDim TmblClose(1)
            
  TmblStr(0) = "Retry"
  TmblStr(1) = "Cancel"
  TmblClose(0) = True
  TmblClose(1) = True

 Case vbYesNoCancel
  ReDim TmblStr(2)
  ReDim TmblClose(2)

  TmblStr(0) = "Yes"
  TmblStr(1) = "No"
  TmblStr(2) = "Cancel"
  TmblClose(0) = True
  TmblClose(1) = True
  TmblClose(2) = True

 Case vbAbortRetryIgnore
  ReDim TmblStr(2)
  ReDim TmblClose(2)
            
  TmblStr(0) = "Abort"
  TmblStr(1) = "Retry"
  TmblStr(2) = "Ignore"
  TmblClose(0) = True
  TmblClose(1) = True
  TmblClose(2) = True

Case Else
 ReDim TmblStr(0)
 ReDim TmblClose(0)

 TmblStr(0) = "OK"
 TmblClose(0) = True
End Select

sMsg = Prompt

With fMessage
 .TeksTombol = TmblStr
 .TmblTerminate = TmblClose
                
 'Gambar Ikon
 TipeIcon = (Buttons And PlusIcons)
  Select Case TipeIcon
   Case vbCritical
    .IkonMsg = vbCritical
   Case vbExclamation
    .IkonMsg = vbExclamation
   Case vbInformation
    .IkonMsg = vbInformation
   Case vbQuestion
    .IkonMsg = vbQuestion
   Case Else
    .IkonMsg = Empty
   End Select
        
   'allignnya jg bisa diset (vbmsgboxright)
   If (Buttons And vbMsgBoxRight) Then
    .lblPesan.Alignment = vbRightJustify
   Else
    .lblPesan.Alignment = vbLeftJustify
   End If
        
   'Taro pesan ke form Msgbox
   .TampilkanPsn Prompt
      
   'tentuin tombol yg dipake u/ nutup msgbox (diambil dr tombol yg diklik-lpressedbtn-)
   Select Case .PressedBtn
    Case 0
    Select Case TipeTmbl
     Case vbAbortRetryIgnore
      TendriMsg = vbAbort
     Case vbOKCancel
      TendriMsg = vbOK
     Case vbRetryCancel
      TendriMsg = vbRetry
     Case vbYesNo
      TendriMsg = vbYes
     Case vbYesNoCancel
      TendriMsg = vbYes
     Case Else
      TendriMsg = vbOK
    End Select
                
    Case 1
    Select Case TipeTmbl
     Case vbAbortRetryIgnore
      TendriMsg = vbRetry
     Case vbOKCancel
      TendriMsg = vbCancel
     Case vbRetryCancel
      TendriMsg = vbCancel
     Case vbYesNo
      TendriMsg = vbNo
     Case vbYesNoCancel
      TendriMsg = vbNo
    End Select
            
    Case 2
    Select Case TipeTmbl
     Case vbAbortRetryIgnore
      TendriMsg = vbIgnore
     Case vbYesNoCancel
      TendriMsg = vbCancel
     End Select
   End Select
   
End With
    
Set fMessage = Nothing
End Function

'''PLEASE RATE'''''PLEASE RATE'''''PLEASE RATE'''
