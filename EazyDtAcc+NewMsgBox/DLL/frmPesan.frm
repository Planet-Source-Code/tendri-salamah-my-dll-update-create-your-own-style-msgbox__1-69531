VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4050
   Icon            =   "frmPesan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   "frmPesan.frx":0442
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   1200
      Width           =   2250
   End
   Begin VB.PictureBox picBtns 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   4050
      TabIndex        =   2
      Top             =   720
      Width           =   4050
      Begin Tendri_S.MyButton cmButton 
         Height          =   375
         Index           =   1
         Left            =   1178
         TabIndex        =   4
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         SPN             =   "MyButtonDefSkin"
         Text            =   "Button"
         TextColorEnabled=   12632064
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Tendri_S.MyButton cmButton 
         Height          =   375
         Index           =   2
         Left            =   2018
         TabIndex        =   5
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         SPN             =   "MyButtonDefSkin"
         Text            =   "Button"
         TextColorEnabled=   12632064
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Tendri_S.MyButton cmButton 
         Height          =   375
         Index           =   3
         Left            =   2858
         TabIndex        =   6
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         SPN             =   "MyButtonDefSkin"
         Text            =   "Button"
         TextColorEnabled=   12632064
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Tendri_S.MyButton cmButton 
         Height          =   375
         Index           =   0
         Left            =   338
         TabIndex        =   7
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         SPN             =   "MyButtonDefSkin"
         Text            =   "Button"
         TextColorEnabled=   12632064
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Pic1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblPesan 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2955
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'%#########################################%'
'%Author   : Tendri S (20)                 %'
'%Date     : October 05, 2007              %'
'%Location : Bekasi, Indonesia             %'
'%Email    : mizz_daeng@plasa.com          %'
'%Please Do Not Removes Any Copyrights     %'
'%#########################################%'
Option Explicit

Public Event BtnKliked(ByVal lBtn As Long)
Private Buttons As Integer
Private sTeksTombol() As String
Private bTombolClose() As Boolean
Private lIkon As Long
Private lPressedBtn As Long

Private Const critical As Long = 101
Private Const question As Long = 102
Private Const informasi As Long = 103
Private Const seru As Long = 104
Private Const GapTombol As Long = 105
Private Const GapLabelNormal As Long = 106

Public Property Let JudulMsg(ByVal baru As String)
If baru = "" Then Me.Caption = App.Title
Me.Caption = baru
End Property

Public Property Let IkonMsg(ByVal baru As Long)
lIkon = baru
    
'Load gambar ikon dr resources
If lIkon = vbCritical Then
 'Pic1.Picture = LoadPicture(App.Path & "\Resources\MSGBOX01.ico")
 Pic1.Picture = LoadResPicture(critical, vbResIcon)
 Pic1.Visible = True
ElseIf lIkon = vbQuestion Then
 Pic1.Picture = LoadResPicture(question, vbResIcon)
 Pic1.Visible = True
ElseIf lIkon = vbInformation Then
 Pic1.Picture = LoadResPicture(informasi, vbResIcon)
 Pic1.Visible = True
ElseIf lIkon = vbExclamation Then
 'Pic1.Picture = LoadPicture(App.Path & "\Resources\MSGBOX03.ico")
 Pic1.Picture = LoadResPicture(seru, vbResIcon)
 Pic1.Visible = True
Else
 Pic1.Visible = False
End If
End Property

Public Property Let TeksTombol(baru() As String)
Dim ljml As Long
Dim lbts As Long
    
sTeksTombol = baru
    
'set visible tombol2nya
For ljml = cmButton.LBound To cmButton.UBound
 cmButton(ljml).Visible = False
Next
For ljml = 0 To UBound(sTeksTombol)
 cmButton(ljml).Visible = True
Next
End Property

Public Property Let TmblTerminate(baru() As Boolean)
bTombolClose = baru
End Property

Public Property Get PressedBtn() As Long
PressedBtn = lPressedBtn
End Property

'Display Message
Public Sub TampilkanPsn(sPesan As String)
Dim lPemisah As Long
Dim lbts As Long
Dim ljml As Long
    
lblPesan.Caption = sPesan
    
'tinggi label pesan bertambah heightnya otomatis sesuai teks pesannya
lPemisah = picBtns.Top - (lblPesan.Top + lblPesan.Height)
If lPemisah < GapLabelNormal Then
 Me.Height = Me.Height + (GapLabelNormal - lPemisah)
End If
    
'set posisi tombol2 sesuai dgn lebar form
ljml = UBound(sTeksTombol) + 1
Select Case ljml
 Case 1
  cmButton.Item(0).Left = (Me.Width - cmButton(0).Width - lbts) \ 2
  cmButton.Item(0).Text = sTeksTombol(0)
 Case 2
  cmButton.Item(0).Left = ((Me.Width - GapTombol - lbts) \ 2) - cmButton(0).Width
  cmButton.Item(1).Left = GapTombol + cmButton(0).Width + cmButton(0).Left
  cmButton.Item(0).Text = sTeksTombol(0)
  cmButton.Item(1).Text = sTeksTombol(1)
 Case 3
  cmButton(1).Left = (Me.Width - cmButton(0).Width - lbts) \ 2
  cmButton(0).Left = cmButton(1).Left - cmButton(0).Width - GapTombol
  cmButton(2).Left = cmButton(1).Left + cmButton(1).Width + GapTombol
  cmButton(0).Text = sTeksTombol(0)
  cmButton(1).Text = sTeksTombol(1)
  cmButton(2).Text = sTeksTombol(2)
 End Select
    
 Me.Show 1
End Sub

Private Sub cmButton_Click(Index As Integer)
'index tombol yg telah diklik
lPressedBtn = Index
    
'Unload atau raise event sesuai arraynya
If bTombolClose(lPressedBtn) Then
 Unload Me
' MyPath = App.Path

Else
 RaiseEvent BtnKliked(Index)
End If
End Sub

Private Sub cmButton_MouseHover(Index As Integer)
'Dim ljml1 As Long
'Dim ljml2 As Long
'ljml1 = UBound(sTeksTombol)
''
'For ljml2 = 0 To UBound(sTeksTombol) + 1
' cmButton(0).TextColorEnabled = &H4000&
' cmButton(0).FontSize = 9
' If ljml2 = 2 Then
'
' End If
'Next
End Sub

Private Sub cmButton_MouseOut(Index As Integer)
'Dim ljml As Long
'Dim ljml2 As Long
'ljml1 = UBound(sTeksTombol)
''
'For ljml2 = 0 To UBound(sTeksTombol) + 1
'' cmButton(ljml1).TextColorEnabled = &HC0C000
'' If ljml2 = 2 Then
''
'' End If
''Next
'Next
End Sub

Private Sub Form_Load()

End Sub

'''PLEASE RATE'''''PLEASE RATE'''''PLEASE RATE'''
