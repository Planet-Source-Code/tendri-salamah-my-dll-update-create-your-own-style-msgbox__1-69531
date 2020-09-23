VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBayarResep 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Billing Resep"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmByrResep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tGrandTtl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox tDiskon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5760
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox tTtl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox tSubttl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5850
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5010
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox cbKdObt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   210
      TabIndex        =   25
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox cbKdPasien 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2093
      TabIndex        =   23
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ComboBox cbKdPtugas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2093
      TabIndex        =   21
      Top             =   600
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5093
      TabIndex        =   20
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   32768
      CalendarForeColor=   12632256
      CalendarTitleBackColor=   8421376
      CalendarTitleForeColor=   16776960
      CalendarTrailingForeColor=   8421440
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   62783491
      CurrentDate     =   39301
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8760
      Top             =   5280
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E3DFE0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   7500
      TabIndex        =   15
      Top             =   6105
      Width           =   7560
      Begin prjApotik.MyButton cmNew 
         Height          =   615
         Left            =   2760
         TabIndex        =   6
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         SPN             =   "MyButtonDefSkin"
         Text            =   "Input"
         TextColorEnabled=   12632064
         TextColorDisabled=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmByrResep.frx":044A
      End
      Begin prjApotik.MyButton cmDel 
         Height          =   615
         Left            =   5400
         TabIndex        =   8
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         SPN             =   "MyButtonDefSkin"
         Text            =   "Delete"
         TextColorEnabled=   12632064
         TextColorDisabled=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmByrResep.frx":089C
      End
      Begin prjApotik.MyButton cmEdit 
         Height          =   615
         Left            =   4080
         TabIndex        =   7
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         SPN             =   "MyButtonDefSkin"
         Text            =   "Edit"
         TextColorEnabled=   12632064
         TextColorDisabled=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmByrResep.frx":0CDA
      End
      Begin prjApotik.MyButton cmBrowse 
         Height          =   615
         Left            =   1320
         TabIndex        =   9
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         SPN             =   "MyButtonDefSkin"
         Text            =   "Browse"
         TextColorEnabled=   12632064
         TextColorDisabled=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmByrResep.frx":112C
      End
      Begin prjApotik.MyButton cmExit 
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         SPN             =   "MyButtonDefSkin"
         Text            =   "Exit"
         TextColorEnabled=   12632064
         TextColorDisabled=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmByrResep.frx":157E
      End
      Begin prjApotik.MyButton cmSave 
         Height          =   615
         Left            =   4800
         TabIndex        =   16
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         SPN             =   "MyButtonDefSkin"
         Text            =   "Save"
         TextColorEnabled=   12632064
         TextColorDisabled=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmByrResep.frx":19D0
      End
      Begin prjApotik.MyButton cmCancel 
         Height          =   615
         Left            =   6120
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         SPN             =   "MyButtonDefSkin"
         Text            =   "Cancel"
         TextColorEnabled=   12632064
         TextColorDisabled=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmByrResep.frx":A0712
      End
      Begin prjApotik.MyButton cmSave2 
         Height          =   615
         Left            =   5400
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         SPN             =   "MyButtonDefSkin"
         Text            =   "Save"
         TextColorEnabled=   12632064
         TextColorDisabled=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmByrResep.frx":A0C54
      End
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   600
      Picture         =   "frmByrResep.frx":13F996
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   14
      Top             =   6120
      Width           =   2250
   End
   Begin VB.TextBox tNmObt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox tNmPasien 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5093
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmByrResep.frx":141EEC
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox tNm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5093
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox tNoFak 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2093
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox tHrg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3330
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Width           =   1575
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2655
      Left            =   240
      TabIndex        =   31
      Top             =   2160
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   -2147483632
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Obat"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Obat"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Harga Obat"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Subtotal"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ListView LVB 
      Height          =   1335
      Left            =   0
      TabIndex        =   38
      Top             =   4800
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   -2147483632
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No Faktur"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tgl Faktur"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Kode Petugas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Kode Pasien"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   4680
      TabIndex        =   37
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Diskon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   4680
      TabIndex        =   35
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   4680
      TabIndex        =   33
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   5850
      TabIndex        =   30
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   5010
      TabIndex        =   28
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Harga Obat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   3330
      TabIndex        =   26
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Nama Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   3773
      TabIndex        =   24
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Kode Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   773
      TabIndex        =   22
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tgl Faktur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   3773
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Nama Obat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   1650
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Kode Obat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Kode Petugas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   773
      TabIndex        =   12
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Nama Petugas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   3773
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "No Faktur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   773
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBayarResep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'%###############################%'
'%Programmed by Tendri S (20)    %'
'%Created on August 05, 2007     %'
'%Update on October 04, 2007     %'
'%with my Updated DataAccessDLL  %'
'%###############################%'
Dim Buttons As Long
Dim lst As ListItem
Dim total As Double

Dim rs As New Recordset
Dim xx As String
Dim RecArray()
Sub RecordArray()
ReDim RecArray(0 To 6)

RecArray(0) = Me.tNoFak
RecArray(1) = Me.DTPicker1
RecArray(2) = Me.cbKdPtugas
RecArray(3) = Me.cbKdPasien
RecArray(4) = Me.tDiskon
RecArray(5) = Me.tTtl
End Sub

Private Sub cmSave_Click()
Dim masuk, hps, tgl As String

Call RecordArray

If AksesDB.CekData("ambilreseph", "nofak", tNoFak) = True Then
 MB.TendriMsg "Duplication Data..!", vbExclamation, "Message"
 Exit Sub
End If

'FOR ADDING NEW DATA, THERE ARE 9 LINES CODE DEPEND ON
'HOM MANY FIELDS THAT YOU MUST BE ADDED TO DATABASE, THAT
'YOU HAVE TO WRITE IT WITH COMMON SQL SINTAKS
'masuk = "insert into ambilreseph values "
'masuk = masuk & "('" & Trim(tNoFak) & "',"
'masuk = masuk & "'" & Me.DTPicker1.Value & "',"
'masuk = masuk & " '" & Trim(cbKdPtugas) & "',"
'masuk = masuk & " '" & Trim(cbKdPasien) & "',"
'masuk = masuk & "'" & Trim(cbKdObt) & "',"
'masuk = masuk & Val(Format(tDiskon, "##0")) & ","
'masuk = masuk & Val(Format(tGrandTtl, "##0")) & ")"
'cn.Execute masuk

'WITH MY DLL
xx = "select * from ambilreseph where nofak=''"
Call AksesDB.InsertData(xx, RecArray)

'With LV
' For I = 1 To .ListItems.Count
'  masuk = "insert into ambilresepd "
'  masuk = masuk & "(nofak,kdobat,hrg,qty)"
'  masuk = masuk & "values "
'  masuk = masuk & "('" & tNoFak & "'"
'  masuk = masuk & ",'" & .ListItems.Item(I).Text & "'"
'  masuk = masuk & "," & Val(Format(.ListItems.Item(I).SubItems(2), "##0")) & ""
'  masuk = masuk & "," & .ListItems.Item(I).SubItems(3) & ")"
'  cn.Execute masuk
'
'  masuk = "update obat set hrgobat="
'  masuk = masuk & Val(Format(.ListItems.Item(I).SubItems(2), "##0")) & ""
'  masuk = masuk & " where kdobat='" & .ListItems.Item(I).Text & "'"
'  cn.Execute masuk
' Next
'End With

For I = 1 To LV.ListItems.Count
 ReDim RecArray(0 To 3)
 RecArray(0) = Me.tNoFak
 RecArray(1) = LV.ListItems.Item(I).Text
 RecArray(2) = LV.ListItems.Item(I).SubItems(2)
 RecArray(3) = LV.ListItems.Item(I).SubItems(3)
 xx = "select * from ambilresepd where nofak=''"
 Call AksesDB.InsertData(xx, RecArray)

 ReDim RecArray(0 To 1)
 RecArray(0) = LV.ListItems.Item(I).Text
 RecArray(1) = LV.ListItems.Item(I).SubItems(2)
 yy = "select kdobat,hrgobat from obat where kdobat='" & LV.ListItems.Item(I).Text & "'"
 Call AksesDB.UpdateData(yy, RecArray, "Hrg Obat")
Next
 
MB.TendriMsg "Data is completely saved..", vbInformation, "Tendri Data Access"
kosong Me
aktif Me, False
LV.ListItems.Clear

cmNew.Visible = True
cmEdit.Visible = True
cmDel.Visible = True
cmCancel.Visible = False
End Sub

Sub totalkan()
total = 0

For I = 1 To LV.ListItems.Count
 total = total + LV.ListItems(I).SubItems(4)
Next

tTtl = total
End Sub

Private Sub cbKdObt_Click()
tNmObt = AksesDB.GetData("obat", "nmobat", "kdobat", cbKdObt)
tHrg = AksesDB.GetData("obat", "hrgobat", "kdobat", cbKdObt)
Text1.SetFocus
End Sub

Private Sub cbKdPasien_Click()
tNmPasien = AksesDB.GetData("Pasien", "nmpasien", "kdpasien", cbKdPasien)
End Sub

Private Sub cbKdPtugas_Click()
tNm = AksesDB.GetData("Petugas", "nmptugas", "kdptugas", cbKdPtugas)
End Sub

Private Sub cmBrowse_Click()
Dim rs As New Recordset
Dim xx As String
Dim dt As String
dt = InputBox("Masukkan No Faktur, Tgl Faktur, Kode Petugas yang dicari :", _
"Input")

xx = "select * from ambilreseph where "
xx = xx & "left(nofak, " & Len(dt) & ")='" & dt & "' "
xx = xx & " or day(tgl)=" & Val(dt) & " or "
xx = xx & " left(kdptugas," & Len(dt) & ")='" & dt & "'"
xx = xx & " order by nofak"
'Set rs = New adodb.Recordset
'rs.Open xx, cn, adOpenDynamic, adLockOptimistic
Set rs = AksesDB.SearchData(xx)

LVB.ListItems.Clear

If rs.State = 0 Then Set rs = AksesDB.SearchData(xx)

If rs.State <> 0 Then
 Do Until rs.EOF
  Set lst = LVB.ListItems.Add(, , rs(0))
  lst.SubItems(1) = rs(1)
  lst.SubItems(2) = rs(2)
  lst.SubItems(3) = rs(3)
 
  rs.MoveNext
 Loop
End If

LVB.Visible = True
End Sub
Private Sub cmCancel_Click()
cmNew.Visible = True
Me.cmEdit.Visible = True
cmDel.Visible = True

MB.TendriMsg "Cancelled..", vbInformation, "Message"

kosong Me
aktif Me, False
WarnaNORMAL Me
End Sub
Private Sub cmDel_Click()
psn = MB.TendriMsg("Are you sure want to Delete Data?", vbYesNo + vbQuestion, "Confirm")
If psn = vbNo Then Exit Sub

'8 LINES CODE
'xx = "delete from ambilreseph where nofak="
'xx = xx & "'" & Trim(tNoFak) & "'"
'Cn.Execute xx
'
'Do While Not LV.ListItems.Clear
' hapus = "delete from ambilresepd where kdobat='"
' hapus = hapus & "'" & Trim(LV.SelectedItem.Text) & "'"
' Cn.Execute hapus
'Loop

'4 LINES CODE with my dll
'Do While Not LV.ListItems.Count <> 0
 Call AksesDB.DeleteData("ambilresepd", "nofak", tNoFak)
'Loop
Call AksesDB.DeleteData("ambilreseph", "nofak", tNoFak)

If tNoFak = "" Then Exit Sub

MB.TendriMsg "Data is completely removed..", vbInformation, "Tendri Data Access"
kosong Me
aktif Me, False
LV.ListItems.Clear
End Sub

Private Sub cmEdit_Click()
cmNew.Visible = False
Me.cmEdit.Visible = False
cmDel.Visible = False
cmSave.Visible = False

aktif Me, True
tNoFak.Enabled = False
End Sub

Private Sub cmExit_Click()
End
End Sub

Private Sub cmNew_Click()
cmNew.Visible = False
Me.cmEdit.Visible = False
cmDel.Visible = False
cmSave2.Visible = False
cmCancel.Visible = True

aktif Me, True
tNoFak.SetFocus
End Sub

'UPDATE DATA
'You can change this data access code with methods of my dll
'-----------------------------------------------------------
'Private Sub cmSave2_Click()
'If cekKosong(Me) = True Then Exit Sub
'
'hps = "delete from ambilreseph where nofak='" & tNoFak & "'"
'cn.Execute hps
'hps = "delete from ambilresepd where nofak='" & tNoFak & "'"
'cn.Execute hps
'
'masuk = "insert into ambilreseph values "
'masuk = masuk & "('" & Trim(tNoFak) & "',"
'masuk = masuk & "'" & Me.DTPicker1.Value & "',"
'masuk = masuk & " '" & Trim(cbKdObt) & "',"
'masuk = masuk & "'" & Trim(cbKdPtugas) & "',"
'masuk = masuk & Val(Format(tDiskon, "##0")) & ","
'masuk = masuk & Val(Format(tGrandTtl, "##0")) & ")"
''MsgBox masuk
'cn.Execute masuk
'
'With LV
' For I = 1 To .ListItems.Count
'  masuk = "insert into ambilresepd values ('" & tNoFak & "'"
'  masuk = masuk & ",'" & .ListItems.Item(I).Text & "'"
'  masuk = masuk & "," & Val(Format(.ListItems.Item(I).SubItems(2), "##0")) & ""
'  masuk = masuk & "," & .ListItems.Item(I).SubItems(3) & ")"
'  cn.Execute masuk
'
'  masuk = "update obat set hrgobat="
'  masuk = masuk & Val(Format(.ListItems.Item(I).SubItems(2), "##0")) & ""
'  masuk = masuk & " where kdobat='" & .ListItems.Item(I).Text & "'"
'  'MsgBox masuk
'  cn.Execute masuk
' Next
'End With
'
'MsgBox "Data has been updated.."
'kosong Me
'aktif Me, False
'
'cmNew.Visible = True
'cmEdit.Visible = True
'cmDel.Visible = True
'End Sub

Private Sub Form_Load()
bukaDB
kosong Me
aktif Me, False
DTPicker1 = Date
Call AksesDB.FillListCombo("kdptugas", "petugas", cbKdPtugas)
Call AksesDB.FillListCombo("kdpasien", "pasien", cbKdPasien)
Call AksesDB.FillListCombo("kdobat", "obat", cbKdObt)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set AksesDB = Nothing
End Sub

Private Sub LVB_DblClick()
Dim rs As New Recordset
Dim cari As String

If LVB.ListItems.Count = 0 Then LVB.Visible = False: Exit Sub

Me.tNoFak = LVB.SelectedItem.Text
Me.DTPicker1 = LVB.SelectedItem.SubItems(1)
Me.cbKdPtugas = LVB.SelectedItem.SubItems(2)
Me.cbKdPasien = LVB.SelectedItem.SubItems(3)
tNm = AksesDB.GetData("petugas", "nmptugas", "kdptugas", LVB.SelectedItem.SubItems(2))
tNmPasien = AksesDB.GetData("pasien", "nmpasien", "kdpasien", LVB.SelectedItem.SubItems(3))

cari = "select * from ambilresepd where nofak='" & tNoFak & "'"
'Set xx = New ADODB.Recordset
'xx.Open cari, cn
Set rs = AksesDB.SearchData(cari)

LV.ListItems.Clear

If rs.State = 0 Then Set rs = AksesDB.SearchData(cari)

If rs.State <> 0 Then
 Do Until rs.EOF
  Set lst = LV.ListItems.Add(, , IIf(IsNull(rs(1)), "", rs(1)))
  lst.SubItems(1) = AksesDB.GetData("obat", "nmobat", "kdobat", IIf(IsNull(rs(1)), "", rs(1)))
  lst.SubItems(2) = AksesDB.GetData("obat", "hrgobat", "kdobat", IIf(IsNull(rs(1)), "", rs(1)))
  lst.SubItems(3) = IIf(IsNull(rs(3)), "", rs(3))
  lst.SubItems(4) = rs(3) * rs(2)
 
  rs.MoveNext
 Loop
End If

LVB.Visible = False
totalkan

cari = "select diskon,total from ambilreseph where nofak='" & tNoFak & "'"
'Set xx = New ADODB.Recordset
'xx.Open CARI, cn
Set rs = AksesDB.SearchData(cari)

If rs.State = 0 Then Set rs = AksesDB.SearchData(cari)

If rs.State <> 0 Then
 tDiskon = rs(0)
 tGrandTtl = rs(1)
End If
End Sub

Private Sub tDiskon_Change()
tGrandTtl = Val(tTtl) - Val(tDiskon)
End Sub

Private Sub Text1_Change()
Me.tSubttl = Val(Me.tHrg) * Val(Me.Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Set lst = LV.ListItems.Add(, , cbKdObt)
 lst.SubItems(1) = tNmObt
 lst.SubItems(2) = tHrg
 lst.SubItems(3) = Text1
 lst.SubItems(4) = tSubttl
 cbKdObt = ""
 tNmObt = ""
 tHrg = ""
 Text1 = ""
 tSubttl = ""
 cbKdObt.SetFocus
 totalkan
End If
End Sub



