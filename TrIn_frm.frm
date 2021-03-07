VERSION 5.00
Object = "{DF2BBE39-40A8-433B-A279-073F48DA94B6}#1.0#0"; "axvlc.dll"
Begin VB.Form TrIn_frm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pintu Masuk"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   12945
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pintu Masuk"
      Height          =   7815
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdMasuk 
         Caption         =   "Test Masuk"
         Height          =   495
         Left            =   8880
         TabIndex        =   2
         Top             =   6840
         Width           =   975
      End
      Begin VB.TextBox txtmasuk 
         Height          =   375
         Left            =   7680
         TabIndex        =   1
         Text            =   "0001"
         Top             =   6960
         Width           =   975
      End
      Begin VB.Data datacekMasuk 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin AXVLCCtl.VLCPlugin2 cam1 
         Height          =   6855
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   7215
         AutoLoop        =   0   'False
         AutoPlay        =   -1  'True
         Toolbar         =   -1  'True
         ExtentWidth     =   12726
         ExtentHeight    =   12091
         MRL             =   ""
         Object.Visible         =   -1  'True
         Volume          =   0
         StartTime       =   0
         BaseURL         =   ""
         BackColor       =   0
         FullscreenEnabled=   -1  'True
         Branding        =   -1  'True
      End
      Begin VB.Label lblRfMasuk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RFID"
         Height          =   195
         Left            =   7680
         TabIndex        =   11
         Top             =   5400
         Width           =   375
      End
      Begin VB.Label lblNopolMasuk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nopol"
         Height          =   195
         Left            =   7680
         TabIndex        =   10
         Top             =   5760
         Width           =   420
      End
      Begin VB.Label lblTglmasuk1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   195
         Left            =   7680
         TabIndex        =   9
         Top             =   6120
         Width           =   585
      End
      Begin VB.Label lbljammasuk1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam"
         Height          =   195
         Left            =   7680
         TabIndex        =   8
         Top             =   6480
         Width           =   285
      End
      Begin VB.Label lblpintumasuk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pintu"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lbltransmasuk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "no trans"
         Height          =   195
         Left            =   10560
         TabIndex        =   6
         Top             =   5280
         Width           =   570
      End
      Begin VB.Label lblfilebmp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "no trans"
         Height          =   195
         Left            =   10680
         TabIndex        =   5
         Top             =   5880
         Width           =   570
      End
      Begin VB.Image img1 
         BorderStyle     =   1  'Fixed Single
         Height          =   4560
         Left            =   7560
         Picture         =   "TrIn_frm.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   4800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capture IN"
         Height          =   195
         Left            =   7560
         TabIndex        =   4
         Top             =   360
         Width           =   765
      End
   End
End
Attribute VB_Name = "TrIn_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
