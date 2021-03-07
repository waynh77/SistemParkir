VERSION 5.00
Object = "{DF2BBE39-40A8-433B-A279-073F48DA94B6}#1.0#0"; "axvlc.dll"
Begin VB.Form TrOUT_frm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pintu Keluar"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   12885
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pintu Keluar"
      Height          =   7935
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   9240
         TabIndex        =   2
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton CmdKeluar 
         Caption         =   "Test Keluar"
         Height          =   495
         Left            =   10440
         TabIndex        =   1
         Top             =   6840
         Width           =   855
      End
      Begin VB.Data Data2 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin AXVLCCtl.VLCPlugin2 cam2 
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
         Volume          =   50
         StartTime       =   0
         BaseURL         =   ""
         BackColor       =   0
         FullscreenEnabled=   -1  'True
         Branding        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capture OUT"
         Height          =   195
         Left            =   7560
         TabIndex        =   11
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "no trans"
         Height          =   195
         Left            =   11040
         TabIndex        =   10
         Top             =   5280
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "no trans"
         Height          =   195
         Left            =   11160
         TabIndex        =   9
         Top             =   5640
         Width           =   570
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   4560
         Left            =   7560
         Picture         =   "TrOUT_frm.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   4800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam keluar"
         Height          =   195
         Index           =   6
         Left            =   7680
         TabIndex        =   8
         Top             =   6600
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal keluar"
         Height          =   195
         Index           =   1
         Left            =   7680
         TabIndex        =   7
         Top             =   6240
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RFID"
         Height          =   195
         Index           =   11
         Left            =   7680
         TabIndex        =   6
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nopol"
         Height          =   195
         Index           =   9
         Left            =   7680
         TabIndex        =   5
         Top             =   5880
         Width           =   420
      End
      Begin VB.Label lblpintukeluar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pintu"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   825
      End
   End
End
Attribute VB_Name = "TrOUT_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
