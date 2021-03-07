VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form TrLast_frm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Transaksi Terakhir"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   12945
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaksi Terakhir"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.Data DataTrans 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   1140
      End
      Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
         Bindings        =   "TrLast_frm.frx":0000
         Height          =   6855
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7215
         _Version        =   196614
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   12726
         _ExtentY        =   12091
         _StockProps     =   79
         Caption         =   "Data Transaksi Terakhir"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Trans"
         Height          =   195
         Index           =   10
         Left            =   7680
         TabIndex        =   8
         Top             =   5160
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IN OUT"
         Height          =   195
         Index           =   8
         Left            =   7680
         TabIndex        =   7
         Top             =   5880
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gate Name"
         Height          =   195
         Index           =   7
         Left            =   7680
         TabIndex        =   6
         Top             =   5520
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nopol"
         Height          =   195
         Index           =   4
         Left            =   7680
         TabIndex        =   5
         Top             =   6600
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RFID"
         Height          =   195
         Index           =   3
         Left            =   7680
         TabIndex        =   4
         Top             =   6240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   195
         Index           =   2
         Left            =   7680
         TabIndex        =   3
         Top             =   6960
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam "
         Height          =   195
         Index           =   0
         Left            =   9240
         TabIndex        =   2
         Top             =   6960
         Width           =   330
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   4560
         Left            =   7680
         Picture         =   "TrLast_frm.frx":0018
         Stretch         =   -1  'True
         Top             =   480
         Width           =   4800
      End
   End
End
Attribute VB_Name = "TrLast_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
