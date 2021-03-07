VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   15645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.CommandButton Command3 
         Caption         =   "CARI"
         Height          =   315
         Left            =   13440
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbpintu2 
         Height          =   315
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbPintu3 
         Height          =   315
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdCetak2 
         Caption         =   "CETAK"
         Height          =   315
         Left            =   13440
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cmbsort2 
         Height          =   315
         Left            =   12120
         TabIndex        =   12
         Top             =   1440
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtp3 
         Height          =   315
         Left            =   8160
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   289406977
         CurrentDate     =   42473
      End
      Begin VB.CheckBox cekcari2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tanggal Masuk"
         Height          =   255
         Left            =   6720
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Data Data5 
         Caption         =   "Data5"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   9000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ComboBox cmbsts2 
         Height          =   315
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtcari2 
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cmbcari2 
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin SSDataWidgets_B.SSDBGrid SSDBGrid2 
         Bindings        =   "Form1.frx":0000
         Height          =   7695
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   10815
         _Version        =   196614
         AllowUpdate     =   0   'False
         SelectTypeRow   =   1
         RowHeight       =   423
         Columns.Count   =   8
         Columns(0).Width=   2937
         Columns(0).Caption=   "Tanggal Jam Masuk"
         Columns(0).Name =   "Tanggal Jam Masuk"
         Columns(0).DataField=   "tgljammasuk"
         Columns(0).DataType=   7
         Columns(0).NumberFormat=   "dd MMM yyyy HH:mm:ss"
         Columns(0).FieldLen=   256
         Columns(1).Width=   2090
         Columns(1).Caption=   "Pintu Masuk"
         Columns(1).Name =   "Pintu Masuk"
         Columns(1).DataField=   "pintumasuk"
         Columns(1).FieldLen=   256
         Columns(2).Width=   3016
         Columns(2).Caption=   "Tanggal Jam Keluar"
         Columns(2).Name =   "Tanggal Jam Keluar"
         Columns(2).DataField=   "tgljamkeluar"
         Columns(2).DataType=   7
         Columns(2).NumberFormat=   "dd MMM yyyy HH:mm:ss"
         Columns(2).FieldLen=   256
         Columns(3).Width=   1958
         Columns(3).Caption=   "Pintu Keluar"
         Columns(3).Name =   "Pintu Keluar"
         Columns(3).DataField=   "pintukeluar"
         Columns(3).FieldLen=   256
         Columns(4).Width=   2858
         Columns(4).Caption=   "Nama"
         Columns(4).Name =   "Nama"
         Columns(4).DataField=   "nama"
         Columns(4).FieldLen=   256
         Columns(5).Width=   2302
         Columns(5).Caption=   "Nopol"
         Columns(5).Name =   "Nopol"
         Columns(5).DataField=   "nopol"
         Columns(5).FieldLen=   256
         Columns(6).Width=   2090
         Columns(6).Caption=   "RFID"
         Columns(6).Name =   "RFID"
         Columns(6).DataField=   "rfid"
         Columns(6).FieldLen=   256
         Columns(7).Width=   1349
         Columns(7).Caption=   "Status"
         Columns(7).Name =   "Status"
         Columns(7).DataField=   "sts"
         Columns(7).FieldLen=   256
         _ExtentX        =   19076
         _ExtentY        =   13573
         _StockProps     =   79
         Caption         =   "Transaksi Kartu Member"
         BackColor       =   16777215
      End
      Begin MSComCtl2.DTPicker dtp4 
         Height          =   315
         Left            =   12120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   289406977
         CurrentDate     =   42473
      End
      Begin MSComCtl2.DTPicker dtp7 
         Height          =   315
         Left            =   9480
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   289406978
         CurrentDate     =   42473
      End
      Begin MSComCtl2.DTPicker dtp8 
         Height          =   315
         Left            =   13440
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   289406978
         CurrentDate     =   42473
      End
      Begin VB.Image img6 
         BorderStyle     =   1  'Fixed Single
         Height          =   2640
         Left            =   11160
         Picture         =   "Form1.frx":0014
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   3360
      End
      Begin VB.Image img5 
         BorderStyle     =   1  'Fixed Single
         Height          =   2760
         Left            =   11160
         Picture         =   "Form1.frx":3E6B0
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   3480
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "MASUK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   11160
         TabIndex        =   22
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "KELUAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   11160
         TabIndex        =   21
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pintu Masuk"
         Height          =   195
         Left            =   11040
         TabIndex        =   20
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pintu Keluar"
         Height          =   195
         Left            =   11040
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblJmlData 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data"
         Height          =   255
         Left            =   11040
         TabIndex        =   18
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By"
         Height          =   255
         Left            =   11040
         TabIndex        =   17
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Keluar"
         Height          =   195
         Left            =   10920
         TabIndex        =   10
         Top             =   240
         Width           =   1080
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
