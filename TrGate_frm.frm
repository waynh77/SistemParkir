VERSION 5.00
Object = "{DF2BBE39-40A8-433B-A279-073F48DA94B6}#1.0#0"; "axvlc.dll"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form TrGate_frm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Transaksi Pintu"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15510
   Icon            =   "TrGate_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   15510
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   15901
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "PINTU"
      TabPicture(0)   =   "TrGate_frm.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(1)"
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "TRANSAKSI KARTU MASTER"
      TabPicture(1)   =   "TrGate_frm.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "TRANSAKSI KARTU MEMBER"
      TabPicture(2)   =   "TrGate_frm.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "Data5"
      Tab(2).Control(3)=   "SSDBGrid2"
      Tab(2).Control(4)=   "Label13"
      Tab(2).Control(5)=   "Label12"
      Tab(2).Control(6)=   "img5"
      Tab(2).Control(7)=   "img6"
      Tab(2).ControlCount=   8
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   -64080
         TabIndex        =   70
         Top             =   1080
         Width           =   3975
         Begin VB.CommandButton Command3 
            Caption         =   "CARI"
            Height          =   315
            Left            =   2520
            TabIndex        =   75
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cmbpintu2 
            Height          =   315
            Left            =   1200
            TabIndex        =   74
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cmbPintu3 
            Height          =   315
            Left            =   1200
            TabIndex        =   73
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton cmdCetak2 
            Caption         =   "CETAK"
            Height          =   315
            Left            =   2520
            TabIndex        =   72
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cmbsort2 
            Height          =   315
            Left            =   1200
            TabIndex        =   71
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pintu Masuk"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pintu Keluar"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblJmlData 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Data"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   1440
            Width           =   3615
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort By"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   -74880
         TabIndex        =   59
         Top             =   360
         Width           =   14775
         Begin MSComCtl2.DTPicker dtp3 
            Height          =   315
            Left            =   8160
            TabIndex        =   63
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98566145
            CurrentDate     =   42473
         End
         Begin VB.CheckBox cekcari2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tanggal Masuk"
            Height          =   255
            Left            =   6720
            TabIndex        =   64
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbsts2 
            Height          =   315
            Left            =   600
            TabIndex        =   62
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtcari2 
            Height          =   315
            Left            =   3840
            TabIndex        =   61
            Top             =   240
            Width           =   2775
         End
         Begin VB.ComboBox cmbcari2 
            Height          =   315
            Left            =   2520
            TabIndex        =   60
            Top             =   240
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtp7 
            Height          =   315
            Left            =   9480
            TabIndex        =   65
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98566146
            CurrentDate     =   42473
         End
         Begin MSComCtl2.DTPicker dtp4 
            Height          =   315
            Left            =   12000
            TabIndex        =   67
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98566145
            CurrentDate     =   42473
         End
         Begin MSComCtl2.DTPicker dtp8 
            Height          =   315
            Left            =   13320
            TabIndex        =   68
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98566146
            CurrentDate     =   42473
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Keluar"
            Height          =   195
            Left            =   10800
            TabIndex        =   69
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Data Data5 
         Caption         =   "Data5"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -67680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3840
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   8295
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   14775
         Begin VB.ComboBox cmbSort 
            Height          =   315
            Left            =   11520
            TabIndex        =   58
            Top             =   2040
            Width           =   2655
         End
         Begin VB.CommandButton Command4 
            Caption         =   "CETAK"
            Height          =   315
            Left            =   10560
            TabIndex        =   55
            Top             =   3000
            Width           =   3615
         End
         Begin VB.Data Data6 
            Caption         =   "Data6"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   7800
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1800
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.ComboBox cmbpintu 
            Height          =   315
            Left            =   11520
            TabIndex        =   49
            Top             =   1680
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker dtp1 
            Height          =   315
            Left            =   11520
            TabIndex        =   46
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98566145
            CurrentDate     =   42473
         End
         Begin VB.CheckBox cekcari 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tanggal"
            Height          =   255
            Left            =   10560
            TabIndex        =   45
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cmbSts 
            Height          =   315
            Left            =   11400
            TabIndex        =   44
            Top             =   240
            Width           =   2775
         End
         Begin VB.CommandButton Command2 
            Caption         =   "CARI"
            Height          =   315
            Left            =   10560
            TabIndex        =   40
            Top             =   2400
            Width           =   3615
         End
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   11760
            TabIndex        =   39
            Top             =   600
            Width           =   2415
         End
         Begin VB.ComboBox cmbCari 
            Height          =   315
            Left            =   10560
            TabIndex        =   38
            Top             =   600
            Width           =   1215
         End
         Begin VB.Data DataTrans 
            Caption         =   "Data3"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   7800
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2520
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Data Data3 
            Caption         =   "Data3"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   7800
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2160
            Visible         =   0   'False
            Width           =   1140
         End
         Begin MSComCtl2.DTPicker dtp2 
            Height          =   315
            Left            =   11520
            TabIndex        =   47
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98566145
            CurrentDate     =   42473
         End
         Begin MSComCtl2.DTPicker dtp5 
            Height          =   315
            Left            =   12840
            TabIndex        =   53
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98566146
            CurrentDate     =   42473
         End
         Begin MSComCtl2.DTPicker dtp6 
            Height          =   315
            Left            =   12840
            TabIndex        =   54
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98566146
            CurrentDate     =   42473
         End
         Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
            Bindings        =   "TrGate_frm.frx":0D1E
            Height          =   7695
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   10095
            _Version        =   196614
            AllowUpdate     =   0   'False
            SelectTypeRow   =   1
            RowHeight       =   423
            Columns.Count   =   7
            Columns(0).Width=   2514
            Columns(0).Caption=   "Tanggal"
            Columns(0).Name =   "Tanggal"
            Columns(0).DataField=   "tgljam"
            Columns(0).DataType=   7
            Columns(0).NumberFormat=   "dd MMM yyyy"
            Columns(0).FieldLen=   256
            Columns(1).Width=   1931
            Columns(1).Caption=   "Jam"
            Columns(1).Name =   "Jam"
            Columns(1).DataField=   "tgljam"
            Columns(1).DataType=   7
            Columns(1).NumberFormat=   "HH:mm:ss"
            Columns(1).FieldLen=   256
            Columns(2).Width=   1111
            Columns(2).Caption=   "In/Out"
            Columns(2).Name =   "In/Out"
            Columns(2).DataField=   "sts"
            Columns(2).FieldLen=   256
            Columns(3).Width=   2275
            Columns(3).Caption=   "Nama Pintu"
            Columns(3).Name =   "Nama Pintu"
            Columns(3).DataField=   "namapintu"
            Columns(3).FieldLen=   256
            Columns(4).Width=   3757
            Columns(4).Caption=   "Nama"
            Columns(4).Name =   "Nama"
            Columns(4).DataField=   "nama"
            Columns(4).FieldLen=   256
            Columns(5).Width=   2275
            Columns(5).Caption=   "Nopol"
            Columns(5).Name =   "Nopol"
            Columns(5).DataField=   "nopol"
            Columns(5).FieldLen=   256
            Columns(6).Width=   3200
            Columns(6).Caption=   "RFID No"
            Columns(6).Name =   "RFID No"
            Columns(6).DataField=   "rfid"
            Columns(6).FieldLen=   256
            _ExtentX        =   17806
            _ExtentY        =   13573
            _StockProps     =   79
            Caption         =   "Transaksi Kartu Master"
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort By"
            Height          =   255
            Left            =   10560
            TabIndex        =   57
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label lblJml 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Data"
            Height          =   255
            Left            =   10560
            TabIndex        =   56
            Top             =   2760
            Width           =   3615
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Pintu"
            Height          =   255
            Left            =   10560
            TabIndex        =   50
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   255
            Left            =   10560
            TabIndex        =   43
            Top             =   240
            Width           =   855
         End
         Begin VB.Image img3 
            BorderStyle     =   1  'Fixed Single
            Height          =   2880
            Left            =   10560
            Picture         =   "TrGate_frm.frx":0D32
            Stretch         =   -1  'True
            Top             =   5160
            Width           =   3600
         End
         Begin VB.Label lbljamtr 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jam "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   13560
            TabIndex        =   30
            Top             =   4800
            Width           =   585
         End
         Begin VB.Label lbltgltr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10560
            TabIndex        =   29
            Top             =   4800
            Width           =   975
         End
         Begin VB.Label lblrfidtrans 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RFID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10560
            TabIndex        =   28
            Top             =   4080
            Width           =   660
         End
         Begin VB.Label lblNopoltr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nopol"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10560
            TabIndex        =   27
            Top             =   4440
            Width           =   705
         End
         Begin VB.Label lblgate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gate Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10560
            TabIndex        =   26
            Top             =   3360
            Width           =   1380
         End
         Begin VB.Label lblinout 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IN OUT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10560
            TabIndex        =   25
            Top             =   3720
            Width           =   900
         End
         Begin VB.Label lblnotrans 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Trans"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7800
            TabIndex        =   24
            Top             =   3840
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pintu Keluar"
         Height          =   8295
         Index           =   1
         Left            =   -68280
         TabIndex        =   13
         Top             =   480
         Width           =   8175
         Begin MSCommLib.MSComm comKeluar 
            Left            =   720
            Top             =   840
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
         Begin VB.Data Data2 
            Caption         =   "Data1"
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
            Top             =   240
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.CommandButton CmdKeluar 
            Caption         =   "On"
            Height          =   495
            Left            =   7080
            TabIndex        =   14
            Top             =   4080
            Width           =   855
         End
         Begin AXVLCCtl.VLCPlugin2 cam2 
            Height          =   3615
            Left            =   240
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   600
            Width           =   5775
            AutoLoop        =   0   'False
            AutoPlay        =   -1  'True
            Toolbar         =   0   'False
            ExtentWidth     =   10186
            ExtentHeight    =   6376
            MRL             =   ""
            Object.Visible         =   -1  'True
            Volume          =   0
            StartTime       =   0
            BaseURL         =   ""
            BackColor       =   0
            FullscreenEnabled=   0   'False
            Branding        =   -1  'True
         End
         Begin VB.Label lblNamaKeluar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6120
            TabIndex        =   37
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keluar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   6120
            TabIndex        =   36
            Top             =   2880
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Masuk"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   6120
            TabIndex        =   35
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jam"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6120
            TabIndex        =   34
            Top             =   2520
            Width           =   510
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6120
            TabIndex        =   33
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Masuk"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   4440
            Width           =   480
         End
         Begin VB.Image img4 
            BorderStyle     =   1  'Fixed Single
            Height          =   3000
            Left            =   240
            Picture         =   "TrGate_frm.frx":3F3CE
            Stretch         =   -1  'True
            Top             =   4680
            Width           =   3480
         End
         Begin VB.Label lblpintukeluar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Pintu"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   825
         End
         Begin VB.Label lblnopolkeluar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nopol"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6120
            TabIndex        =   21
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label lblrfkeluar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RFID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6120
            TabIndex        =   20
            Top             =   600
            Width           =   660
         End
         Begin VB.Label lbltglkeluar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal keluar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6120
            TabIndex        =   19
            Top             =   3240
            Width           =   1785
         End
         Begin VB.Label lbljamkeluar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jam keluar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6120
            TabIndex        =   18
            Top             =   3600
            Width           =   1320
         End
         Begin VB.Image img2 
            BorderStyle     =   1  'Fixed Single
            Height          =   3000
            Left            =   4440
            Picture         =   "TrGate_frm.frx":7DA6A
            Stretch         =   -1  'True
            Top             =   4680
            Width           =   3480
         End
         Begin VB.Label lblfilebmp2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no trans"
            Height          =   195
            Left            =   2520
            TabIndex        =   17
            Top             =   6600
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lbltranskeluar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no trans"
            Height          =   195
            Left            =   1080
            TabIndex        =   16
            Top             =   5760
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keluar"
            Height          =   195
            Left            =   4440
            TabIndex        =   15
            Top             =   4440
            Width           =   450
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pintu Masuk"
         Height          =   8295
         Index           =   0
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   6255
         Begin VB.Data Data4 
            Caption         =   "Data4"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   4320
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   960
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.CommandButton Command1 
            Caption         =   "On"
            Height          =   495
            Left            =   4920
            TabIndex        =   12
            Top             =   7080
            Width           =   975
         End
         Begin MSCommLib.MSComm comMasuk 
            Left            =   2280
            Top             =   720
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
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
         Begin VB.Data datacekMasuk 
            Caption         =   "Data3"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   2160
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   240
            Visible         =   0   'False
            Width           =   1140
         End
         Begin AXVLCCtl.VLCPlugin2 cam1 
            Height          =   3615
            Left            =   240
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   600
            Width           =   5775
            AutoLoop        =   0   'False
            AutoPlay        =   -1  'True
            Toolbar         =   0   'False
            ExtentWidth     =   10186
            ExtentHeight    =   6376
            MRL             =   ""
            Object.Visible         =   -1  'True
            Volume          =   0
            StartTime       =   0
            BaseURL         =   ""
            BackColor       =   0
            FullscreenEnabled=   0   'False
            Branding        =   -1  'True
         End
         Begin VB.Label lblnamaMasuk 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nopol"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3960
            TabIndex        =   41
            Top             =   5280
            Width           =   705
         End
         Begin VB.Label lblRfMasuk 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RFID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3960
            TabIndex        =   11
            Top             =   4920
            Width           =   660
         End
         Begin VB.Label lblNopolMasuk 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nopol"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3960
            TabIndex        =   10
            Top             =   5640
            Width           =   705
         End
         Begin VB.Label lblTglmasuk1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3960
            TabIndex        =   9
            Top             =   6000
            Width           =   975
         End
         Begin VB.Label lbljammasuk1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jam"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3960
            TabIndex        =   8
            Top             =   6360
            Width           =   510
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
            Left            =   1080
            TabIndex        =   6
            Top             =   6240
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lblfilebmp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no trans"
            Height          =   195
            Left            =   1560
            TabIndex        =   5
            Top             =   7200
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Image img1 
            BorderStyle     =   1  'Fixed Single
            Height          =   2880
            Left            =   240
            Picture         =   "TrGate_frm.frx":BC106
            Stretch         =   -1  'True
            Top             =   4800
            Width           =   3480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capture IN"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   4560
            Width           =   765
         End
      End
      Begin SSDataWidgets_B.SSDBGrid SSDBGrid2 
         Bindings        =   "TrGate_frm.frx":FA7A2
         Height          =   7575
         Left            =   -74880
         TabIndex        =   42
         Top             =   1200
         Width           =   10575
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
         _ExtentX        =   18653
         _ExtentY        =   13361
         _StockProps     =   79
         Caption         =   "Transaksi Kartu Member"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   -63960
         TabIndex        =   52
         Top             =   6000
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   -63960
         TabIndex        =   51
         Top             =   3120
         Width           =   945
      End
      Begin VB.Image img5 
         BorderStyle     =   1  'Fixed Single
         Height          =   2880
         Left            =   -64080
         Picture         =   "TrGate_frm.frx":FA7B6
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   3960
      End
      Begin VB.Image img6 
         BorderStyle     =   1  'Fixed Single
         Height          =   2880
         Left            =   -64080
         Picture         =   "TrGate_frm.frx":138E52
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   3960
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSAKSI GATE/PINTU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "TrGate_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cekMasuk_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub cmdCetak2_Click()
With CtkTransMember
    .Show
    .WindowState = 2
    .Data1.DatabaseName = db
    .Data1.Connect = dbCon
    .Data1.RecordSource = Data5.RecordSource
End With
End Sub

Private Sub CmdKeluar_Click()
If CmdKeluar.Caption = "On" Then
    CmdKeluar.Caption = "Off"
    bukaPortKeluar
Else
    CmdKeluar.Caption = "On"
    kosongKeluar
    If comKeluar.PortOpen Then
        comKeluar.PortOpen = False
    End If
End If

End Sub

Sub ProsesKeluar()
datacekMasuk.RecordSource = "Select * from msmember where rfid=" & Trim(kutip(lblrfkeluar.Caption))
datacekMasuk.Refresh
With datacekMasuk.Recordset
    If .BOF Then
        playsound = sndPlaySound(App.Path + "\voice\notavail.wav", 1)
        'MsgBox "RFID tidak diketemukan"
    Else
        If !tglexp < Now Then
            playsound = sndPlaySound(App.Path + "\voice\exp.wav", 1)
        Else
            lblnopolkeluar.Caption = !nopol
            lbltglkeluar.Caption = Format(Now, "dd mmm yyyy")
            lbljamkeluar.Caption = Format(Now, "HH:MM:SS")
            lblNamaKeluar.Caption = !nama
            If !master = 0 Then
                Data4.RecordSource = "select * from trmember where rfid=" + kutip(!rfid) + " and sts=0"
                Data4.Refresh
                If Data4.Recordset.BOF Then
                    playsound = sndPlaySound(App.Path + "\voice\diarea.wav", 1)
                    Exit Sub
                Else
                    lbltranskeluar.Caption = Data4.Recordset!notrans
                    
                    'tampil gbrmasuk
                    If Dir$(App.Path + "\gbr\" + Data4.Recordset!notrans + "-i.jpg") <> "" Then
                        img4.Picture = LoadPicture(App.Path + "\gbr\" + Data4.Recordset!notrans + "-i.jpg")
                    Else
                        img4.Picture = LoadPicture(App.Path + "\logo eq.jpg")
                    End If
                    Label7.Caption = Format(Data4.Recordset!tgljammasuk, "dd MMM yyyy")
                    Label8.Caption = Format(Data4.Recordset!tgljammasuk, "HH:mm:ss")
                End If
            Else
                lbltranskeluar.Caption = "OUT-" + Format(Now, "yyyymmddHHMMSS") + "-" + !rfid
            End If
            
            'simpan transaksi
            If !master = 0 Then
                Data4.Recordset.Edit
                'Data4.Recordset!notrans = lbltransmasuk.Caption
                'Data4.Recordset!rfid = lblrfkeluar.Caption
                Data4.Recordset!tgljamkeluar = Now
                Data4.Recordset!sts = -1
                Data4.Recordset!pintukeluar = lblpintukeluar.Caption
                Data4.Recordset.Update
            Else
                DataTrans.Recordset.AddNew
                DataTrans.Recordset!notrans = lbltranskeluar.Caption
                DataTrans.Recordset!rfid = lblrfkeluar.Caption
                DataTrans.Recordset!tgljam = Now
                DataTrans.Recordset!sts = "OUT"
                DataTrans.Recordset!namapintu = lblpintukeluar.Caption
                DataTrans.Recordset.Update
            End If
            
            'capture gambar
            Clipboard.Clear
            cam2.video.takeSnapshot
            Dim tgljam As String
            tgljam = Now
            
            Dim sFile As String
            sFile = Dir(App.Path & "\*.bmp")
            Do Until sFile = ""
                If tgljam = FileDateTime(sFile) Then
                    lblfilebmp2.Caption = sFile
                End If
            sFile = Dir
            Loop
            
            'rename file bmp
            Dim Filename As String
            Dim NewFileName As String
            Dim fileJpg As String
            
            Filename = App.Path & "\" & lblfilebmp2.Caption
            NewFileName = App.Path & "\" & lbltranskeluar.Caption & ".bmp"
            fileJpg = App.Path & "\gbr\" & lbltranskeluar.Caption & "-o.jpg"
            
            Name Filename As NewFileName
                        
            'convert bmp to jpg
            dib = FreeImage_LoadEx(NewFileName)
            If (dib) Then
               Call FreeImage_SaveEx(dib, fileJpg)
               Call FreeImage_Unload(dib)
            End If
            
            'delete file bmp
            DeleteFile (NewFileName)
                    
            'tampil hasil capture
            If Dir$(fileJpg) <> "" Then
                img2.Picture = LoadPicture(fileJpg)
            Else
                img2.Picture = LoadPicture(App.Path + "\logo eq.jpg")
            End If
            
            'play sound
            playsound = sndPlaySound(Data2.Recordset!lokasifile, 1)
            
            'open gate
            BukaPintu2
            
            'update grid
            Data3.Refresh
            Data5.Refresh
            
            'bukaPortKeluar
        End If
    End If
End With

End Sub

Sub ProsesMasuk()
datacekMasuk.RecordSource = "Select * from msmember where rfid=" & Trim(kutip(lblRfMasuk.Caption))
datacekMasuk.Refresh
With datacekMasuk.Recordset
    If .BOF Then
        'MsgBox "RFID tidak diketemukan"
        playsound = sndPlaySound(App.Path + "\voice\notavail.wav", 1)
        'txtmasuk.Text = ""
        'txtmasuk.SetFocus
    Else
        If !tglexp < Now Then
            playsound = sndPlaySound(App.Path + "\voice\exp.wav", 1)
        Else
            lblRfMasuk.Caption = !rfid
            lblNopolMasuk.Caption = !nopol
            lblTglmasuk1.Caption = Format(Now, "dd mmm yyyy")
            lbljammasuk1.Caption = Format(Now, "HH:MM:SS")
            lblnamaMasuk.Caption = !nama
            If !master = 0 Then
                Data4.RecordSource = "select * from trmember where rfid=" + kutip(!rfid) + " and sts=0"
                Data4.Refresh
                If Data4.Recordset.BOF Then
                    lbltransmasuk.Caption = Format(Now, "yyyymmddHHMMSS") + "-" + !rfid
                Else
                    playsound = sndPlaySound(App.Path + "\voice\diarea.wav", 1)
                    Exit Sub
                End If
            Else
                lbltransmasuk.Caption = "IN-" + Format(Now, "yyyymmddHHMMSS") + "-" + !rfid
            End If
            
            'simpan transaksi
            If !master = 0 Then
                Data4.Recordset.AddNew
                Data4.Recordset!notrans = lbltransmasuk.Caption
                Data4.Recordset!rfid = lblRfMasuk.Caption
                Data4.Recordset!tgljammasuk = Now
                Data4.Recordset!sts = 0
                Data4.Recordset!pintuMasuk = lblpintumasuk.Caption
                Data4.Recordset.Update
            Else
                DataTrans.Recordset.AddNew
                DataTrans.Recordset!notrans = lbltransmasuk.Caption
                DataTrans.Recordset!rfid = lblRfMasuk.Caption
                DataTrans.Recordset!tgljam = Now
                DataTrans.Recordset!sts = "IN"
                DataTrans.Recordset!namapintu = lblpintumasuk.Caption
                DataTrans.Recordset.Update
            End If
            
            'capture gambar
            Clipboard.Clear
            cam1.video.takeSnapshot
            Dim tgljam As String
            tgljam = Now
            
            Dim sFile As String
            sFile = Dir(App.Path & "\*.bmp")
            Do Until sFile = ""
                If tgljam = FileDateTime(sFile) Then
                    lblfilebmp.Caption = sFile
                End If
            sFile = Dir
            Loop
            
            'rename file bmp
            Dim Filename As String
            Dim NewFileName As String
            Dim fileJpg As String
            
            Filename = App.Path & "\" & lblfilebmp.Caption
            NewFileName = App.Path & "\" & lbltransmasuk.Caption & ".bmp"
            fileJpg = App.Path & "\gbr\" & lbltransmasuk.Caption & "-i.jpg"
            
            Name Filename As NewFileName
                        
            'convert bmp to jpg
            dib = FreeImage_LoadEx(NewFileName)
            If (dib) Then
               Call FreeImage_SaveEx(dib, fileJpg)
               Call FreeImage_Unload(dib)
            End If
            
            'delete file bmp
            DeleteFile (NewFileName)
                    
            'tampil hasil capture
            If Dir$(fileJpg) <> "" Then
                img1.Picture = LoadPicture(fileJpg)
            Else
                img1.Picture = LoadPicture(App.Path + "\logo eq.jpg")
            End If
            
            'play sound
            playsound = sndPlaySound(Data1.Recordset!lokasifile, 1)
            
            'open gate
            bukaPintu
            
            'update grid
            Data3.Refresh
            Data5.Refresh
            
            'bukaPortMasuk

        End If
    End If
End With
End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Private Sub comKeluar_OnComm()
kosongKeluar
If comKeluar.PortOpen Then
    lblrfkeluar.Caption = Mid(comKeluar.Input, 2, 10)
    comKeluar.PortOpen = False
    If Len(lblrfkeluar.Caption) = 10 Then
        'MsgBox "Baca rfid sukses"
        ProsesKeluar
        'BukaPintu2
        bukaPortKeluar
    End If
Else
    'MsgBox "Port Close"
    bukaPortKeluar
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "On" Then
    Command1.Caption = "Off"
    bukaPortMasuk
Else
    Command1.Caption = "On"
    kosongMasuk
    If comMasuk.PortOpen Then
        comMasuk.PortOpen = False
    End If
End If
End Sub

Sub bukaPortMasuk()
    comMasuk.CommPort = Data1.Recordset!comMasuk
    comMasuk.PortOpen = True
    comMasuk.RTSEnable = True
    comMasuk.RThreshold = 14
    comMasuk.InputLen = 0
End Sub

Sub bukaPortKeluar()
    comKeluar.CommPort = Data2.Recordset!comKeluar
    comKeluar.PortOpen = True
    comKeluar.RTSEnable = True
    comKeluar.RThreshold = 14
    comKeluar.InputLen = 0
End Sub

Sub bukaPintu()
        comMasuk.CommPort = Data1.Recordset!comMasuk
        comMasuk.PortOpen = True

        If comMasuk.PortOpen = False Then
           MsgBox "Comm Port tidak berfungsi !", 16, "Kesalahan"
            Return
        End If

        comMasuk.RTSEnable = True
        comMasuk.DTREnable = False
        Sleep 200
        comMasuk.DTREnable = True
        Sleep 200
        comMasuk.PortOpen = False

End Sub

Sub BukaPintu2()
        comKeluar.CommPort = Data2.Recordset!comKeluar
        comKeluar.PortOpen = True

        If comKeluar.PortOpen = False Then
           MsgBox "Comm Port tidak berfungsi !", 16, "Kesalahan"
            Return
        End If

        comKeluar.RTSEnable = True
        comKeluar.DTREnable = False
        Sleep 200
        comKeluar.DTREnable = True
        Sleep 200
        comKeluar.PortOpen = False
End Sub

Private Sub Command2_Click() 'cari transaksi kartu master
Dim sql As String
sql = "select * from view_trans1 "
Select Case cmbSts.ListIndex
Case 0
    sql = sql + "where sts like '*' "
Case 1
    sql = sql + "where sts='IN' "
Case 2
    sql = sql + "where sts='OUT' "
End Select

Select Case cmbCari.ListIndex
Case 0
    sql = sql + "and nama like " & kutip("*" + txtCari.Text + "*")
Case 1
    sql = sql + "and nopol like " & kutip("*" + txtCari.Text + "*")
Case 2
    sql = sql + "and rfid like " & kutip("*" + txtCari.Text + "*")
End Select

If cekcari.Value = 1 Then
    If dtp1.Value = dtp2.Value Then
        sql = sql + " and format(tgljam,'mm/dd/yyyy') = '" + Format(dtp1.Value, "mm/dd/yyyy") + "'"
        'sql = sql + " and format(tgljam,'mm/dd/yyyy') = '" + Format(dtp1.Value, "mm/dd/yyyy") + "'"
    Else
        sql = sql + " and format(tgljam,'mm/dd/yyyy HH:mm:ss') >= '" + Format(dtp1.Value, "mm/dd/yyyy ") + _
        Format(dtp5.Value, "HH:mm:ss") + _
        "' and format(tgljam,'mm/dd/yyyy HH:mm:ss') <= '" + Format(dtp2.Value + 1, "mm/dd/yyyy ") + _
        Format(dtp6.Value, "HH:mm:ss") + "'"
    End If
End If

If cmbpintu.ListIndex > 0 Then
    sql = sql + " and namapintu=" & kutip(cmbpintu)
End If

Select Case cmbSort.ListIndex
Case 0
    sql = sql + " order by tgljam"
Case 1
    sql = sql + " order by sts"
Case 2
    sql = sql + " order by namapintu"
Case 3
    sql = sql + " order by nama"
Case 4
    sql = sql + " order by nopol"
Case 5
    sql = sql + " order by rfid"

End Select

Data3.RecordSource = sql
Data3.Refresh
Dim jmldata As Integer
jmldata = Data3.Recordset.RecordCount
lblJml.Caption = "Jumlah Data = " & jmldata
End Sub

Private Sub Command3_Click() 'cari transaksi member
Dim sql As String
sql = "select * from view_trmember where "
Select Case cmbsts2.ListIndex
Case 0
    sql = sql + "sts like '*' "
Case 1
    sql = sql + "sts=true "
Case 2
    sql = sql + "sts=0 "
End Select

Select Case cmbcari2.ListIndex
Case 0
    sql = sql + "and nama like " & kutip("*" + txtcari2.Text + "*")
Case 1
    sql = sql + "and nopol like " & kutip("*" + txtcari2.Text + "*")
Case 2
    sql = sql + "and rfid like " & kutip("*" + txtcari2.Text + "*")
End Select

If cekcari2.Value = 1 Then
    If dtp3.Value = dtp4.Value Then
        sql = sql + " and format(tgljammasuk,'mm/dd/yyyy') = '" + Format(dtp3.Value, "mm/dd/yyyy") + "'"
    Else
        sql = sql + " and format(tgljammasuk,'mm/dd/yyyy HH:mm:ss') >= '" + Format(dtp3.Value, "mm/dd/yyyy ") + _
        Format(dtp7.Value, "HH:mm:ss") + _
        "' and format(tgljamkeluar,'mm/dd/yyyy HH:mm:ss') <= '" + Format(dtp4.Value + 1, "mm/dd/yyyy ") + _
        Format(dtp8.Value, "HH:mm:ss") + "'"
    End If
End If

If cmbpintu2.ListIndex > 0 Then
    sql = sql + " and pintuMasuk=" & kutip(cmbpintu2)
End If

If cmbPintu3.ListIndex > 0 Then
    sql = sql + " and pintukeluar=" & kutip(cmbPintu3)
End If

Select Case cmbsort2.ListIndex
Case 0
    sql = sql + " order by tgljammasuk"
Case 1
    sql = sql + " order by pintumasuk"
Case 2
    sql = sql + " order by tgljamkeluar"
Case 3
    sql = sql + " order by pintukeluar"
Case 4
    sql = sql + " order by nama"
Case 5
    sql = sql + " order by nopol"
Case 6
    sql = sql + " order by rfid"
Case 7
    sql = sql + " order by sts"

End Select

Data5.RecordSource = sql
Data5.Refresh
Dim jmldata As Integer
jmldata = Data5.Recordset.RecordCount
lblJmlData.Caption = "Jumlah Data = " & jmldata
End Sub

Private Sub Command4_Click()
With CtkTransMaster
    .Show
    .WindowState = 2
    .Data1.DatabaseName = db
    .Data1.Connect = dbCon
    .Data1.RecordSource = Data3.RecordSource
End With
End Sub

Private Sub comMasuk_OnComm()
kosongMasuk
If comMasuk.PortOpen Then
    lblRfMasuk.Caption = Mid(comMasuk.Input, 2, 10)
    comMasuk.PortOpen = False
    If Len(lblRfMasuk.Caption) = 10 Then
        'MsgBox "Baca rfid sukses"
        ProsesMasuk
        'bukaPintu
        bukaPortMasuk
    End If
Else
    'MsgBox "Port Close"
    bukaPortMasuk
End If
End Sub

Private Sub Data3_Reposition()
IsiData
End Sub

Private Sub Data5_Reposition()
If Data5.Recordset.RecordCount > 0 Then
    If IsNull(Data5.Recordset!pintukeluar) Then
        img5.Picture = LoadPicture(App.Path + "\logo eq.jpg")
        img6.Picture = LoadPicture(App.Path + "\logo eq.jpg")
    Else
        If Dir$(App.Path + "\gbr\" + Data5.Recordset!notrans & "-i.jpg") <> "" Then
            img5.Picture = LoadPicture(App.Path + "\gbr\" + Data5.Recordset!notrans & "-i.jpg")
        Else
            img5.Picture = LoadPicture(App.Path + "\logo eq.jpg")
        End If
        If Dir$(App.Path + "\gbr\" + Data5.Recordset!notrans & "-o.jpg") <> "" Then
            img6.Picture = LoadPicture(App.Path + "\gbr\" + Data5.Recordset!notrans & "-o.jpg")
        Else
            img6.Picture = LoadPicture(App.Path + "\logo eq.jpg")
        End If
        
    End If
End If
End Sub

Private Sub Form_Load()
bukadaTa
Data1.DatabaseName = db
Data1.Connect = dbCon
Data1.RecordSource = "select * from view_masuk"
Data1.Refresh

Data2.DatabaseName = db
Data2.Connect = dbCon
Data2.RecordSource = "select * from view_keluar"
Data2.Refresh

datacekMasuk.DatabaseName = db
datacekMasuk.Connect = dbCon

DataTrans.DatabaseName = db
DataTrans.Connect = dbCon
DataTrans.RecordSource = "select * from trans"
DataTrans.Refresh

Data4.DatabaseName = db
Data4.Connect = dbCon
Data4.RecordSource = "select * from trmember"
Data4.Refresh

Tampil

Data3.DatabaseName = db
Data3.Connect = dbCon
Data3.RecordSource = "select * from view_trans1  order by tgljam desc"
Data3.Refresh
Dim jmldata As Integer
jmldata = Data3.Recordset.RecordCount
lblJml.Caption = "Jumlah Data = " & jmldata

Data5.DatabaseName = db
Data5.Connect = dbCon
Data5.RecordSource = "select * from view_trmember order by tgljamkeluar desc"
Data5.Refresh
jmldata = Data5.Recordset.RecordCount
lblJmlData.Caption = "Jumlah Data = " & jmldata

kosongMasuk
kosongKeluar

Data6.DatabaseName = db
Data6.Connect = dbCon
Data6.RecordSource = "select * from mspintu"
Data6.Refresh

isiCombo

dtp1.Value = Now
dtp2.Value = Now
dtp3.Value = Now
dtp4.Value = Now
End Sub

Sub isiCombo()
'isi combo status
cmbSts.Clear
cmbSts.AddItem ("Semua")
cmbSts.AddItem ("IN")
cmbSts.AddItem ("OUT")
cmbSts.ListIndex = 0

cmbsts2.Clear
cmbsts2.AddItem ("Semua")
cmbsts2.AddItem ("Selesai")
cmbsts2.AddItem ("Belum Selesai")
cmbsts2.ListIndex = 0


'isi combo cari master
cmbCari.Clear
cmbCari.AddItem ("Nama")
cmbCari.AddItem ("Nopol")
cmbCari.AddItem ("RFID")
cmbCari.ListIndex = 0

cmbcari2.Clear
cmbcari2.AddItem ("Nama")
cmbcari2.AddItem ("Nopol")
cmbcari2.AddItem ("RFID")
cmbcari2.ListIndex = 0

'isi combo pintu
cmbpintu.Clear
With Data6.Recordset
    If Not .BOF Then
        .MoveFirst
            cmbpintu.AddItem "Semua"
        Do While Not .EOF
            cmbpintu.AddItem !namapintu
            .MoveNext
        Loop
        cmbpintu.ListIndex = 0
    End If
End With

cmbpintu2.Clear
Data6.RecordSource = "select * from mspintu where jenispintu='Pintu Masuk'"
Data6.Refresh
With Data6.Recordset
    If Not .BOF Then
        .MoveFirst
            cmbpintu2.AddItem "Semua"
        Do While Not .EOF
            cmbpintu2.AddItem !namapintu
            .MoveNext
        Loop
        cmbpintu2.ListIndex = 0
    End If
End With

cmbPintu3.Clear
Data6.RecordSource = "select * from mspintu where jenispintu='Pintu Keluar'"
Data6.Refresh
With Data6.Recordset
    If Not .BOF Then
        .MoveFirst
            cmbPintu3.AddItem "Semua"
        Do While Not .EOF
            cmbPintu3.AddItem !namapintu
            .MoveNext
        Loop
        cmbPintu3.ListIndex = 0
    End If
End With

cmbSort.Clear
cmbSort.AddItem "Tanggal Jam"
cmbSort.AddItem "Status In/Out"
cmbSort.AddItem "Nama Pintu"
cmbSort.AddItem "Nama"
cmbSort.AddItem "NOPOL"
cmbSort.AddItem "RFID"
cmbSort.ListIndex = 0

cmbsort2.Clear
cmbsort2.AddItem "Tanggal Jam Masuk"
cmbsort2.AddItem "Pintu Masuk"
cmbsort2.AddItem "Tanggal Jam Keluar"
cmbsort2.AddItem "Pintu Keluar"
cmbsort2.AddItem "Nama"
cmbsort2.AddItem "Nopol"
cmbsort2.AddItem "RFID"
cmbsort2.AddItem "Status"
cmbsort2.ListIndex = 0

End Sub

Sub kosongMasuk()
lblRfMasuk.Caption = ""
lblNopolMasuk.Caption = ""
lblTglmasuk1.Caption = ""
lbljammasuk1.Caption = ""
lblnamaMasuk.Caption = ""
End Sub

Sub kosongKeluar()
lblrfkeluar.Caption = ""
lblnopolkeluar.Caption = ""
lbltglkeluar.Caption = ""
lbljamkeluar.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
lblNamaKeluar.Caption = ""
End Sub

Sub Tampil()
With Data1.Recordset
If Not .BOF Then
    lblpintumasuk.Caption = !namapintumasuk
    cam1.playlist.items.Clear
    cam1.playlist.Add ("file:///" + !lokasisetting)
    cam1.playlist.play
    
End If
End With

With Data2.Recordset
If Not .BOF Then
    lblpintukeluar.Caption = !namapintukeluar
    cam2.playlist.items.Clear
    cam2.playlist.Add ("file:///" + !lokasisetting)
    cam2.playlist.play
End If
End With

End Sub

Sub IsiData()
With Data3.Recordset
    If .RecordCount > 0 Then
        lblnotrans.Caption = !notrans
        lblgate.Caption = !namapintu
        lblinout.Caption = !sts
        lblrfidtrans.Caption = !rfid
        lblNopoltr.Caption = !nopol
        lbltgltr.Caption = Format(!tgljam, "dd MMM yyyy")
        lbljamtr.Caption = Format(!tgljam, "HH:mm:ss")
        Dim status As String
        status = Trim(!sts)
        If status = "IN" Then
            If Dir$(App.Path + "\gbr\" + !notrans + "-i.jpg") <> "" Then
                img3.Picture = LoadPicture(App.Path + "\gbr\" + !notrans + "-i.jpg")
            Else
                img3.Picture = LoadPicture(App.Path + "\logo eq.jpg")
            End If
        Else
            If Dir$(App.Path + "\gbr\" + !notrans + "-o.jpg") <> "" Then
                img3.Picture = LoadPicture(App.Path + "\gbr\" + !notrans + "-o.jpg")
            Else
                img3.Picture = LoadPicture(App.Path + "\logo eq.jpg")
            End If
        End If
    End If
End With
End Sub

