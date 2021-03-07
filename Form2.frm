VERSION 5.00
Object = "{DF2BBE39-40A8-433B-A279-073F48DA94B6}#1.0#0"; "axvlc.dll"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18420
   LinkTopic       =   "Form2"
   ScaleHeight     =   8895
   ScaleWidth      =   18420
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   15901
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "PINTU"
      TabPicture(0)   =   "Form2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "TRANSAKSI KARTU MASTER"
      TabPicture(1)   =   "Form2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "TRANSAKSI KARTU MEMBER"
      TabPicture(2)   =   "Form2.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pintu Masuk"
         Height          =   8295
         Index           =   0
         Left            =   -74880
         TabIndex        =   43
         Top             =   480
         Width           =   6255
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
         Begin VB.CommandButton Command1 
            Caption         =   "On"
            Height          =   495
            Left            =   4920
            TabIndex        =   44
            Top             =   7080
            Width           =   975
         End
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
         Begin MSCommLib.MSComm comMasuk 
            Left            =   2280
            Top             =   720
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
         Begin AXVLCCtl.VLCPlugin2 cam1 
            Height          =   3615
            Left            =   240
            TabIndex        =   45
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capture IN"
            Height          =   195
            Left            =   240
            TabIndex        =   54
            Top             =   4560
            Width           =   765
         End
         Begin VB.Image img1 
            BorderStyle     =   1  'Fixed Single
            Height          =   2880
            Left            =   240
            Picture         =   "Form2.frx":0054
            Stretch         =   -1  'True
            Top             =   4800
            Width           =   3480
         End
         Begin VB.Label lblfilebmp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no trans"
            Height          =   195
            Left            =   1560
            TabIndex        =   53
            Top             =   7200
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lbltransmasuk 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no trans"
            Height          =   195
            Left            =   1080
            TabIndex        =   52
            Top             =   6240
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lblpintumasuk 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Pintu"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   825
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
            TabIndex        =   50
            Top             =   6360
            Width           =   510
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
            TabIndex        =   49
            Top             =   6000
            Width           =   975
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
            TabIndex        =   48
            Top             =   5640
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
            TabIndex        =   47
            Top             =   4920
            Width           =   660
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
            TabIndex        =   46
            Top             =   5280
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pintu Keluar"
         Height          =   8295
         Index           =   1
         Left            =   -68520
         TabIndex        =   26
         Top             =   480
         Width           =   8175
         Begin VB.CommandButton CmdKeluar 
            Caption         =   "On"
            Height          =   495
            Left            =   7080
            TabIndex        =   27
            Top             =   4080
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
            Left            =   3000
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   240
            Visible         =   0   'False
            Width           =   1140
         End
         Begin MSCommLib.MSComm comKeluar 
            Left            =   720
            Top             =   840
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
         Begin AXVLCCtl.VLCPlugin2 cam2 
            Height          =   3615
            Left            =   240
            TabIndex        =   28
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keluar"
            Height          =   195
            Left            =   4440
            TabIndex        =   42
            Top             =   4440
            Width           =   450
         End
         Begin VB.Label lbltranskeluar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no trans"
            Height          =   195
            Left            =   1080
            TabIndex        =   41
            Top             =   5760
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lblfilebmp2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no trans"
            Height          =   195
            Left            =   2520
            TabIndex        =   40
            Top             =   6600
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Image img2 
            BorderStyle     =   1  'Fixed Single
            Height          =   3000
            Left            =   4440
            Picture         =   "Form2.frx":3E6F0
            Stretch         =   -1  'True
            Top             =   4680
            Width           =   3480
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
            TabIndex        =   39
            Top             =   3600
            Width           =   1320
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
            TabIndex        =   38
            Top             =   3240
            Width           =   1785
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
            TabIndex        =   37
            Top             =   600
            Width           =   660
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
            TabIndex        =   36
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label lblpintukeluar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Pintu"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   825
         End
         Begin VB.Image img4 
            BorderStyle     =   1  'Fixed Single
            Height          =   3000
            Left            =   240
            Picture         =   "Form2.frx":7CD8C
            Stretch         =   -1  'True
            Top             =   4680
            Width           =   3480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Masuk"
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   4440
            Width           =   480
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
            TabIndex        =   32
            Top             =   2520
            Width           =   510
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
            TabIndex        =   31
            Top             =   1800
            Width           =   795
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
            TabIndex        =   30
            Top             =   2880
            Width           =   780
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
            TabIndex        =   29
            Top             =   960
            Width           =   705
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   8295
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   14415
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
         Begin VB.ComboBox cmbCari 
            Height          =   315
            Left            =   10560
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   11760
            TabIndex        =   9
            Top             =   600
            Width           =   2415
         End
         Begin VB.CommandButton Command2 
            Caption         =   "CARI"
            Height          =   315
            Left            =   10560
            TabIndex        =   8
            Top             =   2400
            Width           =   3615
         End
         Begin VB.ComboBox cmbSts 
            Height          =   315
            Left            =   11400
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   2775
         End
         Begin VB.CheckBox cekcari 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tanggal"
            Height          =   255
            Left            =   10560
            TabIndex        =   6
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cmbpintu 
            Height          =   315
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1680
            Width           =   2655
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
         Begin VB.CommandButton Command4 
            Caption         =   "CETAK"
            Height          =   315
            Left            =   10560
            TabIndex        =   3
            Top             =   3000
            Width           =   3615
         End
         Begin VB.ComboBox cmbSort 
            Height          =   315
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   2040
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker dtp1 
            Height          =   315
            Left            =   11520
            TabIndex        =   5
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   47579137
            CurrentDate     =   42473
         End
         Begin MSComCtl2.DTPicker dtp2 
            Height          =   315
            Left            =   11520
            TabIndex        =   11
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   47579137
            CurrentDate     =   42473
         End
         Begin MSComCtl2.DTPicker dtp5 
            Height          =   315
            Left            =   12840
            TabIndex        =   12
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   47579138
            CurrentDate     =   42473
         End
         Begin MSComCtl2.DTPicker dtp6 
            Height          =   315
            Left            =   12840
            TabIndex        =   13
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   47579138
            CurrentDate     =   42473
         End
         Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
            Bindings        =   "Form2.frx":BB428
            Height          =   7695
            Left            =   240
            TabIndex        =   14
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
            BackColor       =   16777215
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
            TabIndex        =   25
            Top             =   3840
            Visible         =   0   'False
            Width           =   1095
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
            TabIndex        =   24
            Top             =   3720
            Width           =   900
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
            TabIndex        =   23
            Top             =   3360
            Width           =   1380
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
            TabIndex        =   22
            Top             =   4440
            Width           =   705
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
            TabIndex        =   21
            Top             =   4080
            Width           =   660
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
            TabIndex        =   20
            Top             =   4800
            Width           =   975
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
            TabIndex        =   19
            Top             =   4800
            Width           =   585
         End
         Begin VB.Image img3 
            BorderStyle     =   1  'Fixed Single
            Height          =   2880
            Left            =   10560
            Picture         =   "Form2.frx":BB43C
            Stretch         =   -1  'True
            Top             =   5160
            Width           =   3600
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   255
            Left            =   10560
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Pintu"
            Height          =   255
            Left            =   10560
            TabIndex        =   17
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label lblJml 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Data"
            Height          =   255
            Left            =   10560
            TabIndex        =   16
            Top             =   2760
            Width           =   3615
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort By"
            Height          =   255
            Left            =   10560
            TabIndex        =   15
            Top             =   2160
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
