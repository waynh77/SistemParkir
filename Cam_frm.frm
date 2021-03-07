VERSION 5.00
Object = "{DF2BBE39-40A8-433B-A279-073F48DA94B6}#1.0#0"; "axvlc.dll"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Cam_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Kamera"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14505
   Icon            =   "Cam_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   14505
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Width           =   1260
   End
   Begin VB.TextBox txtnama 
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   13920
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   3480
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPlay 
      Height          =   495
      Left            =   9000
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Cam_frm.frx":0CCA
      Top             =   720
      Width           =   4815
   End
   Begin AXVLCCtl.VLCPlugin2 cam1 
      Height          =   4455
      Left            =   7560
      TabIndex        =   1
      Top             =   2160
      Width           =   6735
      AutoLoop        =   0   'False
      AutoPlay        =   -1  'True
      Toolbar         =   -1  'True
      ExtentWidth     =   11880
      ExtentHeight    =   7858
      MRL             =   ""
      Object.Visible         =   -1  'True
      Volume          =   50
      StartTime       =   0
      BaseURL         =   ""
      BackColor       =   16777215
      FullscreenEnabled=   -1  'True
      Branding        =   -1  'True
   End
   Begin SSDataWidgets_B.SSDBGrid dgv1 
      Bindings        =   "Cam_frm.frx":0D30
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   7215
      _Version        =   196614
      AllowUpdate     =   0   'False
      SelectTypeRow   =   1
      MaxSelectedRows =   1
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   5001
      Columns(0).Caption=   "Nama Kamera"
      Columns(0).Name =   "Nama Kamera"
      Columns(0).DataField=   "Namacam"
      Columns(0).FieldLen=   256
      Columns(1).Width=   7064
      Columns(1).Caption=   "Lokasi File Settingan"
      Columns(1).Name =   "Lokasi File Settingan"
      Columns(1).DataField=   "lokasisetting"
      Columns(1).FieldLen=   256
      _ExtentX        =   12726
      _ExtentY        =   10186
      _StockProps     =   79
      Caption         =   "Data Kamera"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA KAMERA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Kamera"
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   6840
      MouseIcon       =   "Cam_frm.frx":0D44
      MousePointer    =   99  'Custom
      Picture         =   "Cam_frm.frx":104E
      ToolTipText     =   "Keluar"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   6000
      MouseIcon       =   "Cam_frm.frx":1D18
      MousePointer    =   99  'Custom
      Picture         =   "Cam_frm.frx":2022
      ToolTipText     =   "Hapus"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   5160
      MouseIcon       =   "Cam_frm.frx":2CEC
      MousePointer    =   99  'Custom
      Picture         =   "Cam_frm.frx":2FF6
      ToolTipText     =   "Edit"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4320
      MouseIcon       =   "Cam_frm.frx":4978
      MousePointer    =   99  'Custom
      Picture         =   "Cam_frm.frx":4C82
      ToolTipText     =   "Tambah"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   13680
      MouseIcon       =   "Cam_frm.frx":594C
      MousePointer    =   99  'Custom
      Picture         =   "Cam_frm.frx":5C56
      ToolTipText     =   "Batal"
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   12960
      MouseIcon       =   "Cam_frm.frx":6920
      MousePointer    =   99  'Custom
      Picture         =   "Cam_frm.frx":6C2A
      ToolTipText     =   "Simpan"
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7680
      MouseIcon       =   "Cam_frm.frx":78F4
      MousePointer    =   99  'Custom
      Picture         =   "Cam_frm.frx":7BFE
      ToolTipText     =   "Stop"
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lokasi File Setting"
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7680
      MouseIcon       =   "Cam_frm.frx":88C8
      MousePointer    =   99  'Custom
      Picture         =   "Cam_frm.frx":8BD2
      ToolTipText     =   "Test Kamera"
      Top             =   1440
      Width           =   480
   End
End
Attribute VB_Name = "Cam_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim namacamOld As String
Private Sub Command1_Click()
cdg1.InitDir = App.Path
cdg1.Filter = "*.xspf"
cdg1.ShowOpen
txtPlay.Text = cdg1.Filename
End Sub

Sub TombolAwal()
txtnama.Enabled = False
txtPlay.Enabled = False
Command1.Enabled = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = True
Image6.Visible = True
Image7.Visible = True
dgv1.Enabled = True
End Sub

Sub TombolSimpan()
txtnama.Enabled = True
txtPlay.Enabled = True
Command1.Enabled = True
Image3.Visible = True
Image4.Visible = True
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
dgv1.Enabled = False
End Sub

Private Sub Data1_Reposition()
Isi
End Sub

Private Sub Form_Load()
bukadaTa
Data1.DatabaseName = db
Data1.Connect = dbCon
Data1.RecordSource = "select * from mscam"
Data1.Refresh
Image2.Visible = False
TombolAwal
Data2.DatabaseName = db
Data2.Connect = dbCon
End Sub

Private Sub Image1_Click()
cam1.playlist.items.Clear
cam1.playlist.Add ("file:///" + txtPlay.Text)
cam1.playlist.play
Image1.Visible = False
Image2.Visible = True
End Sub

Private Sub Image2_Click()
cam1.playlist.stop
Image1.Visible = True
Image2.Visible = False
End Sub

Private Sub Image3_Click()
If txtnama.Text = "" Or txtPlay.Text = "" Then
    MsgBox ("Data tidak valid")
Else
    If tambah = True Then
        Data2.RecordSource = "select * from mscam where namacam='" + txtnama.Text + "'"
        Data2.Refresh
        If Data2.Recordset.BOF Then
            With Data1.Recordset
                .AddNew
                !namacam = txtnama.Text
                !lokasisetting = txtPlay.Text
                .Update
            End With
            MsgBox ("Berhasil simpan data")
            Data1.Refresh
            TombolAwal
            Isi
        Else
            MsgBox ("Nama kamera sudah ada, Silahkan masukan yang lain")
        End If
    Else
        Data2.RecordSource = "select * from mscam where namacam='" + txtnama.Text + "' and namacam<>'" + namacamOld + "'"
        Data2.Refresh
        If Data2.Recordset.BOF Then
            With Data1.Recordset
                .Edit
                !namacam = txtnama.Text
                !lokasisetting = txtPlay.Text
                .Update
            End With
            MsgBox ("Berhasil update data")
            Data1.RecordSource = "select * from mscam"
            Data1.Refresh
            TombolAwal
            Isi
        Else
            MsgBox ("Nama kamera sudah ada, Silahkan masukan yang lain")
        End If
    End If
End If

End Sub

Private Sub Image4_Click()
TombolAwal
End Sub

Private Sub Image5_Click()
tambah = True
txtnama.Text = ""
txtPlay.Text = ""
TombolSimpan
txtnama.SetFocus
End Sub

Private Sub Image6_Click()
If Not Data1.Recordset.BOF Then
    tambah = False
    TombolSimpan
    txtnama.SetFocus
    namacamOld = txtnama.Text
Else
    MsgBox ("Data Kosong")
End If
End Sub

Private Sub Image7_Click()
Dim tny As String
If Data1.Recordset.BOF Then
MsgBox ("Data kosong")
Else
tny = MsgBox("Apakah anda yakin?", vbYesNo, "Hapus Data")
If tny = vbYes Then
    Data1.Recordset.Delete
    Data1.RecordSource = "Select * from mscam"
    Data1.Refresh
End If
End If
End Sub

Private Sub Image8_Click()
Unload Me
End Sub

Sub Isi()
With Data1.Recordset
    If Not .BOF And Image5.Visible = True Then
        txtnama.Text = !namacam
        txtPlay.Text = !lokasisetting
    End If
End With
End Sub
