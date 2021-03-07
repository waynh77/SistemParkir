VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Member_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Member"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10020
   Icon            =   "Member_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10020
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   1  'ODBCCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Text            =   "1"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan RFID"
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox cmbsort 
      Height          =   315
      Left            =   1080
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Data datacek 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cari"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   6615
      Begin VB.TextBox txtcari 
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   3375
      End
      Begin VB.ComboBox cmbcari 
         Height          =   315
         ItemData        =   "Member_frm.frx":0CCA
         Left            =   120
         List            =   "Member_frm.frx":0CD7
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.Image imgCari 
         Height          =   480
         Left            =   6000
         MouseIcon       =   "Member_frm.frx":0CF9
         MousePointer    =   99  'Custom
         Picture         =   "Member_frm.frx":1003
         ToolTipText     =   "Cari Data"
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TRANSAKSI MEMBER"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txtalamat 
         Height          =   495
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtnama 
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   360
         Width           =   3135
      End
      Begin VB.CheckBox cekMaster 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Master"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   1920
         Width           =   1455
      End
      Begin MSCommLib.MSComm Com1 
         Left            =   4800
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CheckBox cekBlokir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Blokir"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtCat 
         Height          =   855
         Left            =   5640
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtNopol 
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtRfid 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1440
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dtpDaftar 
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   136183809
         CurrentDate     =   42460
      End
      Begin MSComCtl2.DTPicker dtpExp 
         Height          =   375
         Left            =   7800
         TabIndex        =   9
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   136183809
         CurrentDate     =   42460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         Height          =   195
         Index           =   4
         Left            =   4680
         TabIndex        =   10
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Exp"
         Height          =   195
         Index           =   3
         Left            =   7200
         TabIndex        =   5
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Daftar"
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   4
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Pol"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   3
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RFID No"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   630
      End
   End
   Begin SSDataWidgets_B.SSDBGrid Grid1 
      Bindings        =   "Member_frm.frx":1CCD
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   9735
      _Version        =   196614
      AllowUpdate     =   0   'False
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   17171
      _ExtentY        =   7435
      _StockProps     =   79
      Caption         =   "DATA MEMBER"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Member_frm.frx":1CE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Member_frm.frx":29BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Member_frm.frx":3695
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Member_frm.frx":436F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Member_frm.frx":5D01
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Member_frm.frx":69DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Member_frm.frx":76B5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port No"
      Height          =   195
      Index           =   8
      Left            =   3240
      TabIndex        =   23
      Top             =   3480
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "Member_frm.frx":838F
      MousePointer    =   99  'Custom
      Picture         =   "Member_frm.frx":8699
      ToolTipText     =   "Cetak"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   21
      Top             =   3480
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   9360
      MouseIcon       =   "Member_frm.frx":9363
      MousePointer    =   99  'Custom
      Picture         =   "Member_frm.frx":966D
      ToolTipText     =   "Keluar"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image imgHapus 
      Height          =   480
      Left            =   8040
      MouseIcon       =   "Member_frm.frx":A337
      MousePointer    =   99  'Custom
      Picture         =   "Member_frm.frx":A641
      ToolTipText     =   "Hapus"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image imgEdit 
      Height          =   480
      Left            =   7440
      MouseIcon       =   "Member_frm.frx":B30B
      MousePointer    =   99  'Custom
      Picture         =   "Member_frm.frx":B615
      ToolTipText     =   "Edit"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image imgTambah 
      Height          =   480
      Left            =   6840
      MouseIcon       =   "Member_frm.frx":CF97
      MousePointer    =   99  'Custom
      Picture         =   "Member_frm.frx":D2A1
      ToolTipText     =   "Tambah"
      Top             =   2760
      Width           =   480
   End
End
Attribute VB_Name = "Member_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim NoRfid As String

Private Sub cmdScan_Click()
If cmdScan.Caption = "Scan RFID" Then
    cmdScan.Caption = "Stop Scan"
    BukaPort
    txtRfid.Text = ""
Else
    cmdScan.Caption = "Scan RFID"
    txtRfid.Text = ""
    If Com1.PortOpen Then
        Com1.PortOpen = False
    End If
End If
End Sub

Sub BukaPort()
    Com1.CommPort = Text1.Text
    Com1.PortOpen = True
    Com1.RTSEnable = True
    Com1.RThreshold = 14
    Com1.InputLen = 0
End Sub

Private Sub Com1_OnComm()
txtRfid.Text = ""
If Com1.PortOpen Then
    txtRfid.Text = Mid(Com1.Input, 2, 10)
    Com1.PortOpen = False
Else
    MsgBox "Port Close"
End If
End Sub

Private Sub Data1_Reposition()
Isi
End Sub

Private Sub Form_Load()
bukadaTa
Data1.DatabaseName = db
Data1.Connect = dbCon
Data1.RecordSource = "msmember"
Data1.Refresh
datacek.DatabaseName = db
datacek.Connect = dbCon
cmbCari.ListIndex = 0
BtnAwal
isiCombo
End Sub

Sub isiCombo()
With cmbSort
    .Clear
    .AddItem ("RFID")
    .AddItem ("Nama")
    .AddItem ("Alamat")
    .AddItem ("Nopol")
    .AddItem ("Tgl Daftar")
    .AddItem ("Tgl Expired")
    .ListIndex = 0
End With
End Sub

Sub BtnAwal()
Frame1.Enabled = False
Frame2.Visible = True
imgTambah.Visible = True
imgEdit.ToolTipText = "Edit"
imgEdit.Picture = ImageList1.ListImages(4).Picture
imgHapus.ToolTipText = "Hapus"
Grid1.Enabled = True
End Sub

Sub btnSimpan()
Frame1.Enabled = True
Frame2.Visible = False
imgTambah.Visible = False
imgEdit.ToolTipText = "Simpan"
imgEdit.Picture = ImageList1.ListImages(7).Picture
imgHapus.ToolTipText = "Batal"
Grid1.Enabled = False
End Sub

Sub Kosong()
txtRfid.Text = ""
txtNopol.Text = "-"
txtCat.Text = "-"
txtalamat.Text = "-"
dtpDaftar.Value = Now
dtpExp.Value = Now + 1
txtnama.Text = ""
cekBlokir.Value = 0
cekMaster.Value = 0
End Sub

Sub Isi()
With Data1.Recordset
    If Not .BOF And imgTambah.Visible = True Then
        txtRfid.Text = !rfid
        txtNopol.Text = !nopol
        txtCat.Text = !catatan
        dtpDaftar.Value = !tgldaftar
        dtpExp.Value = !tglexp
        cekBlokir.Value = !blokir * -1
        txtnama.Text = !nama
        cekMaster.Value = !master * -1
        txtalamat.Text = !alamat
    End If
End With
End Sub

Private Sub Image1_Click()
With DataMember_rpt
    .Data1.DatabaseName = db
    .Data1.Connect = dbCon
    .Data1.RecordSource = Data1.RecordSource
    .Show
    '.Top = 0
    '.Left = 0
End With
End Sub

Private Sub Image2_Click(Index As Integer)
Unload Me
End Sub

Private Sub imgCari_Click()
Dim sql As String
sql = "select * from msmember "
Select Case cmbCari.ListIndex
Case 0
    sql = sql + "where nopol like "
Case 1
    sql = sql + "where nama like "
Case 2
    sql = sql + "where alamat like "
End Select
sql = sql & "'*" & txtCari.Text & "*'"

Select Case cmbSort.ListIndex
Case 0
    sql = sql + " order by rfid"
Case 1
    sql = sql + " order by nama"
Case 2
    sql = sql + " order by alamat"
Case 3
    sql = sql + " order by nopol"
Case 4
    sql = sql + " order by tgldaftar"
Case 5
    sql = sql + " order by tglexp"
End Select

Data1.RecordSource = sql
Data1.Refresh
End Sub

Private Sub imgEdit_Click()
If imgEdit.ToolTipText = "Edit" Then
    btnSimpan
    tambah = False
    txtRfid.SetFocus
    NoRfid = txtRfid.Text
Else
    If txtRfid.Text = "" Then
        MsgBox "RFID tidak boleh kosong"
        txtRfid.SetFocus
    Else 'simpan
        If tambah = True Then
            'cek no rfid sudah ada atau belum
            datacek.RecordSource = "Select * from msmember where rfid=" & kutip(txtRfid.Text)
            datacek.Refresh
            If Not datacek.Recordset.BOF Then
                MsgBox "Nomor RFID sudah ada"
            Else
                With Data1.Recordset
                    .AddNew
                    !rfid = txtRfid.Text
                    !nopol = txtNopol.Text
                    !catatan = txtCat.Text
                    !tgldaftar = dtpDaftar.Value
                    !tglexp = dtpExp.Value
                    !blokir = cekBlokir.Value * -1
                    !nama = txtnama.Text
                    !master = cekMaster.Value * -1
                    !alamat = txtalamat.Text
                    .Update
                End With
                MsgBox "Berhasil tambah data"
                BtnAwal
                Data1.Refresh
            End If
        Else 'update/edit
            'cek no rfid
            datacek.RecordSource = "Select * from msmember where rfid=" & kutip(txtRfid.Text) & _
            " and rfid<>" & kutip(NoRfid)
            datacek.Refresh
            If Not datacek.Recordset.BOF Then
                MsgBox "Nomor RFID sudah ada"
            Else
                With Data1.Recordset
                    .Edit
                    '!rfid = txtRfid.Text
                    !nopol = txtNopol.Text
                    !catatan = txtCat.Text
                    !tgldaftar = dtpDaftar.Value
                    !tglexp = dtpExp.Value
                    !blokir = cekBlokir.Value * -1
                    !nama = txtnama.Text
                    !master = cekMaster.Value * -1
                    !alamat = txtalamat.Text
                    .Update
                End With
                MsgBox "Berhasil update data"
                BtnAwal
                Data1.Refresh
            End If
        End If
    End If
End If
End Sub

Private Sub imgHapus_Click()
If imgHapus.ToolTipText = "Hapus" Then
    Dim tny As String
    tny = MsgBox("Apakah anda yakin?", vbYesNo, "Hapus")
    If tny = vbYes Then
        With Data1.Recordset
            If .BOF Or .RecordCount = 1 Then
                MsgBox "Tidak dapat hapus data (data kosong)"
            Else
                .Delete
                .MoveFirst
                MsgBox "Berhasil hapus data"
                Isi
            End If
        End With
    End If
Else 'batal
    BtnAwal
    Data1.Refresh
End If
End Sub

Private Sub imgTambah_Click()
Kosong
btnSimpan
txtRfid.SetFocus
tambah = True
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then imgCari_Click
End Sub
