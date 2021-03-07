VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Suara_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Suara"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12930
   Icon            =   "Suara_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   12930
   Begin VB.Data datacek 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Index           =   1
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Suara"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lokasi File"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   750
      End
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Suara_frm.frx":0CCA
      Height          =   2295
      Left            =   5280
      TabIndex        =   1
      Top             =   840
      Width           =   7455
      _Version        =   196614
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   4075
      Columns(0).Caption=   "Nama Suara"
      Columns(0).Name =   "Nama Suara"
      Columns(0).DataField=   "namasuara"
      Columns(0).FieldLen=   256
      Columns(1).Width=   7911
      Columns(1).Caption=   "Lokasi File Suara"
      Columns(1).Name =   "Lokasi File Suara"
      Columns(1).DataField=   "lokasifile"
      Columns(1).FieldLen=   256
      _ExtentX        =   13150
      _ExtentY        =   4048
      _StockProps     =   79
      Caption         =   "Data Suara"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   2520
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
            Picture         =   "Suara_frm.frx":0CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Suara_frm.frx":19B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Suara_frm.frx":2692
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Suara_frm.frx":336C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Suara_frm.frx":4CFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Suara_frm.frx":59D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Suara_frm.frx":66B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4680
      MouseIcon       =   "Suara_frm.frx":738C
      MousePointer    =   99  'Custom
      Picture         =   "Suara_frm.frx":7696
      ToolTipText     =   "Test Sound"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   2880
      MouseIcon       =   "Suara_frm.frx":8360
      MousePointer    =   99  'Custom
      Picture         =   "Suara_frm.frx":866A
      ToolTipText     =   "Tambah"
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   3720
      MouseIcon       =   "Suara_frm.frx":9334
      MousePointer    =   99  'Custom
      Picture         =   "Suara_frm.frx":963E
      ToolTipText     =   "Edit"
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   4560
      MouseIcon       =   "Suara_frm.frx":AFC0
      MousePointer    =   99  'Custom
      Picture         =   "Suara_frm.frx":B2CA
      ToolTipText     =   "Hapus"
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   12240
      MouseIcon       =   "Suara_frm.frx":BF94
      MousePointer    =   99  'Custom
      Picture         =   "Suara_frm.frx":C29E
      ToolTipText     =   "Keluar"
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA SUARA"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1200
   End
End
Attribute VB_Name = "Suara_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tambah As Boolean
Dim NamaSuara As String

Private Sub Command1_Click()
cdg1.InitDir = App.Path
cdg1.Filter = ""
AddFilter cdg1, "File Suara", "*.wav"
cdg1.ShowOpen
Text1(1).Text = cdg1.Filename
End Sub

Private Sub Data1_Reposition()
Isi
End Sub

Private Sub Form_Load()
Call bukadaTa
Data1.DatabaseName = db
Data1.Connect = dbCon
datacek.DatabaseName = db
datacek.Connect = dbCon
isiGrid
BtnAwal
End Sub

Sub isiGrid()
Data1.RecordSource = "select * from mssuara"
Data1.Refresh
End Sub

Sub BtnAwal()
Image5.Visible = True
Image6.ToolTipText = "Edit"
Image6.Picture = ImageList1.ListImages(4).Picture
Image7.ToolTipText = "Hapus"
Frame1.Enabled = False
SSDBGrid1.Enabled = True
End Sub

Sub btnSimpan()
Image5.Visible = False
Image6.ToolTipText = "Simpan"
Image6.Picture = ImageList1.ListImages(7).Picture
Image7.ToolTipText = "Batal"
Frame1.Enabled = True
SSDBGrid1.Enabled = False
End Sub

Private Sub Image1_Click()
If Text1(1).Text <> "" Then
playsound = sndPlaySound(Text1(1).Text, 1)
End If
End Sub

Private Sub Image5_Click()
    tambah = True
    btnSimpan
    Kosong
    Text1(0).SetFocus
End Sub

Private Sub Image6_Click()
If Image6.ToolTipText = "Edit" Then
    tambah = False
    btnSimpan
    Text1(0).SetFocus
    NamaSuara = Text1(0).Text
Else
    If Text1(0).Text = "" Or Text1(1).Text = "" Then
        MsgBox "Data belum lengkap"
    Else
        If tambah = True Then
            'cek nama suara sudah ada atau belum
            datacek.RecordSource = "select * from mssuara where namasuara=" & kutip(Text1(0).Text)
            datacek.Refresh
            If datacek.Recordset.BOF Then
                With Data1.Recordset
                    .AddNew
                    !NamaSuara = Text1(0).Text
                    !lokasifile = Text1(1).Text
                    .Update
                End With
                MsgBox "Berhasil tambah data"
                Data1.Refresh
                BtnAwal
                Isi
            Else
                MsgBox "Nama suara sudah ada"
            End If

        Else 'edit
            'cek nama suara sudah ada atau belum
            datacek.RecordSource = "select * from mssuara where namasuara=" & kutip(Text1(0).Text) & _
            " and namasuara<>" & kutip(NamaSuara)
            datacek.Refresh
            If datacek.Recordset.BOF Then
                With Data1.Recordset
                    .Edit
                    !NamaSuara = Text1(0).Text
                    !lokasifile = Text1(1).Text
                    .Update
                End With
                MsgBox "Berhasil update data"
                Data1.Refresh
                BtnAwal
                Isi
            Else
                MsgBox "User name sudah ada"
            End If
        End If
    End If
End If
End Sub

Private Sub Image7_Click()
    If Image7.ToolTipText = "Batal" Then
        BtnAwal
        Isi
    Else
    Dim tny As String
    tny = MsgBox("Apakah anda yakin?", vbYesNo, "Hapus")
    If tny = vbYes Then
        With Data1.Recordset
            If .BOF Then
                MsgBox "Tidak dapat hapus data (data kosong)"
            Else
                .Delete
                .MoveFirst
                MsgBox "Berhasil hapus data"
                Kosong
                Isi
            End If
        End With
    End If
    End If
End Sub

Private Sub Image8_Click()
Unload Me
End Sub

Sub Kosong()
Text1(0).Text = ""
Text1(1).Text = ""
End Sub

Sub Isi()
With Data1.Recordset
    If Not .BOF And Image5.Visible = True Then
        Text1(0).Text = !NamaSuara
        Text1(1).Text = !lokasifile
    End If
End With
End Sub
