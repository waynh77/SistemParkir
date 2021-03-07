VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Pintu_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Pintu"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   Icon            =   "Pintu_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   4665
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Index           =   0
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      Begin VB.ComboBox cmbjenis 
         Height          =   315
         ItemData        =   "Pintu_frm.frx":0CCA
         Left            =   1440
         List            =   "Pintu_frm.frx":0CD4
         TabIndex        =   6
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtnama 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pintu"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pintu"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Data datacek 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   960
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
            Picture         =   "Pintu_frm.frx":0CF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pintu_frm.frx":19CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pintu_frm.frx":26A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pintu_frm.frx":3381
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pintu_frm.frx":4D13
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pintu_frm.frx":59ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pintu_frm.frx":66C7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Pintu_frm.frx":73A1
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   4335
      _Version        =   196614
      AllowUpdate     =   0   'False
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "Nama Pintu"
      Columns(0).Name =   "User Name"
      Columns(0).DataField=   "namapintu"
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Jenis Pintu"
      Columns(1).Name =   "Password"
      Columns(1).DataField=   "jenispintu"
      Columns(1).FieldLen=   256
      _ExtentX        =   7646
      _ExtentY        =   6800
      _StockProps     =   79
      Caption         =   "DATA PINTU"
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   2760
      MouseIcon       =   "Pintu_frm.frx":73B8
      MousePointer    =   99  'Custom
      Picture         =   "Pintu_frm.frx":76C2
      ToolTipText     =   "Tambah"
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   3360
      MouseIcon       =   "Pintu_frm.frx":838C
      MousePointer    =   99  'Custom
      Picture         =   "Pintu_frm.frx":8696
      ToolTipText     =   "Edit"
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   3960
      MouseIcon       =   "Pintu_frm.frx":A018
      MousePointer    =   99  'Custom
      Picture         =   "Pintu_frm.frx":A322
      ToolTipText     =   "Hapus"
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   4080
      MouseIcon       =   "Pintu_frm.frx":AFEC
      MousePointer    =   99  'Custom
      Picture         =   "Pintu_frm.frx":B2F6
      ToolTipText     =   "Keluar"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PINTU"
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
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Pintu_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim Pintu1 As String

Sub Kosong()
txtnama.Text = ""
End Sub

Sub Isi()
If Not Data1(0).Recordset.BOF And Image2(1).Visible = True Then
    txtnama.Text = Data1(0).Recordset!namapintu
    cmbjenis.Text = Data1(0).Recordset!jenispintu
End If
End Sub

Private Sub Data1_Reposition(Index As Integer)
Select Case Index
Case 0
If Image2(1).Visible = True Then
    Isi
End If
End Select
End Sub

Private Sub Form_Load()
Call bukadaTa
Data1(0).DatabaseName = db
Data1(0).Connect = dbCon
isiGrid1
Frame1.Enabled = False
btnAwal
datacek.DatabaseName = db
datacek.Connect = dbCon
cmbjenis.ListIndex = 0
End Sub

Sub isiGrid1()
Data1(0).RecordSource = "Select * from mspintu"
Data1(0).Refresh
End Sub

Sub btnAwal()
Image2(1).Visible = True
Image2(1).ToolTipText = "Tambah"
Image2(0).Picture = ImageList1.ListImages(4).Picture
Image2(0).ToolTipText = "Edit"
Image2(2).ToolTipText = "Hapus"
Frame1.Enabled = False
SSDBGrid1(0).Enabled = True
End Sub

Sub btnSimpan()
Image2(1).Visible = False
Image2(1).ToolTipText = "Tambah"
Image2(0).Picture = ImageList1.ListImages(7).Picture
Image2(0).ToolTipText = "Simpan"
Image2(2).ToolTipText = "Batal"
Frame1.Enabled = True
SSDBGrid1(0).Enabled = False
End Sub

Private Sub Image2_Click(Index As Integer)
Select Case Index
Case 0 'edit
    If Image2(0).ToolTipText = "Edit" Then
        tambah = False
        btnSimpan
        Pintu1 = txtnama.Text
        txtnama.SetFocus
    Else
        If txtnama.Text = "" Or cmbjenis.Text = "" Then
            MsgBox ("Data belum lengkap")
        Else
            If tambah = True Then 'simpan tambah
                'cek user sudah ada atau belum
                datacek.RecordSource = "select * from mspintu where namapintu=" & kutip(txtnama.Text)
                datacek.Refresh
                If datacek.Recordset.BOF Then
                    With Data1(0).Recordset
                        .AddNew
                        !namapintu = txtnama.Text
                        !jenispintu = cmbjenis.Text
                        .Update
                    End With
                    MsgBox "Berhasil tambah data"
                    btnAwal
                    Isi
                Else
                    MsgBox "User name sudah ada"
                End If
            Else 'Update data
                'cek user sudah ada atau belum
                datacek.RecordSource = "select * from mspintu where namapintu=" & kutip(txtnama.Text) & _
                " and namapintu<>" & kutip(Pintu1)
                datacek.Refresh
                If datacek.Recordset.BOF Then
                    With Data1(0).Recordset
                        .Edit
                        !namapintu = txtnama.Text
                        !jenispintu = cmbjenis.Text
                        .Update
                    End With
                    MsgBox "Berhasil update data"
                    btnAwal
                    Isi
                Else
                    MsgBox "User name sudah ada"
                End If
            End If
        End If
    End If
Case 1 'tambah
    Kosong
    Frame1.Enabled = True
    tambah = True
    SSDBGrid1(0).Enabled = False
    btnSimpan
    txtnama.SetFocus
Case 2 'hapus
    If Image2(2).ToolTipText = "Batal" Then
        btnAwal
        Isi
    Else
    Dim tny As String
    tny = MsgBox("Apakah anda yakin?", vbYesNo, "Hapus")
    If tny = vbYes Then
        With Data1(0).Recordset
            If .BOF Or .RecordCount = 1 Then
                MsgBox "Tidak dapat hapus data (data kosong/hanya 1)"
            Else
                .Delete
                .MoveFirst
                MsgBox "Berhasil hapus data"
                Isi
            End If
        End With
    End If
    End If
Case 3 'keluar
    Unload Me
End Select
End Sub

Private Sub cmbjenis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Image2_Click (0)
End If
End Sub

Private Sub txtnama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbjenis.SetFocus
End If
End Sub


