VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form SetApp_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengaturan Aplikasi"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6075
   Icon            =   "SetApp_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6075
   Begin VB.Data Data1 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PINTU KELUAR"
      Height          =   2415
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   5655
      Begin VB.CheckBox Cek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktif"
         Height          =   255
         Index           =   7
         Left            =   4800
         TabIndex        =   26
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtComKeluar 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   1800
         Width           =   495
      End
      Begin VB.CheckBox Cek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktif"
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox Cek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktif"
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox Cek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktif"
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox cmb6 
         DataField       =   " "
         DataSource      =   " "
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Text            =   " "
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox cmb5 
         DataField       =   " "
         DataSource      =   " "
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Text            =   " "
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cmb4 
         DataField       =   " "
         DataSource      =   " "
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Text            =   " "
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Com Port"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pintu"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kamera"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suara"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PINTU MASUK"
      Height          =   2295
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5655
      Begin VB.CheckBox Cek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktif"
         Height          =   255
         Index           =   6
         Left            =   4800
         TabIndex        =   25
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtcomMasuk 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   1800
         Width           =   495
      End
      Begin VB.CheckBox Cek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktif"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox Cek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktif"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox Cek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktif"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cmb3 
         DataField       =   " "
         DataSource      =   " "
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Text            =   " "
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox cmb2 
         DataField       =   " "
         DataSource      =   " "
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Text            =   " "
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cmb1 
         DataField       =   " "
         DataSource      =   " "
         Height          =   315
         ItemData        =   "SetApp_frm.frx":0CCA
         Left            =   1680
         List            =   "SetApp_frm.frx":0CCC
         TabIndex        =   9
         Text            =   " "
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Com Port"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suara"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kamera"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pintu"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   825
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   5040
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
            Picture         =   "SetApp_frm.frx":0CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetApp_frm.frx":19A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetApp_frm.frx":2682
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetApp_frm.frx":335C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetApp_frm.frx":4CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetApp_frm.frx":59C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetApp_frm.frx":66A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   5400
      MouseIcon       =   "SetApp_frm.frx":737C
      MousePointer    =   99  'Custom
      Picture         =   "SetApp_frm.frx":7686
      ToolTipText     =   "Keluar"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   4800
      MouseIcon       =   "SetApp_frm.frx":8350
      MousePointer    =   99  'Custom
      Picture         =   "SetApp_frm.frx":865A
      ToolTipText     =   "Hapus"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   4200
      MouseIcon       =   "SetApp_frm.frx":9324
      MousePointer    =   99  'Custom
      Picture         =   "SetApp_frm.frx":962E
      ToolTipText     =   "Edit"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PENGATURAN APLIKASI"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "SetApp_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
bukadaTa
Data1.DatabaseName = db
Data1.Connect = dbCon

IsiCmb

Data4.DatabaseName = db
Data4.Connect = dbCon
Data4.RecordSource = "select * from msPengaturan"
Data4.Refresh

Isi
BtnAwal
End Sub

Sub IsiCmb()
Data1.RecordSource = "select namapintu from mspintu where jenispintu='Pintu Masuk'"
Data1.Refresh
cmb1.Clear
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            cmb1.AddItem (!namapintu)
            .MoveNext
        Loop
        cmb1.ListIndex = 0
    End If
End With

Data1.RecordSource = "select namapintu from mspintu where jenispintu='Pintu Keluar'"
Data1.Refresh
cmb4.Clear
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            cmb4.AddItem (!namapintu)
            .MoveNext
        Loop
        cmb4.ListIndex = 0
    End If
End With

Data1.RecordSource = "select namacam from mscam"
Data1.Refresh
cmb2.Clear
cmb5.Clear
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            cmb2.AddItem (!namacam)
            cmb5.AddItem (!namacam)
            .MoveNext
        Loop
        cmb2.ListIndex = 0
        cmb5.ListIndex = 0
    End If
End With

Data1.RecordSource = "select namasuara from mssuara"
Data1.Refresh
cmb3.Clear
cmb6.Clear
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            cmb3.AddItem (!NamaSuara)
            cmb6.AddItem (!NamaSuara)
            .MoveNext
        Loop
        cmb3.ListIndex = 0
        cmb6.ListIndex = 0
    End If
End With

End Sub

Sub BtnAwal()
Image6.ToolTipText = "Edit"
Image6.Picture = ImageList1.ListImages(4).Picture
Image7.ToolTipText = "Batal"
Image7.Visible = False
Frame1(0).Enabled = False
Frame1(1).Enabled = False
Image6.Left = 4800
End Sub

Sub btnSimpan()
Image6.ToolTipText = "Simpan"
Image6.Picture = ImageList1.ListImages(7).Picture
Image7.ToolTipText = "Batal"
Image7.Visible = True
Frame1(0).Enabled = True
Frame1(1).Enabled = True
Image6.Left = 4200
End Sub

Private Sub Image2_Click(Index As Integer)
Unload Me
End Sub

Private Sub Image5_Click()

End Sub

Private Sub Image6_Click()
If Image6.ToolTipText = "Edit" Then
    btnSimpan
Else
    With Data4.Recordset
        .Edit
        !namapintumasuk = cmb1.Text
        !kameramasuk = cmb2.Text
        !suaramasuk = cmb3.Text
        !namapintukeluar = cmb4.Text
        !kamerakeluar = cmb5.Text
        !suarakeluar = cmb6.Text
        !comMasuk = txtcomMasuk.Text
        !comKeluar = txtComKeluar.Text
        If Cek(0).Value = 0 Then
            !stspintumasuk = "False"
        Else
            !stspintumasuk = "True"
        End If
        If Cek(1).Value = 0 Then
            !stskameramasuk = "False"
        Else
            !stskameramasuk = "True"
        End If
        If Cek(2).Value = 0 Then
            !stssuaramasuk = "False"
        Else
            !stssuaramasuk = "True"
        End If
        If Cek(3).Value = 0 Then
            !stspintukeluar = "False"
        Else
            !stspintukeluar = "True"
        End If
        If Cek(4).Value = 0 Then
            !stskamerakeluar = "False"
        Else
            !stskamerakeluar = "True"
        End If
        If Cek(5).Value = 0 Then
            !stssuarakeluar = "False"
        Else
            !stssuarakeluar = "True"
        End If
        If Cek(6).Value = 0 Then
            !stscommasuk = "False"
        Else
            !stscommasuk = "True"
        End If
        If Cek(7).Value = 0 Then
            !stscomkeluar = "False"
        Else
            !stscomkeluar = "True"
        End If
        .Update
    End With
    Data4.Refresh
    MsgBox "Berhasil update data"
    BtnAwal
    Isi
End If
End Sub

Sub Isi()
With Data4.Recordset
    If Not .BOF Then
        cmb1.Text = !namapintumasuk
        cmb2.Text = !kameramasuk
        cmb3.Text = !suaramasuk
        cmb4.Text = !namapintukeluar
        cmb5.Text = !kamerakeluar
        cmb6.Text = !suarakeluar
        txtcomMasuk.Text = !comMasuk
        txtComKeluar.Text = !comKeluar
        Cek(0).Value = IIf(!stspintumasuk = "True", 1, 0)
        Cek(1).Value = IIf(!stskameramasuk = "True", 1, 0)
        Cek(2).Value = IIf(!stssuaramasuk = "True", 1, 0)
        Cek(3).Value = IIf(!stspintukeluar = "True", 1, 0)
        Cek(4).Value = IIf(!stskamerakeluar = "True", 1, 0)
        Cek(5).Value = IIf(!stssuarakeluar = "True", 1, 0)
        Cek(6).Value = IIf(!stscommasuk = "True", 1, 0)
        Cek(7).Value = IIf(!stscomkeluar = "True", 1, 0)
    End If
End With
End Sub

Private Sub Image7_Click()
    BtnAwal
    Isi
End Sub
