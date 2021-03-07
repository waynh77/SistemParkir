VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form User_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengaturan Pengguna"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9450
   Icon            =   "User_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9450
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   3435
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   3240
      Width           =   9015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   9720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datacek 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   9720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11280
      Top             =   1080
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
            Picture         =   "User_frm.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":19A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":267E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":3358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":4CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":59C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":669E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   4800
      TabIndex        =   3
      Top             =   960
      Width           =   4335
      Begin VB.TextBox txtUser 
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtPass 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Index           =   1
      Left            =   12000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "User_frm.frx":7378
      Height          =   2535
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4335
      _Version        =   196614
      AllowUpdate     =   0   'False
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "User Name"
      Columns(0).Name =   "User Name"
      Columns(0).DataField=   "username"
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Password"
      Columns(1).Name =   "Password"
      Columns(1).DataField=   "pass"
      Columns(1).FieldLen=   256
      _ExtentX        =   7646
      _ExtentY        =   4471
      _StockProps     =   79
      Caption         =   "Data User"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Index           =   0
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PENGATURAN PENGGUNA"
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
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   8640
      MouseIcon       =   "User_frm.frx":738F
      MousePointer    =   99  'Custom
      Picture         =   "User_frm.frx":7699
      ToolTipText     =   "Keluar"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   8640
      MouseIcon       =   "User_frm.frx":8363
      MousePointer    =   99  'Custom
      Picture         =   "User_frm.frx":866D
      ToolTipText     =   "Hapus"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   8040
      MouseIcon       =   "User_frm.frx":9337
      MousePointer    =   99  'Custom
      Picture         =   "User_frm.frx":9641
      ToolTipText     =   "Edit"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   7440
      MouseIcon       =   "User_frm.frx":AFC3
      MousePointer    =   99  'Custom
      Picture         =   "User_frm.frx":B2CD
      ToolTipText     =   "Tambah"
      Top             =   2520
      Width           =   480
   End
End
Attribute VB_Name = "User_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim User1 As String

Sub Kosong()
txtUser.Text = ""
txtPass.Text = ""
End Sub

Sub Isi()
If Not Data1(0).Recordset.BOF Then
    txtUser.Text = Data1(0).Recordset!UserName
    txtPass.Text = Data1(0).Recordset!pass
End If
End Sub


Private Sub Data1_Reposition(Index As Integer)
Select Case Index
Case 0
If Image2(1).Visible = True Then
    isiGrid2
    Isi
End If
End Select
End Sub

Private Sub Form_Load()
Call bukadaTa
Data1(1).DatabaseName = db
Data1(1).Connect = dbCon
Data1(0).DatabaseName = db
Data1(0).Connect = dbCon
isiGrid1
'isiGrid2
Frame1.Enabled = False
BtnAwal
datacek.DatabaseName = db
datacek.Connect = dbCon
'SSDBGrid1(1).AllowUpdate = False
Data2.DatabaseName = db
Data2.Connect = dbCon
End Sub

Sub isiGrid1()
Data1(0).RecordSource = "Select * from msuser"
Data1(0).Refresh
End Sub

Sub isiGrid2()
With Data1(1)
    .RecordSource = "Select * from view_usermenu where username=" & kutip(Data1(0).Recordset!UserName)
    .Refresh
    If .Recordset.RecordCount > 0 Then
        Dim baris As Integer
        List1.Clear
        .Recordset.MoveFirst
        baris = 0
        Do While Not .Recordset.EOF
            List1.AddItem .Recordset!idmenu & " - " & .Recordset!namamenu & " :" & .Recordset!keterangan
            If .Recordset!sts = "True" Then
                List1.Selected(baris) = True
            End If
            baris = baris + 1
            .Recordset.MoveNext
        Loop
    End If
End With
End Sub

Sub listKosong()
With Data1(1)
    .RecordSource = "Select * from view_usermenu where username=" & kutip(Data1(0).Recordset!UserName)
    .Refresh
    If .Recordset.RecordCount > 0 Then
        Dim baris As Integer
        List1.Clear
        .Recordset.MoveFirst
        baris = 0
        Do While Not .Recordset.EOF
            List1.AddItem .Recordset!idmenu & " - " & .Recordset!namamenu & " :" & .Recordset!keterangan
            baris = baris + 1
            .Recordset.MoveNext
        Loop
    End If
End With
End Sub

Sub BtnAwal()
Image2(1).Visible = True
Image2(1).ToolTipText = "Tambah"
Image2(0).Picture = ImageList1.ListImages(4).Picture
Image2(0).ToolTipText = "Edit"
Image2(2).ToolTipText = "Hapus"
Frame1.Enabled = False
SSDBGrid1(0).Enabled = True
List1.Enabled = False
End Sub

Sub btnSimpan()
Image2(1).Visible = False
Image2(1).ToolTipText = "Tambah"
Image2(0).Picture = ImageList1.ListImages(7).Picture
Image2(0).ToolTipText = "Simpan"
Image2(2).ToolTipText = "Batal"
Frame1.Enabled = True
SSDBGrid1(0).Enabled = False
List1.Enabled = True
End Sub

Private Sub Image2_Click(Index As Integer)
Dim Id As Integer
Dim idmenu As String
Select Case Index
Case 0 'edit
    If Image2(0).ToolTipText = "Edit" Then
        tambah = False
        btnSimpan
        User1 = txtUser.Text
        txtUser.SetFocus
    Else
        If txtUser.Text = "" Or txtPass.Text = "" Then
            MsgBox ("Data belum lengkap")
        Else
            If tambah = True Then 'simpan tambah
                'cek user sudah ada atau belum
                datacek.RecordSource = "select * from msuser where username=" & kutip(txtUser.Text)
                datacek.Refresh
                If datacek.Recordset.BOF Then
                    With Data1(0).Recordset
                        .AddNew
                        !UserName = txtUser.Text
                        !pass = txtPass.Text
                        .Update
                    End With
                    
                    For i = 0 To List1.ListCount - 1
                        Data2.RecordSource = "select * from trusermenu order by id desc"
                        Data2.Refresh
                        If Data2.Recordset.BOF Then
                            Id = 1
                        Else
                            Id = Data2.Recordset!Id + 1
                        End If
                        
                        idmenu = Left(List1.List(i), 5)
                        
                        Data2.Recordset.AddNew
                        Data2.Recordset!Id = Id
                        Data2.Recordset!UserName = txtUser.Text
                        Data2.Recordset!idmenu = idmenu
                        If List1.Selected(i) = True Then
                            Data2.Recordset!sts = "True"
                        Else
                            Data2.Recordset!sts = 0
                        End If
                        Data2.Recordset.Update
                        
                    Next
                    
                    MsgBox "Berhasil tambah data"
                    BtnAwal
                    Isi
                    isiGrid2
                Else
                    MsgBox "User name sudah ada"
                End If
            Else 'Update data
                'cek user sudah ada atau belum
                datacek.RecordSource = "select * from msuser where username=" & kutip(txtUser.Text) & _
                " and username<>" & kutip(User1)
                datacek.Refresh
                If datacek.Recordset.BOF Then
                    With Data1(0).Recordset
                        .Edit
                        !UserName = txtUser.Text
                        !pass = txtPass.Text
                        .Update
                    End With
                    For i = 0 To List1.ListCount - 1
                        idmenu = Left(List1.List(i), 5)
                        Data2.RecordSource = "select * from trusermenu where username=" & kutip(User1) & _
                                                " and idmenu=" & idmenu
                        Data2.Refresh
                        If Data2.Recordset.RecordCount > 0 Then
                            Data2.Recordset.Edit
                        Else
                            Data2.RecordSource = "select * from trusermenu order by id desc"
                            Data2.Refresh
                            If Data2.Recordset.BOF Then
                                Id = 1
                            Else
                                Id = Data2.Recordset!Id + 1
                            End If

                            Data2.Recordset.AddNew
                            Data2.Recordset!Id = Id
                        End If
                        Data2.Recordset!UserName = txtUser.Text
                        Data2.Recordset!idmenu = idmenu
                        If List1.Selected(i) = True Then
                            Data2.Recordset!sts = "True"
                        Else
                            Data2.Recordset!sts = 0
                        End If
                        Data2.Recordset.Update
                    Next
                    MsgBox "Berhasil update data"
                    BtnAwal
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
    txtUser.SetFocus
    listKosong
Case 2 'hapus
    If Image2(2).ToolTipText = "Batal" Then
        BtnAwal
        Isi
    Else
    Dim tny As String
    tny = MsgBox("Apakah anda yakin?", vbYesNo, "Hapus")
    If tny = vbYes Then
        With Data1(0).Recordset
            If .BOF Or .RecordCount = 1 Then
                MsgBox "Tidak dapat hapus data (data kosong/hanya 1)"
            Else
                Data2.RecordSource = "select * from trusermenu where username=" + kutip(txtUser.Text)
                Data2.Refresh
                Do While Not Data2.Recordset.EOF
                    Data2.Recordset.Delete
                    Data2.Recordset.MoveNext
                Loop
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

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Image2_Click (0)
End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPass.SetFocus
End If
End Sub
