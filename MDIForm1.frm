VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "AKSES KONTROL"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13950
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   13890
      TabIndex        =   6
      Top             =   9465
      Width           =   13950
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   3600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tampil Menu Samping"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   120
         MouseIcon       =   "MDIForm1.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "MDIForm1.frx":0FD4
         ToolTipText     =   "Show Menu"
         Top             =   100
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      Height          =   9465
      Left            =   0
      ScaleHeight     =   9405
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      Begin VB.PictureBox PicOriginal 
         AutoSize        =   -1  'True
         Height          =   64860
         Left            =   120
         Picture         =   "MDIForm1.frx":1C9E
         ScaleHeight     =   64800
         ScaleWidth      =   1.21200e5
         TabIndex        =   3
         Top             =   6840
         Visible         =   0   'False
         Width           =   1.21260e5
      End
      Begin VB.PictureBox PicStretched 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   -240
         ScaleHeight     =   705
         ScaleWidth      =   13950
         TabIndex        =   2
         Top             =   6600
         Visible         =   0   'False
         Width           =   13950
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Hide Menu"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         MouseIcon       =   "MDIForm1.frx":AA3D2
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   8520
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   480
         MouseIcon       =   "MDIForm1.frx":AA6DC
         MousePointer    =   99  'Custom
         Picture         =   "MDIForm1.frx":AA9E6
         ToolTipText     =   "Hide Menu"
         Top             =   7920
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Pintu"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         MouseIcon       =   "MDIForm1.frx":AB6B0
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   480
         MouseIcon       =   "MDIForm1.frx":AB9BA
         MousePointer    =   99  'Custom
         Picture         =   "MDIForm1.frx":ABCC4
         ToolTipText     =   "Transaksi Pintu"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Member"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         MouseIcon       =   "MDIForm1.frx":AC98E
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         MouseIcon       =   "MDIForm1.frx":ACC98
         MousePointer    =   99  'Custom
         Picture         =   "MDIForm1.frx":ACFA2
         ToolTipText     =   "Transaksi Member"
         Top             =   120
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   10155
      Width           =   13950
      _ExtentX        =   24606
      _ExtentY        =   847
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3863
            Text            =   "www.EQ_PARK.com"
            TextSave        =   "www.EQ_PARK.com"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   4577
            Text            =   "User Name"
            TextSave        =   "User Name"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   2
            Picture         =   "MDIForm1.frx":ADC6C
            TextSave        =   "16.39"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2752
            Picture         =   "MDIForm1.frx":AE3E6
            TextSave        =   "18/11/2016"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu adm_mnu 
      Caption         =   "&Admin"
      Begin VB.Menu user_mnu 
         Caption         =   "&1. Pengaturan Pengguna"
      End
      Begin VB.Menu alat_mnu 
         Caption         =   "&2. Pengaturan Alat"
         Begin VB.Menu cam_mnu 
            Caption         =   "&a. Kamera"
         End
         Begin VB.Menu suara_mnu 
            Caption         =   "&b. Suara"
         End
         Begin VB.Menu pintu_mnu 
            Caption         =   "&c. Pintu"
         End
      End
      Begin VB.Menu app_mnu 
         Caption         =   "&3. Pengaturan Aplikasi"
      End
   End
   Begin VB.Menu Trans_mnu 
      Caption         =   "&Transaksi"
      Begin VB.Menu TrMember_mnu 
         Caption         =   "&1. Transaksi Member"
      End
      Begin VB.Menu gate_mnu 
         Caption         =   "&2. Transaksi Pintu"
      End
   End
   Begin VB.Menu lap_mnu 
      Caption         =   "&Laporan"
      Begin VB.Menu Lmember_mnu 
         Caption         =   "&1. Laporan Member"
      End
      Begin VB.Menu lpintu_mnu 
         Caption         =   "&2. Laporan Pintu"
         Begin VB.Menu lmaster_mnu 
            Caption         =   "Kartu Master"
         End
         Begin VB.Menu Lgatemember_mnu 
            Caption         =   "Kartu Member"
         End
      End
   End
   Begin VB.Menu tool_mnu 
      Caption         =   "&Peralatan"
      Begin VB.Menu pswd_mnu 
         Caption         =   "&1. Ganti Password"
      End
      Begin VB.Menu Logout_mnu 
         Caption         =   "&2. LogOut"
      End
      Begin VB.Menu Kalk_mnu 
         Caption         =   "&3. Kalkulator"
      End
   End
   Begin VB.Menu win_mnu 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu x_mnu 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub app_mnu_Click()
With SetApp_frm
    .Show
    .Top = 0
    .Left = 0
End With
End Sub

Private Sub cam_mnu_Click()
With Cam_frm
    .Show
    .Top = 0
    .Left = 0
End With
End Sub

Private Sub db_mnu_Click()

End Sub

Private Sub gate_mnu_Click()
With TrGate_frm
    .Show
    .Top = 0
    .Left = 0
    .WindowState = 2
    'Image3_Click
End With
End Sub

Private Sub Image1_Click()
TrMember_mnu_Click
End Sub

Private Sub Image2_Click()
gate_mnu_Click
End Sub

Private Sub Image3_Click()
Picture1.Visible = False
Picture2.Visible = True
MDIForm_Resize
End Sub


Private Sub Image4_Click()
Picture2.Visible = False
Picture1.Visible = True
MDIForm_Resize
End Sub

Private Sub Kalk_mnu_Click()
Call Shell("calc", vbNormalFocus)
End Sub

Private Sub Lgatemember_mnu_Click()
With LTrMember
    .Show
End With
End Sub

Private Sub lmaster_mnu_Click()
LTrans_frm.Show
End Sub

Private Sub Lmember_mnu_Click()
LMember_frm.Show
End Sub

Private Sub Logout_mnu_Click()
Dim tny As String
tny = MsgBox("Apakah anda yakin?", vbYesNo, "Logout")
If tny = vbYes Then
    Login_frm.Show
    Login_frm.Text1(0).Text = ""
    Login_frm.Text1(1).Text = ""
    Unload Me
End If
End Sub

Private Sub lpintu_mnu_Click()
'
End Sub

Private Sub MDIForm_Activate()
cekUser
End Sub

Private Sub MDIForm_Load()
bukadaTa
Picture2.Visible = False
Picture1.Visible = True
End Sub

Sub cekUser()
Dim idmenu As Integer
Data1.DatabaseName = db
Data1.Connect = dbCon
Data1.RecordSource = "Select * from trusermenu where username=" + kutip(StatusBar1.Panels(2).Text) + " order by idmenu"
Data1.Refresh
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        
        If !sts = 0 Then
            adm_mnu.Visible = False
        End If
        .MoveNext
        
        If adm_mnu.Visible = True And !sts = 0 Then
            user_mnu.Visible = False
        End If
        .MoveNext
        
        If adm_mnu.Visible = True And !sts = 0 Then
            alat_mnu.Visible = False
        End If
        .MoveNext

        If adm_mnu.Visible = True And alat_mnu.Visible = True And !sts = 0 Then
            cam_mnu.Visible = False
        End If
        .MoveNext
        
        If adm_mnu.Visible = True And alat_mnu.Visible = True And !sts = 0 Then
            suara_mnu.Visible = False
        End If
        .MoveNext
        
        If adm_mnu.Visible = True And alat_mnu.Visible = True And !sts = 0 Then
            pintu_mnu.Visible = False
        End If
        .MoveNext
    
        If adm_mnu.Visible = True And !sts = 0 Then
            app_mnu.Visible = False
        End If
        .MoveNext
    
        If !sts = 0 Then
            Trans_mnu.Visible = False
        End If
        .MoveNext
    
        If Trans_mnu.Visible = True And !sts = 0 Then
            TrMember_mnu.Visible = False
        End If
        If !sts = 0 Then
           Image1.Enabled = False
        End If
        .MoveNext
    
        If Trans_mnu.Visible = True And !sts = 0 Then
            gate_mnu.Visible = False
        End If
        If !sts = 0 Then
            Image2.Enabled = False
        End If
        .MoveNext
    
        If !sts = 0 Then
            lap_mnu.Visible = False
        End If
        .MoveNext
    
        If lap_mnu.Visible = True And !sts = 0 Then
            Lmember_mnu.Visible = False
        End If
        .MoveNext
    
        If lap_mnu.Visible = True And !sts = 0 Then
            lpintu_mnu.Visible = False
        End If
        .MoveNext
    
        If lap_mnu.Visible = True And lpintu_mnu.Visible = True And !sts = 0 Then
            lmaster_mnu.Visible = False
        End If
        .MoveNext
    
        If lap_mnu.Visible = True And lpintu_mnu.Visible = True And !sts = 0 Then
            Lgatemember_mnu.Visible = False
        End If
        .MoveNext
    End If
End With
End Sub

Private Sub MDIForm_Resize()
Dim client_rect As RECT
Dim client_hwnd As Long

    PicStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight

    ' Copy the original picture into picStretched.
    PicStretched.PaintPicture _
        PicOriginal.Picture, _
        0, 0, _
        PicStretched.ScaleWidth, _
        PicStretched.ScaleHeight, _
        0, 0, _
        PicOriginal.ScaleWidth, _
        PicOriginal.ScaleHeight
        
        
    ' Set the MDI form's picture.
    Picture = PicStretched.Image

    ' Invalidate the picture.
    client_hwnd = FindWindowEx(Me.hWnd, 0, "MDIClient", vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1

    PicStretched.Visible = False
End Sub

Private Sub pintu_mnu_Click()
With Pintu_frm
    .Show
    .Top = 0
    .Left = 0
End With
End Sub

Private Sub pswd_mnu_Click()
GantiPass_frm.Show
End Sub

Private Sub suara_mnu_Click()
With Suara_frm
    .Show
    .Top = 0
    .Left = 0
End With
End Sub

Private Sub TrMember_mnu_Click()
With Member_frm
    .Show
    .Top = 0
    .Left = 0
End With

End Sub

Private Sub user_mnu_Click()
With User_frm
    .Show
    .Top = 0
    .Left = 0
End With
End Sub

Private Sub x_mnu_Click()
Dim tny As String
tny = MsgBox("Apakah anda yakin?", vbYesNo, "Keluar")
If tny = vbYes Then
    End
End If
End Sub
