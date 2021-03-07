VERSION 5.00
Begin VB.Form Login_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   Icon            =   "Login_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keluar"
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   5520
      MouseIcon       =   "Login_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "Login_frm.frx":0FD4
      ToolTipText     =   "Keluar"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   4800
      MouseIcon       =   "Login_frm.frx":1C9E
      MousePointer    =   99  'Custom
      Picture         =   "Login_frm.frx":1FA8
      ToolTipText     =   "Login"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   2160
      MouseIcon       =   "Login_frm.frx":2C72
      MousePointer    =   99  'Custom
      Picture         =   "Login_frm.frx":2F7C
      ToolTipText     =   "Info Aplikasi"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "www.EQ-PARK.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   120
      Picture         =   "Login_frm.frx":3C46
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Login_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call bukadaTa
Data1.DatabaseName = db
Data1.Connect = dbCon
End Sub

Private Sub Image2_Click(Index As Integer)
Select Case Index
Case 0
Case 1
    Dim sql As String
    sql = "Select * from msuser where username=" & kutip(Text1(0).Text) & " and pass=" & kutip(Text1(1).Text) '& "'"
    Data1.RecordSource = sql
    Data1.Refresh
    With Data1.Recordset
        If .BOF Then
            MsgBox ("Username/Password tidak benar")
        Else
            MDIForm1.StatusBar1.Panels(2).Text = UCase(Text1(0).Text)
            MDIForm1.Show
            Unload Me
        End If
    End With
Case 2
    Unload Me
End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0
If KeyAscii = 13 Then
    Text1(1).SetFocus
End If
Case 1
If KeyAscii = 13 Then
    Image2_Click (1)
End If
End Select
End Sub
