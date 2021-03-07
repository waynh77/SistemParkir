VERSION 5.00
Begin VB.Form GantiPass_frm 
   Caption         =   "Ganti Password"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6045
   Icon            =   "GantiPass_frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "BATAL"
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Konfirmasi"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password Baru"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1065
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password Lama"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1125
   End
End
Attribute VB_Name = "GantiPass_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Data1.RecordSource = "Select * from msuser where username=" + kutip(MDIForm1.StatusBar1.Panels(2).Text) + _
    " and pass=" + kutip(Text1(0).Text)
    Data1.Refresh
    With Data1.Recordset
        If .BOF Then
            MsgBox ("Password lama salah")
        Else
            If Text1(1).Text <> Text1(2).Text Then
                MsgBox ("Password baru tidak dapat dikonfirmasi")
            Else
                .Edit
                !pass = Text1(2).Text
                .Update
                MsgBox ("Berhasil update password")
                Unload Me
            End If
        End If
    End With
Case 1
    Unload Me
End Select
End Sub

Private Sub Form_Load()
bukadaTa
Data1.DatabaseName = db
Data1.Connect = dbCon
End Sub
