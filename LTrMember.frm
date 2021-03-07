VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form LTrMember 
   Caption         =   "Lap. Transaksi Member"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5190
   Icon            =   "LTrMember.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtp1 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46792705
      CurrentDate     =   42466
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46792705
      CurrentDate     =   42466
   End
   Begin VB.Label Label1 
      Caption         =   "s/d Tanggal"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sort By"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   510
   End
End
Attribute VB_Name = "LTrMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sql As String
sql = "select * from view_trmember where format(tgljammasuk,'MM/dd/yyyy')>=" + kutip(Format(dtp1.Value, "MM/dd/yyyy")) + _
    " and format(tgljamkeluar,'MM/dd/yyyy')<=" + kutip(Format(dtp2.Value, "MM/dd/yyyy"))

    Select Case Combo1.ListIndex
        Case 0
            sql = sql + " order by tgljammasuk"
        Case 1
            sql = sql + " order by pintumasuk"
        Case 2
            sql = sql + " order by tgljamkeluar"
        Case 3
            sql = sql + " order by pintukeluar"
        Case 4
            sql = sql + " order by nama"
        Case 5
            sql = sql + " order by nopol"
        Case 6
            sql = sql + " order by rfid"
        Case 7
            sql = sql + " order by sts"
    End Select
    
    Data1.RecordSource = sql
    Data1.Refresh


    Dim ttl As Integer
    ttl = Data1.Recordset.RecordCount
    If ttl > 0 Then
        With CtkTransMember
            .Show
            .WindowState = 2
            .Label18.Caption = "Periode : " & Format(dtp1.Value, "dd MMM yyyy") + " s/d " + Format(dtp2.Value, "dd MMM yyyy")
            .Label18.Visible = True
            .Data1.DatabaseName = db
            .Data1.Connect = dbCon
            .Data1.RecordSource = Data1.RecordSource
            Unload Me
        End With
    Else
        MsgBox ("Data tidak diketemukan")
    End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call bukadaTa
Data1.DatabaseName = db
Data1.Connect = dbCon

dtp1.Value = Now
dtp2.Value = Now

isiCmb

End Sub

Sub isiCmb()
Combo1.Clear
Combo1.AddItem "Tanggal Jam Masuk"
Combo1.AddItem "Pintu Masuk"
Combo1.AddItem "Tanggal Jam Keluar"
Combo1.AddItem "Pintu Keluar"
Combo1.AddItem "Nama"
Combo1.AddItem "Nopol"
Combo1.AddItem "RFID"
Combo1.AddItem "Status"
Combo1.ListIndex = 0
End Sub
