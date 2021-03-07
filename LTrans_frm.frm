VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form LTrans_frm 
   Caption         =   "Laporan Transaksi Kartu Master"
   ClientHeight    =   1290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   1935
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
      Top             =   1440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtp1 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   47841281
      CurrentDate     =   42466
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   47841281
      CurrentDate     =   42466
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sort By"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "s/d Tanggal"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "LTrans_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sql As String
sql = "select * from view_trans1 where format(tgljam,'MM/dd/yyyy')>=" + kutip(Format(dtp1.Value, "MM/dd/yyyy")) + _
    " and format(tgljam,'MM/dd/yyyy')<=" + kutip(Format(dtp2.Value, "MM/dd/yyyy"))

    Select Case Combo1.ListIndex
    Case 0
        sql = sql + " order by tgljam"
    Case 1
        sql = sql + " order by sts"
    Case 2
        sql = sql + " order by namapintu"
    Case 3
        sql = sql + " order by nama"
    Case 4
        sql = sql + " order by nopol"
    Case 5
        sql = sql + " order by RFID"
    End Select
    
    data1.RecordSource = sql
    data1.Refresh


    Dim ttl As Integer
    ttl = data1.Recordset.RecordCount
    If ttl > 0 Then
        With LapTrans_rpt
            .Show
            .WindowState = 2
            .Label2.Caption = "Periode : " & Format(dtp1.Value, "dd MMM yyyy") + " s/d " + Format(dtp2.Value, "dd MMM yyyy")
            .Data.DatabaseName = db
            .Data.Connect = dbCon
            .Data.RecordSource = data1.RecordSource
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
bukadaTa
data1.DatabaseName = db
data1.Connect = dbCon

dtp1.Value = Now
dtp2.Value = Now

Combo1.Clear
Combo1.AddItem "Tanggal Jam"
Combo1.AddItem "Status In/Out"
Combo1.AddItem "Nama Pintu"
Combo1.AddItem "Nama"
Combo1.AddItem "Nopol"
Combo1.AddItem "RFID"
Combo1.ListIndex = 0

End Sub
