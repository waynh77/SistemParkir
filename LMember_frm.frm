VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form LMember_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Data Member"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5700
   ControlBox      =   0   'False
   Icon            =   "LMember_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtp1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   211812353
      CurrentDate     =   42466
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tanggal Daftar"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   211812353
      CurrentDate     =   42466
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sort By"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "s/d Tanggal"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "LMember_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sql As String
    sql = "select * from msmember "
    
    If Check1.Value = 1 Then
    sql = sql + "where format(tgldaftar,'MM/dd/yyyy')>=" & kutip(Format(dtp1.Value, "MM/dd/yyyy")) & _
                        " and format(tgldaftar,'MM/dd/yyyy')<=" & kutip(Format(dtp2.Value + 1, "MM/dd/yyyy"))
    End If
    
    Select Case Combo1.ListIndex
    Case 0
        sql = sql + " order by nama"
    Case 1
        sql = sql + " order by alamat"
    Case 2
        sql = sql + " order by nopol"
    Case 3
        sql = sql + " order by tgldaftar"
    Case 4
        sql = sql + " order by tglexp"
    Case 5
        sql = sql + " order by RFID"
    Case 6
        sql = sql + " order by Blokir"
    Case 7
        sql = sql + " order by master"
    Case 8
        sql = sql + " order by Catatan"
    End Select
    
    data1.RecordSource = sql
    data1.Refresh
    
    Dim ttl As Integer
    ttl = data1.Recordset.RecordCount
    If ttl > 0 Then
        With DataMember_rpt
            .data1.DatabaseName = db
            .data1.Connect = dbCon
            .data1.RecordSource = data1.RecordSource
            .Show
            .WindowState = 2
            '.Label2.Caption = "Periode : " & Format(dtp1.Value, "dd MMM yyyy") + " s/d " + Format(dtp2.Value, "dd MMM yyyy")
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
Combo1.AddItem "Nama"
Combo1.AddItem "Alamat"
Combo1.AddItem "Nopol"
Combo1.AddItem "Tanggal Daftar"
Combo1.AddItem "Tanggal Expired"
Combo1.AddItem "RFID"
Combo1.AddItem "Blokir"
Combo1.AddItem "Kartu Master"
Combo1.AddItem "Catatan"
Combo1.ListIndex = 0

End Sub
