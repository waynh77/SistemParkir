VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LapMember_rpt 
   Caption         =   "Laporan Data Member"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11850
   MDIChild        =   -1  'True
   _ExtentX        =   20902
   _ExtentY        =   12303
   SectionData     =   "LapMember_rpt.dsx":0000
End
Attribute VB_Name = "LapMember_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Label9.Caption = "Tanggal Cetak : " + Format(Now, "d MMM yyyy HH:mm:ss")
End Sub

Private Sub Detail_Format()
With Data.Recordset
    If !blokir = 0 Then
        Field6.Text = "Tidak"
    Else
        Field6.Text = "Ya"
    End If
End With
End Sub
