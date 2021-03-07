VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} DataMember_rpt 
   Caption         =   "Data Member"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13230
   Icon            =   "DataMember_rpt.dsx":0000
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   23336
   _ExtentY        =   12541
   SectionData     =   "DataMember_rpt.dsx":0CCA
End
Attribute VB_Name = "DataMember_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
If Field8.Text = 0 Then
    Field8.Text = "Tdk"
Else
    Field8.Text = "Ya"
End If
If Field9.Text = 0 Then
    Field9.Text = "Tdk"
Else
    Field9.Text = "Ya"
End If
End Sub

Private Sub PageHeader_Format()
Label11.Caption = Format(Now, "dd MMM yyyy HH:mm:ss")
End Sub
