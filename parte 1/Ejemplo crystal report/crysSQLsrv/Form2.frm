VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form2"
   ScaleHeight     =   4905
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   645
      Left            =   2520
      TabIndex        =   0
      Top             =   1140
      Width           =   1065
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private con As ADODB.Connection
Private rs As ADODB.Recordset
Private CrysApp As New CRAXDDRT.Application
Private CrysRep As New CRAXDDRT.Report

Private Sub Command1_Click()
    Set CrysRep = CrysApp.OpenReport("e:\emp_rep.rpt")
    Call CrysRep.Database.AddOLEDBSource("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=sa;Initial Catalog=hr_data;Data Source=KRPDD156", "employee_master")
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    If con.State = 1 Then con.Close
    con.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=sa;Initial Catalog=hr_data;Data Source=KRPDD156"
    con.Open
    If rs.State = 1 Then rs.Close
    rs.Open "select emp_code,emp_name from employee_master where emp_grade='S'", con
    If Not rs.EOF Then
        With CrysRep
            Call .Database.Tables(1).SetDataSource(rs)
            .DiscardSavedData
        End With
    End If
    Call CrysRep.ParameterFields(1).AddCurrentValue("empgrade")
    Form1.CRViewer.ReportSource = CrysRep
    Form1.CRViewer.ViewReport
    Form1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rs.State = 1 Then rs.Close
    If con.State = 1 Then con.Close
End Sub


