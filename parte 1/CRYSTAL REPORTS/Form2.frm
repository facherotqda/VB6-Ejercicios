VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Reportes"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6705
   LinkTopic       =   "Form2"
   ScaleHeight     =   2925
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir Pacientes"
      Height          =   1095
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir Clientes"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As ADODB.Connection
Private rs As ADODB.Recordset
Private CrysApp As New CRAXDDRT.Application
Private CrysRep As New CRAXDDRT.Report

Private Sub Command1_Click()

 Set CrysRep = CrysApp.OpenReport("C:\Users\KingFat\Desktop\CRYSTAL REPORTS\Report1DbConsultas.rpt")
    Call CrysRep.Database.AddOLEDBSource("Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Consultas;Data Source=KINGFAT-PC", "CLIENTES")
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    If con.State = 1 Then con.Close
    con.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Consultas;Data Source=KINGFAT-PC"
    con.Open
    If rs.State = 1 Then rs.Close
    'selecciono en la consulta lo que quiero mostrar
    'rs.Open "select * from CLIENTES where POBLACIÓN='MADRID'", con
    rs.Open "select * from CLIENTES", con
    If Not rs.EOF Then
        With CrysRep
            Call .Database.Tables(1).SetDataSource(rs)
            .DiscardSavedData
        End With
    End If
    
    'Call CrysRep.ParameterFields(1).AddCurrentValue("CLIENTES")
    
    Form1.CRViewer.ReportSource = CrysRep
    Form1.CRViewer.ViewReport
    Form1.Show

End Sub

Private Sub Command2_Click()

Set CrysRep = CrysApp.OpenReport("C:\Users\KingFat\Desktop\CRYSTAL REPORTS\CrystalReportsHospital.rpt")
Call CrysRep.Database.AddOLEDBSource("Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=KINGFAT-PC", "Hospital")
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset

If con.State = 1 Then con.Close
con.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=KINGFAT-PC"

con.Open

If rs.State = 1 Then rs.Close
rs.Open "Select * from Hospital", con
    If Not rs.EOF Then
        With CrysRep
        Call .Database.Tables(1).SetDataSource(rs)
        .DiscardSavedData
        End With
        
    End If
         
    Call CrysRep.ParameterFields(1).AddCurrentValue("parHospital")
    
    Form3.CRViewer1.ReportSource = CrysRep
    Form3.CRViewer1.ViewReport
    Form3.Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rs.State = 1 Then rs.Close
    If con.State = 1 Then con.Close
End Sub

