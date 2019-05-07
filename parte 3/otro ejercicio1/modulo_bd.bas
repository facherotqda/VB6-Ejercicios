Attribute VB_Name = "modulo_bd"
Option Explicit

Global cn As New ADODB.Connection
Global rs_sp As New ADODB.Recordset
Global rs_gasty As New ADODB.Recordset


Sub main()

Set cn = New ADODB.Connection

With cn
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=sistran;Data Source=KINGFAT-PC"
    .Open
    If .State = 1 Then
    MsgBox "Conectado a la bd"
    Else
    MsgBox "No Conectado"
    End If
    
End With

Form1.Show
End Sub
