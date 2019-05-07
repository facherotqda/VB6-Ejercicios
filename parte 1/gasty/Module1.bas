Attribute VB_Name = "Module1"
Option Explicit

Global rs As New ADODB.Recordset
Global cn As New ADODB.Connection
Dim estado As Boolean




Sub main()
Set cn = New ADODB.Connection
With cn
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Consultas;Data Source=KINGFAT-PC"
        .Open
         If .State = 1 Then
    
          MsgBox "Conectado a la Bd", vbInformation, "CONECTADO"
          estado = True
    
         Else
    
         MsgBox "Error en la coneccion", vbInformation, "ERROR"
         estado = False
         End If
End With
If estado = True Then
Form1.Show
End If
End Sub


