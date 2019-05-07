Attribute VB_Name = "ModuloSentencias"
Option Explicit
Dim estado As Boolean

Sub main()

With Conec
    .CursorLocation = adUseClient
    .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=OFFFF;Data Source=KINGFAT-PC"
    
    If .State = 1 Then
    
    MsgBox "Conectado a la Bd", vbInformation, "CONECTADO"
    estado = True
    
    Else
    
    MsgBox "Error en la coneccion", vbInformation, "ERROR"
    estado = False
    End If
    
End With   'Fin del Conec (coneccion a base de datos)

If estado = True Then
    frmBaseDatos.Show
    
End If
    
End Sub

