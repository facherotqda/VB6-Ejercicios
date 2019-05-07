Attribute VB_Name = "Modulo_Conex"
Option Explicit

Global cn As New ADODB.Connection
Global rs_tablaLibro As New ADODB.Recordset
Global rs_SpLibro As New ADODB.Recordset


Sub main()

Dim estado As Boolean
Set cn = New ADODB.Connection

With cn
        .CursorLocation = adUseClient
        .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Prueba;Data Source=KINGFAT-PC"
        
            If .State = 1 Then
            MsgBox "Conectando", vbInformation, "OK"
            estado = True
         Else
         MsgBox "No se pudo conectar a la BD", vbInformation, "Error"
            estado = False
         End If
End With
        
If estado = True Then

Form1.Show
End If
End Sub

 Sub CrearTablaLibro()


        With rs_tablaLibro
        
        If .State = 1 Then .Close
            .Open "SELECT * FROM Libro", cn, adOpenStatic, adLockOptimistic
            
            
        End With
        
End Sub

Sub LLamadaASpLibro()

    With rs_SpLibro
    
    If .State = 1 Then .Close
    
    .Open "execute sp_Ordenar", cn, adOpenStatic, adLockOptimistic
    
    End With
End Sub



