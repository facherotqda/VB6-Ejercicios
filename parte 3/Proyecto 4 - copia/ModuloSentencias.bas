Attribute VB_Name = "ModuloSentencias"
Option Explicit

Sub main()

With BaseDatos
'se Conecta a la BD
.CursorLocation = adUseClient
.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=OFFFF;Data Source=KINGFAT-PC"


If .State = 1 Then
 MsgBox "CONECTADO A LA BD"
 Else
 MsgBox "No esta Conectada"
End If
 Proyecto_bd.frmLogin.Show


End With
End Sub


Sub AbrirTablaUsuarios()

With rsTablaUsuarios
If .State = 1 Then .Close ' si esta abierta la tabla Usuarios que se cierre

.Open "Select * from Usuarios", BaseDatos, adOpenStatic, adLockOptimistic

End With

End Sub

Sub AbrirClientes()

With rsTablaClientes
 If .State = 1 Then .Close
 .Open "Select *from Clientes", BaseDatos, adOpenStatic, adLockOptimistic
 
End With
End Sub

Sub CrearNuevoRegistro() ' Crea un nuevo registro de Usuario



End Sub


