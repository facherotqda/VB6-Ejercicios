Attribute VB_Name = "ModuloSentencias"
Option Explicit

Public listdelMenu As ListItem

Sub main()
'conectar a la base de datos

With bd
.CursorLocation = adUseClient
.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=OFFFF;Data Source=KINGFAT-PC"
 MsgBox "Estas conectado a sqlServer"
 LoginForm.Show
End With

End Sub

'Conectores a tablas independientes
Sub AbrirTablaPersonas()

With RsTablaPersonas

If .State = 1 Then .Close 'Si esta abierta la tabla, que se cierre
.Open "Select * from Personas", bd, adOpenStatic, adLockOptimistic 'llama a la tabla Personas

End With
End Sub




