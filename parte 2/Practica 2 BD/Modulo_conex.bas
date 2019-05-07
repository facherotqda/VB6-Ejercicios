Attribute VB_Name = "Modulo_conex"
Option Explicit
Dim estado As Boolean

Global cn As New ADODb.Connection
Global rs As New ADODb.Recordset
Global cmd As New ADODb.Command
Global rs_ej As New ADODb.Recordset
Global rs_spAgregar As New ADODb.Recordset
Global rs_spEliminar As New ADODb.Recordset
Global vCodigoCliente As String


'coneccion a la base de datos
Sub main()
Set cn = New ADODb.Connection
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

abrirTablaClientes
Planilla_ABM.Show
End If
End Sub

'conectores a tablas independientes

Sub abrirTablaClientes()

    With rs
    
    If .State = 1 Then .Close
        .Open "SELECT [CÓDIGO CLIENTE],empresa,dirección,población,teléfono,responsable FROM CLIENTES", cn, adOpenStatic, adLockOptimistic
        'abrir con SP
        '.Open "execute sp_MOSTRAR", cn, adOpenStatic, adLockOptimistic
        
        ' .Open "execute sp_BuscarID 'ct02'", cn, adOpenStatic, adLockOptimistic
    End With
End Sub







