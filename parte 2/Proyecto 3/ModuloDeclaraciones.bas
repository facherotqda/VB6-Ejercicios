Attribute VB_Name = "ModuloDeclaraciones"
Option Explicit
'Variables de conexion a Base de Datos
Global bd As New ADODB.Connection 'declaro una variable de tipo ADODB.Connection

'Variables RecordSet
Global RsTablaPersonas As New ADODB.Recordset
