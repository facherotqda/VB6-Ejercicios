Attribute VB_Name = "ModuloDeclaraciones"
Option Explicit

Global BaseDatos As New ADODB.Connection ' variable de conexion
Global rsTablaClientes As New ADODB.Recordset 'variable recordset para tabla CLIENTES
Global rsTablaUsuarios As New ADODB.Recordset 'variable recordset para tabla USUARIOS

Global vCodigoCliente As Integer 'variable global para guardar el ID cliente
