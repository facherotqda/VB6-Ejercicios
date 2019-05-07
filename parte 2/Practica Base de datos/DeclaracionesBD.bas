Attribute VB_Name = "DeclaracionesBD"
Option Explicit

Global Conec As New ADODB.Connection
Global rsCrearTabla As New ADODB.Recordset
Global cmCrearTabla As New ADODB.Command
Global cmCrearProcedure As New ADODB.Command
Global cmEjecutarProcedure As New ADODB.Command

Global rsParaProcedure As New ADODB.Recordset


