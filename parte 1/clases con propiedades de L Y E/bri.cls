VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Nombre As String
Private m_apellido As String
Private m_edad As Integer


Public Property Get Nombre() As String
    Nombre = "MUESTOR PROPIEDAD DEL OBJETO"
End Property

Public Property Let Nombre(ByVal value As String)
    Nombre = value
End Property

Public Property Get EDAD() As Integer
EDAD = m_edad
End Property

Public Property Let EDAD(ByVal value As Integer)
m_edad = value
End Property



Sub mostrarDatos()

MsgBox m_Nombre & m_edad

End Sub
