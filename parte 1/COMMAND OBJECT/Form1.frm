VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_Eliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton btn_Actualizar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Text            =   "text_email"
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Text            =   "text_tel"
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Text            =   "text_nom"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Text            =   "text_cod"
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton btn_Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "e-mail"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Telefono"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim CM As ADODB.Command



Private Sub btn_Actualizar_Click()

Set CM = New ADODB.Command

With CM
    .ActiveConnection = cn
    .CommandType = adCmdText
    .CommandText = "UPDATE Hospital SET   Nombre=?, Direccion=?, Telefono=? WHERE Hospital_Cod = ? "
    .Prepared = True
End With

CM.Parameters.Append CM.CreateParameter("Nombre", adVarChar, adParamInput, 20, Text2.Text)
CM.Parameters.Append CM.CreateParameter("Direccion", adVarChar, adParamInput, 20, Text3.Text)
CM.Parameters.Append CM.CreateParameter("Telefono", adVarChar, adParamInput, 20, Text4.Text)
CM.Parameters.Append CM.CreateParameter("Codigo", adInteger, adParamInput, 10, Text1.Text)

CM.Execute

Set CM = Nothing

End Sub

Private Sub btn_Cancelar_Click()
Form_Unload (1)
End Sub

Private Sub btn_Eliminar_Click()

Set CM = New ADODB.Command

With CM
.ActiveConnection = cn
.CommandType = adCmdText
.CommandText = "DELETE Hospital where Hospital_Cod = ?"
.Prepared = True

End With


CM.Parameters.Append CM.CreateParameter("Codigo", adInteger, adParamInput, 10, Text1.Text)

CM.Execute

Set CM = Nothing

End Sub

Private Sub btn_Guardar_Click()
Set CM = New ADODB.Command
        
    With CM
    
    .ActiveConnection = cn
    .CommandType = adCmdText
    .CommandText = "INSERT INTO Hospital (Hospital_Cod, Nombre, Direccion, Telefono) VALUES (?,?,?,?)"
     '.CommandText = "INSERT INTO Hospital (Hospital_Cod, Nombre, Direccion, Telefono)" &_
     '"VALUES (?,?,?,?)"  &_ es para concatenar la linea de abajo
    .Prepared = True
    
    End With
    
   
    CM.Parameters.Append CM.CreateParameter("Codigo", adInteger, adParamInput, 2, Text1.Text)
    CM.Parameters.Append CM.CreateParameter("Nombre", adVarChar, adParamInput, 30, Text2.Text)
    CM.Parameters.Append CM.CreateParameter("Apellidos", adVarChar, adParamInput, 40, Text3.Text)
    CM.Parameters.Append CM.CreateParameter("Telefono", adVarChar, adParamInput, 12, Text4.Text)
    
    
    'Ejecutamos el comando
    CM.Execute
    Set CM = Nothing
End Sub

Private Sub Form_Load()

Set cn = New ADODB.Connection
    With cn
            .CursorLocation = adUseClient
            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=KINGFAT-PC"
            .Open
             If .State = 1 Then
             MsgBox "CONECTADO"
             Else
             MsgBox "NO SE PUDO CONECTAR A LA BD"
             End If
    End With


End Sub



Private Sub Form_Unload(Cancel As Integer)

If (MsgBox("Seguro desea Salir", vbCritical + vbYesNo) = vbYes) Then
End
Else
Cancel = 1
End If
End Sub
