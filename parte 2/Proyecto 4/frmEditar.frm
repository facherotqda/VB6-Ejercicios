VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditar 
   Caption         =   "Editar Usuario"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker FechaPicker 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   51314689
      CurrentDate     =   42961
   End
   Begin VB.TextBox txtApellido 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de Nacimiento"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAceptar_Click()
'validamos los textboxs nombre y apellido
If txtNombre.Text = "" Then MsgBox "Error Ingrese un Nombre", vbInformation, "Aviso": txtNombre.SetFocus: Exit Sub
If txtApellido.Text = "" Then MsgBox "Error Ingrese un Apellido", vbInformation, "Aviso": txtApellido.SetFocus: Exit Sub
'validar si se agrega una fecha de nacimiento

If FechaPicker.CheckBox = True Then
  If Not IsNull(FechaPicker.Value) Then
    MsgBox " Está con el CheckBox marcado", vbInformation 'con el ckeck
    Else: MsgBox " Está con el CheckBox Desmarcado ", vbInformation 'sin el ckeck
    'falta poner sentencia para que retome el foco en la fecha de nacimiento
  End If
End If


'Editar El CLIENTE
With rsTablaClientes
.Requery
 .Find "idCliente=' " & Val(vCodigoCliente) & "'"
'.Find "Nombre='" & Trim(txtNombre.Text) & "'"
  
     !Nombre = txtNombre.Text
     !Apellido = txtApellido.Text
     
     
     'deberia ir la fecha de nacimiento
      '!FechaDeNacimiento = FechaPicker.Value
     !FechaNacimiento = FechaPicker.Value
     
.UpdateBatch '*Actualiza todos los registros modificados en el conjunto de registros
             ' desconectado.
             
             '*Update actualizar su registro actual.


    'se debe ACTUALIZAR la grilla y darle ESTILO
    Set frmClientes.GrillaUsuariosGrid.DataSource = rsTablaClientes
    frmClientes.EstiloGrilla
    MsgBox "USUARIO EDITADO", vbInformation, "Satisfactoriamente GG "
    vCodigoCliente = 0 'solo sirve si de elimina o se edita un registro
    Unload Me
 
     
End With ' fin del update


End Sub

Private Sub btnCancelar_Click()
vCodigoCliente = 0
Unload Me
End Sub

Private Sub Form_Load()
CargarCliente

End Sub

Sub CargarCliente()

With rsTablaClientes

    .Find "idCliente=' " & Val(vCodigoCliente) & "'"
    'llenamos los campos con los datos del vcodigoCliente (ya tomado previamente)
    txtNombre = !Nombre
    txtApellido = !Apellido
    FechaPicker.Value = !FechaNacimiento
End With
End Sub




