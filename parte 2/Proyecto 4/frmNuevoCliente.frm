VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNuevoCliente 
   BackColor       =   &H00404040&
   Caption         =   "Nuevo Cliente"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker FechaPicker 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   107741185
      CurrentDate     =   42958
      MaxDate         =   402133
      MinDate         =   10959
   End
   Begin VB.TextBox txtApellido 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Nacimiento"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmNuevoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Sub ValidarDatosClientes()
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


'validamos si el Cliente se encuentra(si existe)
With rsTablaClientes
.Requery
.Find "Nombre='" & Trim(txtNombre.Text) & "'"
     If .EOF Then 'sino encontro ninguna similitud
     'debemos agregrar un usuario nuevo
     .AddNew
     !Nombre = txtNombre.Text
     !Apellido = txtApellido.Text
     !FechaIngreso = Now 'Date solo fecha
     
     'deberia ir la fecha de nacimiento
      '!FechaDeNacimiento = FechaPicker.Value
     !FechaNacimiento = FechaPicker.Value
     
.Update
    'se debe ACTUALIZAR la grilla y darle ESTILO
    Set frmClientes.GrillaUsuariosGrid.DataSource = rsTablaClientes
    frmClientes.EstiloGrilla
    MsgBox "USUARIO CREADO", vbInformation, "Creado Satisfactoriamente"
    Unload Me
     Else
         'si el usuario existe entonces
     MsgBox "El usuario EXISTE", vbInformation, "INGRESE OTRO USUARIO"
     txtNombre.Text = ""
     txtApellido.Text = ""
     txtNombre.SetFocus

     End If 'finaliza la busqueda del usuario
     End With

End Sub

Private Sub btnAceptar_Click()
ValidarDatosClientes

End Sub

Private Sub btnCancelar_Click()
Unload Me
End Sub




Private Sub Form_Load()

End Sub
