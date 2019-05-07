VERSION 5.00
Begin VB.Form EditarCliente 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextEmpresa 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TextDireccion 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox TextPoblacion 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox TextTel 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox TextCod_Cliente 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox TextResponsable 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton CommandAceptar 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "EMPRESA"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "DIRECCION"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "POBLACION"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "TELEFONO"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "RESPONSABLE"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "COD CLIENTE"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "EditarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandAceptar_Click()

With rs_spAgregar
       .Open "Execute sp_EditarCliente '" & TextCod_Cliente.Text & "','" & TextEmpresa.Text & "','" & TextDireccion.Text & "','" & TextPoblacion.Text & "','" & TextTel.Text & "','" & TextResponsable.Text & "' ", cn, adOpenStatic, adLockOptimistic
        MsgBox "SE EDITO SATISFACTORIAMENTE UN CLIENTE "
        
End With

abrirTablaClientes


'vuelvo a cargar la planilla
Set Planilla_ABM.DataGrid1.DataSource = rs
 'le vuelvo a dar formato
 Planilla_ABM.FormatoGrilla
 Planilla_ABM.CargarCombo
 
Unload Me

End Sub

Private Sub Form_Load()

CargarCliente

End Sub

Sub CargarCliente()

With rs
       .Find "[CÓDIGO CLIENTE] ='" & vCodigoCliente & " ' "
       'igualamos campos
        
        
        TextEmpresa.Text = !empresa
        TextDireccion.Text = !DIRECCIÓN
        TextPoblacion.Text = !POBLACIÓN
        TextTel.Text = !TELÉFONO
        TextResponsable.Text = !responsable
        TextCod_Cliente.Text = ![CÓDIGO CLIENTE]
        
End With


End Sub
