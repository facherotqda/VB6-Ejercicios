VERSION 5.00
Begin VB.Form NuevoCLiente 
   Caption         =   "NUEVO CLIENTE"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1815
      TabIndex        =   3
      Top             =   2085
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1815
      TabIndex        =   4
      Top             =   2685
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   885
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1815
      TabIndex        =   0
      Top             =   285
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1815
      TabIndex        =   5
      Top             =   3285
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1815
      TabIndex        =   2
      Top             =   1530
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "COD CLIENTE"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "RESPONSABLE"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "TELEFONO"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "POBLACION"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "DIRECCION"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "EMPRESA"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "NuevoCLiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()


Dim id As String

'variable del textbox a una variable string (id)
Dim empresa As String
Dim direccion As String
Dim poblacion As String
Dim telefono As String
Dim responsable As String
Dim cod_Cliente As String


cod_Cliente = Trim(Text5.Text)
empresa = Trim(Text4.Text)
direccion = Trim(Text3.Text)
poblacion = Trim(Text6.Text)
telefono = Trim(Text1.Text)
responsable = Trim(Text2.Text)

'
'On Error GoTo 3
With rs
'vamos a hacer una validacion usando metodos dentro de un recordset para validar si existe un dato
    .Requery
    .Find "[CÓDIGO CLIENTE] ='" & cod_Cliente & " ' "
   
      If .EOF Then 'sino encontro nada
      'agregamos un cliente con un recordset que llama a un sp
      With rs_spAgregar
       .Open "Execute sp_AgregarCliente '" & cod_Cliente & "','" & empresa & "','" & direccion & "','" & poblacion & "','" & telefono & "','" & responsable & "' ", cn, adOpenStatic, adLockOptimistic
        MsgBox "SE agrego SATISFACTORIAMENTE UN CLIENTE "
      End With
    Else
        MsgBox "NO SE PUDO AGREGAR CLIENTE "
    End If
    
 End With
 Unload Me
'3 If Err Then
'Unload Me
'End If

End Sub

Private Sub Form_Load()
  
End Sub

Private Sub Text1_Change()

End Sub
