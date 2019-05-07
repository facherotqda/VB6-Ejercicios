VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton commIngresar 
      Caption         =   "Ingresar"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtClave 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave:"
      BeginProperty Font 
         Name            =   "Crackvetica"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Crackvetica"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Login App"
      BeginProperty Font 
         Name            =   "Crackvetica"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   6240
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9360
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub commIngresar_Click()




'Buscamos el usuario

With rsTablaUsuarios
.Requery 'actualizamos la tabla
.Find "Usuario='" & Trim(txtUsuario.Text) & "'"

  If .EOF Then

    MsgBox "No se encontro el usuario", vbInformation, "Aviso"
    txtUsuario.Text = ""
    txtUsuario.SetFocus
    Exit Sub
    
    Else
    'si encontro el usuario
      If !Clave = Trim(txtClave.Text) Then
      'si es correcto
      frmMdiCENTRAL.Show
      'cierro FormLogin
       Unload Me
      
    'si la contraseña es incorrecta
      Else
      MsgBox "La CLAVE ES ERRONEA", vbInformation, "EROR"
      txtClave.Text = ""
      txtClave.SetFocus
      Exit Sub
      
    End If
 End If
 .Close 'cerramos rsTablaUsuarios
 
End With 'rsTablaUsuarios



End Sub

Private Sub Form_Load()
AbrirTablaUsuarios

End Sub



