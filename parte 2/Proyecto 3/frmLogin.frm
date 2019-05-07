VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LoginForm 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loggin App"
   ClientHeight    =   4635
   ClientLeft      =   5370
   ClientTop       =   3330
   ClientWidth     =   9270
   LinkTopic       =   "Login App"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleMode       =   0  'User
   ScaleWidth      =   9270
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4080
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      DataField       =   "*"
      DataMember      =   "*"
      DataSource      =   "*"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Long
Dim i As Long

Private Sub VerificacionUser()

If Text1.Text = "Tobogan" And Text2.Text = "todasputas1" Then
MsgBox "Usuario Confirmado"

' muestro porcentaje en el label, cargando datos de minimo y maximo
'For x = ProgressBar1.Min To ProgressBar1.Max
'Label4.Caption = x
'
'DoEvents
'ProgressBar1.Value = x
'Next x
ProgressBar1.Visible = True
For i = 0 To ProgressBar1.Max
ProgressBar1.Value = i

Label4 = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & "%"
DoEvents
Next

MenuPrincipal.Show
Unload Me
Else
MsgBox "Error en ingreso de datos"

End If
End Sub

Private Sub Command1_Click(Index As Integer)
Call VerificacionUser
End Sub

Private Sub Command2_Click(Index As Integer)
Dim mensaje
mensaje = (MsgBox("¿Desea salir del programa?", vbQuestion + vbYesNo, "Salir de la APP"))
If mensaje = vbYes Then
Unload Me
End If

End Sub

Private Sub Form_Load()

'With ProgressBar1
'
'.Max = 5000
'.Min = 0
'.Value = 0
'End With

ProgressBar1.Max = 10000
ProgressBar1.Visible = False



End Sub
