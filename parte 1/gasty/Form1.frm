VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Suma"
      Height          =   735
      Left            =   4320
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Persona 2"
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MOSTRAR mod Cls"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENTER"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   3840
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gasty As String
Dim gasty2 As String
Dim PersonaGG As Persona
Dim Per As Persona




Private Sub Command1_Click()
gasty = "cabro"
Label1.Caption = gasty

End Sub

Private Sub Command2_Click()

Set PersonaGG = New Persona

'PersonaGG.edad = 5
'PersonaGG.nombre = "Gasty PRO"
'PersonaGG.estado = True

PersonaGG.nombres = "Gasty PENE"




'MsgBox PersonaGG.edad & " " & PersonaGG.nombre

MsgBox PersonaGG.nombres

End Sub



Private Sub Command3_Click()

Set Per = CrearPersona("HOLAA")

MsgBox Per.nombres


End Sub

Private Sub Command4_Click()

Dim resultado As Double

resultado = Module1.SumarDato(1, 5)

MsgBox resultado

End Sub

Private Sub Form_Load()
 Form1.Show
 Label2.Caption = MiPromedio(2, 3)
 
 If Label1.Caption Like "Label1" Then
 
 Label1.Caption = "ULTRA GAY"
 gasty2 = "REGAY"
 
 Label1.Caption = CambiarNombre(gasty2)
 End If
 
 End Sub


