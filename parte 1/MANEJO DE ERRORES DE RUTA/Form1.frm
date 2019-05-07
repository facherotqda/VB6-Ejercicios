VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Error3"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Error2"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Error1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer

'-------------------------------------------
'Label1 = ""
'On Error GoTo 88    '[Si ponemos On Error Resume Next sí se mostraría el texto]
'i = Rnd * 10 ^ 16    '[Esta línea genera el error, i demasiado grande]
'Label1 = Label1 & "Esta instrucción no se llega a ejecutar"
'88 If Err Then MsgBox ("Se ha producido un error. Tipo de error = " & Err & " Descripción: " & Err.Description)
'Label1 = Label1 & "La ejecución continúa"
'Label2 = Err.Description

'-------------------------------------------
'On Error GoTo Gestionaerror
'Dim i As Integer
'i = Rnd * 10 ^ 16    '[Esta línea genera el error]
'Label1 = "La ejecución continúa aquí debido al Resume Next" & vbCrLf
'Label1 = Label1 & i    '[Devuelve cero ya que fue imposible asignarle valor tipo integer]
'Gestionaerror:
'If Err.Number <> 0 Then
'    GestiónError
'    Resume Next
'End If
'-----------------------------------------------

'On Error GoTo 4
'i = 1 / 0
'
'4 If Err Then
'MsgBox "ACA SI"
'End If
'i = 1 / 0
'--------------------------------------------------

'On Error GoTo GestionaGG
'i = 1 / 0
'MsgBox "hola"
'GestionaGG:
'If Err.Number <> 0 Then
'    GestiónError
'    Resume Next
'End If


On Error Resume Next

i = 1 / 0
i = "asd"
MsgBox "gg"
Label1.Caption = i
End Sub

Sub GestiónError()
MsgBox ("Se ha producido un error. Tipo de error = " & Err.Number & ". Descripción: " & Err.Description)
End Sub
