VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public asd As Boolean

' VARIABLE CONSTANTE (NO CAMBIA DE VALOR DURANTE LA EJECUCION DEL PROGRAMA)
 Const pi As Double = 3.1416
 Dim variosNombres As New bri

Private Sub Form_Load()

Combo1.AddItem "ASD"
Combo1.AddItem "sad"


Combo1.ListIndex = 1

Check1.value = 1
Check1.Enabled = True

'llenar list con un item del combo
List1.AddItem Combo1.List(0)

List1.AddItem pi

'ADMITE BOOL AND INT (0 y 1)
'Option1.Value = True
Option1.value = 1

Set variosNombres = New bri

List1.AddItem variosNombres.Nombre

MsgBox variosNombres.Nombre

MsgBox "aca va la funcion :ASD " + das(variosNombres)

variosNombres.EDAD = 10

MsgBox "la edad" & " " & variosNombres.EDAD

'variosNombres.mostrarDatos

End Sub

Function das(hola As bri) As String

Set hola = New bri
das = hola.Nombre


End Function


