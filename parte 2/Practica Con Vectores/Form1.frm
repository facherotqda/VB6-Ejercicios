VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   3
      Left            =   3840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Integer
For i = 0 To 3
Text1(i).Text = Empty
Next i
End Sub

Private Sub Form_Load()

Dim matrizGG(10) As Integer
Dim i As Integer

For i = 0 To 9

'MsgBox "Hola " & i
matrizGG(i) = i + 1
Next i

For i = 0 To 9
MsgBox matrizGG(i)
Next i


End Sub

