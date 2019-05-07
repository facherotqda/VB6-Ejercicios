VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Text            =   "c:\cosa.txt"
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cargar ListBox"
      Height          =   615
      Left            =   7920
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim variable As String
variable = "2545 , 2345"
'se crea el archivo con la direccion
Open "C:\cosa.txt" For Append As #1
Print #1, variable
Close #1
End Sub

Private Sub Command2_Click()
Dim miVariable As String

Open "c:\cosa.txt" For Input As #1

While Not EOF(1)
Line Input #1, miVariable
Wend
Close #1

MsgBox miVariable
End Sub

Private Sub Command3_Click()
On Error GoTo errSub

Dim n_File As Integer
Dim Linea As String

List1.Clear
'Número de archivo libre
n_File = FreeFile
    
'Abre el archivo para leer los datos
Open Text1.Text For Input As n_File

   'Recorre linea a linea el mismo y añade las lineas al control List
    Do While Not EOF(n_File)
        'Lee la linea
        Line Input #n_File, Linea
        List1.AddItem Linea
    Loop
    
Exit Sub
errSub:
'error
MsgBox "Número de error: " & Err.Number & vbNewLine & _
       "Descripción del error: " & Err.Description, vbCritical
End Sub

