VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MenuPrincipal 
   Caption         =   "Menu Principal"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   16905
   ScaleHeight     =   9060
   ScaleWidth      =   16905
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListView3 
      Height          =   2895
      Left            =   240
      TabIndex        =   16
      Top             =   600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1320
      TabIndex        =   13
      Top             =   120
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   7935
      Left            =   8640
      TabIndex        =   11
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   13996
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Personales:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   8055
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   4800
         TabIndex        =   10
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Caption         =   "Trabajo a realizar :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   7695
         Begin VB.ListBox List1 
            Height          =   1620
            Left            =   3240
            TabIndex        =   8
            Top             =   360
            Width           =   3855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   600
            TabIndex        =   7
            Top             =   1080
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Celular            o Telefono:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Completo:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
      Caption         =   "FECHA ACTUAL:"
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
      Left            =   12480
      TabIndex        =   15
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Labeldia 
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
      Left            =   14520
      TabIndex        =   14
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Ordenar por:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu cliente 
      Caption         =   "Cliente"
      Index           =   0
   End
   Begin VB.Menu cliente 
      Caption         =   "Gastos E Ingresos"
      Index           =   1
   End
   Begin VB.Menu turnos 
      Caption         =   "Turnos"
      Index           =   3
   End
End
Attribute VB_Name = "MenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()

If Combo1.ListIndex = -1 Then
MsgBox "No se selecciono ningun Trabajo", vbCritical, "ERROR"
Else
List1.AddItem (Combo1.list(Combo1.ListIndex))
End If

End Sub

Private Sub Command2_Click()

If Text1.Text = "" Or Text2.Text = "" Or List1.ListCount = 0 Then

MsgBox "FALTA COMPLENTAR CAMPOS", vbCritical, "FALTAN DATOS"

Else

'se recorre la lista y se muestra lo que tiene

Dim i As Integer
For i = 0 To List1.ListCount - 1
    MsgBox List1.list(i)
Next i

 

'se agregan datos al listview

    Set Module1.listdelMenu = ListView2.ListItems.Add(, , Text1.Text)
               
                listdelMenu.SubItems(1) = (Text2.Text)
               ' recorro la lista para agregar al listview de la columna 3
                Dim cadena As String
                For i = 0 To List1.ListCount - 1
                cadena = List1.list(i) + " " + cadena
                
                 
                Next i
                listdelMenu.SubItems(2) = (cadena)
            


End If


'debo limpiar todos los controles
' HACEEER
Call LimpiarFrameDatos

'creo el fichero .txt y valido si existe
'hacer
 Call VerificarFichero
    If VerificarFichero = True Then
    MsgBox "FICHERO EXISTE"
    Else
    
    Open "c:\asd.txt" For Output As #1
    Close #1
    
    MsgBox " SE CREO FICHERO"
    End If
'debo guardar los datos en un .txt de forma automatica
'HACEEER

End Sub

Private Sub Form_Load()



With ListView2.ColumnHeaders
.Add , , "Nombre Completo", Width / 6.5
.Add , , "Numero de Cel o Tel", Width / 6.5
.Add , , "Trabajos a Realizar", Width / 2.5
End With

With Combo1
.AddItem ("*- Corte ")
.AddItem ("*- Peinado ")
.AddItem ("*- Baño de Creatina ")
.AddItem ("*- Baño de Creama ")
.AddItem ("*- Alisado ")
.AddItem ("*- Color ")
.AddItem ("*- Reflejos ")
.AddItem ("*- Brushing ")
.AddItem ("*- Botox ")
.AddItem ("*- Desgastado de Puntas ")
End With

With ListView3.ColumnHeaders
.Add , , "Nombre Completo", Width / 1.5
.Add , , "Numero de CLIENTE", Width / 1.5
End With


Labeldia.Caption = Date

End Sub

Private Function VerificarFichero() As Boolean
On Error Resume Next

 Open "c:\asd.txt" For Input As #1
  Close #1

If Err.Number <> 0 Then

End If
VerificarFichero = False
End Function

Private Sub LimpiarFrameDatos()

Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
List1.Clear
End Sub

Private Sub ListView2_DblClick()

Form2.Show
End Sub


