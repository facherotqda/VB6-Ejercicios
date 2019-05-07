VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   Caption         =   "MDIForm1"
   ClientHeight    =   10125
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   17760
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu forms 
      Caption         =   "Formularios"
      Begin VB.Menu formulario1 
         Caption         =   "Formulario 1"
      End
      Begin VB.Menu formulario2 
         Caption         =   "Formulario 2"
      End
      Begin VB.Menu formulario3 
         Caption         =   "Formulario 3"
      End
      Begin VB.Menu formulario4 
         Caption         =   "Formulario 4"
      End
   End
   Begin VB.Menu carga 
      Caption         =   "Carga"
      Begin VB.Menu Alta 
         Caption         =   "Alta"
      End
      Begin VB.Menu baja 
         Caption         =   "Baja"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Alta_Click()

Form3.Show

End Sub





Private Sub formulario1_Click()
Form1.Show

End Sub

Private Sub formulario4_Click()
Form4.Show

End Sub



