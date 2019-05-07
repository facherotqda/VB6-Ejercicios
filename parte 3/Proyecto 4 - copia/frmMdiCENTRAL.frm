VERSION 5.00
Begin VB.MDIForm frmMdiCENTRAL 
   Appearance      =   0  'Flat
   BackColor       =   &H80000012&
   Caption         =   "MDIForm1"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   17760
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu clientees 
      Caption         =   "Clientes"
      Index           =   1
   End
End
Attribute VB_Name = "frmMdiCENTRAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub clientees_Click(Index As Integer)


frmClientes.Show

End Sub



Private Sub MDIForm_Load()
AbrirClientes

End Sub


