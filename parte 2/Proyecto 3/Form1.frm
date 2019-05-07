VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Menu"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10920
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Agregar Datos"
      Height          =   3735
      Left            =   9600
      TabIndex        =   4
      Top             =   600
      Width           =   4455
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Edad:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "DNI"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Buscar Item"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5530
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

ListView1.ListItems.Remove (ListView1.SelectedItem.Index)

End Sub

Private Sub Command2_Click()
Dim itmX As ListItem

Set itmX = ListView1.FindItem(Text1.Text, lvwText, , lvwPartial)
    If itmX Is Nothing Then
    
    MsgBox "Item No encontrado", vbCritical
    Else
    
    ListView1.ListItems(itmX.Index).Selected = True
    ListView1.SetFocus
    End If
End Sub

Private Sub Form_Load()

Dim list As ListItem

With ListView1.ColumnHeaders
.Add , , "Dni ", Width / 5.5
.Add , , "Edad ", Width / 5.5
.Add , , "Nombre Completo ", Width / 5.5
End With

Set list = ListView1.ListItems.Add(, , "16105044")
    list.SubItems(1) = ("54")
    list.SubItems(2) = ("Monica Liliana")
    
Set list = ListView1.ListItems.Add(, , "34755008")
    list.SubItems(1) = ("27")
    list.SubItems(2) = ("Martin Arrua")

Set list = ListView1.ListItems.Add(, , "14963565")
    list.SubItems(1) = ("56")
    list.SubItems(2) = ("Raul Arrua")
    
Set list = ListView1.ListItems.Add(, , "1810522")
    list.SubItems(1) = ("30")
    list.SubItems(2) = ("Lucas Daniel Carbone")
    
End Sub



