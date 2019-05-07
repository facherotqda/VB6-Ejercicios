VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNuevoCliente 
   BackColor       =   &H00404040&
   Caption         =   "Nuevo Cliente"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   5475
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker txtfecha 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51118081
      CurrentDate     =   42958
   End
   Begin VB.TextBox txtApellido 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmNuevoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

