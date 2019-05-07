VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar Imagen"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      Format          =   52494337
      CurrentDate     =   43153
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre Image1"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "FechaNac"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Edad"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gdb As Variant
Dim fs As Variant
Private Sub Command1_Click()

CommonDialog1.DialogTitle = "Seleccione una Imagen solo Jpg, Png o Gif ... "
CommonDialog1.ShowSave

RutaOrigen = CommonDialog1.FileName
Image1.Picture = LoadPicture(RutaOrigen)
ArchivoNombre = CommonDialog1.FileTitle

Label4.Caption = ArchivoNombre

End Sub


Private Sub Command2_Click()


'Guardar la imagen seleccionada en una ruta

If ArchivoNombre = "" Then
    Else
        RutaDestino = "C:\Users\KingFat\Desktop\Cargar Imagen a un form\ImagenesClientes\"
        Set gdb = Nothing
        Set fs = CreateObject("Scripting.FileSystemObject")
        fs.copyfile RutaOrigen, RutaDestino
End If



End Sub

Private Sub Form_Load()

RutaOrigen = "C:\Users\KingFat\Desktop\Cargar Imagen a un form\no_disponible.jpg"
Image1.Picture = LoadPicture(RutaOrigen)

End Sub
