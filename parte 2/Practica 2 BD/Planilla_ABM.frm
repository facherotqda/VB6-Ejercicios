VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Planilla_ABM 
   Caption         =   "Planilla ABM"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17805
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   17805
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   735
      Left            =   2280
      TabIndex        =   8
      Top             =   8640
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   1296
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Planilla_ABM.frx":0000
      Left            =   5400
      List            =   "Planilla_ABM.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   7560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   7560
      Width           =   4095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7215
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   14640
      _ExtentX        =   25823
      _ExtentY        =   12726
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MV Boli"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LISTA DE CLIENTES"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Eliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Modificar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Agregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUSCAR"
      Height          =   495
      Left            =   13200
      TabIndex        =   4
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "BUSCAR POR:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   7560
      Width           =   1935
   End
End
Attribute VB_Name = "Planilla_ABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim estado As Boolean

Private Sub Agregar_Click()
NuevoCLiente.Show , Me
End Sub

Private Sub Command1_Click()

Dim id As String

'variable del textbox a una variable string (id)
id = Trim(Text1.Text)
  With rs_ej
    
    If .State = 1 Then .Close
       
          
          'DE ESTA MANERA ES LA IDEAL - funca
          .Open "execute sp_BuscarID '" & id & "' ", cn, adOpenStatic, adLockOptimistic
          
          '"execute sp_buscarid '" & id & "'"
          'si funca
         '.Open "execute sp_BuscarID 'ct01' ", cn, adOpenStatic, adLockOptimistic
    End With

Set DataGrid2.DataSource = rs_ej

End Sub


Public Sub RefrescarGrilla()
Set DataGrid1.DataSource = rs
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Eliminar_Click()

vCodigoCliente = DataGrid1.Columns(0).Text
MsgBox vCodigoCliente
ELIMINARCliente.Show vbModal
'vCodigoCliente = DataGrid1.Columns(0).Text
'MsgBox vCodigoCliente
'
'If MsgBox("se eliminara el Usuario " & vCodigoCliente, vbInformation + vbYesNo, "AVISO") = vbYes Then
'
'With rs_spEliminar
'
'
'        .Requery
'        .Find "[CÓDIGO CLIENTE] =' " & vCodigoCliente & " ' "
'        If .EOF Then 'sino encuentra al cliente
'
'
'        .Open " Execute sp_ElimarCliente '" & vCodigoCliente & "' ", cn, adOpenStatic, adLockOptimistic
'        MsgBox "Se ELIMINO CLIENTE CORRECTAMENTE"
'
'
'        Else
'
'        MsgBox "NO EXISTE EL CLIENTE"
'        End If
'
'End With
'
'
'End If

End Sub

Private Sub Form_Load()

'Set DataGrid1.DataSource = rs //lo pusismos en RefrescarGrilla
RefrescarGrilla
FormatoGrilla
CargarCombo

End Sub

Sub FormatoGrilla()

'Tamaños
DataGrid1.Columns(0).Width = Width / 18
DataGrid1.Columns(1).Width = Width / 5
DataGrid1.Columns(2).Width = Width / 6
DataGrid1.Columns(3).Width = Width / 10
DataGrid1.Columns(4).Width = Width / 9
DataGrid1.Columns(5).Width = Width / 6

'Nombre de Columnas
DataGrid1.Columns(0).Caption = "ID"
DataGrid1.Columns(1).Caption = "EMPRESA"
DataGrid1.Columns(2).Caption = "DIRECCION"
DataGrid1.Columns(3).Caption = "POBLACION"
DataGrid1.Columns(4).Caption = "TELEFONO"
DataGrid1.Columns(5).Caption = "RESPONSABLE"

End Sub

Sub CargarCombo()
Combo1.AddItem "ID"
Combo1.AddItem "EMPRESA"
Combo1.AddItem "DIRECCION"
Combo1.AddItem "POBLACION"
Combo1.AddItem "TELEFONO"
Combo1.AddItem "RESPONSABLE"

Combo1.ListIndex = 0

End Sub



Private Sub Modificar_Click()

vCodigoCliente = DataGrid1.Columns(0).Text
MsgBox vCodigoCliente
EditarCliente.Show vbModal
End Sub
