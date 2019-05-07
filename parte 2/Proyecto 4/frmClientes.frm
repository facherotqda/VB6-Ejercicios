VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmClientes 
   AutoRedraw      =   -1  'True
   Caption         =   "Clientes"
   ClientHeight    =   10425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10425
   ScaleWidth      =   17760
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoFiltrarClientes 
      Height          =   495
      Left            =   11160
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Salir"
      Height          =   495
      Left            =   10440
      TabIndex        =   10
      Top             =   8400
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ver Usuario"
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton btnEliminar 
      Caption         =   "Eliminar Usurio"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton btnModificar 
      Caption         =   "Modificar Usuario"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton btnRegistrarUsuario 
      Caption         =   "Registrar Usuario"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quitar Filto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox cmbOrden 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10200
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid GrillaUsuariosGrid 
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   12938
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Appearance      =   0
      BackColor       =   -2147483636
      BorderStyle     =   0
      Enabled         =   -1  'True
      ForeColor       =   -2147483639
      HeadLines       =   1
      RowHeight       =   22
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenar por:"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Clientes LoockArt"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Picture         =   "frmClientes.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14280
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub btnEliminar_Click()

'verificar si la tabla esta vacia
With rsTablaClientes
If .RecordCount = 0 Then Exit Sub
End With

'obtengo el codigo de usuario
vCodigoCliente = GrillaUsuariosGrid.Columns(0).Text

'Preguntar si se elimina el usuario seleccionado
If MsgBox("Se Elimina a " & GrillaUsuariosGrid.Columns(1).Text & " " & GrillaUsuariosGrid.Columns(2).Text, vbInformation + vbYesNo, "Aviso") = vbYes Then

With rsTablaClientes
.Requery
.Find "idCliente=' " & Val(vCodigoCliente) & "'"
.Delete
.Requery
EstiloGrilla
vCodigoCliente = 0
End With
End If
End Sub

Private Sub btnModificar_Click()

'verificar si la tabla esta vacia
With rsTablaClientes
If .RecordCount = 0 Then Exit Sub
End With

'obtengo el codigo de usuario
vCodigoCliente = GrillaUsuariosGrid.Columns(0).Text

MsgBox vCodigoCliente

frmEditar.Show vbModal

End Sub

Private Sub btnRegistrarUsuario_Click()
frmNuevoCliente.Show vbModal
End Sub





Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Resize()
'para que tome las dimenciones iguales de imagen1 y form
Image1.Width = Me.ScaleWidth
Image1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Load()
' que la imagen se mueva al tamaño del form
Image1.Move 0, 0, Me.Width, Me.Height

'cargo la grilla con los datos de la tabla Cliente y doy estilo
Set GrillaUsuariosGrid.DataSource = rsTablaClientes
EstiloGrilla

'cargo combobox
cmbOrden.AddItem "ID"
cmbOrden.AddItem "Nombre"
cmbOrden.AddItem "Apellido"
cmbOrden.AddItem "Fecha de Registro"

cmbOrden.ListIndex = 0
End Sub

Sub EstiloGrilla()

'tamaños
GrillaUsuariosGrid.Columns(0).Width = 500
GrillaUsuariosGrid.Columns(1).Width = 2000
GrillaUsuariosGrid.Columns(2).Width = 2000
GrillaUsuariosGrid.Columns(3).Width = 3000
GrillaUsuariosGrid.Columns(4).Width = 3000

'caption'
GrillaUsuariosGrid.Columns(0).Caption = "ID"
GrillaUsuariosGrid.Columns(1).Caption = "Nombre"
GrillaUsuariosGrid.Columns(2).Caption = "Apellido"
GrillaUsuariosGrid.Columns(3).Caption = "Fecha de Registro"
GrillaUsuariosGrid.Columns(4).Caption = "Fecha de Nacimiento"

'alineacion
GrillaUsuariosGrid.Columns(3).Alignment = dbgCenter
GrillaUsuariosGrid.Columns(4).Alignment = dbgCenter
'headFont
GrillaUsuariosGrid.HeadFont.Bold = True

End Sub

