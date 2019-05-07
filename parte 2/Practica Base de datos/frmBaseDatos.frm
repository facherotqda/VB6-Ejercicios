VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBaseDatos 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5106
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
   Begin VB.TextBox txtLocalidad 
      Height          =   405
      Left            =   1920
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtDni 
      Height          =   405
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtNombre 
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Localidad"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Dni"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmBaseDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
'CREAMOS UNA TABLA
'cmCrearTabla.ActiveConnection = Conec 'variable que guarda la coneccion
'cmCrearTabla.CommandType = adCmdText
'cmCrearTabla.CommandText = "CREATE TABLE PERSONA (NAME varchar(30), NAME1 varchar(30))"
'cmCrearTabla.Execute


'CREAMOS UN STORED PROCEDURE
cmCrearProcedure.ActiveConnection = Conec
cmCrearProcedure.CommandType = adCmdText
cmCrearProcedure.CommandText = "ALTER PROCEDURE sp_ALGO as begin select * from persona end "
cmCrearProcedure.Execute



'EJECUTAMOS EL SP
cmEjecutarProcedure.ActiveConnection = Conec
cmEjecutarProcedure.CommandType = adCmdStoredProc
cmEjecutarProcedure.CommandText = "sp_ALGO"

Set rsParaProcedure = cmEjecutarProcedure.Execute

DataGrid1.DataSource = rsParaProcedure

End Sub


