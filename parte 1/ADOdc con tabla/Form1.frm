VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1200
      Top             =   3000
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
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
      Caption         =   ""
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
   Begin VB.Label Label3 
      Caption         =   "Dato 3"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Dato 2"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Dato 1"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As String

Private Sub Form_Load()

cn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=KINGFAT-PC"


Adodc1.ConnectionString = cn

Adodc1.CursorType = adOpenDynamic
Adodc1.RecordSource = "Hospital"
Adodc1.Refresh


Set Text1.DataSource = Adodc1
Set Text2.DataSource = Adodc1
Set Text3.DataSource = Adodc1

Text1.DataField = "Hospital_Cod"
Text2.DataField = "Nombre"
Text3.DataField = "Telefono"


End Sub

