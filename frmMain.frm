VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario mamalon"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExec 
      Caption         =   "Ejecutar"
      Height          =   495
      Left            =   9600
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtSQL 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11175
   End
   Begin MSDataGridLib.DataGrid gridResult 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
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
            LCID            =   2058
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
            LCID            =   2058
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables globales
Dim con As Connection
Dim rs As Recordset

Dim server As String
Dim user As String
Dim password As String
Dim database As String
Dim StatusQuery As Boolean

Private Sub Form_Load()
    Set con = New Connection
    Set rs = New Recordset
    
    server = "192.168.1.70"
    user = "sa"
    password = "123"
    database = "DBTest"
    StatusQuery = False
    
    openMSSQL con, "PROVIDER=SQLOLEDB; DATA SOURCE=" & server & "; UID=" & user & "; PWD=" & password & "; DATABASE=" & database
    If Err.Description <> "" Then
        MsgBox "Error al conectarse, Detalle: " + Err.Description
    End If
End Sub

Private Sub Form_Terminate()
    If StatusQuery Then
        rs.Close
    End If
    closeMSSQL con
End Sub

Private Sub btnExec_Click()
    If StatusQuery Then
        rs.Close
    End If
    
    'configurar recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    
    execMSSQL con, rs, txtSQL.Text
    If Err.Description <> "" Then
        MsgBox "Error al ejecutar SQL, Detalle: " + Err.Description
        Exit Sub
    End If
    
    Set gridResult.DataSource = rs
    StatusQuery = True
End Sub

