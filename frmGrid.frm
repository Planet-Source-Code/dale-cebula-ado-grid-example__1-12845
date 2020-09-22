VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   450
      Left            =   4410
      TabIndex        =   3
      Top             =   4905
      Width           =   1000
   End
   Begin VB.CommandButton cmdConString 
      Caption         =   "Build Conn String"
      Height          =   450
      Left            =   5580
      TabIndex        =   2
      Top             =   4905
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4290
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   7567
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
            LCID            =   2057
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
            LCID            =   2057
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
   Begin VB.TextBox txtConstring 
      Height          =   300
      Left            =   90
      TabIndex        =   0
      Top             =   4455
      Width           =   6765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************
'   Refereneces set to:
'   ActiveX Data Objects 2.5 lib
'   OLE DB service component 1.0 lib
'   ADO Example Dale Cebula 2000
'
'   Demonstrates how to use ADO to retrieve a recordset and populate
'   the data bound grid. Should work with SQL Server as well as MS Access
'*****************************************************************************
Private ConnctionString$
Private SQL$


'   Sub to populate Grid
Private Sub PopulateGrid()
Dim adoRs As New ADODB.Recordset
Dim adocn As New ADODB.Connection

    adocn.CursorLocation = adUseClient
    adocn.Open ConnctionString$

'   Open the connection
    adoRs.Open SQL$, adocn, adOpenStatic
'   set the grid DS to the active connection
    Set DataGrid1.DataSource = adoRs
    Set adoRs = Nothing: Set adocn = Nothing
End Sub

Private Sub cmdConString_Click()
On Error Resume Next
Dim msDLink As New DataLinks
If txtConstring.Text = "" Then
    MsgBox "You must enter an SQL statement", vbOKOnly: Exit Sub
Else
    
    ConnctionString$ = msDLink.PromptNew
    SQL$ = txtConstring.Text
    Call PopulateGrid
End If


End Sub
