VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmLateList 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8340
   ClientLeft      =   360
   ClientTop       =   1695
   ClientWidth     =   11370
   Icon            =   "frmLateList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11370
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid dbgLate 
      Bindings        =   "frmLateList.frx":0442
      Height          =   7515
      Left            =   120
      OleObjectBlob   =   "frmLateList.frx":0458
      TabIndex        =   0
      Top             =   660
      Width           =   11115
   End
   Begin VB.Data datLate 
      Caption         =   "LateList"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   225
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   150
      Visible         =   0   'False
      Width           =   2715
   End
End
Attribute VB_Name = "frmLateList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dCurrentDate As Date

Private Sub Form_Load()
frmLateList.Top = 0
frmLateList.Left = 0

Dim sSQL As String
Dim sDate As String

dCurrentDate = Date

sDate = dCurrentDate

frmLateList.Caption = "Late List For " & sDate & ""

sSQL = "SELECT Rentals.Account, Rentals.CustName, Rentals.Phone, Rentals.Barcode, Rentals.Title, Rentals.Rented, Rentals.DueBack FROM Rentals WHERE '" & sDate & "' > Rentals.DueBack"

datLate.DatabaseName = (App.Path & "\MovieDB.mdb")
datLate.RecordSource = sSQL
End Sub

