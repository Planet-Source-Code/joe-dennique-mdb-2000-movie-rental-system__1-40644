VERSION 5.00
Begin VB.Form frmCheckIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movie Return System"
   ClientHeight    =   5655
   ClientLeft      =   1815
   ClientTop       =   1560
   ClientWidth     =   8445
   Icon            =   "frmCheckIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8445
   Begin VB.TextBox txtRetType 
      DataField       =   "Type"
      DataSource      =   "datReturns"
      Height          =   315
      Left            =   4140
      TabIndex        =   39
      Top             =   7560
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.TextBox txtType 
      DataField       =   "Type"
      DataSource      =   "datRentals"
      Height          =   315
      Left            =   4140
      TabIndex        =   38
      Top             =   7140
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration..."
      Height          =   1635
      Left            =   6420
      TabIndex        =   33
      Top             =   5820
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox txtOther 
         DataField       =   "SpecLate"
         DataSource      =   "datConfig"
         Height          =   315
         Left            =   180
         TabIndex        =   37
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox txtKids 
         DataField       =   "KidsLate"
         DataSource      =   "datConfig"
         Height          =   315
         Left            =   180
         TabIndex        =   36
         Top             =   900
         Width           =   1515
      End
      Begin VB.TextBox txtOld 
         DataField       =   "OldLate"
         DataSource      =   "datConfig"
         Height          =   315
         Left            =   180
         TabIndex        =   35
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox txtNew 
         DataField       =   "NewLate"
         DataSource      =   "datConfig"
         Height          =   315
         Left            =   180
         TabIndex        =   34
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Data datConfig 
      Caption         =   "Config"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6300
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Frame fraLate 
      Caption         =   "Late Charges..."
      Height          =   975
      Left            =   2160
      TabIndex        =   31
      Top             =   6720
      Visible         =   0   'False
      Width           =   1755
      Begin VB.TextBox txtLateAmount 
         DataField       =   "AmountDue"
         DataSource      =   "datLate"
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox txtLateAccount 
         DataField       =   "Account"
         DataSource      =   "datLate"
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Data datLate 
      Caption         =   "Late Charges"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5820
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.TextBox txtReturned 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4140
      TabIndex        =   30
      Top             =   6720
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Frame fraReturns 
      Caption         =   "Returns..."
      Height          =   2265
      Left            =   165
      TabIndex        =   23
      Top             =   5670
      Visible         =   0   'False
      Width           =   1740
      Begin VB.TextBox txtRetReturned 
         DataField       =   "Returned"
         DataSource      =   "datReturns"
         Height          =   315
         Left            =   150
         TabIndex        =   29
         Top             =   1800
         Width           =   1440
      End
      Begin VB.TextBox txtRetDue 
         DataField       =   "DueBack"
         DataSource      =   "datReturns"
         Height          =   315
         Left            =   150
         TabIndex        =   28
         Top             =   1500
         Width           =   1440
      End
      Begin VB.TextBox txtRetRent 
         DataField       =   "Rented"
         DataSource      =   "datReturns"
         Height          =   315
         Left            =   150
         TabIndex        =   27
         Top             =   1200
         Width           =   1440
      End
      Begin VB.TextBox txtRetTitle 
         DataField       =   "Title"
         DataSource      =   "datReturns"
         Height          =   315
         Left            =   150
         TabIndex        =   26
         Top             =   900
         Width           =   1440
      End
      Begin VB.TextBox txtRetBarcode 
         DataField       =   "Barcode"
         DataSource      =   "datReturns"
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   600
         Width           =   1440
      End
      Begin VB.TextBox txtRetAccount 
         DataField       =   "Account"
         DataSource      =   "datReturns"
         Height          =   315
         Left            =   150
         TabIndex        =   24
         Top             =   300
         Width           =   1440
      End
   End
   Begin VB.Data datReturns 
      Caption         =   "Returns"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2115
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5820
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Frame fraNoMovies 
      Caption         =   "No Movies Rented Out..."
      Height          =   1215
      Left            =   60
      TabIndex        =   21
      Top             =   1380
      Visible         =   0   'False
      Width           =   7440
      Begin VB.Label lblNone 
         Alignment       =   2  'Center
         Caption         =   "No Movies Rented Out"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   6915
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   5400
      TabIndex        =   12
      Top             =   5040
      Width           =   915
   End
   Begin VB.Frame fraNav 
      Caption         =   "Record Navigation..."
      Height          =   765
      Left            =   1020
      TabIndex        =   18
      Top             =   3900
      Width           =   5415
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return"
         Height          =   315
         Left            =   4350
         TabIndex        =   20
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   315
         Left            =   3075
         TabIndex        =   9
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   315
         Left            =   2100
         TabIndex        =   10
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev."
         Height          =   315
         Left            =   1125
         TabIndex        =   11
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Find Rented Movie..."
      Height          =   765
      Left            =   1020
      TabIndex        =   17
      Top             =   4800
      Width           =   4140
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   315
         Left            =   3075
         TabIndex        =   2
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   2790
      End
   End
   Begin VB.Frame fraRented 
      Caption         =   "Rented Movies..."
      Height          =   3585
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   7290
      Begin VB.TextBox txtPhone 
         DataField       =   "Phone"
         DataSource      =   "datRentals"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2580
         TabIndex        =   44
         Top             =   1260
         Width           =   4515
      End
      Begin VB.TextBox txtName 
         DataField       =   "CustName"
         DataSource      =   "datRentals"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2580
         TabIndex        =   43
         Top             =   840
         Width           =   4515
      End
      Begin VB.TextBox txtDueBack 
         DataField       =   "DueBack"
         DataSource      =   "datRentals"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2550
         TabIndex        =   7
         Top             =   3075
         Width           =   4515
      End
      Begin VB.TextBox txtRented 
         DataField       =   "Rented"
         DataSource      =   "datRentals"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2550
         TabIndex        =   6
         Top             =   2625
         Width           =   4515
      End
      Begin VB.TextBox txtTitle 
         DataField       =   "Title"
         DataSource      =   "datRentals"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2550
         TabIndex        =   5
         Top             =   2175
         Width           =   4515
      End
      Begin VB.TextBox txtBarcode 
         DataField       =   "Barcode"
         DataSource      =   "datRentals"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2550
         TabIndex        =   4
         Top             =   1725
         Width           =   4515
      End
      Begin VB.TextBox txtAccount 
         DataField       =   "Account"
         DataSource      =   "datRentals"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2550
         TabIndex        =   3
         Top             =   375
         Width           =   4515
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   1260
         Width           =   2265
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   2265
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   19
         Top             =   2175
         Width           =   2265
      End
      Begin VB.Label lblDueBack 
         Caption         =   "Due Back:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   16
         Top             =   3075
         Width           =   2265
      End
      Begin VB.Label lblRented 
         Caption         =   "Rented:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   15
         Top             =   2625
         Width           =   2265
      End
      Begin VB.Label lblBarcode 
         Caption         =   "Barcode :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   1725
         Width           =   2265
      End
      Begin VB.Label lblAccount 
         Caption         =   "Account Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   375
         Width           =   2265
      End
   End
   Begin VB.Data datRentals 
      Caption         =   "Rentals"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2115
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6270
      Visible         =   0   'False
      Width           =   1890
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
Dim sBarcode As String
Dim sFind As String

If txtSearch.Text = "" Then
MsgBox "Barcode Needed to run search query", vbExclamation, "Barcode Needed"
txtSearch.SetFocus
Exit Sub
End If

sBarcode = txtSearch.Text
sFind = "Barcode LIKE '" & sBarcode & "*'"

datRentals.Recordset.FindFirst sFind

If datRentals.Recordset.NoMatch Then
MsgBox "Barcode Not Found", vbExclamation, "No Match"
txtSearch.Text = ""
txtSearch.SetFocus
End If

ReturnMovie
MsgBox "Movie has been returned", vbExclamation, "Movie Returned"
End Sub

Private Sub cmdFirst_Click()
datRentals.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
datRentals.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
datRentals.Recordset.MoveNext
If datRentals.Recordset.EOF Then
MsgBox "You are viewing the last record", vbExclamation, "Last Record"
datRentals.Recordset.MovePrevious
End If
End Sub

Private Sub cmdPrev_Click()
datRentals.Recordset.MovePrevious
If datRentals.Recordset.BOF Then
MsgBox "You are viewing the first record", vbExclamation, "First Record"
datRentals.Recordset.MoveNext
End If
End Sub

Private Sub cmdReturn_Click()
ReturnMovie
End Sub

Private Sub Form_Load()
frmCheckIn.Top = 0
frmCheckIn.Left = 0

datRentals.DatabaseName = (App.Path & "\MovieDB.mdb")
datReturns.DatabaseName = (App.Path & "\MovieDB.mdb")
datReturns.RecordSource = ("Returns")
datLate.DatabaseName = (App.Path & "\MovieDB.mdb")
datLate.RecordSource = ("Latecharge")
datConfig.DatabaseName = (App.Path & "\MovieDB.mdb")
datConfig.RecordSource = ("config")

CheckIt
CountMe
End Sub

Private Sub CountMe()
Dim i As Integer

datRentals.Refresh
i = datRentals.Recordset.RecordCount

If i = 0 Then
HideNav
Exit Sub
End If

ShowNav
End Sub

Private Sub HideNav()
cmdFirst.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdReturn.Enabled = False
cmdFind.Enabled = False
fraNoMovies.Visible = True
End Sub

Private Sub ShowNav()
cmdFirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdReturn.Enabled = True
cmdFind.Enabled = True
fraNoMovies.Visible = False
End Sub

Private Sub ReturnMovie()
Dim sDate As Date

sDate = Date
datRentals.Recordset.Edit
txtReturned.Text = sDate

CheckLate
Returned
DeleteMe

CheckIt
End Sub

Private Sub CheckIt()
Dim sSQL As String

sSQL = "SELECT * FROM Rentals"

datRentals.RecordSource = (sSQL)
datRentals.Refresh
CountMe
End Sub

Private Sub Returned()
datReturns.Recordset.AddNew

txtRetAccount.Text = txtAccount.Text
txtRetBarcode.Text = txtBarcode.Text
txtRetRent.Text = txtRented.Text
txtRetTitle.Text = txtTitle.Text
txtRetType.Text = txtType.Text
txtRetDue.Text = txtDueBack.Text
txtRetReturned.Text = txtReturned.Text

datReturns.Recordset.Update
datReturns.Refresh
End Sub

Private Sub DeleteMe()
datRentals.Recordset.Delete
datRentals.Refresh
End Sub

Private Sub CheckLate()
Dim sAcct As String
Dim sLocate As String
Dim dStart As Date
Dim dEnd As Date
Dim c As Currency
Dim i As Integer
Dim t As Currency
Dim p As Currency
Dim a As Currency

If txtType.Text = "New" Then
c = txtNew.Text
ElseIf txtType.Text = "Old" Then
c = txtOld.Text
ElseIf txtType.Text = "Kids" Then
c = txtKids.Text
ElseIf txtType.Text = "Other" Then
c = txtOther.Text
End If

If txtReturned.Text < txtDueBack.Text Then
Exit Sub
End If

If txtReturned.Text = txtDueBack.Text Then
Exit Sub
End If

If txtReturned.Text > txtDueBack.Text Then

sAcct = txtAccount.Text
sLocate = "Account LIKE '" & sAcct & "'"
dStart = txtDueBack.Text
dEnd = Date

i = DateDiff("d", dStart, dEnd)
t = c * i

datLate.Recordset.FindFirst sLocate

If datLate.Recordset.NoMatch Then
datLate.Recordset.AddNew
txtLateAccount.Text = txtAccount.Text
txtLateAmount.Text = t
datLate.Recordset.Update
datLate.Refresh
Exit Sub
End If


datLate.Recordset.Edit
p = txtLateAmount.Text
a = p + t
txtLateAmount.Text = a
datLate.Recordset.Update
datLate.Refresh
End If
End Sub
