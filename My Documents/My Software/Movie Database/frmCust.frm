VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts"
   ClientHeight    =   8445
   ClientLeft      =   495
   ClientTop       =   1665
   ClientWidth     =   11865
   Icon            =   "frmCust.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11865
   Begin VB.Frame fraPhone 
      Caption         =   "Search By Phone Number..."
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   5460
      TabIndex        =   46
      Top             =   3060
      Width           =   6225
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find"
         Height          =   315
         Left            =   5100
         TabIndex        =   49
         Top             =   300
         Width           =   990
      End
      Begin VB.TextBox txtPhoneSearch 
         Height          =   285
         Left            =   150
         TabIndex        =   48
         Top             =   300
         Width           =   3765
      End
      Begin VB.CommandButton cmdFindPhone 
         Caption         =   "Find"
         Height          =   315
         Left            =   4050
         TabIndex        =   47
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Data datLate 
      Caption         =   "Late Charges"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8220
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2700
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Frame fraLate 
      Caption         =   "Pay Late Charges..."
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   6480
      TabIndex        =   44
      Top             =   1680
      Width           =   5115
      Begin VB.CommandButton cmdPay 
         Caption         =   "Pay"
         Default         =   -1  'True
         Height          =   315
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.TextBox txtPay 
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   3675
      End
   End
   Begin VB.Data datReturns 
      Caption         =   "Returns"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2700
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame fraReturned 
      Caption         =   "Rental History..."
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   240
      TabIndex        =   42
      Top             =   5760
      Width           =   11415
      Begin MSDBGrid.DBGrid dbgHistory 
         Bindings        =   "frmCust.frx":058A
         Height          =   1155
         Left            =   120
         OleObjectBlob   =   "frmCust.frx":05A3
         TabIndex        =   43
         Top             =   300
         Width           =   11115
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search By Account Number..."
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   180
      TabIndex        =   41
      Top             =   3060
      Width           =   5145
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   315
         Left            =   4050
         TabIndex        =   2
         Top             =   300
         Width           =   990
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   3765
      End
   End
   Begin VB.Data datRentals 
      Caption         =   "Rentals"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Data datCust 
      Caption         =   "Customer"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   10800
      TabIndex        =   22
      Top             =   7740
      Width           =   990
   End
   Begin VB.Frame fraMaint 
      Caption         =   "Account Maintenance..."
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   4740
      TabIndex        =   39
      Top             =   7500
      Width           =   5340
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   315
         Left            =   4275
         TabIndex        =   21
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   3225
         TabIndex        =   20
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Height          =   315
         Left            =   2175
         TabIndex        =   24
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   1125
         TabIndex        =   23
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   315
         Left            =   75
         TabIndex        =   19
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Frame fraNav 
      Caption         =   "Navigation..."
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   240
      TabIndex        =   38
      Top             =   7500
      Width           =   4290
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   315
         Left            =   3225
         TabIndex        =   16
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   315
         Left            =   2175
         TabIndex        =   17
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev."
         Height          =   315
         Left            =   1125
         TabIndex        =   18
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   315
         Left            =   75
         TabIndex        =   15
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Frame fraStats 
      Caption         =   "Customer Statistics..."
      ForeColor       =   &H00FF0000&
      Height          =   1290
      Left            =   6450
      TabIndex        =   35
      Top             =   225
      Width           =   5190
      Begin MSMask.MaskEdBox mskLate 
         DataField       =   "AmountDue"
         DataSource      =   "datLate"
         Height          =   315
         Left            =   3360
         TabIndex        =   45
         Top             =   900
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRented 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3375
         TabIndex        =   13
         Top             =   300
         Width           =   1665
      End
      Begin VB.TextBox txtCurrent 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3375
         TabIndex        =   14
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label lblLate 
         Caption         =   "Late Charges:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   40
         Top             =   900
         Width           =   3090
      End
      Begin VB.Label lblCurrent 
         Caption         =   "Movies Rented Current:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   37
         Top             =   600
         Width           =   3090
      End
      Begin VB.Label lblMovieCount 
         Caption         =   "Movies Rented Total:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   36
         Top             =   300
         Width           =   2790
      End
   End
   Begin VB.Frame fraRentals 
      Caption         =   "Rental Information By Customer..."
      ForeColor       =   &H00FF0000&
      Height          =   1725
      Left            =   210
      TabIndex        =   25
      Top             =   3930
      Width           =   11490
      Begin MSDBGrid.DBGrid dbgRental 
         Bindings        =   "frmCust.frx":0F79
         Height          =   1275
         Left            =   150
         OleObjectBlob   =   "frmCust.frx":0F92
         TabIndex        =   26
         Top             =   300
         Width           =   11190
      End
   End
   Begin VB.Frame fraCust 
      Caption         =   "Customer Information..."
      ForeColor       =   &H00FF0000&
      Height          =   2790
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   6090
      Begin VB.TextBox txtPhone 
         DataField       =   "Phone"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2475
         TabIndex        =   10
         Top             =   2400
         Width           =   3465
      End
      Begin VB.TextBox txtPost 
         DataField       =   "Post"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2475
         TabIndex        =   9
         Top             =   2100
         Width           =   3465
      End
      Begin VB.TextBox txtProv 
         DataField       =   "Prov"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2475
         TabIndex        =   8
         Top             =   1800
         Width           =   3465
      End
      Begin VB.TextBox txtCity 
         DataField       =   "City"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2475
         TabIndex        =   7
         Top             =   1500
         Width           =   3465
      End
      Begin VB.TextBox txtAddr 
         DataField       =   "Addr"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2475
         TabIndex        =   6
         Top             =   1200
         Width           =   3465
      End
      Begin VB.TextBox txtName 
         DataField       =   "CustName"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2475
         TabIndex        =   5
         Top             =   900
         Width           =   3465
      End
      Begin VB.TextBox txtDriver 
         DataField       =   "DriverLic"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2475
         TabIndex        =   4
         Top             =   600
         Width           =   3465
      End
      Begin VB.TextBox txtAccount 
         DataField       =   "CustAccount"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2475
         TabIndex        =   3
         Top             =   300
         Width           =   3465
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone Number:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   34
         Top             =   2400
         Width           =   2190
      End
      Begin VB.Label lbldriver 
         Caption         =   "Drivers Licence:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   33
         Top             =   600
         Width           =   2190
      End
      Begin VB.Label lblPost 
         Caption         =   "Postal Code:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   32
         Top             =   2100
         Width           =   2190
      End
      Begin VB.Label lblProv 
         Caption         =   "Province:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   31
         Top             =   1800
         Width           =   2190
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   30
         Top             =   1500
         Width           =   2190
      End
      Begin VB.Label lblAddr 
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   1200
         Width           =   2190
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   900
         Width           =   2190
      End
      Begin VB.Label lblAccount 
         Caption         =   "Account Number:"
         BeginProperty Font 
            Name            =   "MS Mincho"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   27
         Top             =   300
         Width           =   2190
      End
   End
End
Attribute VB_Name = "frmCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
datCust.Recordset.AddNew
HideEdit
txtAccount.SetFocus
cmdUpdate.Default = True
End Sub

Private Sub cmdCancel_Click()
    datCust.Recordset.CancelUpdate
    MsgBox "Account update canceled", vbExclamation, "Update Canceled"
    datCust.Refresh
    ShowEdit
    CountMe
    cmdFind.Default = True
    GetRentals
    GetRentals
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDel_Click()
datCust.Recordset.Delete
MsgBox "Customer account has been deleted", vbExclamation, "Account Deleted"
datCust.Refresh
ShowEdit
CountMe
End Sub

Private Sub cmdEdit_Click()
    datCust.Recordset.Edit
    HideEdit
    txtAccount.SetFocus
    cmdUpdate.Default = True
End Sub

Private Sub cmdFind_Click()
Dim sName As String
Dim sFind As String

sName = txtSearch.Text
sFind = "CustAccount LIKE '" & sName & "*'"

If txtSearch.Text = "" Then
MsgBox "Account number needed to complete search", vbExclamation, "Account Number Needed"
txtSearch.SetFocus
datCust.Refresh
End If

datCust.Recordset.FindFirst sFind

If datCust.Recordset.NoMatch Then
MsgBox "No account was found with that number", vbExclamation, "Account Not Found"
txtSearch.Text = ""
txtSearch.SetFocus
datCust.Refresh
End If

txtSearch.Text = ""
txtSearch.SetFocus
datCust.Refresh

GetRentals
End Sub

Private Sub cmdFindNext_Click()
Dim sPhone As String
Dim sFind As String

sPhone = txtPhoneSearch.Text
sFind = "Phone LIKE '" & sPhone & "*'"

If sPhone = "" Then
MsgBox "You need to enter a phone number", vbExclamation, "Phone Number Needed"
txtPhoneSearch.SetFocus
datCust.Refresh
End If

datCust.Recordset.FindNext sFind

If datCust.Recordset.NoMatch Then
MsgBox "There are no more customers with that phone number", vbExclamation, "No more matches"
txtPhoneSearch.Text = ""
txtSearch.SetFocus
datCust.Refresh
End If

datCust.Refresh
GetRentals
End Sub

Private Sub cmdFindPhone_Click()
Dim sPhone As String
Dim sFind As String

sPhone = txtPhoneSearch.Text
sFind = "Phone LIKE '" & sPhone & "*'"

If sPhone = "" Then
MsgBox "You need to enter a phone number", vbExclamation, "Phone Number Needed"
txtPhoneSearch.SetFocus
datCust.Refresh
End If

datCust.Recordset.FindFirst sFind
datCust.Refresh

GetRentals
End Sub

Private Sub cmdFirst_Click()
    datCust.Recordset.MoveFirst
    
GetRentals
End Sub

Private Sub cmdLast_Click()
    datCust.Recordset.MoveLast
    
GetRentals
End Sub

Private Sub cmdNext_Click()
    datCust.Recordset.MoveNext
    
    If datCust.Recordset.EOF Then
    MsgBox "You are viewing the last record", vbInformation, "Last Record"
    datCust.Recordset.MovePrevious
    End If

GetRentals
End Sub

Private Sub cmdPay_Click()
Dim l As Currency
Dim p As Currency
Dim t As Currency

If txtPay.Text = "" Then
MsgBox "you must enter a $ amount", vbExclamation, "Enter Amount"
txtPay.SetFocus
Exit Sub
End If

datLate.Recordset.Edit
l = mskLate.Text
p = txtPay.Text
t = l - p
mskLate.Text = t
datLate.Recordset.Update
datLate.Refresh

txtPay.Text = ""
txtSearch.SetFocus
End Sub

Private Sub cmdPrev_Click()
    datCust.Recordset.MovePrevious
    
    If datCust.Recordset.BOF Then
    MsgBox "You are viewing the first record", vbInformation, "First Record"
    datCust.Recordset.MoveNext
    End If
    
GetRentals
End Sub

Private Sub cmdUpdate_Click()
    If txtAccount.Text = "" Then
    MsgBox "Account number cannot be left blank", vbExclamation, "Need Account Number"
    txtAccount.SetFocus
    Exit Sub
    End If
    
    datCust.Recordset.Update
    datCust.Refresh
    
    MsgBox "Account information updated", vbExclamation, "Account Updated"
    
    CountMe
    
    ShowEdit
    cmdFind.Default = True
    GetRentals
End Sub

Private Sub GetRentals()
Dim sqlRent As String
Dim sqlReturn As String
Dim sqlLate As String
Dim sAccount As String
Dim sDate As Date
Dim t As Integer
Dim r As Integer
Dim c As Integer

sAccount = txtAccount.Text
sDate = Date
On Error Resume Next
sqlRent = "SELECT Rentals.Barcode, Rentals.Title, Rentals.Rented, Rentals.DueBack FROM Rentals WHERE Account LIKE '" & sAccount & "'"

sqlReturn = "SELECT Returns.Barcode, Returns.Title, Returns.Rented, Returns.DueBack, Returns.Returned FROM Returns WHERE Account LIKE '" & sAccount & "'"
            
sqlLate = "SELECT latecharge.AmountDue FROM Latecharge WHERE Account LIKE '" & sAccount & "'"

datRentals.RecordSource = sqlRent
datRentals.Refresh

datReturns.RecordSource = sqlReturn
datReturns.Refresh

datLate.RecordSource = sqlLate
datLate.Refresh

CountRentals
cmdPay.Enabled = True

If mskLate.Text = "" Then
cmdPay.Enabled = False
End If

End Sub


Private Sub Form_Load()
frmCust.Top = 0
frmCust.Left = 0

datCust.DatabaseName = (App.Path & "\MovieDB.mdb")
datCust.RecordSource = ("Accounts")

datRentals.DatabaseName = (App.Path & "\MovieDB.mdb")
datReturns.DatabaseName = (App.Path & "\MovieDB.mdb")
datLate.DatabaseName = (App.Path & "\Moviedb.mdb")

cmdCancel.Visible = False
cmdUpdate.Visible = False

CountMe

GetRentals
End Sub

Private Sub CountMe()
Dim i As Integer

ShowEdit

datCust.Refresh

i = datCust.Recordset.RecordCount

If i = 0 Then
MsgBox "There are no accounts in the database.", vbExclamation, "No Accounts"
HideButtons
Exit Sub
End If

End Sub

Private Sub HideButtons()
cmdFirst.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdDel.Enabled = False
cmdEdit.Enabled = False
cmdFind.Enabled = False
End Sub

Private Sub HideEdit()
cmdFirst.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdDel.Enabled = False
cmdAdd.Enabled = False
cmdEdit.Enabled = False
cmdClose.Enabled = False
cmdFind.Enabled = False
cmdCancel.Visible = True
cmdUpdate.Visible = True
txtAccount.Enabled = True
txtDriver.Enabled = True
txtName.Enabled = True
txtAddr.Enabled = True
txtCity.Enabled = True
txtProv.Enabled = True
txtPost.Enabled = True
txtPhone.Enabled = True
End Sub

Private Sub ShowEdit()
cmdFirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdDel.Enabled = True
cmdAdd.Enabled = True
cmdEdit.Enabled = True
cmdClose.Enabled = True
cmdFind.Enabled = True
cmdCancel.Visible = False
cmdUpdate.Visible = False
txtAccount.Enabled = False
txtDriver.Enabled = False
txtName.Enabled = False
txtAddr.Enabled = False
txtCity.Enabled = False
txtProv.Enabled = False
txtPost.Enabled = False
txtPhone.Enabled = False
End Sub

Private Sub CountRentals()
Dim c As Integer
Dim r As Integer
Dim t As Integer

c = dbgRental.ApproxCount
r = dbgHistory.ApproxCount

t = c + r

txtCurrent.Text = c
txtRented.Text = t
End Sub
