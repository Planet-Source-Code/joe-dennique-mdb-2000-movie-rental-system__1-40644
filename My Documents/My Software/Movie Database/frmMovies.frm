VERSION 5.00
Begin VB.Form frmMovies 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movie DataBase"
   ClientHeight    =   6180
   ClientLeft      =   1800
   ClientTop       =   1935
   ClientWidth     =   8550
   Icon            =   "frmMovies.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8550
   Begin VB.Frame fraSearch 
      Caption         =   "Search Database by Barcode..."
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4500
      TabIndex        =   27
      Top             =   4875
      Width           =   3840
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Default         =   -1  'True
         Height          =   315
         Left            =   2850
         TabIndex        =   2
         Top             =   225
         Width           =   915
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   2640
      End
   End
   Begin VB.Frame fraCount 
      Caption         =   "Total Movies in database..."
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   150
      TabIndex        =   25
      Top             =   4875
      Width           =   4140
      Begin VB.TextBox txtCount 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1050
         TabIndex        =   26
         Top             =   225
         Width           =   1965
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   7425
      TabIndex        =   18
      Top             =   5700
      Width           =   915
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Database Editing..."
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   4500
      TabIndex        =   24
      Top             =   3900
      Width           =   3840
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   2400
         TabIndex        =   16
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   315
         Left            =   1425
         TabIndex        =   17
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Height          =   315
         Left            =   2400
         TabIndex        =   15
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   1425
         TabIndex        =   14
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   315
         Left            =   450
         TabIndex        =   13
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame fraNav 
      Caption         =   "Database Navigation..."
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   150
      TabIndex        =   23
      Top             =   3900
      Width           =   4140
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   315
         Left            =   3075
         TabIndex        =   10
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   315
         Left            =   2100
         TabIndex        =   11
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev."
         Height          =   315
         Left            =   1125
         TabIndex        =   12
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame fraMovie 
      Caption         =   "Movie Information..."
      ForeColor       =   &H00FF0000&
      Height          =   3600
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   8190
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select Type"
         Height          =   315
         Left            =   4275
         TabIndex        =   6
         Top             =   1275
         Width           =   1290
      End
      Begin VB.TextBox txtType 
         DataField       =   "type"
         DataSource      =   "datMovie"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2550
         TabIndex        =   7
         Top             =   1275
         Width           =   1665
      End
      Begin VB.Data datMovie 
         Caption         =   "Movies"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.TextBox txtNotes 
         DataField       =   "Notes"
         DataSource      =   "datMovie"
         Enabled         =   0   'False
         Height          =   1800
         Left            =   2580
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1620
         Width           =   5490
      End
      Begin VB.TextBox txtRating 
         DataField       =   "Rating"
         DataSource      =   "datMovie"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2550
         TabIndex        =   5
         Top             =   975
         Width           =   5490
      End
      Begin VB.TextBox txtTitle 
         DataField       =   "Title"
         DataSource      =   "datMovie"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2550
         TabIndex        =   4
         Top             =   675
         Width           =   5490
      End
      Begin VB.TextBox txtBarcode 
         DataField       =   "barcode"
         DataSource      =   "datMovie"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2550
         TabIndex        =   3
         Top             =   375
         Width           =   2565
      End
      Begin VB.Label lblType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Type:"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1260
         Width           =   2190
      End
      Begin VB.Label lblNotes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Notes:"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1560
         Width           =   2190
      End
      Begin VB.Label lblRating 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ESRB Rating:"
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
         Left            =   225
         TabIndex        =   21
         Top             =   975
         Width           =   2190
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Title:"
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
         Left            =   225
         TabIndex        =   20
         Top             =   675
         Width           =   2190
      End
      Begin VB.Label lblBarcode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Barcode:"
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
         Left            =   225
         TabIndex        =   19
         Top             =   375
         Width           =   2190
      End
   End
End
Attribute VB_Name = "frmMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    datMovie.Recordset.AddNew
    AllowEdit
    txtBarcode.SetFocus
    cmdUpdate.Default = True
End Sub

Private Sub cmdCancel_Click()
    datMovie.Recordset.CancelUpdate
    DisAllowEdit
    datMovie.Refresh
    txtSearch.Text = ""
    txtSearch.SetFocus
    cmdFind.Default = True
    CountMe
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    datMovie.Recordset.Delete
    MsgBox "Movie has been removed from the database", vbExclamation, "Movie Removed"
    datMovie.Refresh
    CountMe
End Sub

Private Sub cmdEdit_Click()
    datMovie.Recordset.Edit
    AllowEdit
    txtBarcode.SetFocus
    cmdUpdate.Default = True
End Sub

Private Sub cmdFind_Click()
Dim sName As String
Dim sFind As String

sName = txtSearch.Text
sFind = "barcode LIKE '" & sName & "*'"

If txtSearch.Text = "" Then
MsgBox "barcode needed for search", vbExclamation, "Barcode Needed"
txtSearch.SetFocus
Exit Sub
End If

datMovie.Recordset.FindFirst sFind
If datMovie.Recordset.NoMatch Then
MsgBox "There was no movie found with that barcode", vbExclamation, "No Match"
txtSearch.Text = ""
txtSearch.SetFocus
End If

End Sub

Private Sub cmdFirst_Click()
    datMovie.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
    datMovie.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
    datMovie.Recordset.MoveNext
    If datMovie.Recordset.EOF Then
    MsgBox "You are viewing the last mocie in the database", vbExclamation, "Last Movie"
    datMovie.Recordset.MovePrevious
    End If
End Sub

Private Sub cmdPrev_Click()
    datMovie.Recordset.MovePrevious
    If datMovie.Recordset.BOF Then
    MsgBox "You are viewing the first mocie in the database", vbExclamation, "First Movie"
    datMovie.Recordset.MoveNext
    End If
End Sub



Private Sub cmdSelect_Click()
frmType.Show vbModal
End Sub

Private Sub cmdUpdate_Click()
    If txtBarcode.Text = "" Then
    MsgBox "Barcode information needed", vbExclamation, "Barcode Needed"
    txtBarcode.SetFocus
    Exit Sub
    End If
    
    datMovie.Recordset.Update
    MsgBox "Movie information updated", vbExclamation, "Information Updated"
    DisAllowEdit
    txtSearch.Text = ""
    txtSearch.SetFocus
    cmdFind.Default = True
    CountMe
End Sub

Private Sub Form_Load()
frmMovies.Top = 0
frmMovies.Left = 0

    datMovie.DatabaseName = (App.Path & "\movieDB.mdb")
    datMovie.RecordSource = ("Movies")
    
    'cmdCancel.Visible = False
    'cmdUpdate.Visible = False
    'cmdSelect.Visible = False
    
    DisAllowEdit
    
    CountMe
End Sub


Private Sub CountMe()
Dim i As Integer

datMovie.Refresh

i = datMovie.Recordset.RecordCount

txtCount.Text = i

If i = 0 Then
MsgBox "There are no movies in the database", vbExclamation, "No Movies"
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

Private Sub AllowEdit()
cmdFirst.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdClose.Enabled = False
cmdFind.Enabled = False
cmdEdit.Visible = False
cmdAdd.Visible = False
cmdDel.Visible = False
cmdCancel.Visible = True
cmdUpdate.Visible = True
txtBarcode.Enabled = True
txtTitle.Enabled = True
txtRating.Enabled = True
txtNotes.Enabled = True
cmdSelect.Visible = True
End Sub

Private Sub DisAllowEdit()
cmdFirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdClose.Enabled = True
cmdFind.Enabled = True
cmdEdit.Visible = True
cmdAdd.Visible = True
cmdDel.Visible = True
cmdCancel.Visible = False
cmdUpdate.Visible = False
txtBarcode.Enabled = False
txtTitle.Enabled = False
txtRating.Enabled = False
txtNotes.Enabled = False
cmdSelect.Visible = False
End Sub
