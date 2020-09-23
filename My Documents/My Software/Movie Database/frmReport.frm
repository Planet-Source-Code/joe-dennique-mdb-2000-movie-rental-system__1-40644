VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Receipt"
   ClientHeight    =   8940
   ClientLeft      =   2295
   ClientTop       =   1620
   ClientWidth     =   6735
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   6735
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox mskInv 
      DataField       =   "SubTotal"
      DataSource      =   "datReport"
      Height          =   240
      Left            =   4425
      TabIndex        =   18
      Top             =   7200
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Data datReport 
      Caption         =   "Report"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   2565
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "GST"
      DataSource      =   "datReport"
      Height          =   240
      Left            =   4425
      TabIndex        =   19
      Top             =   7650
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      DataField       =   "PST"
      DataSource      =   "datReport"
      Height          =   240
      Left            =   4425
      TabIndex        =   20
      Top             =   8100
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      DataField       =   "TotalCost"
      DataSource      =   "datReport"
      Height          =   240
      Left            =   4425
      TabIndex        =   21
      Top             =   8550
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblInvPhone 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Phone"
      DataSource      =   "datReport"
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
      Left            =   2040
      TabIndex        =   25
      Top             =   1560
      Width           =   4590
   End
   Begin VB.Label lblInvName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "CustName"
      DataSource      =   "datReport"
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
      Left            =   2040
      TabIndex        =   24
      Top             =   1080
      Width           =   4530
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone #:"
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
      TabIndex        =   23
      Top             =   1560
      Width           =   1590
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   22
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Label lclInvAccount 
      BackColor       =   &H00FFFFFF&
      DataField       =   "CustAccount"
      DataSource      =   "datReport"
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
      Left            =   2025
      TabIndex        =   17
      Top             =   600
      Width           =   3510
   End
   Begin VB.Label lblInvID 
      BackColor       =   &H00FFFFFF&
      DataField       =   "InvoiceID"
      DataSource      =   "datReport"
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
      Left            =   2025
      TabIndex        =   16
      Top             =   150
      Width           =   3570
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total:"
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
      Left            =   2625
      TabIndex        =   15
      Top             =   8475
      Width           =   1590
   End
   Begin VB.Label lblPST 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P.S.T.:"
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
      Left            =   2625
      TabIndex        =   14
      Top             =   8025
      Width           =   1590
   End
   Begin VB.Label lblGST 
      BackColor       =   &H00FFFFFF&
      Caption         =   "G.S.T.:"
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
      Left            =   2625
      TabIndex        =   13
      Top             =   7575
      Width           =   1590
   End
   Begin VB.Label lblSub 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sub-Total:"
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
      Left            =   2625
      TabIndex        =   12
      Top             =   7125
      Width           =   1590
   End
   Begin VB.Label lblRental10 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental10"
      DataSource      =   "datReport"
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
      TabIndex        =   11
      Top             =   6450
      Width           =   3615
   End
   Begin VB.Label lblRental9 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental9"
      DataSource      =   "datReport"
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
      TabIndex        =   10
      Top             =   6000
      Width           =   3540
   End
   Begin VB.Label lblRental8 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental8"
      DataSource      =   "datReport"
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
      TabIndex        =   9
      Top             =   5550
      Width           =   3540
   End
   Begin VB.Label lblRental7 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental7"
      DataSource      =   "datReport"
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
      TabIndex        =   8
      Top             =   5100
      Width           =   3540
   End
   Begin VB.Label lblRental6 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental6"
      DataSource      =   "datReport"
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
      TabIndex        =   7
      Top             =   4650
      Width           =   3540
   End
   Begin VB.Label lblRental5 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental5"
      DataSource      =   "datReport"
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
      TabIndex        =   6
      Top             =   4200
      Width           =   3540
   End
   Begin VB.Label lblRental4 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental4"
      DataSource      =   "datReport"
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
      TabIndex        =   5
      Top             =   3750
      Width           =   3540
   End
   Begin VB.Label lblRental3 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental3"
      DataSource      =   "datReport"
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
      TabIndex        =   4
      Top             =   3300
      Width           =   3540
   End
   Begin VB.Label lblRental2 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental2"
      DataSource      =   "datReport"
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
      TabIndex        =   3
      Top             =   2850
      Width           =   3540
   End
   Begin VB.Label lblRental1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Rental1"
      DataSource      =   "datReport"
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
      TabIndex        =   2
      Top             =   2400
      Width           =   3540
   End
   Begin VB.Label lblAccount 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Account #:"
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
      TabIndex        =   1
      Top             =   600
      Width           =   1590
   End
   Begin VB.Label lblEntry 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Receipt #:"
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
      TabIndex        =   0
      Top             =   150
      Width           =   1590
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
datReport.DatabaseName = (App.Path & "\movieDB.mdb")
datReport.RecordSource = ("Invoice")

FindMe
End Sub

Private Sub FindMe()
Dim sEntry As String
Dim sFind As String

datReport.Refresh

sEntry = frmSystem.txtInvReport.Text
sFind = "InvoiceID LIKE " & sEntry & ""

If frmSystem.txtInvReport.Text = "" Then
Exit Sub
End If

datReport.Recordset.FindFirst sFind

End Sub

Private Sub Form_Terminate()
datReport.DatabaseName = ""
datReport.RecordSource = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
datReport.DatabaseName = ""
datReport.RecordSource = ""
End Sub

