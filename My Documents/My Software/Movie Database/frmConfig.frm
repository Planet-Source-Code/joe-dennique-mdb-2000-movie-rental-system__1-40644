VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   6270
   ClientLeft      =   3150
   ClientTop       =   2565
   ClientWidth     =   6645
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6645
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   315
      Left            =   2475
      TabIndex        =   15
      Top             =   5775
      Width           =   1155
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   315
      Left            =   3825
      TabIndex        =   16
      Top             =   5775
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   5175
      TabIndex        =   17
      Top             =   5775
      Width           =   1155
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration..."
      ForeColor       =   &H00FF0000&
      Height          =   5535
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   6315
      Begin VB.TextBox txtSpecLate 
         DataField       =   "SpecLate"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   12
         ToolTipText     =   "G.S.T. as a decimal 0.07 is 7%"
         Top             =   4350
         Width           =   1830
      End
      Begin VB.TextBox txtKidsLate 
         DataField       =   "KidsLate"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   11
         ToolTipText     =   "G.S.T. as a decimal 0.07 is 7%"
         Top             =   3975
         Width           =   1830
      End
      Begin VB.TextBox txtOldLate 
         DataField       =   "OldLate"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   10
         ToolTipText     =   "G.S.T. as a decimal 0.07 is 7%"
         Top             =   3600
         Width           =   1830
      End
      Begin VB.TextBox txtNewLate 
         DataField       =   "NewLate"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   9
         ToolTipText     =   "G.S.T. as a decimal 0.07 is 7%"
         Top             =   3225
         Width           =   1830
      End
      Begin VB.TextBox txtAltDay 
         DataField       =   "SpecialDay"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   8
         ToolTipText     =   "Number of days an other movie is allowed to be rented"
         Top             =   2850
         Width           =   1830
      End
      Begin VB.TextBox txtKidsDay 
         DataField       =   "KidsDay"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   7
         ToolTipText     =   "Number of days a kids movie is allowed to be rented"
         Top             =   2475
         Width           =   1830
      End
      Begin VB.TextBox txtOldDay 
         DataField       =   "OldDay"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   6
         ToolTipText     =   "Number of days an old movie is allowed to be rented"
         Top             =   2100
         Width           =   1830
      End
      Begin VB.TextBox txtNewDay 
         DataField       =   "NewDay"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   5
         ToolTipText     =   "Number of days a new movie is allowed to be rented"
         Top             =   1725
         Width           =   1830
      End
      Begin VB.TextBox txtPST 
         DataField       =   "PST"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   14
         ToolTipText     =   "P.S.T. as a decimal 0.07 is 7%"
         Top             =   5100
         Width           =   1830
      End
      Begin VB.TextBox txtGST 
         DataField       =   "GST"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3675
         TabIndex        =   13
         ToolTipText     =   "G.S.T. as a decimal 0.07 is 7%"
         Top             =   4725
         Width           =   1830
      End
      Begin VB.TextBox txtKids 
         DataField       =   "KidsMovie"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3705
         TabIndex        =   3
         ToolTipText     =   "Rental Cost for kids movie"
         Top             =   1020
         Width           =   1830
      End
      Begin VB.TextBox txtOld 
         DataField       =   "OldMovie"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3705
         TabIndex        =   2
         ToolTipText     =   "Rental Cost for old movie"
         Top             =   660
         Width           =   1830
      End
      Begin VB.TextBox txtNew 
         DataField       =   "NewMovie"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3705
         TabIndex        =   1
         ToolTipText     =   "Rental Cost for new movie"
         Top             =   300
         Width           =   1830
      End
      Begin VB.TextBox txtSpec 
         DataField       =   "Special"
         DataSource      =   "datConfig"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3705
         TabIndex        =   4
         ToolTipText     =   "Rental Cost for other movie"
         Top             =   1380
         Width           =   1830
      End
      Begin VB.Label Label2 
         Caption         =   "Other Movie Late Charge:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   75
         TabIndex        =   31
         Top             =   4350
         Width           =   3435
      End
      Begin VB.Label Label1 
         Caption         =   "Kids Movie Late Charge:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   75
         TabIndex        =   30
         Top             =   3975
         Width           =   3060
      End
      Begin VB.Label lblOldLate 
         Caption         =   "Old Movie Late Charge:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   75
         TabIndex        =   29
         Top             =   3600
         Width           =   3060
      End
      Begin VB.Label lblNewLate 
         Caption         =   "New Movie Late Charge:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   75
         TabIndex        =   28
         Top             =   3225
         Width           =   3060
      End
      Begin VB.Label lblAltDay 
         Caption         =   "Other Movie Length:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   75
         TabIndex        =   27
         Top             =   2850
         Width           =   2775
      End
      Begin VB.Label lblKidsDay 
         Caption         =   "Kids Movie Length:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   75
         TabIndex        =   26
         Top             =   2475
         Width           =   2775
      End
      Begin VB.Label lblOldDay 
         Caption         =   "Old Movie Length:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   75
         TabIndex        =   25
         Top             =   2100
         Width           =   2775
      End
      Begin VB.Label lblNewDay 
         Caption         =   "New Movie Length:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   75
         TabIndex        =   24
         Top             =   1725
         Width           =   2775
      End
      Begin VB.Label lblPST 
         Caption         =   "PST  (#.##):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   45
         TabIndex        =   23
         Top             =   5100
         Width           =   2175
      End
      Begin VB.Label lblGST 
         Caption         =   "GST  (#.##):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   45
         TabIndex        =   22
         Top             =   4740
         Width           =   2175
      End
      Begin VB.Label lblSpec 
         Caption         =   "Other Movie Cost:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   45
         TabIndex        =   21
         Top             =   1380
         Width           =   2355
      End
      Begin VB.Label lblKids 
         Caption         =   "Kids Movie Cost:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   45
         TabIndex        =   20
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label lblOld 
         Caption         =   "Old Movie Cost:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   45
         TabIndex        =   19
         Top             =   660
         Width           =   2175
      End
      Begin VB.Label lblNew 
         Caption         =   "New Movie Cost:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   45
         TabIndex        =   18
         Top             =   300
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdEdit_Click()
AllowEdit
End Sub

Private Sub cmdUpdate_Click()
WriteData
DisAllowEdit
End Sub

Private Sub Form_Load()
frmConfig.Top = 0
frmConfig.Left = 0

ShowData

If txtNew.Text = "" Then
MsgBox "Please complete the configuration information.", vbExclamation, "Information needed"
End If

DisAllowEdit
End Sub


Private Sub AllowEdit()
cmdClose.Enabled = False
cmdEdit.Enabled = False
cmdUpdate.Enabled = True
txtNew.Enabled = True
txtOld.Enabled = True
txtKids.Enabled = True
txtSpec.Enabled = True
txtGST.Enabled = True
txtPST.Enabled = True
txtNewDay.Enabled = True
txtOldDay.Enabled = True
txtKidsDay.Enabled = True
txtAltDay.Enabled = True
txtNewLate.Enabled = True
txtOldLate.Enabled = True
txtKidsLate.Enabled = True
txtSpecLate.Enabled = True
txtNew.SetFocus
cmdUpdate.Default = True
End Sub

Private Sub DisAllowEdit()
cmdClose.Enabled = True
cmdEdit.Enabled = True
cmdUpdate.Enabled = False
txtNew.Enabled = False
txtOld.Enabled = False
txtKids.Enabled = False
txtSpec.Enabled = False
txtGST.Enabled = False
txtPST.Enabled = False
txtNewDay.Enabled = False
txtOldDay.Enabled = False
txtKidsDay.Enabled = False
txtAltDay.Enabled = False
txtNewLate.Enabled = False
txtOldLate.Enabled = False
txtKidsLate.Enabled = False
txtSpecLate.Enabled = False
cmdEdit.Default = True
End Sub

Private Sub ShowData()
txtNew.Text = Read_Ini("settings", "New")
txtOld.Text = Read_Ini("Settings", "Old")
txtKids.Text = Read_Ini("Settings", "Kids")
txtSpec.Text = Read_Ini("Settings", "Other")
txtNewDay.Text = Read_Ini("Settings", "NewDays")
txtOldDay.Text = Read_Ini("Settings", "OldDays")
txtKidsDay.Text = Read_Ini("Settings", "KidsDays")
txtAltDay.Text = Read_Ini("Settings", "OtherDays")
txtNewLate.Text = Read_Ini("Settings", "NewLate")
txtOldLate.Text = Read_Ini("Settings", "OldLate")
txtKidsLate.Text = Read_Ini("Settings", "KidsLate")
txtSpecLate.Text = Read_Ini("Settings", "OtherLate")
txtGST.Text = Read_Ini("Settings", "GST")
txtPST.Text = Read_Ini("Settings", "PST")
End Sub

Private Sub WriteData()
Write_Ini "Settings", "New", txtNew.Text
Write_Ini "Settings", "Old", txtOld.Text
Write_Ini "Settings", "Kids", txtKids.Text
Write_Ini "Settings", "Other", txtSpec.Text
Write_Ini "Settings", "NewDays", txtNewDay.Text
Write_Ini "Settings", "OldDays", txtOldDay.Text
Write_Ini "Settings", "KidsDays", txtKidsDay.Text
Write_Ini "Settings", "OtherDays", txtAltDay.Text
Write_Ini "Settings", "NewLate", txtNewLate.Text
Write_Ini "Settings", "OldLate", txtOldLate.Text
Write_Ini "Settings", "KidsLate", txtKidsLate.Text
Write_Ini "Settings", "OtherLate", txtSpecLate.Text
Write_Ini "Settings", "GST", txtGST.Text
Write_Ini "Settings", "PST", txtPST.Text
End Sub
