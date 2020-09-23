VERSION 5.00
Begin VB.Form frmType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rental Type Selection"
   ClientHeight    =   1410
   ClientLeft      =   3630
   ClientTop       =   5805
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optOld 
      Caption         =   "Old"
      Height          =   240
      Left            =   1500
      TabIndex        =   6
      Top             =   675
      Width           =   1140
   End
   Begin VB.OptionButton optKids 
      Caption         =   "Kids"
      Height          =   240
      Left            =   2850
      TabIndex        =   5
      Top             =   675
      Width           =   1140
   End
   Begin VB.OptionButton OptOther 
      Caption         =   "Other"
      Height          =   240
      Left            =   4200
      TabIndex        =   4
      Top             =   675
      Width           =   1140
   End
   Begin VB.OptionButton optNew 
      Caption         =   "New"
      Height          =   240
      Left            =   150
      TabIndex        =   2
      Top             =   675
      Value           =   -1  'True
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Select"
      Height          =   315
      Left            =   5700
      TabIndex        =   1
      Top             =   900
      Width           =   765
   End
   Begin VB.TextBox txtType 
      Height          =   285
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      Caption         =   "Select a rental type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   1575
      TabIndex        =   3
      Top             =   75
      Width           =   3690
   End
End
Attribute VB_Name = "frmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
If optNew.Value = True Then
txtType.Text = "New"
ElseIf optOld.Value = True Then
txtType.Text = "Old"
ElseIf optKids.Value = True Then
txtType.Text = "Kids"
ElseIf OptOther.Value = True Then
txtType.Text = "Other"
End If
frmMovies.txtType.Text = frmType.txtType.Text

frmType.Hide
End Sub
