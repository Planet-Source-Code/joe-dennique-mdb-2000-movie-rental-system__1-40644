VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H00000000&
   Caption         =   "Movie DataBase 2002"
   ClientHeight    =   8235
   ClientLeft      =   2070
   ClientTop       =   2325
   ClientWidth     =   8625
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMain.frx":08CA
   WindowState     =   2  'Maximized
   Begin VB.Data datCheck 
      Align           =   1  'Align Top
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   8625
   End
   Begin MSComDlg.CommonDialog cmnMain 
      Left            =   150
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilstMain 
      Left            =   150
      Top             =   2025
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6254
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7704
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8874
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":9150
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":95A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":99F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbarmain 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   7995
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9578
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "7:36 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/11/2002"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilstMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEY_EXIT"
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEY_CONFIG"
            Object.ToolTipText     =   "Configure Settings"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEY_SYSTEM"
            Object.ToolTipText     =   "Start Retal System"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEY_MOVIE"
            Object.ToolTipText     =   "View / Edit Movie Database"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEY_LATELIST"
            Object.ToolTipText     =   "View / Edit Late List"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEY_CUST"
            Object.ToolTipText     =   "View / Edit Customer Database"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEY_RETURN"
            Object.ToolTipText     =   "Open Return System"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEY_PRINT"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KEY_ABOUT"
            Object.ToolTipText     =   "About this software"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuF_PrintS 
         Caption         =   "&Printer Setup"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuF_Print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuF_Space 
         Caption         =   "-"
      End
      Begin VB.Menu mnuF_Exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "&Open"
      Begin VB.Menu mnuO_Accounts 
         Caption         =   "&Accounts"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuO_DB 
         Caption         =   "&Database"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuO_Late 
         Caption         =   "&Late List"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuO_Return 
         Caption         =   "&Return System"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuO_Rent 
         Caption         =   "&Rental System"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Setup"
      Begin VB.Menu mnuS_Config 
         Caption         =   "&Configuration"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuH_About 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPrint As Boolean

Private Sub MDIForm_Load()
Dim z As Integer

INISetup App.Path & "\Config.dat", 5500

OpenDB

datCheck.Refresh
z = datCheck.Recordset.RecordCount

If z = 0 Then
Load frmConfig
CloseDB
Exit Sub
End If

CloseDB
End Sub

Private Sub mnuF_Exit_Click()
Unload Me
End Sub

Private Sub mnuF_Print_Click()

On Error GoTo Print_Error
ActiveForm.PrintForm

Print_Error:
Beep
Exit Sub
End Sub

Private Sub mnuF_PrintS_Click()
cmnMain.ShowPrinter
End Sub

Private Sub mnuH_About_Click()
    Load frmAbout
End Sub

Private Sub mnuO_Accounts_Click()
Load frmCust
End Sub

Private Sub mnuO_DB_Click()
    Load frmMovies
End Sub

Private Sub mnuO_Late_Click()
    Load frmLateList
End Sub

Private Sub mnuO_Rent_Click()
    Load frmSystem
End Sub

Private Sub mnuO_Return_Click()
    Load frmCheckIn
End Sub

Private Sub mnuS_Config_Click()
    Load frmConfig
End Sub

Private Sub OpenDB()
datCheck.DatabaseName = (App.Path & "\movieDB.mdb")
datCheck.RecordSource = ("Config")
End Sub

Private Sub CloseDB()
datCheck.DatabaseName = ""
datCheck.RecordSource = ""
End Sub

Public Sub ProcessMessage(ByVal sMessage As String)
Dim bPrint As Boolean

On Error GoTo ProcessMessage_ERROR
Select Case UCase$(sMessage)
Case "KEY_EXIT"
    Unload Me
Case "KEY_CONFIG"
    Load frmConfig
Case "KEY_SYSTEM"
    Load frmSystem
Case "KEY_MOVIE"
    Load frmMovies
Case "KEY_LATELIST"
    Load frmLateList
Case "KEY_CUST"
    Load frmCust
Case "KEY_RETURN"
    Load frmCheckIn
Case "KEY_PRINT"
    ActiveForm.PrintForm
Case "KEY_ABOUT"
    Load frmAbout
End Select

ProcessMessage_EXIT:
    Exit Sub
    
ProcessMessage_ERROR:
    Beep
    Resume ProcessMessage_EXIT
    
End Sub

Private Sub tbarMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    ProcessMessage Button.Key
End Sub
