VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmSystem 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rental System"
   ClientHeight    =   8895
   ClientLeft      =   1125
   ClientTop       =   435
   ClientWidth     =   10620
   Icon            =   "frmSystem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   10620
   Begin VB.TextBox txtRentName 
      DataField       =   "CustName"
      DataSource      =   "datRental"
      Height          =   315
      Left            =   5100
      TabIndex        =   121
      Top             =   8640
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtRentPhone 
      DataField       =   "Phone"
      DataSource      =   "datRental"
      Height          =   315
      Left            =   5100
      TabIndex        =   120
      Top             =   8340
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame fraRentType 
      BackColor       =   &H00FF8080&
      Caption         =   "Rental Type..."
      Height          =   1035
      Left            =   120
      TabIndex        =   106
      Top             =   9120
      Visible         =   0   'False
      Width           =   9915
      Begin VB.TextBox txtRentType 
         DataField       =   "Type"
         DataSource      =   "datRental"
         Height          =   285
         Left            =   4440
         TabIndex        =   117
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtType10 
         Height          =   285
         Left            =   8820
         TabIndex        =   116
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtType9 
         Height          =   285
         Left            =   7800
         TabIndex        =   115
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtType8 
         Height          =   285
         Left            =   6840
         TabIndex        =   114
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtType7 
         Height          =   285
         Left            =   5880
         TabIndex        =   113
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtType6 
         Height          =   285
         Left            =   4920
         TabIndex        =   112
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtType5 
         Height          =   285
         Left            =   3960
         TabIndex        =   111
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtType4 
         Height          =   285
         Left            =   3000
         TabIndex        =   110
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtType3 
         Height          =   285
         Left            =   2040
         TabIndex        =   109
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtType2 
         Height          =   285
         Left            =   1080
         TabIndex        =   108
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtType1 
         Height          =   285
         Left            =   120
         TabIndex        =   107
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraBarcode 
      BackColor       =   &H00FF8080&
      Caption         =   "Barcode Info."
      Height          =   4215
      Left            =   3825
      TabIndex        =   92
      Top             =   3210
      Visible         =   0   'False
      Width           =   1065
      Begin VB.TextBox txtIsBack 
         DataField       =   "IsBack"
         DataSource      =   "datRental"
         Height          =   315
         Left            =   150
         TabIndex        =   105
         Top             =   3450
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtRentBar 
         DataField       =   "Barcode"
         DataSource      =   "datRental"
         Height          =   315
         Left            =   150
         TabIndex        =   103
         Top             =   3750
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode10 
         Height          =   315
         Left            =   150
         TabIndex        =   102
         Top             =   3000
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode9 
         Height          =   315
         Left            =   150
         TabIndex        =   101
         Top             =   2700
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode8 
         Height          =   315
         Left            =   150
         TabIndex        =   100
         Top             =   2400
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode7 
         Height          =   315
         Left            =   150
         TabIndex        =   99
         Top             =   2100
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode6 
         Height          =   315
         Left            =   150
         TabIndex        =   98
         Top             =   1800
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode5 
         Height          =   315
         Left            =   150
         TabIndex        =   97
         Top             =   1500
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode4 
         Height          =   315
         Left            =   150
         TabIndex        =   96
         Top             =   1200
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode3 
         Height          =   315
         Left            =   150
         TabIndex        =   95
         Top             =   900
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode2 
         Height          =   315
         Left            =   150
         TabIndex        =   94
         Top             =   600
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtBarcode1 
         Height          =   315
         Left            =   150
         TabIndex        =   93
         Top             =   300
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin VB.TextBox txtEntryID 
      DataField       =   "InvoiceID"
      DataSource      =   "datInvoice"
      Height          =   315
      Left            =   4275
      TabIndex        =   91
      Top             =   7860
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtInvReport 
      Height          =   315
      Left            =   4275
      TabIndex        =   90
      Top             =   7560
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtReturnMe 
      DataField       =   "DueBack"
      DataSource      =   "datRental"
      Height          =   315
      Left            =   4275
      TabIndex        =   89
      Top             =   8460
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtRentMe 
      DataField       =   "Title"
      DataSource      =   "datRental"
      Height          =   315
      Left            =   4275
      TabIndex        =   88
      Top             =   8160
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H00FF8080&
      Caption         =   "Date Info."
      Height          =   990
      Left            =   6975
      TabIndex        =   84
      Top             =   7785
      Visible         =   0   'False
      Width           =   1140
      Begin VB.TextBox txtNow 
         Height          =   315
         Left            =   150
         TabIndex        =   87
         Top             =   300
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox txtRentDay 
         Height          =   315
         Left            =   150
         TabIndex        =   86
         Top             =   600
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.Frame fraRentTitle 
      BackColor       =   &H00FF8080&
      Caption         =   "Rental Title Info."
      Height          =   1215
      Left            =   150
      TabIndex        =   72
      Top             =   7710
      Visible         =   0   'False
      Width           =   3915
      Begin VB.TextBox txtRentAccount 
         DataField       =   "Account"
         DataSource      =   "datRental"
         Height          =   315
         Left            =   2850
         TabIndex        =   83
         Top             =   825
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle5 
         Height          =   315
         Left            =   1050
         TabIndex        =   82
         Top             =   525
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle4 
         Height          =   315
         Left            =   1050
         TabIndex        =   81
         Top             =   225
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle3 
         Height          =   315
         Left            =   150
         TabIndex        =   80
         Top             =   825
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle2 
         Height          =   315
         Left            =   150
         TabIndex        =   79
         Top             =   525
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle1 
         Height          =   315
         Left            =   150
         TabIndex        =   78
         Top             =   225
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle10 
         Height          =   315
         Left            =   2850
         TabIndex        =   77
         Top             =   225
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle9 
         Height          =   315
         Left            =   1950
         TabIndex        =   76
         Top             =   825
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle8 
         Height          =   315
         Left            =   1950
         TabIndex        =   75
         Top             =   525
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle7 
         Height          =   315
         Left            =   1950
         TabIndex        =   74
         Top             =   225
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtTitle6 
         Height          =   315
         Left            =   1050
         TabIndex        =   73
         Top             =   825
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.Frame fraLength 
      BackColor       =   &H00FF8080&
      Caption         =   "Rental Length Info."
      Height          =   2115
      Left            =   5100
      TabIndex        =   61
      Top             =   6135
      Visible         =   0   'False
      Width           =   1590
      Begin VB.TextBox txtRentDate 
         DataField       =   "Rented"
         DataSource      =   "datRental"
         Height          =   315
         Left            =   75
         TabIndex        =   85
         Top             =   1725
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie10 
         Height          =   315
         Left            =   750
         TabIndex        =   71
         Top             =   1425
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie9 
         Height          =   315
         Left            =   750
         TabIndex        =   70
         Top             =   1125
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie8 
         Height          =   315
         Left            =   750
         TabIndex        =   69
         Top             =   825
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie7 
         Height          =   315
         Left            =   750
         TabIndex        =   68
         Top             =   525
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie5 
         Height          =   315
         Left            =   75
         TabIndex        =   67
         Top             =   1425
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie4 
         Height          =   315
         Left            =   75
         TabIndex        =   66
         Top             =   1125
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie3 
         Height          =   315
         Left            =   75
         TabIndex        =   65
         Top             =   825
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie2 
         Height          =   315
         Left            =   75
         TabIndex        =   64
         Top             =   525
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie1 
         Height          =   315
         Left            =   75
         TabIndex        =   63
         Top             =   225
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtMovie6 
         Height          =   315
         Left            =   750
         TabIndex        =   62
         Top             =   225
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin VB.Frame fraConfig 
      BackColor       =   &H00FF8080&
      Caption         =   "Config Info."
      Height          =   3315
      Left            =   5100
      TabIndex        =   50
      Top             =   2685
      Visible         =   0   'False
      Width           =   1065
      Begin VB.TextBox txtNew 
         DataField       =   "NewMovie"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   60
         Top             =   225
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtOld 
         DataField       =   "OldMovie"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   59
         Top             =   525
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtKids 
         DataField       =   "KidsMovie"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   58
         Top             =   825
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtSpec 
         DataField       =   "Special"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   57
         Top             =   1125
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtCfgGST 
         DataField       =   "GST"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   56
         Top             =   2625
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtCfgPST 
         DataField       =   "PST"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   55
         Top             =   2925
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtKidsDay 
         DataField       =   "KidsDay"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   54
         Top             =   2025
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtNewDay 
         DataField       =   "NewDay"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   53
         Top             =   1425
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtAltday 
         DataField       =   "SpecialDay"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   52
         Top             =   2325
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtOldDay 
         DataField       =   "OldDay"
         DataSource      =   "datConfig"
         Height          =   285
         Left            =   75
         TabIndex        =   51
         Top             =   1725
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.Frame fraInvCount 
      BackColor       =   &H00FF8080&
      Caption         =   "Invoice Information"
      Height          =   4365
      Left            =   150
      TabIndex        =   31
      Top             =   3135
      Visible         =   0   'False
      Width           =   3465
      Begin VB.TextBox txtInvName 
         DataField       =   "CustName"
         DataSource      =   "datInvoice"
         Height          =   315
         Left            =   1860
         TabIndex        =   119
         Top             =   3600
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtInvPhone 
         DataField       =   "Phone"
         DataSource      =   "datInvoice"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1860
         TabIndex        =   118
         Top             =   3240
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtrental1 
         DataField       =   "Rental1"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   49
         Top             =   225
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtrental2 
         DataField       =   "Rental2"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtrental3 
         DataField       =   "Rental3"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   47
         Top             =   975
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtrental4 
         DataField       =   "Rental4"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   46
         Top             =   1350
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtrental5 
         DataField       =   "Rental5"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   45
         Top             =   1725
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtInvAccount 
         DataField       =   "Custaccount"
         DataSource      =   "datInvoice"
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         TabIndex        =   44
         Top             =   3975
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtInvMovie 
         DataField       =   "totalMovie"
         DataSource      =   "datInvoice"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1875
         TabIndex        =   43
         Top             =   225
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtInvTotal 
         DataField       =   "TotalCost"
         DataSource      =   "datInvoice"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1875
         TabIndex        =   42
         Top             =   1725
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtRental6 
         DataField       =   "Rental6"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   41
         Top             =   2100
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtRental7 
         DataField       =   "Rental7"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   40
         Top             =   2475
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtRental8 
         DataField       =   "Rental8"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   39
         Top             =   2850
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtRental9 
         DataField       =   "Rental9"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   38
         Top             =   3225
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtRental10 
         DataField       =   "Rental10"
         DataSource      =   "datInvoice"
         Height          =   285
         Left            =   75
         TabIndex        =   37
         Top             =   3600
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtAfter 
         DataField       =   "TotalCost"
         DataSource      =   "datInvoice"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1875
         TabIndex        =   36
         Top             =   2100
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtBefore 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1875
         TabIndex        =   35
         Top             =   2475
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtInvSub 
         DataField       =   "Subtotal"
         DataSource      =   "datInvoice"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1875
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtInvGST 
         DataField       =   "GST"
         DataSource      =   "datInvoice"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1875
         TabIndex        =   33
         Top             =   975
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtInvPST 
         DataField       =   "PST"
         DataSource      =   "datInvoice"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1875
         TabIndex        =   32
         Top             =   1350
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Data datRental 
      Caption         =   "Rentals"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4425
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2235
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Data datConfig 
      Caption         =   "Config"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   9540
      TabIndex        =   8
      Top             =   8385
      Width           =   915
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   315
      Left            =   8550
      TabIndex        =   6
      Top             =   7935
      Width           =   915
   End
   Begin VB.CommandButton cmdComplete 
      Caption         =   "Complete"
      Height          =   315
      Left            =   9540
      TabIndex        =   7
      Top             =   7935
      Width           =   915
   End
   Begin VB.Frame fraTotal 
      BackColor       =   &H00FF8080&
      Caption         =   "Totals..."
      Height          =   1905
      Left            =   7020
      TabIndex        =   18
      Top             =   5685
      Width           =   3465
      Begin MSMask.MaskEdBox mskTotal 
         Height          =   315
         Left            =   2025
         TabIndex        =   27
         Top             =   1500
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRented 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2025
         TabIndex        =   23
         Top             =   300
         Width           =   1290
      End
      Begin MSMask.MaskEdBox mskSub 
         Height          =   315
         Left            =   2025
         TabIndex        =   28
         Top             =   600
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskGST 
         Height          =   315
         Left            =   2025
         TabIndex        =   29
         Top             =   900
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPST 
         Height          =   315
         Left            =   2025
         TabIndex        =   30
         Top             =   1200
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00FF8080&
         Caption         =   "Sub-Total:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   26
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00FF8080&
         Caption         =   "Total:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   150
         TabIndex        =   22
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label lblPST 
         BackColor       =   &H00FF8080&
         Caption         =   "PST:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   21
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblGST 
         BackColor       =   &H00FF8080&
         Caption         =   "GST:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   150
         TabIndex        =   20
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblRented 
         BackColor       =   &H00FF8080&
         Caption         =   "Total Movies Rented:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   150
         TabIndex        =   19
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.Data datInvoice 
      Caption         =   "Invoice"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2220
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame fraList 
      BackColor       =   &H00FF8080&
      Caption         =   "Rental List..."
      Height          =   2340
      Left            =   7020
      TabIndex        =   16
      Top             =   2910
      Width           =   3465
      Begin VB.ListBox lstRental 
         Enabled         =   0   'False
         Height          =   2010
         Left            =   75
         TabIndex        =   17
         Top             =   225
         Width           =   3315
      End
   End
   Begin VB.Data datMovie 
      Caption         =   "Movies"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Data datCust 
      Caption         =   "Customers"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2220
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Frame fraMovie 
      BackColor       =   &H00FF8080&
      Caption         =   "Movie Search..."
      Height          =   1965
      Left            =   5880
      TabIndex        =   9
      Top             =   150
      Width           =   4590
      Begin VB.TextBox txtBarcode 
         DataField       =   "Barcode"
         DataSource      =   "datMovie"
         Height          =   285
         Left            =   1200
         TabIndex        =   104
         Top             =   1425
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtType 
         DataField       =   "Type"
         DataSource      =   "datMovie"
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Top             =   1035
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   3525
         TabIndex        =   5
         Top             =   1500
         Width           =   915
      End
      Begin VB.TextBox txtTitle 
         DataField       =   "Title"
         DataSource      =   "datMovie"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         Top             =   675
         Width           =   3240
      End
      Begin VB.CommandButton cmdFindBar 
         Caption         =   "Find"
         Height          =   315
         Left            =   3525
         TabIndex        =   4
         Top             =   225
         Width           =   915
      End
      Begin VB.TextBox txtFindBar 
         Height          =   285
         Left            =   75
         TabIndex        =   3
         Top             =   225
         Width           =   3315
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FF8080&
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
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   750
         Width           =   1440
      End
   End
   Begin VB.Frame fraCust 
      BackColor       =   &H00FF8080&
      Caption         =   "Customer Info..."
      ForeColor       =   &H00000000&
      Height          =   1965
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5490
      Begin VB.TextBox txtAccount 
         DataField       =   "CustAccount"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1725
         TabIndex        =   24
         Top             =   1425
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtPhone 
         DataField       =   "Phone"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1725
         TabIndex        =   13
         Top             =   1050
         Width           =   3615
      End
      Begin VB.TextBox txtName 
         DataField       =   "CustName"
         DataSource      =   "datCust"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1725
         TabIndex        =   12
         Top             =   675
         Width           =   3615
      End
      Begin VB.CommandButton cmdFindCust 
         Caption         =   "Find"
         Height          =   315
         Left            =   4425
         TabIndex        =   2
         Top             =   225
         Width           =   915
      End
      Begin VB.TextBox txtFindCust 
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   4140
      End
      Begin VB.Label lblPhone 
         BackColor       =   &H00FF8080&
         Caption         =   "Phone:"
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
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FF8080&
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
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   675
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dCurrentDate As Date
Dim dDueDate As Date


Private Sub cmdAdd_Click()

If txtTitle.Text = "" Then
MsgBox "You must select a movie", vbExclamation, "No Movie Selected"
txtFindBar.SetFocus
Exit Sub
End If

CheckDueDate
CountRentals
End Sub

Private Sub cmdClose_Click()
    If txtrental1.Text = "" Then
    ClearAllEntries
    Unload Me
    Exit Sub
    End If
    datInvoice.Recordset.CancelUpdate
    ClearAllEntries
    Unload Me
End Sub

Private Sub cmdComplete_Click()
Dim sTotal As String

sTotal = mskTotal.Text

txtInvAccount.Text = txtAccount.Text
txtInvMovie.Text = txtRented.Text
txtInvSub.Text = mskSub.Text
txtInvGST.Text = mskGST.Text
txtInvPST.Text = mskPST.Text
txtInvTotal.Text = mskTotal.Text
txtInvName.Text = txtName.Text
txtInvPhone.Text = txtPhone.Text

If txtInvAccount.Text = "" Then
MsgBox "No account was selected", vbExclamation, "Account Nuber Needed"
txtFindCust.SetFocus
Exit Sub
End If

If txtrental1.Text = "" Then
MsgBox "You have not selected any movies to rent.", vbExclamation, "Movie selection needed"
txtFindBar.SetFocus
Exit Sub
End If

datInvoice.Recordset.Update
datInvoice.Refresh
'MsgBox "The total of this purchase is " & sTotal & " ", vbExclamation, "Total"

RentalComplete
ClearAllEntries
txtFindCust.SetFocus

Load frmReport
End Sub

Private Sub cmdFindBar_Click()
Dim sName As String
Dim sFind As String

If txtFindBar.Text = "" Then
MsgBox "Barcode needed to complete search", vbExclamation, "Barcode Needed"
txtFindBar.SetFocus
Exit Sub
End If

datMovie.RecordSource = ("Movies")
datMovie.Refresh

sName = txtFindBar.Text
sFind = "barcode LIKE '" & sName & "*'"

datMovie.Recordset.FindFirst sFind

If datMovie.Recordset.NoMatch Then
MsgBox "Movie was not found in database", vbExclamation, "No Match"
datMovie.RecordSource = ("Dummy")
datMovie.Refresh
End If
End Sub

Private Sub cmdFindCust_Click()
Dim sName As String
Dim sFind As String

If txtFindCust.Text = "" Then
MsgBox "Customer Account Number Needed", vbExclamation, "Account Number Needed"
txtFindCust.SetFocus
Exit Sub
End If

datCust.RecordSource = ("Accounts")
datCust.Refresh

sName = txtFindCust.Text
sFind = "CustAccount LIKE '" & sName & "*'"

datCust.Recordset.FindFirst sFind

If datCust.Recordset.NoMatch Then
MsgBox "Customer account not found", vbExclamation, "No Match"
datCust.RecordSource = ("Dummy")
datCust.Refresh
End If

End Sub

Private Sub cmdReset_Click()
If txtrental1.Text = "" Then
ClearAllEntries
Exit Sub
End If
datInvoice.Recordset.CancelUpdate
datInvoice.Refresh
ClearAllEntries

txtFindCust.SetFocus
End Sub

Private Sub Form_Load()
frmSystem.Top = 0
frmSystem.Left = 0

Dim sDate As String

sDate = Date

txtRentDay = sDate
dCurrentDate = sDate
txtNow = sDate

datCust.DatabaseName = (App.Path & "\movieDB.mdb")

datMovie.DatabaseName = (App.Path & "\movieDB.mdb")

datInvoice.DatabaseName = (App.Path & "\movieDB.mdb")

datConfig.DatabaseName = (App.Path & "\movieDB.mdb")
datConfig.RecordSource = ("Config")

datRental.DatabaseName = (App.Path & "\movieDB.mdb")
datRental.RecordSource = ("Rentals")

ClearAllEntries
End Sub

Private Sub ClearAllEntries()
datCust.RecordSource = ("Dummy")
datMovie.RecordSource = ("Dummy")
datInvoice.RecordSource = ("Dummy")
datMovie.Refresh
datCust.Refresh
datInvoice.Refresh

lstRental.Clear
txtrental1.Text = ""
txtrental2.Text = ""
txtrental3.Text = ""
txtrental4.Text = ""
txtrental5.Text = ""
txtRental6.Text = ""
txtRental7.Text = ""
txtRental8.Text = ""
txtRental9.Text = ""
txtRental10.Text = ""
txtInvAccount.Text = ""
txtInvMovie.Text = ""
txtInvTotal.Text = ""
txtFindCust.Text = ""
txtFindBar.Text = ""
txtRented.Text = "0"
mskSub.Text = "0.00"
mskGST.Text = "0.00"
mskPST.Text = "0.00"
mskTotal.Text = "0.00"
End Sub

Private Sub CountMe()

If txtrental2.Text = "" Then
txtRented = 1
Exit Sub
ElseIf txtrental3.Text = "" Then
txtRented = 2
Exit Sub
ElseIf txtrental4.Text = "" Then
txtRented = 3
Exit Sub
ElseIf txtrental5.Text = "" Then
txtRented = 4
Exit Sub
ElseIf txtRental6.Text = "" Then
txtRented = 5
Exit Sub
ElseIf txtRental7.Text = "" Then
txtRented = 6
Exit Sub
ElseIf txtRental8.Text = "" Then
txtRented = 7
Exit Sub
ElseIf txtRental9.Text = "" Then
txtRented = 8
Exit Sub
ElseIf txtRental10.Text = "" Then
txtRented = 9
Exit Sub
End If

txtRented = 10
End Sub


Private Sub CalculateRental()
Dim c As Currency
Dim g As Currency
Dim p As Currency
Dim t As Currency
Dim x As Currency
'Dim i As Currency
Dim s As Currency
Dim b As Currency

If txtType = "Kids" Then
c = txtKids.Text
End If

If txtType = "New" Then
c = txtNew.Text
End If

If txtType = "Old" Then
c = txtOld.Text
End If

If txtType = "Other" Then
c = txtSpec.Text
End If

g = txtCfgGST.Text
p = txtCfgPST.Text
s = mskSub.Text
txtBefore.Text = s


t = g + p
x = 1 + t


b = s + c
mskSub.Text = "$" & b & ""
mskGST.Text = "$" & b * g & ""
mskPST.Text = "$" & b * p & ""
txtAfter.Text = b * x
mskTotal.Text = "$" & txtAfter.Text & ""
txtInvMovie.Text = txtRented.Text
txtInvTotal.Text = mskTotal.Text
End Sub

Private Sub CountRentals()
Dim sMovie As String

sMovie = txtTitle.Text

If txtrental1.Text = "" Then
datInvoice.RecordSource = ("Invoice")
datInvoice.Refresh
datInvoice.Recordset.AddNew
txtrental1.Text = txtTitle.Text
txtTitle1.Text = txtTitle.Text
txtBarcode1.Text = txtBarcode.Text
txtInvReport.Text = txtEntryID.Text
txtType1.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
ElseIf txtrental2.Text = "" Then
txtrental2.Text = txtTitle.Text
txtTitle2.Text = txtTitle.Text
txtBarcode2.Text = txtBarcode.Text
txtType2.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
ElseIf txtrental3.Text = "" Then
txtrental3.Text = txtTitle.Text
txtTitle3.Text = txtTitle.Text
txtBarcode3.Text = txtBarcode.Text
txtType3.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
ElseIf txtrental4.Text = "" Then
txtrental4.Text = txtTitle.Text
txtTitle4.Text = txtTitle.Text
txtBarcode4.Text = txtBarcode.Text
txtType4.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
ElseIf txtrental5.Text = "" Then
txtrental5.Text = txtTitle.Text
txtTitle5.Text = txtTitle.Text
txtBarcode5.Text = txtBarcode.Text
txtType5.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
ElseIf txtRental6.Text = "" Then
txtRental6.Text = txtTitle.Text
txtTitle6.Text = txtTitle.Text
txtBarcode6.Text = txtBarcode.Text
txtType6.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
ElseIf txtRental7.Text = "" Then
txtRental7.Text = txtTitle.Text
txtTitle7.Text = txtTitle.Text
txtBarcode7.Text = txtBarcode.Text
txtType7.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
ElseIf txtRental8.Text = "" Then
txtRental8.Text = txtTitle.Text
txtTitle8.Text = txtTitle.Text
txtBarcode8.Text = txtBarcode.Text
txtType8.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
ElseIf txtRental9.Text = "" Then
txtRental9.Text = txtTitle.Text
txtTitle9.Text = txtTitle.Text
txtBarcode9.Text = txtBarcode.Text
txtType9.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
ElseIf txtRental10.Text = "" Then
txtRental10.Text = txtTitle.Text
txtTitle10.Text = txtTitle.Text
txtBarcode10.Text = txtBarcode.Text
txtType10.Text = txtType.Text
lstRental.AddItem sMovie
CountMe
CalculateRental
Exit Sub
End If

MsgBox "You have reached your 10 movie limit.", vbExclamation, "Limit Exceeded"
End Sub

Private Sub CheckDueDate()
Dim sDue As String
Dim r As String

GetDueDate "d"

r = dDueDate
sDue = r

If txtMovie1.Text = "" Then
txtMovie1.Text = sDue
Exit Sub
ElseIf txtMovie2.Text = "" Then
txtMovie2.Text = sDue
Exit Sub
ElseIf txtMovie3.Text = "" Then
txtMovie3.Text = sDue
Exit Sub
ElseIf txtMovie4.Text = "" Then
txtMovie4.Text = sDue
Exit Sub
ElseIf txtMovie5.Text = "" Then
txtMovie5.Text = sDue
Exit Sub
ElseIf txtMovie6.Text = "" Then
txtMovie6.Text = sDue
Exit Sub
ElseIf txtMovie7.Text = "" Then
txtMovie7.Text = sDue
Exit Sub
ElseIf txtMovie8.Text = "" Then
txtMovie8.Text = sDue
Exit Sub
ElseIf txtMovie9.Text = "" Then
txtMovie9.Text = sDue
Exit Sub
ElseIf txtMovie10.Text = "" Then
txtMovie10.Text = sDue
Exit Sub
End If

End Sub

Private Sub RentalComplete()
If txtMovie1.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle1.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie1.Text
txtRentBar.Text = txtBarcode1.Text
txtRentType.Text = txtType1.Text
datRental.Recordset.Update
End If

If txtMovie2.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle2.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie2.Text
txtRentBar.Text = txtBarcode2.Text
txtRentType.Text = txtType2.Text
datRental.Recordset.Update
End If

If txtMovie3.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle3.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie3.Text
txtRentBar.Text = txtBarcode3.Text
txtRentType.Text = txtType3.Text
datRental.Recordset.Update
End If

If txtMovie4.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle4.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie4.Text
txtRentBar.Text = txtBarcode4.Text
txtRentType.Text = txtType4.Text
datRental.Recordset.Update
End If

If txtMovie5.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle5.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie5.Text
txtRentBar.Text = txtBarcode5.Text
txtRentType.Text = txtType5.Text
datRental.Recordset.Update
End If

If txtMovie6.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle6.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie6.Text
txtRentBar.Text = txtBarcode6.Text
txtRentType.Text = txtType6.Text
datRental.Recordset.Update
End If

If txtMovie7.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle7.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie7.Text
txtRentBar.Text = txtBarcode7.Text
txtRentType.Text = txtType7.Text
datRental.Recordset.Update
End If

If txtMovie8.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle8.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie8.Text
txtRentBar.Text = txtBarcode8.Text
txtRentType.Text = txtType8.Text
datRental.Recordset.Update
End If

If txtMovie9.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle9.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie9.Text
txtRentBar.Text = txtBarcode9.Text
txtRentType.Text = txtType9.Text
datRental.Recordset.Update
End If

If txtMovie10.Text = "" Then
Exit Sub
Else
datRental.Recordset.AddNew
txtRentMe.Text = txtTitle10.Text
txtRentAccount.Text = txtAccount.Text
txtRentName.Text = txtName.Text
txtRentPhone.Text = txtPhone.Text
txtRentDate.Text = txtNow.Text
txtReturnMe.Text = txtMovie10.Text
txtRentBar.Text = txtBarcode10.Text
txtRentType.Text = txtType10.Text
datRental.Recordset.Update
End If
End Sub

Private Sub GetDueDate(ByVal sInterval As String)
Dim i As Integer

If txtType.Text = "New" Then
i = txtNewDay.Text
dDueDate = DateAdd(sInterval, i, dCurrentDate)
ElseIf txtType.Text = "Old" Then
i = txtOldDay.Text
dDueDate = DateAdd(sInterval, i, dCurrentDate)
ElseIf txtType.Text = "Kids" Then
i = txtKidsDay.Text
dDueDate = DateAdd(sInterval, i, dCurrentDate)
ElseIf txtType.Text = "Other" Then
i = txtAltDay.Text
dDueDate = DateAdd(sInterval, i, dCurrentDate)
End If

End Sub
