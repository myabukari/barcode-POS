VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmDebtors 
   BackColor       =   &H00C29E21&
   Caption         =   "Debtors Payments"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13140
   Icon            =   "frmTransactions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   13140
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgAccounts 
      Height          =   1815
      Left            =   240
      TabIndex        =   39
      Top             =   1920
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   12754465
      ForeColor       =   14869218
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorBkg    =   15522944
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
      FormatString    =   $"frmTransactions.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComctlLib.ListView LstInvoice 
      Height          =   7095
      Left            =   9480
      TabIndex        =   38
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   12515
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   14802911
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "InvoiceNumber"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TransactionDescription"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "AmountDue"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AccID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "BalCD"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   240
      TabIndex        =   32
      Top             =   7800
      Width           =   12735
      Begin VB.Frame Frame5 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   600
         TabIndex        =   33
         Top             =   0
         Width           =   9975
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Save"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   12754465
            Gradient        =   1
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmTransactions.frx":0395
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdDelete 
            Height          =   375
            Left            =   1800
            TabIndex        =   50
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Delete"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   12754465
            Gradient        =   1
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmTransactions.frx":07E7
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdSave1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelete1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "D&elete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   1335
         End
         Begin lvButton_H.lvButtons_H cmdFind 
            Height          =   375
            Left            =   3480
            TabIndex        =   51
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Find"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   12754465
            Gradient        =   1
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmTransactions.frx":0981
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   5160
            TabIndex        =   52
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Clear"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   12754465
            Gradient        =   1
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmTransactions.frx":0DD3
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdFind1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Find"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdClear1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   240
            Width           =   1335
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   8280
            TabIndex        =   53
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "E&xit"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   12754465
            Gradient        =   1
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmTransactions.frx":2684
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdReport 
            Height          =   375
            Left            =   6720
            TabIndex        =   54
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Report"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   12754465
            Gradient        =   1
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmTransactions.frx":2AD6
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdReport1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdExit1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8280
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.TextBox txtInvoiceNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F0D1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtPhoneNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F0D1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   27
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtCreditLimit 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F0D1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   25
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F0D1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Caption         =   "Payments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4695
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   12735
      Begin VB.TextBox txtArrears 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10320
         TabIndex        =   44
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtBalCD 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         TabIndex        =   41
         Top             =   720
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C29E21&
         Caption         =   "PAYMENTS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   8895
         Begin MSComctlLib.ListView LstPayments 
            Height          =   2295
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4048
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   14802911
            BackColor       =   0
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PaymentDate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PaymentMode"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ChequeNumber"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "AmountPaid"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Arrears"
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin VB.TextBox txtBalBD 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         TabIndex        =   10
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtAmountPaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox cboPaymentMode 
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTransactions.frx":2DF0
         Left            =   1680
         List            =   "frmTransactions.frx":2DFD
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtChequeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpPaymentDate 
         Height          =   315
         Left            =   6480
         TabIndex        =   9
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14336310
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22020099
         CurrentDate     =   39115
      End
      Begin MSComCtl2.DTPicker dtpChequeDueDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   47
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14336310
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22020099
         CurrentDate     =   39115
      End
      Begin VB.Label lblchequeDate 
         BackColor       =   &H00C29E21&
         Caption         =   "ChequeDue Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Arrears:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9360
         TabIndex        =   45
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C29E21&
         Caption         =   "Bal c/d:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   40
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C29E21&
         Caption         =   "Bal b/d:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C29E21&
         Caption         =   "PaymentDate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C29E21&
         Caption         =   "AmountPaid:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblChequeNo 
         BackColor       =   &H00C29E21&
         Caption         =   "ChequeNo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C29E21&
         Caption         =   "PaymentMode:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   9135
      Begin VB.TextBox txtAmountDue 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpTransTime 
         Height          =   315
         Left            =   6480
         TabIndex        =   5
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14336310
         Format          =   22020098
         CurrentDate     =   39115
      End
      Begin MSComCtl2.DTPicker dtpTransDate 
         Height          =   315
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14336310
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22020099
         CurrentDate     =   39115
      End
      Begin VB.ComboBox cboTransDescription 
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTransactions.frx":2E1F
         Left            =   1680
         List            =   "frmTransactions.frx":2E29
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C29E21&
         Caption         =   "TransactionTime:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C29E21&
         Caption         =   "TransactionDate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C29E21&
         Caption         =   "AmountDue:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Transaction Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame framAllInvoice 
      BackColor       =   &H00C29E21&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4455
      Left            =   9360
      TabIndex        =   43
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox txtAllInvoice 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9360
      TabIndex        =   42
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C29E21&
      Caption         =   "InvoiceNo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C29E21&
      Caption         =   "PhoneNumber:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C29E21&
      Caption         =   "CreditLevel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C29E21&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C29E21&
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   240
      TabIndex        =   23
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "frmDebtors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Accountid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer
Dim BalCD As Double, AccTransid As String, xflag As Boolean, AccountTransid As String

Private Sub cboPaymentMode_Click()
If Me.cboPaymentMode = "Cheque" Then
Me.txtChequeNo.Enabled = True
Me.dtpChequeDueDate.Enabled = True
Me.lblchequeDate.Enabled = True
Me.lblChequeNo.Enabled = True
End If
End Sub

Private Sub cboPaymentMode_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtChequeNo.SetFocus
End If
End Sub

Private Sub cboPaymentMode_LostFocus()
If Me.cboPaymentMode = "" And Me.txtAmountPaid <> "" Then
MsgBox "PLEASE SELECT PAYMENTMODE", vbInformation, "PAYMENTMODE"
Me.cboPaymentMode.SetFocus
End If
End Sub

Private Sub cboTransDescription_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtAmountDue.SetFocus
End If
End Sub

Private Sub cmdClear_Click()
  Call ClearCtrls
  Me.LstInvoice.ListItems.Clear
  Me.LstPayments.ListItems.Clear
  Me.flxgAccounts.Visible = False
  Me.txtName.SetFocus
End Sub

Private Sub cmdDelete_Click()
If Me.txtInvoiceNo = "" Then
MsgBox "There is Nothing to be Deleted", vbInformation, "Select InvoiceNo to Delete"
Exit Sub
End If

If MsgBox("ARE YOU SURE  YOU WANT TO DELETE THIS INVOICE NUMBER'S DETAILS?", vbYesNo + vbQuestion, "CONFIRM DELETE") = vbYes Then
   
   Me.cmdDelete.Enabled = False
   
   On Error GoTo SaveError
       
   'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      Me.MousePointer = vbDefault
      Me.cmdDelete.Enabled = True
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   
   cn.BeginTrans
   
    cn.Execute "Delete From  AccountTransaction  Where AccID ='" & Accountid & "' And AccTransID='" & AccTransid & "'", Y
    If Y > 0 Then
    cn.Execute "Delete From  TransPayments  Where AccTransID ='" & AccTransid & "'", Y
    End If
    If Y > 0 Then
    cn.CommitTrans
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     
     ClearCtrls
     Call ListInvoice
     Call ListPayments
   Else
      cn.RollbackTrans
      MsgBox "Sorry, Unable to Delete Products Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
    
   sflag = False
   Me.cmdDelete.Enabled = False
   If cn.State = 1 Then cn.Close
   Me.flxgAccounts.Visible = False
   Me.MousePointer = vbDefault
   
  
   Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Delete Products  Details:Please Try Again!", vbInformation, "Delete Failed"
        Exit Sub
End If
End Sub

Private Sub cmdExit_Click()
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdReport_Click()
frmAccountsRpt.Show
End Sub

Private Sub cmdSave_Click()

Me.txtChequeNo.Enabled = False
Me.dtpChequeDueDate.Enabled = False
Me.lblchequeDate.Enabled = False
Me.lblChequeNo.Enabled = False

If Trim(Me.txtName) = "" Then
   MsgBox "YOU MUST ENTER ACCOUNT NAME.", vbInformation, "ACCOUNT NAME"
   Me.txtName.SetFocus: Exit Sub
End If

If Trim(Me.txtInvoiceNo) = "" Then
   MsgBox "YOU MUST ENTER INVOICE NUMBER.", vbInformation, "ENTER INVOICE NUMBER"
   Me.txtInvoiceNo.SetFocus: Exit Sub
End If


If Trim(Me.cboTransDescription) = "" Then
   MsgBox "YOU MUST ENTER TRANSACTION DESCRIPTION.", vbInformation, "DESCRIPTION"
   Me.cboTransDescription.SetFocus: Exit Sub
End If


If Trim(Me.txtAmountDue) = "" Then
   MsgBox "YOU MUST ENTER AMOUNTDUE OF GOODS ON CREDIT.", vbInformation, "AMOUNTDUE"
   Me.txtAmountDue.SetFocus: Exit Sub
End If


If Val(Me.txtBalCD) = 0 Then
      MsgBox "THE DEBTOR FOR THIS PARTICULAR INVOICE HAS FINISHED PAYING.", vbInformation, "DEBTOR CLEARED"
     Exit Sub
End If

 If Me.txtAmountPaid <> "" And (Val(Me.txtAmountPaid) > Val(Me.txtBalCD)) Then
     MsgBox "AMOUNT PAID SHOULD NOT BE MORE THAN BAL C/D", vbInformation, "LIABILITY"
     Me.txtAmountPaid = "": Me.txtAmountPaid.SetFocus
     Exit Sub
   End If
  


'If Trim(Me.txtAmountPaid) = "" Then
   'MsgBox "YOU MUST ENTER AMOUNTPAID.", vbInformation, "AMOUNTPAID"
   'Me.txtAmountPaid.SetFocus: Exit Sub
'End If



On Error GoTo SaveError
Me.cmdSave.Enabled = False

'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If
'Me.txtStock = Trim(Me.txtTotalQuantity)
'If sflag = False Then
   'save part
'rs.Open "Select ProductName From Products Where ProductName ='" & Trim(Me.txtName) & "'", cn, adOpenForwardOnly, adLockReadOnly
'If rs.RecordCount > 0 Then
   'rs.Close: cn.Close
   'MsgBox "A ProductName Has Already Been Setup with the Name.", vbInformation
   'Me.MousePointer = vbDefault
   'Me.cmdSave.Enabled = True
   'Me.txtName.SetFocus: Exit Sub
   
   'Me.MousePointer = vbDefault
   'Me.cmdSave.Enabled = True
'End If
'rs.Close

   'Call Generate_ProductID(Productid)
   Call Generate_AccountTransID(AccountTransid)
   Call AccountCal
   rs.Open "Select * From AccountTransaction Where InvoiceNo='" & Trim(Me.txtInvoiceNo) & "' And AccID='" & Accountid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount = 0 Then
     
     cn.BeginTrans
     cn.Execute "Insert Into AccountTransaction  ([AccID],[InvoiceNo],[TransactionDescription],[AmountDue],[Date],[BalanceCD],[AccTransID]) select '" & (Accountid) & "','" & Trim(Me.txtInvoiceNo.Text) & "','" & Trim(Me.cboTransDescription.Text) & "','" & Val(Me.txtAmountDue.Text) & "','" & Trim(Me.dtpTransDate) & "','" & Val(Me.txtBalCD.Text) & "','" & (AccountTransid) & "'", Y
   
     If Y > 0 Then
     If Me.cboPaymentMode <> "" And Me.txtAmountPaid <> "" Then
     cn.Execute "Insert Into TransPayments ([AccTransID],[PaymentMode],[ChequeNo],[AmountPaid],[Arrears],[PaymentDate],[ChequeDueDate]) select '" & (AccountTransid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Val(Me.txtArrears.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
     Else
     cn.Execute "Insert Into TransPayments ([AccTransID],[PaymentMode],[ChequeNo],[AmountPaid],[Arrears],[PaymentDate],[ChequeDueDate]) select '" & (AccountTransid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Val(Me.txtArrears.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
     End If
     End If
     If Y > 0 Then
     cn.CommitTrans
     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
       If rs.State = 1 Then rs.Close
       
       
       Call ListPayments
       Call ListInvoice
       Call ClearCtrls
     'Me.flxgAccounts.Visible = False
     Me.txtName.SetFocus
     Else
     cn.RollbackTrans
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     If rs.State = 1 Then rs.Close
     Me.txtName.SetFocus
     End If
     
 Else
     'If Me.cboPaymentMode <> "" And Me.txtAmountPaid <> "" Then
     'cn.Execute "Insert Into TransPayments ([AccTransID],[PaymentMode],[ChequeNo],[AmountPaid],[Arrears],[PaymentDate],[ChequeDueDate]) select '" & (AccTransid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Val(Me.txtArrears.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
     'If Y > 0 Then
      ' MsgBox "Saved Successfully!", vbInformation, "Save Successful"
       
       If rs.State = 1 Then rs.Close
       Call ListPayments
       Call ListInvoice
       Call ClearCtrls
       'Me.flxgAccounts.Visible = False
       
       'End If
     'If rs.State = 1 Then rs.Close
     'End If
    If rs.State = 1 Then rs.Close
  End If
'Else
'edit part

  ' rs.Open "Select ProductName From Products Where ProductName ='" & Trim(Me.txtName) & "' and ProductID<>'" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
   'If rs.RecordCount > 0 Then
      'rs.Close: cn.Close
      'MsgBox "A Product Has Already Been Setup with the Name.", vbInformation
      'Me.txtName.SetFocus: Exit Sub
  ' End If
   'rs.Close
  '
  ' rs.Open "Select ProductInventory.StockLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where  Products.ProductID='" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
   'If rs.RecordCount > 0 Then
      'StockQty = Val(Trim(Me.txtTotalQuantity)) + rs.Fields("StockLevel")
      'rs.Close
   'End If
  
   
  ' cn.BeginTrans
  ' cn.Execute "Update Products Set ProductName ='" & Trim(Me.txtName.Text) & "',BaseUnit='" & Val(Trim(Me.txtbaseunit.Text)) & "',UnitPrice='" & Trim(Me.txtUnitPrice.Text) & "',TotalQuantity='" & Val(Trim(Me.txtTotalQuantity.Text)) & "',PricePerCartone='" & Val(Trim(Me.txtbasePrice.Text)) & "',Discount='" & Val(Trim(Me.txtDiscount.Text)) & "',Manufacturer='" & Trim(Me.txtManufacterer.Text) & "',ExpiryDate='" & Trim(Me.dtpExp) & "',ManufucteryDate='" & Trim(Me.dtpmanu) & "',NoOfCartones='" & Val(Trim(Me.txtNoCartones.Text)) & "' Where ProductID ='" & Productid & "'", Y
  '     If Y > 0 Then
  'cn.Execute "Update ProductInventory Set StockLevel ='" & StockQty & "',ReorderLevel='" & Val(Trim(Me.txtReorder.Text)) & "' Where ProductID ='" & Productid & "'", Y
  '     End If
         
   'If Y > 0 Then
      'cn.CommitTrans
     ' MsgBox "Edit Successful!", vbInformation, "Edit Successful"
     'Clear ctrls and setfocus to Clinic ctrl
      'Call ListProducts
      'Call ClearCtrls
     ' Me.txtName = "": Me.txtbasePrice = "": Me.txtbaseunit = "": Me.txtDiscount = ""
      'Me.txtManufacterer = "": Me.txtNoCartones = "": Me.txtReorder = "": Me.txtStock = ""
     ' Me.txtTotalQuantity = "": Me.txtUnitPrice = ""
     ' Me.txtName.SetFocus
  ' Else
     ' MsgBox "Sorry, Unable to Edit Product Details:Please Try Again!", vbInformation, "Edit Failed"
 '  End If

' End If

'flag = False
'If cn.State = 1 Then cn.Close
'If rs.State = 1 Then rs.Close

'Me.MousePointer = vbDefault
'Me.cmdSave.Enabled = True


Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub

Private Sub flxgAccounts_Click()
Me.txtName = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 0)
Me.txtCreditLimit = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 1)
Me.txtPhoneNo = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 2)
Accountid = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 3)
Me.flxgAccounts.Visible = False
Me.txtInvoiceNo = ""
Me.cmdSave.Enabled = True
sflag = False
oflag = False
Call ListInvoice
Me.framAllInvoice.Caption = Me.txtName & " " & "Invoices"
Me.txtInvoiceNo.SetFocus
End Sub

Private Sub flxgAccounts_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtName = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 0)
Me.txtCreditLimit = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 1)
Me.txtPhoneNo = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 2)
Accountid = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 3)
Me.flxgAccounts.Visible = False

Me.cmdSave.Enabled = True
sflag = False
oflag = False
Call ListInvoice
Me.framAllInvoice.Caption = Me.txtName & " " & "Invoices"
Me.txtInvoiceNo.SetFocus
End If
End Sub

Private Sub Form_Load()
Call ListInvoice
Me.txtChequeNo.Enabled = False
Me.dtpChequeDueDate.Enabled = False
Me.lblchequeDate.Enabled = False
Me.lblChequeNo.Enabled = False
Me.dtpTransDate = Date
Me.dtpTransTime = Date
Me.dtpPaymentDate = Date
  Me.Height = 9330
  Me.Width = 13260
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub LstInvoice_Click()
If Me.txtName <> "" Then

End If
On Error GoTo SaveError
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)
xflag = True
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If
rs.Open "Select * From AccountTransaction Inner Join AccountHolders On AccountTransaction.AccID=AccountHolders.AccID", cn, adOpenForwardOnly, adLockReadOnly


    While Not rs.EOF
        If Me.LstInvoice.SelectedItem.Text = rs("InvoiceNo") Then
                Me.txtInvoiceNo.Text = Trim(rs("InvoiceNo"))
                Me.cboTransDescription.Text = Trim(rs("TransactionDescription"))
                Me.txtAmountDue.Text = Trim(rs("AmountDue"))
                Me.dtpTransDate = Trim(rs("Date"))
                Me.txtBalCD.Text = Trim(rs("BalanceCD"))
                Me.txtBalBD.Text = Trim(rs("BalanceCD"))
                Me.txtName.Text = Trim(rs("Name"))
                Me.txtPhoneNo.Text = Trim(rs("PhoneNumber"))
                Me.txtCreditLimit.Text = Trim(rs("CreditLimit"))
                AccTransid = Trim(rs("AccTransID"))
            rs.MoveNext
        Else
            rs.MoveNext
        End If
    Wend
    rs.Close
    Call ListPayments
    Set rs = Nothing
    sflag = True
    Me.cmdDelete.Enabled = True
    Me.cmdSave.Enabled = True
    xflag = False
    
    
   If Val(Me.txtBalCD) = 0 Then
      MsgBox "THE DEBTOR FOR THIS PARTICULAR INVOICE HAS FINISHED PAYING.", vbInformation, "DEBTOR CLEARED"
    
   End If


    Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
      xflag = False
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "TRY AGAIN"
     Me.txtName.SetFocus
     Exit Sub
End Sub

Private Sub LstPayments_Click()
Me.cmdSave.Enabled = True
End Sub

Private Sub txtAmountDue_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789/|\;:.,()@#$%&*_-'"
If KeyAscii = vbKeyReturn Then
   Me.dtpTransDate.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtAmountDue_LostFocus()
If Me.txtAmountDue <> "" And Me.txtBalCD = "" Then
Me.txtBalCD = Val(Me.txtAmountDue)
End If
End Sub

Private Sub txtAmountPaid_GotFocus()
Me.cmdSave.Enabled = True
If Me.txtAmountDue <> "" And Me.txtBalCD = "" Then
Me.txtBalCD = Val(Me.txtAmountDue)
End If
End Sub

Private Sub txtAmountPaid_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789/|\;:.,()@#$%&*_-'"
If KeyAscii = vbKeyReturn Then
   Me.dtpPaymentDate.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtBalance_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789/|\;:.,()@#$%&*_-'"
If KeyAscii = vbKeyReturn Then
   Me.cmdSave.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
 End If

End Sub

Private Sub txtAmountPaid_LostFocus()
If Me.cboPaymentMode = "" And Me.txtAmountPaid <> "" Then
MsgBox "PLEASE SELECT PAYMENTMODE", vbInformation, "PAYMENTMODE"
Me.cboPaymentMode.SetFocus: Exit Sub
End If
Me.cmdSave.Enabled = True

End Sub

Private Sub txtChequeNo_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtAmountPaid.SetFocus
End If
End Sub

Private Sub txtCreditLimit_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789/|\;:.,()@#$%&*_-'"
If KeyAscii = vbKeyReturn Then
   Me.cmdSave.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub txtInvoiceNo_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.cboTransDescription.SetFocus
End If
End Sub

Private Sub txtInvoiceNo_LostFocus()

On Error GoTo OkError
'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If

rs.Open "Select * From AccountTransaction Where InvoiceNo='" & (Me.txtInvoiceNo) & "' And AccID='" & Accountid & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount = 0 Then
   Me.cboTransDescription = ""
   Me.txtAmountDue = ""
   Me.dtpTransDate = Date
   Me.cboPaymentMode = ""
   Me.txtChequeNo = ""
   Me.dtpPaymentDate = Date
   Me.txtBalBD = ""
   Me.txtBalCD = ""
   Me.txtArrears = ""
   If rs.State = 1 Then rs.Close
   
    Me.LstPayments.ListItems.Clear
 Else
   MsgBox "THE INVOICE NUMBER ENTERED FOR" & " " & Me.txtName & " " & "ALREADY EXIST", vbInformation, "SELECT INVOICE NUMBER FROM" & " " & Me.txtName & " " & "INVOICES OR CHANGE INVOICE NUMBER"
   Me.txtInvoiceNo = ""
   Me.txtInvoiceNo.SetFocus
   If rs.State = 1 Then rs.Close: Exit Sub
End If
Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "Invoice Number", vbInformation, "Displaying"
     Exit Sub
End Sub

Private Sub txtName_Change()
If Me.txtName = "" Then
Exit Sub
End If
On Error GoTo OkError
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      
       MsgBox strg, vbInformation:
      Exit Sub
   End If
If xflag = False Then  ' when true u can view grid to add products,but when false u can edit with no grid shown
   rs.Open "Select * From AccountHolders  Where Name Like '" & Trim(Me.txtName) & "%" & "' Order By Name ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   flxgAccounts.Height = 950 + (285 * (rs.RecordCount - 1))
   
   If flxgAccounts.Height >= 4455 Then
      flxgAccounts.Height = 4455
   End If
    flxgAccounts.Rows = rs.RecordCount + 1
   With flxgAccounts
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("Name")
       .TextMatrix(X, 1) = rs.Fields("CreditLimit")
       .TextMatrix(X, 2) = rs.Fields("PhoneNumber")
       .TextMatrix(X, 3) = rs.Fields("AccID")
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 3
      .RowSel = 1
   End With
   flxgAccounts.Visible = True
   If rs.State = 1 Then rs.Close
   Me.cmdSave.Enabled = True
Else
     flxgAccounts.Visible = False
     'Dquantity = "1"
      If cFlag = True Then
       If rs.State <> 0 Then
          rs.Close
       End If
       cFlag = False: Exit Sub
      End If
     
End If
Else
      flxgAccounts.Visible = False
      'If cn.State = 1 Then cn.Close
      'If rs.State = 1 Then rs.Close
      'oflag = False
      xflag = False: Exit Sub
End If
If rs.State <> 0 Then
   rs.Close
End If

Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "Items In Stock", , "Displaying"
     Exit Sub
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
Me.flxgAccounts.Visible = True
   Me.flxgAccounts.SetFocus
End If
End Sub
Private Sub ListInvoice()
On Error GoTo SaveError
Me.LstInvoice.ListItems.Clear
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select * From AccountTransaction Where AccID='" & (Accountid) & "' Order By InvoiceNo", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.LstInvoice.ListItems.Add(, , Trim(rs!InvoiceNo))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!TransactionDescription)
                List_Item.SubItems(2) = Trim(rs!AmountDue)
                List_Item.SubItems(3) = Trim(rs!Date)
                List_Item.SubItems(4) = Trim(rs!Date)
                'List_Item.SubItems(5) = Trim(rs!AccID)
                 'List_Item.SubItems(6) = Trim(rs!BalanceCD)
                rs.MoveNext
            Loop
        End If
    Next i
    DoEvents
    
  rs.Close
    Set rs = Nothing
    If cn.State = 1 Then cn.Close
    Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "SAVE FAILED"
     Exit Sub
    
End Sub

Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
End Sub
Private Sub ListPayments()
On Error GoTo SaveError
Me.LstPayments.ListItems.Clear
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select TransPayments.* From TransPayments Inner Join AccountTransaction On TransPayments.AccTransID=AccountTransaction.AccTransID Where AccountTransaction.InvoiceNo='" & Trim(Me.txtInvoiceNo) & "' And AccID='" & (Accountid) & "' Order By PaymentDate", cn, adOpenForwardOnly, adLockReadOnly
 For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.LstPayments.ListItems.Add(, , Trim(rs!PaymentDate))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!PaymentMode)
                List_Item.SubItems(2) = Trim(rs!ChequeNo)
                List_Item.SubItems(3) = Trim(rs!AmountPaid)
                List_Item.SubItems(4) = Trim(rs!Arrears)
                'List_Item.SubItems(5) = Trim(rs!AccID)
                rs.MoveNext
            Loop
        End If
    Next i
    DoEvents
    
  rs.Close
    Set rs = Nothing
    If cn.State = 1 Then cn.Close
    Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "SAVE FAILED"
     Exit Sub
    
End Sub
Private Sub AccountCal()

On Error GoTo SaveError
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If

  If Me.txtAmountPaid <> "" And Me.txtBalCD <> "" And (Val(Me.txtAmountPaid) = Val(Me.txtBalCD)) Then
       MsgBox "THE DEBTOR FOR THIS PARTICULAR INVOICE HAS FINISHED PAYING", vbInformation, "PAYMENT COMPLETE"
  End If


 rs.Open "Select * From AccountTransaction Where InvoiceNo='" & Trim(Me.txtInvoiceNo) & "' And AccID='" & (Accountid) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   BalCD = rs.Fields("BalanceCD") - Val(Me.txtAmountPaid)
   Me.txtBalCD = BalCD
   Me.txtBalBD = BalCD
   Me.txtArrears = BalCD
   cn.BeginTrans
   If Me.cboPaymentMode <> "" And Me.txtAmountPaid <> "" Then
     cn.Execute "Insert Into TransPayments ([AccTransID],[PaymentMode],[ChequeNo],[AmountPaid],[Arrears],[PaymentDate],[ChequeDueDate]) select '" & (AccTransid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Val(Me.txtArrears.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
   End If
     If Y > 0 Then
   cn.Execute "Update AccountTransaction Set BalanceCD ='" & BalCD & "' Where InvoiceNo ='" & Trim(Me.txtInvoiceNo) & "' And AccID='" & (Accountid) & "'", Y
   If rs.State = 1 Then rs.Close
     End If
     If Y > 0 Then
     cn.CommitTrans
     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
     If rs.State = 1 Then rs.Close
   Else
     cn.RollbackTrans
     MsgBox "Saved Failed!", vbInformation, "Try Again"
     If rs.State = 1 Then rs.Close
   End If
Else
 If Me.txtAmountDue <> "" Then
   BalCD = Val(Me.txtAmountDue) - Val(Me.txtAmountPaid)
   Me.txtBalCD = BalCD
   Me.txtBalBD = BalCD
   Me.txtArrears = BalCD
   If rs.State = 1 Then rs.Close
 End If
   If rs.State = 1 Then rs.Close
End If

Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     Me.cmdFind.Enabled = True
     MsgBox "Sorry, Try Again!", vbInformation, "Try Again"
     Exit Sub
End Sub
Private Function Generate_AccountTransID(Account_ID As String) As Boolean
Dim strg As String, strg1 As String, strg2 As String, bFlag As Boolean

On Error GoTo SaveError
rs.Open "Select AccTransID From AccountTransaction  order by AccTransID Desc", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
      strg1 = Trim(rs.Fields!AccTransid)

   strg1 = Trim(Str(Val(strg1) + 1))
   strg1 = String$(8 - Len(strg1), "0") & strg1
Else
   strg1 = "00000001"
End If
Account_ID = strg1

If rs.State = 1 Then rs.Close

Generate_AccountTransID = True

Exit Function
SaveError:
     If rs.State = 1 Then rs.Close
    Generate_AccountTransID = False
     Exit Function
     
End Function
