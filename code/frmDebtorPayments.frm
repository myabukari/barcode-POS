VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmDebtorsPayments 
   BackColor       =   &H00C29E21&
   Caption         =   "Debtor's Payments"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   Icon            =   "frmDebtorPayments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   11970
   Begin VB.Frame Frame4 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   7800
      Width           =   11655
      Begin lvButton_H.lvButtons_H cmdPaymentReport 
         Height          =   375
         Left            =   6000
         TabIndex        =   25
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         Caption         =   "View Detailed PaymentReport"
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
         Image           =   "frmDebtorPayments.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdClear 
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
         Image           =   "frmDebtorPayments.frx":0624
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   9480
         TabIndex        =   27
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
         Image           =   "frmDebtorPayments.frx":1ED5
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdSpecifiedDebtor 
         Height          =   375
         Left            =   2400
         TabIndex        =   37
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         Caption         =   "View Report For Specified Debtor"
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
         Image           =   "frmDebtorPayments.frx":2327
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   11655
      Begin VB.TextBox txtReceiptNo 
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
         Left            =   1560
         TabIndex        =   35
         Top             =   1200
         Width           =   2415
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
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   2295
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
         ItemData        =   "frmDebtorPayments.frx":3AE9
         Left            =   1560
         List            =   "frmDebtorPayments.frx":3AF6
         TabIndex        =   9
         Top             =   240
         Width           =   2415
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
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgDebtorPayment 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   4683
         _Version        =   393216
         BackColor       =   16117969
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   $"frmDebtorPayments.frx":3B18
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
         _Band(0).Cols   =   7
      End
      Begin MSComCtl2.DTPicker dtpPaymentDate 
         Height          =   315
         Left            =   9480
         TabIndex        =   11
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   20774915
         CurrentDate     =   39115
      End
      Begin lvButton_H.lvButtons_H cmdCredit 
         Height          =   375
         Left            =   9480
         TabIndex        =   30
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "&Credit"
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
         Image           =   "frmDebtorPayments.frx":3BB6
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dtpBroughtforward 
         Height          =   315
         Left            =   5520
         TabIndex        =   32
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
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
         Format          =   20774915
         CurrentDate     =   39115
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C29E21&
         Caption         =   "Brought Forward on:"
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
         Left            =   4200
         TabIndex        =   38
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "ReceiptNo:"
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
         TabIndex        =   36
         Top             =   1200
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
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
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C29E21&
         Caption         =   "Credit Amount:"
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
         TabIndex        =   13
         Top             =   720
         Width           =   1335
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
         Left            =   8160
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgDebtors 
         Height          =   1335
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   16117969
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   "<Debtor's Name                                           |<DebtorID"
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
         _Band(0).Cols   =   2
      End
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
         Left            =   5520
         TabIndex        =   19
         Top             =   2400
         Width           =   2295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgInvoiceSearch 
         Height          =   1215
         Left            =   1560
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2143
         _Version        =   393216
         BackColor       =   16117969
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   $"frmDebtorPayments.frx":4008
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgInvoices 
         Height          =   1935
         Left            =   5040
         TabIndex        =   7
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3413
         _Version        =   393216
         BackColor       =   16117969
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   "<Invoice Number                 |<TransactionDescription    |<TransactionDate    |<TransactionAmount     |<DebtorID  "
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
         _Band(0).Cols   =   5
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
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   2895
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
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpTransDate 
         Height          =   315
         Left            =   9480
         TabIndex        =   20
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   20774915
         CurrentDate     =   39115
      End
      Begin VB.TextBox txtDebtorsBalance 
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
         Left            =   1560
         TabIndex        =   29
         Top             =   1680
         Width           =   2895
      End
      Begin lvButton_H.lvButtons_H cmdDebit 
         Height          =   375
         Left            =   9480
         TabIndex        =   31
         Top             =   2880
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "&Debit"
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
         Image           =   "frmDebtorPayments.frx":4097
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtInvoiceBalance 
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
         Left            =   1560
         TabIndex        =   33
         Top             =   1200
         Width           =   2895
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
         ItemData        =   "frmDebtorPayments.frx":44E9
         Left            =   1560
         List            =   "frmDebtorPayments.frx":44F3
         TabIndex        =   18
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Balance on InvoiceNo:"
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
         Left            =   360
         TabIndex        =   34
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Debtor Total Balance:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
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
         Left            =   7920
         TabIndex        =   23
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C29E21&
         Caption         =   "Debit Amount:"
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
         Left            =   4200
         TabIndex        =   22
         Top             =   2400
         Width           =   1335
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
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "DebtorName:"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1215
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
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmDebtorsPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Debtorid As Integer
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, AccTransid As String
Dim Balance As Double, TotalCredit As Double, TotalDebit As Double, msg As String, Invoiceflag As Boolean
Dim ShowDebtorflag As Boolean
Private Sub cmdSave_Click()
''Me.txtChequeNo.Enabled = False
'''Me.dtpChequeDueDate.Enabled = False
''Me.lblchequeDate.Enabled = False
''Me.lblChequeNo.Enabled = False
'
'If Trim(Me.txtName) = "" Then
'   MsgBox "YOU MUST ENTER ACCOUNT NAME.", vbInformation, "ACCOUNT NAME"
'   Me.txtName.SetFocus: Exit Sub
'End If
'
'If Trim(Me.txtInvoiceNo) = "" Then
'   MsgBox "YOU MUST ENTER INVOICE NUMBER.", vbInformation, "ENTER INVOICE NUMBER"
'   Me.txtInvoiceNo.SetFocus: Exit Sub
'End If
'
'
'If Trim(Me.cboTransDescription) = "" Then
'   MsgBox "YOU MUST ENTER TRANSACTION DESCRIPTION.", vbInformation, "DESCRIPTION"
'   Me.cboTransDescription.SetFocus: Exit Sub
'End If
'
'
'If Trim(Me.txtAmountDue) = "" Then
'   MsgBox "YOU MUST ENTER AMOUNTDUE OF GOODS ON CREDIT.", vbInformation, "AMOUNTDUE"
'   Me.txtAmountDue.SetFocus: Exit Sub
'End If
'
'
''If Val(Me.txtBalCD) = 0 Then
''      MsgBox "THE DEBTOR FOR THIS PARTICULAR INVOICE HAS FINISHED PAYING.", vbInformation, "DEBTOR CLEARED"
''     Exit Sub
''End If
'
'' If Me.txtAmountPaid <> "" And (Val(Me.txtAmountPaid) > Val(Me.txtBalCD)) Then
''     MsgBox "AMOUNT PAID SHOULD NOT BE MORE THAN BAL C/D", vbInformation, "LIABILITY"
''     Me.txtAmountPaid = "": Me.txtAmountPaid.SetFocus
''     Exit Sub
''   End If
'
'
'
''If Trim(Me.txtAmountPaid) = "" Then
'   'MsgBox "YOU MUST ENTER AMOUNTPAID.", vbInformation, "AMOUNTPAID"
'   'Me.txtAmountPaid.SetFocus: Exit Sub
''End If
'
'
'
'On Error GoTo SaveError
'Me.cmdSave.Enabled = False
'
''Open Connecttion to Server
'bFlag = OpenConnection(cn, strg)
'
'If bFlag = False Then
'   If cn.State = 1 Then cn.Close
'   If rs.State = 1 Then rs.Close
'   Me.MousePointer = vbDefault
'   Me.cmdSave.Enabled = True
'   MsgBox strg, vbInformation:
'   Exit Sub
'End If
''Me.txtStock = Trim(Me.txtTotalQuantity)
''If sflag = False Then
'   'save part
''rs.Open "Select ProductName From Products Where ProductName ='" & Trim(Me.txtName) & "'", cn, adOpenForwardOnly, adLockReadOnly
''If rs.RecordCount > 0 Then
'   'rs.Close: cn.Close
'   'MsgBox "A ProductName Has Already Been Setup with the Name.", vbInformation
'   'Me.MousePointer = vbDefault
'   'Me.cmdSave.Enabled = True
'   'Me.txtName.SetFocus: Exit Sub
'
'   'Me.MousePointer = vbDefault
'   'Me.cmdSave.Enabled = True
''End If
''rs.Close
'
'   'Call Generate_ProductID(Productid)
''   Call Generate_AccountTransID(AccountTransid)
''   Call AccountCal
'   rs.Open "Select * From AccountTransaction Where InvoiceNo='" & Trim(Me.txtInvoiceNo) & "' And AccID='" & Accountid & "'", cn, adOpenForwardOnly, adLockReadOnly
'   If rs.RecordCount = 0 Then
'
'     cn.BeginTrans
'     cn.Execute "Insert Into AccountTransaction  ([AccID],[InvoiceNo],[TransactionDescription],[AmountDue],[Date],[BalanceCD],[AccTransID]) select '" & (Accountid) & "','" & Trim(Me.txtInvoiceNo.Text) & "','" & Trim(Me.cboTransDescription.Text) & "','" & Val(Me.txtAmountDue.Text) & "','" & Trim(Me.dtpTransDate) & "','" & Val(Me.txtBalCD.Text) & "','" & (AccountTransid) & "'", Y
'
'     If Y > 0 Then
'     If Me.cboPaymentMode <> "" And Me.txtAmountPaid <> "" Then
'     cn.Execute "Insert Into TransPayments ([AccTransID],[PaymentMode],[ChequeNo],[AmountPaid],[Arrears],[PaymentDate],[ChequeDueDate]) select '" & (AccountTransid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Val(Me.txtArrears.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
'     Else
'     cn.Execute "Insert Into TransPayments ([AccTransID],[PaymentMode],[ChequeNo],[AmountPaid],[Arrears],[PaymentDate],[ChequeDueDate]) select '" & (AccountTransid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Val(Me.txtArrears.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
'     End If
'     End If
'     If Y > 0 Then
'     cn.CommitTrans
'     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
'       If rs.State = 1 Then rs.Close
'
'
'       Call ListPayments
'       Call ListInvoice
'       Call ClearCtrls
'     'Me.flxgAccounts.Visible = False
'     Me.txtName.SetFocus
'     Else
'     cn.RollbackTrans
'     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
'     If rs.State = 1 Then rs.Close
'     Me.txtName.SetFocus
'     End If
'
' Else
'     'If Me.cboPaymentMode <> "" And Me.txtAmountPaid <> "" Then
'     'cn.Execute "Insert Into TransPayments ([AccTransID],[PaymentMode],[ChequeNo],[AmountPaid],[Arrears],[PaymentDate],[ChequeDueDate]) select '" & (AccTransid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Val(Me.txtArrears.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
'     'If Y > 0 Then
'      ' MsgBox "Saved Successfully!", vbInformation, "Save Successful"
'
'       If rs.State = 1 Then rs.Close
'       Call ListPayments
'       Call ListInvoice
'       Call ClearCtrls
'       'Me.flxgAccounts.Visible = False
'
'       'End If
'     'If rs.State = 1 Then rs.Close
'     'End If
'    If rs.State = 1 Then rs.Close
'  End If
''Else
''edit part
'
'  ' rs.Open "Select ProductName From Products Where ProductName ='" & Trim(Me.txtName) & "' and ProductID<>'" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
'   'If rs.RecordCount > 0 Then
'      'rs.Close: cn.Close
'      'MsgBox "A Product Has Already Been Setup with the Name.", vbInformation
'      'Me.txtName.SetFocus: Exit Sub
'  ' End If
'   'rs.Close
'  '
'  ' rs.Open "Select ProductInventory.StockLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where  Products.ProductID='" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
'   'If rs.RecordCount > 0 Then
'      'StockQty = Val(Trim(Me.txtTotalQuantity)) + rs.Fields("StockLevel")
'      'rs.Close
'   'End If
'
'
'  ' cn.BeginTrans
'  ' cn.Execute "Update Products Set ProductName ='" & Trim(Me.txtName.Text) & "',BaseUnit='" & Val(Trim(Me.txtbaseunit.Text)) & "',UnitPrice='" & Trim(Me.txtUnitPrice.Text) & "',TotalQuantity='" & Val(Trim(Me.txtTotalQuantity.Text)) & "',PricePerCartone='" & Val(Trim(Me.txtbasePrice.Text)) & "',Discount='" & Val(Trim(Me.txtDiscount.Text)) & "',Manufacturer='" & Trim(Me.txtManufacterer.Text) & "',ExpiryDate='" & Trim(Me.dtpExp) & "',ManufucteryDate='" & Trim(Me.dtpmanu) & "',NoOfCartones='" & Val(Trim(Me.txtNoCartones.Text)) & "' Where ProductID ='" & Productid & "'", Y
'  '     If Y > 0 Then
'  'cn.Execute "Update ProductInventory Set StockLevel ='" & StockQty & "',ReorderLevel='" & Val(Trim(Me.txtReorder.Text)) & "' Where ProductID ='" & Productid & "'", Y
'  '     End If
'
'   'If Y > 0 Then
'      'cn.CommitTrans
'     ' MsgBox "Edit Successful!", vbInformation, "Edit Successful"
'     'Clear ctrls and setfocus to Clinic ctrl
'      'Call ListProducts
'      'Call ClearCtrls
'     ' Me.txtName = "": Me.txtbasePrice = "": Me.txtbaseunit = "": Me.txtDiscount = ""
'      'Me.txtManufacterer = "": Me.txtNoCartones = "": Me.txtReorder = "": Me.txtStock = ""
'     ' Me.txtTotalQuantity = "": Me.txtUnitPrice = ""
'     ' Me.txtName.SetFocus
'  ' Else
'     ' MsgBox "Sorry, Unable to Edit Product Details:Please Try Again!", vbInformation, "Edit Failed"
' '  End If
'
'' End If
'
''flag = False
''If cn.State = 1 Then cn.Close
''If rs.State = 1 Then rs.Close
'
''Me.MousePointer = vbDefault
''Me.cmdSave.Enabled = True
'
'
'Exit Sub
'SaveError:
'     If cn.State = 1 Then cn.Close
'     If rs.State = 1 Then rs.Close
'     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
'     Exit Sub

End Sub

Private Sub cboPaymentMode_Click()
If Me.cboPaymentMode = "Cheque" Then
Me.txtChequeNo.Enabled = True
Else
Me.txtChequeNo.Enabled = False
End If
End Sub

Private Sub cmdClear_Click()
Call ClearCtrls
End Sub

Private Sub cmdCredit_Click()
'Me.txtChequeNo.Enabled = False
'Me.dtpChequeDueDate.Enabled = False
'Me.lblchequeDate.Enabled = False
'Me.lblChequeNo.Enabled = False

If Trim(Me.txtName) = "" Then
   MsgBox "YOU MUST ENTER DEBTOR'S NAME.", vbInformation, "DEBTOR'S NAME"
   Me.txtName.SetFocus: Exit Sub
End If

If Trim(Me.txtInvoiceNo) = "" Then
   MsgBox "YOU MUST ENTER INVOICE NUMBER.", vbInformation, "ENTER INVOICE NUMBER"
   Me.txtInvoiceNo.SetFocus: Exit Sub
End If


If Trim(Me.cboPaymentMode) = "" Then
   MsgBox "YOU MUST ENTER MODE OF PAYMENT.", vbInformation, "DESCRIPTION"
   Me.cboPaymentMode.SetFocus: Exit Sub
End If


If Trim(Me.cboPaymentMode) = "Cheque" And Trim(Me.txtChequeNo) = "" Then
   MsgBox "YOU MUST ENTER CHEQUE NUMBER.", vbInformation, "CHEQUE NUMBER"
   Me.txtChequeNo.SetFocus: Exit Sub
End If


'If Val(Me.txtBalCD) = 0 Then
'      MsgBox "THE DEBTOR FOR THIS PARTICULAR INVOICE HAS FINISHED PAYING.", vbInformation, "DEBTOR CLEARED"
'     Exit Sub
'End If

' If Me.txtAmountPaid <> "" And (Val(Me.txtAmountPaid) > Val(Me.txtBalCD)) Then
'     MsgBox "AMOUNT PAID SHOULD NOT BE MORE THAN BAL C/D", vbInformation, "LIABILITY"
'     Me.txtAmountPaid = "": Me.txtAmountPaid.SetFocus
'     Exit Sub
'   End If
  


If Trim(Me.txtAmountPaid) = "" Then
   MsgBox "YOU MUST ENTER AMOUNTPAID.", vbInformation, "AMOUNTPAID"
   Me.txtAmountPaid.SetFocus: Exit Sub
End If

Me.cboTransDescription = "": Me.txtAmountDue = ""

If Me.dtpPaymentDate < Date Then
If MsgBox("ARE YOU SURE  ABOUT THE PAYMENT DATE ENTERED?", vbYesNo + vbQuestion, "CONFIRM PAYMENT DATE") = vbNo Then
Me.dtpPaymentDate.SetFocus
Exit Sub
End If
End If
  
 Randomize
 Me.txtReceiptNo = Round(10000000 * Rnd(), 0) & "/" & Round(1000 * Rnd(), 0)
  
  
On Error GoTo SaveError
'Me.cmdDebit.Enabled = False

'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdDebit.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If

   rs.Open "Select * From DebtorsTransaction Where InvoiceNo='" & Trim(Me.txtInvoiceNo) & "' And DebtorID='" & Debtorid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount = 0 Then
      MsgBox "You cannot Proceed crediting without debiting for the InvoiceNumber specified!", vbInformation, "Credit failed"
   If rs.State = 1 Then rs.Close
 Else
' Format$(dtpPaymentDate, "Medium Date") & " " & Format$(Now, "Medium Time")
     If Me.cboPaymentMode <> "" And Me.txtAmountPaid <> "" Then
      
      Call DebtorsBalance(Debtorid, Balance, Trim(Me.txtInvoiceNo), msg, Val(Trim(Me.txtAmountPaid)))
      If Balance < 0 Then
        MsgBox "The Debtor is paying more than available balance for the InvoiceNo specified", vbInformation, "Credit failed"
        Exit Sub
      ElseIf msg <> "" And msg <> "TotalDebit is less than TotalCredit" And Balance = 0 Then
        MsgBox msg, vbInformation, ""
      ElseIf msg <> "" And msg = "TotalDebit is less than TotalCredit" Then
        MsgBox msg, vbInformation, "The Debtor is paying more than available balance for the InvoiceNo specified"
        Exit Sub
      End If
      
'      Call CalculateBalance(Balance)
       Call Balances(Balance, Val(Trim(Me.txtAmountPaid)), Val(Trim(Me.txtAmountDue)))
      If Balance < 0 Then
        MsgBox "Total Credits will exceed total Debits,transaction cannot proceed", vbInformation, "Proceeding will lead to liability"
        Exit Sub
      End If
      
      cn.BeginTrans
      cn.Execute "Insert Into DebtorsPayments ([InvoiceNo],[Paymentmode],[PaymentDate],PaymentTime,[ChequeNo],[Credit],[Balance],ReceiptNo,BFID) select '" & Trim(Me.txtInvoiceNo) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & (Me.dtpPaymentDate) & "','" & Format$(dtpPaymentDate, "Medium Date") & " " & Format$(Now, "Medium Time") & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Balance & "','" & Trim(Me.txtReceiptNo) & "','BF'", Y
      If Y > 0 Then
        cn.Execute "Update Balances Set Balance ='" & Balance & "' Where BalanceID ='B1'", Y
      End If
        If Y > 0 Then
            cn.CommitTrans
            MsgBox "Credited Successfully!", vbInformation, "Credited Successful"
         Else
            cn.RollbackTrans
            MsgBox "Sorry, Unable to Save PaymentS Details:Please Try Again!", vbInformation, "Save Failed"
            If rs.State = 1 Then rs.Close
            Me.txtName.SetFocus
        End If
       If rs.State = 1 Then rs.Close
       
         Me.cboPaymentMode = "": Me.txtAmountPaid = "": Me.txtChequeNo = "": Me.dtpPaymentDate = Date: Me.txtReceiptNo = ""
         Call DebtorsPayments
         Call DebtorsBalance(Debtorid, Balance)
         Me.txtDebtorsBalance = Balance
         Call DebtorsBalance(Debtorid, Balance, Trim(Me.txtInvoiceNo))
         Me.txtInvoiceBalance = Balance
''       Call ListInvoice
''       Call ClearCtrls
'       'Me.flxgAccounts.Visible = False
'
'       'End If
'     'If rs.State = 1 Then rs.Close
'     End If
'    If rs.State = 1 Then rs.Close
  End If
  If rs.State = 1 Then rs.Close
  End If
  If rs.State = 1 Then rs.Close

  


Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub

Private Sub cmdDebit_Click()


If Trim(Me.txtName) = "" Then
   MsgBox "YOU MUST ENTER DEBTOR'S NAME.", vbInformation, "DEBTOR'S NAME"
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

Me.cboPaymentMode = "": Me.txtAmountPaid = "": Me.txtChequeNo = ""
'If Val(Me.txtBalCD) = 0 Then
'      MsgBox "THE DEBTOR FOR THIS PARTICULAR INVOICE HAS FINISHED PAYING.", vbInformation, "DEBTOR CLEARED"
'     Exit Sub
'End If

' If Me.txtAmountPaid <> "" And (Val(Me.txtAmountPaid) > Val(Me.txtBalCD)) Then
'     MsgBox "AMOUNT PAID SHOULD NOT BE MORE THAN BAL C/D", vbInformation, "LIABILITY"
'     Me.txtAmountPaid = "": Me.txtAmountPaid.SetFocus
'     Exit Sub
'   End If
  

If Me.dtpTransDate < Date Then
If MsgBox("ARE YOU SURE  ABOUT THE TRANSACTION DATE ENTERED?", vbYesNo + vbQuestion, "CONFIRM TRANSACTION DATE") = vbNo Then
Me.dtpTransDate.SetFocus
Exit Sub
End If
End If

'If Trim(Me.txtAmountPaid) = "" Then
   'MsgBox "YOU MUST ENTER AMOUNTPAID.", vbInformation, "AMOUNTPAID"
   'Me.txtAmountPaid.SetFocus: Exit Sub
'End If



On Error GoTo SaveError
'Me.cmdDebit.Enabled = False

'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdDebit.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If

   rs.Open "Select * From DebtorsTransaction Where InvoiceNo='" & Trim(Me.txtInvoiceNo) & "' And DebtorID='" & Debtorid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount = 0 Then
     
         cn.BeginTrans
         cn.Execute "Insert Into DebtorsTransaction  ([DebtorID],[InvoiceNo],[TransactionDescription],[TransactionAmount],[TransactionDate]) select '" & (Debtorid) & "','" & Trim(Me.txtInvoiceNo.Text) & "','" & Trim(Me.cboTransDescription.Text) & "','" & Val(Me.txtAmountDue) & "','" & Trim(Me.dtpTransDate) & "'", Y
         If Y > 0 Then
'         Call CalculateBalance(Balance)
         If rs.State = 1 Then rs.Close
         Call Balances(Balance, Val(Trim(Me.txtAmountPaid)), Val(Trim(Me.txtAmountDue)))
         rs.Open "Select Balance From Balances  Where BalanceID ='B1'", cn, adOpenForwardOnly, adLockReadOnly
          If rs.RecordCount > 0 Then
                cn.Execute "Update Balances Set Balance ='" & Balance & "' Where BalanceID ='B1'", Y
            Else
                cn.Execute "Insert Into Balances  ([BalanceID],[Balance]) select 'B1','" & Balance & "'", Y
          End If
          If rs.State = 1 Then rs.Close
         End If
         'Inserts BF into BFID field of DebtorsPayments table to link it to BalanceBroughtForward table
         'to aid in showing the brought forward balance in reports
         If Y > 0 Then
         cn.Execute "Insert Into DebtorsPayments ([InvoiceNo],[Debit],[Balance],PaymentDate,PaymentTime,BFID) select '" & Trim(Me.txtInvoiceNo) & "','" & Val(Trim(Me.txtAmountDue)) & "','" & Balance & "','" & Trim(Me.dtpTransDate) & "','" & Format$(dtpTransDate, "Medium Date") & " " & Format$(Now, "Medium Time") & "','BF'", Y
         End If
         
           If Y > 0 Then
               cn.CommitTrans
               MsgBox "Debited Successfully!", vbInformation, "Debited Successful"
               If rs.State = 1 Then rs.Close
               
               Me.cboTransDescription = "": Me.txtAmountDue = "": Me.dtpTransDate = Date
                Call DebtorsPayments
                Call DebtorsBalance(Debtorid, Balance)
                Me.txtDebtorsBalance = Balance
                Call DebtorsBalance(Debtorid, Balance, Trim(Me.txtInvoiceNo))
                Me.txtInvoiceBalance = Balance
        '       Call ListPayments
        '       Call ListInvoice
        '       Call ClearCtrls
             '  Me.flxgAccounts.Visible = False
                Me.txtName.SetFocus
            Else
                cn.RollbackTrans
                MsgBox "Sorry, Unable to Debit Debitor's Details:Please Try Again!", vbInformation, "Debit Failed"
                If rs.State = 1 Then rs.Close
                Me.txtName.SetFocus
            End If
    If rs.State = 1 Then rs.Close
 ElseIf rs.RecordCount > 0 And Trim(Me.cboTransDescription) = "B/C" Then
   
   If rs.State = 1 Then rs.Close
         cn.BeginTrans
         Call Balances(Balance, Val(Trim(Me.txtAmountPaid)), Val(Trim(Me.txtAmountDue)))
         
         rs.Open "Select Balance From Balances  Where BalanceID ='B1'", cn, adOpenForwardOnly, adLockReadOnly
          If rs.RecordCount > 0 Then
                cn.Execute "Update Balances Set Balance ='" & Balance & "' Where BalanceID ='B1'", Y
            Else
                cn.Execute "Insert Into Balances  ([BalanceID],[Balance]) select 'B1','" & Balance & "'", Y
          End If
          If rs.State = 1 Then rs.Close
          
          If Y > 0 Then
          cn.Execute "Insert Into DebtorsPayments ([InvoiceNo],[Debit],[Balance],PaymentDate,PaymentTime,Paymentmode,BFID) select '" & Trim(Me.txtInvoiceNo) & "','" & Val(Trim(Me.txtAmountDue)) & "','" & Balance & "','" & Trim(Me.dtpTransDate) & "','" & Format$(dtpTransDate, "Medium Date") & " " & Format$(Now, "Medium Time") & "','B/C','BF'", Y
           End If
             If Y > 0 Then
               cn.CommitTrans
               MsgBox "Debited Successfully!", vbInformation, "Debited Successful"
               If rs.State = 1 Then rs.Close
                Me.cboTransDescription = "": Me.txtAmountDue = "": Me.dtpTransDate = ""
                Call DebtorsPayments
                Call DebtorsBalance(Debtorid, Balance)
                Me.txtDebtorsBalance = Balance
                Call DebtorsBalance(Debtorid, Balance, Trim(Me.txtInvoiceNo))
                Me.txtInvoiceBalance = Balance
             Else
                cn.RollbackTrans
                MsgBox "Sorry, Debit failed:Please Try Again!", vbInformation, "Debit Failed"
                If rs.State = 1 Then rs.Close
             End If
                If rs.State = 1 Then rs.Close
  ElseIf rs.RecordCount > 0 And Trim(Me.cboTransDescription) <> "B/C" Then
             MsgBox "You cannot debit an already existing Invoice Number for sales transaction ", vbInformation, "Invoice Number already exist!"
             If rs.State = 1 Then rs.Close
             Exit Sub
  If rs.State = 1 Then rs.Close
End If


Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub

Private Sub cmdFind_Click()

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPaymentReport_Click()
frmDebtorsPaymentsRpt.Show
End Sub

Private Sub cmdReceipt_Click()
'Randomize
'Me.txtReceiptNo = Round(10000000 * Rnd(), 0) & "/" & Round(1000 * Rnd(), 0)
End Sub

Private Sub cmdSpecifiedDebtor_Click()
frmIndividualDebtorPaymentsRpt.Show
frmIndividualDebtorPaymentsRpt.cboDebtor = Me.txtName
End Sub

Private Sub dtpBroughtforward_CloseUp()
For X = 0 To 6
Me.flxgDebtorPayment.TextMatrix(1, X) = ""
Next
Me.flxgDebtorPayment.Rows = 2
Call DebtorsPayments(Debtorid)
End Sub

Private Sub flxgDebtors_Click()
Me.txtName = Me.flxgDebtors.TextMatrix(Me.flxgDebtors.Row, 0)

Debtorid = Me.flxgDebtors.TextMatrix(Me.flxgDebtors.Row, 1)
Me.flxgDebtors.Visible = False
Invoiceflag = True
Me.txtInvoiceNo = ""
Invoiceflag = False
Me.cmdDebit.Enabled = True
sflag = False
oflag = False

Call DebtorsBalance(Debtorid, Balance)
Me.txtDebtorsBalance = Balance
Call DebtorsPayments(Debtorid)
Call FindInvoiceNo(Debtorid)
'Call ListInvoice
'Me.framAllInvoice.Caption = Me.txtName & " " & "Invoices"
Me.txtInvoiceNo.SetFocus
End Sub

Private Sub MSHFlexGrid2_Click()

End Sub

Private Sub flxgInvoices_Click()
Invoiceflag = True
Me.txtInvoiceNo = Me.flxgInvoices.TextMatrix(Me.flxgInvoices.Row, 0)
Invoiceflag = False
Me.cboTransDescription = Me.flxgInvoices.TextMatrix(Me.flxgInvoices.Row, 1)
Me.dtpTransDate = Me.flxgInvoices.TextMatrix(Me.flxgInvoices.Row, 2)
Call DebtorsBalance(Debtorid, Balance, Trim(Me.txtInvoiceNo))
Me.txtInvoiceBalance = Balance
If Balance = "0" Then
MsgBox "The Debtor has being cleared for this InvoiceNo", vbInformation, "Balance Information"
Me.flxgInvoiceSearch.Visible = False
Exit Sub
End If
End Sub

Private Sub flxgInvoiceSearch_Click()
Me.txtInvoiceNo = Me.flxgInvoiceSearch.TextMatrix(Me.flxgInvoiceSearch.Row, 0)
Me.cboTransDescription = Me.flxgInvoiceSearch.TextMatrix(Me.flxgInvoiceSearch.Row, 1)
Me.dtpTransDate = Me.flxgInvoiceSearch.TextMatrix(Me.flxgInvoiceSearch.Row, 2)
Call DebtorsBalance(Debtorid, Balance, Trim(Me.txtInvoiceNo))
Me.txtInvoiceBalance = Balance
If Balance = "0" Then
MsgBox "The Debtor has being cleared for this InvoiceNo", vbInformation, "Balance Information"
Me.flxgInvoiceSearch.Visible = False
Exit Sub
End If
'      Call DebtorsBalance(Debtorid, Balance, Trim(Me.txtInvoiceNo), msg)
'      If msg <> "" And msg <> "TotalDebit is less than TotalCredit" Then
'        MsgBox msg, vbInformation, "The Debtor has being cleared for this InvoiceNo"
'      ElseIf msg <> "" And msg = "TotalDebit is less than TotalCredit" Then
'        MsgBox msg, vbInformation, "The Debtor is paying more than available balance for the InvoiceNo specified"
'        Exit Sub
'      End If

Me.flxgInvoiceSearch.Visible = False
End Sub

Private Sub Form_Load()
Me.dtpPaymentDate = Date
Me.dtpTransDate = Date
Me.dtpBroughtforward = Date
Me.txtChequeNo.Enabled = False
Me.Height = 9195
  Me.Width = 12090
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub txtInvoiceNo_Change()
If Me.txtInvoiceNo = "" Then
 Me.flxgInvoiceSearch.Visible = False
 Exit Sub
End If
'On Error GoTo OkError
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      
       MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   If Invoiceflag = False Then
   rs.Open "Select * From DebtorsTransaction  Where InvoiceNo Like '" & Trim(Me.txtInvoiceNo) & "%" & "' and DebtorID='" & Debtorid & "' Order By TransactionDate ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
       flxgInvoiceSearch.Height = 950 + (285 * (rs.RecordCount - 1))
       If flxgInvoiceSearch.Height >= 4455 Then
          flxgInvoiceSearch.Height = 4455
       End If
       flxgInvoiceSearch.Rows = rs.RecordCount + 1
       With flxgInvoiceSearch
          For X = 1 To rs.RecordCount
            .TextMatrix(X, 0) = rs.Fields("InvoiceNo")
            .TextMatrix(X, 1) = rs.Fields("TransactionDescription")
            .TextMatrix(X, 2) = rs.Fields("TransactionDate")
            .TextMatrix(X, 3) = rs.Fields("TransactionAmount")
            .TextMatrix(X, 4) = rs.Fields("DebtorID")
            rs.MoveNext
          Next
          .Col = 0
          .Row = 1
          .ColSel = 4
          .RowSel = 1
       End With
       flxgInvoiceSearch.Visible = True
       If rs.State = 1 Then rs.Close
    '   Me.cmdSave.Enabled = True
    Else
         flxgInvoiceSearch.Visible = False
         If rs.State = 1 Then rs.Close
          
         
    End If

If rs.State = 1 Then rs.Close
End If
Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "Debtors could not display", , "Displaying"
     Exit Sub
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

Private Sub txtName_Change()
If Me.txtName = "" Then
Me.flxgDebtors.Visible = False
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
   
If ShowDebtorflag = False Then  ' when true u can view grid to add products,but when false u can edit with no grid shown
   rs.Open "Select * From Debtors  Where DebtorName Like '" & Trim(Me.txtName) & "%" & "' Order By DebtorName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   flxgDebtors.Height = 950 + (285 * (rs.RecordCount - 1))
   
   If flxgDebtors.Height >= 4455 Then
      flxgDebtors.Height = 4455
   End If
    flxgDebtors.Rows = rs.RecordCount + 1
   With flxgDebtors
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("DebtorName")
       .TextMatrix(X, 1) = rs.Fields("DebtorID")
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 1
      .RowSel = 1
   End With
   flxgDebtors.Visible = True
   If rs.State = 1 Then rs.Close
'   Me.cmdSave.Enabled = True
Else
     flxgDebtors.Visible = False
     'Dquantity = "1"
      If cFlag = True Then
       If rs.State <> 0 Then
          rs.Close
       End If
       cFlag = False: Exit Sub
      End If
     
End If
Else
      flxgDebtors.Visible = False
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
    MsgBox "Debtors could not display", , "Displaying"
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
   Me.txtInvoiceNo.SetFocus
End If
End Sub
Private Sub FindInvoiceNo(DebtorNo)
 
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If

   rs.Open "Select * From DebtorsTransaction  Where DebtorID = '" & DebtorNo & "' Order By TransactionDate ", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
   rs.MoveFirst
         Me.flxgInvoices.Rows = rs.RecordCount + 1
         For X = 1 To rs.RecordCount
           Me.flxgInvoices.TextMatrix(X, 0) = rs.Fields("InvoiceNo")
           Me.flxgInvoices.TextMatrix(X, 1) = rs.Fields("TransactionDescription")
           Me.flxgInvoices.TextMatrix(X, 2) = rs.Fields("TransactionDate")
           Me.flxgInvoices.TextMatrix(X, 3) = rs.Fields("TransactionAmount")
           Me.flxgInvoices.TextMatrix(X, 4) = rs.Fields("DebtorID")
           rs.MoveNext
         Next
         If rs.State = 1 Then rs.Close
      End If
   If rs.State = 1 Then rs.Close
   
End Sub
Private Sub CalculateBalance(CalBalance As Double, Optional msg As String)

bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   If rs.State = 1 Then rs.Close
 rs.Open "Select  Sum(Credit) as Credits,Sum(Debit) as Debits From DebtorsPayments   ", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
'   If Me.txtAmountPaid <> "" Then
            If rs.Fields("Credits") <> "" Then
            TotalCredit = Val(Me.txtAmountPaid) + rs.Fields("Credits")
            Else
            TotalCredit = Val(Me.txtAmountPaid)
            End If
            If rs.Fields("Debits") <> "" Then
            TotalDebit = Val(Me.txtAmountDue) + rs.Fields("Debits")
            Else
            TotalDebit = Val(Me.txtAmountDue)
            End If
            CalBalance = TotalDebit - TotalCredit
        If CalBalance = 0 Then
             msg = "Debtors have being cleared"
         ElseIf CalBalance < 0 Then
             msg = "TotalDebit is less than TotalCredit"
        End If
'   Else
'
'   End If
   If rs.State = 1 Then rs.Close
   End If
   If rs.State = 1 Then rs.Close
   
End Sub
Private Sub DebtorsBalance(DebtorNo, CalBalance As Double, Optional InvoiceNo As String, Optional msg As String, Optional CurrentCredit As Double, Optional CurrentDebit As Double)

bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   msg = ""
   If rs.State = 1 Then rs.Close
 If InvoiceNo <> "" Then
 rs.Open "Select  Sum(Credit) as Credits,Sum(Debit) as Debits From (DebtorsPayments Inner Join DebtorsTransaction on DebtorsPayments.InvoiceNo=DebtorsTransaction.InvoiceNo) inner Join Debtors on Debtors.DebtorID=DebtorsTransaction.DebtorID where DebtorsTransaction.DebtorID='" & Debtorid & "' and DebtorsTransaction.InvoiceNo='" & InvoiceNo & "' ", cn, adOpenForwardOnly, adLockReadOnly
 Else
 rs.Open "Select  Sum(Credit) as Credits,Sum(Debit) as Debits From (DebtorsPayments Inner Join DebtorsTransaction on DebtorsPayments.InvoiceNo=DebtorsTransaction.InvoiceNo) inner Join Debtors on Debtors.DebtorID=DebtorsTransaction.DebtorID where DebtorsTransaction.DebtorID='" & Debtorid & "' ", cn, adOpenForwardOnly, adLockReadOnly
 End If
   If rs.RecordCount > 0 Then
'   If Me.txtAmountPaid <> "" Then
            If rs.Fields("Credits") <> "" Then
            TotalCredit = CurrentCredit + rs.Fields("Credits")
            Else
            TotalCredit = CurrentCredit
            End If
            If rs.Fields("Debits") <> "" Then
            TotalDebit = CurrentDebit + rs.Fields("Debits")
            Else
            TotalDebit = CurrentDebit
            End If
            CalBalance = TotalDebit - TotalCredit
        If CalBalance = 0 Then
             msg = "Debtor has being cleared"
         ElseIf CalBalance < 0 Then
             msg = "TotalDebit is less than TotalCredit"
        End If
'   Else
'
'   End If
   If rs.State = 1 Then rs.Close
   End If
   If rs.State = 1 Then rs.Close
End Sub
Private Sub DebtorsPayments(Optional DebtorNo)
bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
     If rs.State = 1 Then rs.Close
'Format$(dtpBroughtforward, "Medium Date") & " " & Format$(Now, "Medium Time")
'Shows debtors debits and credits for dates before dtpBroughtforward, and shows a Brought Forward balance on
'dtpBroughtforward date
   If Me.dtpBroughtforward <> Date And Me.dtpBroughtforward < Date Then
        rs.Open "Select  Balance From (DebtorsPayments Inner Join DebtorsTransaction on DebtorsPayments.InvoiceNo=DebtorsTransaction.InvoiceNo) inner Join Debtors on Debtors.DebtorID=DebtorsTransaction.DebtorID where PaymentDate <='" & Me.dtpBroughtforward & "' order by PaymentDate desc,PaymentTime desc", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
        'Shows the Brought Forward balance
              Me.flxgDebtorPayment.TextMatrix(1, 0) = Me.dtpBroughtforward
              Me.flxgDebtorPayment.TextMatrix(1, 2) = "Brought forward"
              Me.flxgDebtorPayment.TextMatrix(1, 6) = Format$(rs.Fields("Balance"), "#,###.00")
          If rs.State = 1 Then rs.Close
        End If
          If rs.State = 1 Then rs.Close
          flxgRows = Me.flxgDebtorPayment.Rows
   End If
   
'   If rs.State = 1 Then rs.Close
'Shows the preceeding balances after Brought Forward balance
    If Me.dtpBroughtforward <> Date And Me.dtpBroughtforward < Date Then
      rs.Open "Select  * From (DebtorsPayments Inner Join DebtorsTransaction on DebtorsPayments.InvoiceNo=DebtorsTransaction.InvoiceNo) inner Join Debtors on Debtors.DebtorID=DebtorsTransaction.DebtorID where PaymentDate >'" & Me.dtpBroughtforward & "' order by PaymentDate asc", cn, adOpenForwardOnly, adLockReadOnly
      If rs.RecordCount > 0 Then
         rs.MoveFirst
             Me.flxgDebtorPayment.Rows = rs.RecordCount + flxgRows
         For X = flxgRows To rs.RecordCount + (flxgRows - 1)
             If rs.Fields("PaymentDate") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 0) = rs.Fields("PaymentDate")
    '          Else
    '          Me.flxgDebtorPayment.TextMatrix(X, 0) = rs.Fields("TransactionDate")
               End If
               Me.flxgDebtorPayment.TextMatrix(X, 1) = rs.Fields("DebtorName")
               Me.flxgDebtorPayment.TextMatrix(X, 2) = rs.Fields("InvoiceNo")
               If rs.Fields("Paymentmode") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 3) = rs.Fields("Paymentmode")
               Else
               Me.flxgDebtorPayment.TextMatrix(X, 3) = rs.Fields("TransactionDescription")
               End If
               If rs.Fields("Debit") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 4) = Format$(rs.Fields("Debit"), "#,###.00")
               End If
               If rs.Fields("Credit") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 5) = Format$(rs.Fields("Credit"), "#,###.00")
               End If
               If rs.Fields("Balance") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 6) = Format$(rs.Fields("Balance"), "#,###.00")
               End If
           rs.MoveNext
         Next
         If rs.State = 1 Then rs.Close
     End If
     If rs.State = 1 Then rs.Close
   Else
   'shows all balances without Brought Forward balance
    rs.Open "Select  * From (DebtorsPayments Inner Join DebtorsTransaction on DebtorsPayments.InvoiceNo=DebtorsTransaction.InvoiceNo) inner Join Debtors on Debtors.DebtorID=DebtorsTransaction.DebtorID  order by PaymentDate asc", cn, adOpenForwardOnly, adLockReadOnly
     If rs.RecordCount > 0 Then
       rs.MoveFirst
         Me.flxgDebtorPayment.Rows = rs.RecordCount + 1
         For X = 1 To rs.RecordCount
               If rs.Fields("PaymentDate") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 0) = rs.Fields("PaymentDate")
    '          Else
    '          Me.flxgDebtorPayment.TextMatrix(X, 0) = rs.Fields("TransactionDate")
               End If
               Me.flxgDebtorPayment.TextMatrix(X, 1) = rs.Fields("DebtorName")
               Me.flxgDebtorPayment.TextMatrix(X, 2) = rs.Fields("InvoiceNo")
               If rs.Fields("Paymentmode") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 3) = rs.Fields("Paymentmode")
               Else
               Me.flxgDebtorPayment.TextMatrix(X, 3) = rs.Fields("TransactionDescription")
               End If
               If rs.Fields("Debit") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 4) = Format$(rs.Fields("Debit"), "#,###.00")
               End If
               If rs.Fields("Credit") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 5) = Format$(rs.Fields("Credit"), "#,###.00")
               End If
               If rs.Fields("Balance") <> "" Then
               Me.flxgDebtorPayment.TextMatrix(X, 6) = Format$(rs.Fields("Balance"), "#,###.00")
               End If
         rs.MoveNext
         Next
         If rs.State = 1 Then rs.Close
      End If
   If rs.State = 1 Then rs.Close
   End If
   If rs.State = 1 Then rs.Close
End Sub

Private Sub Balances(CalBalance As Double, Optional CurrentCredit As Double, Optional CurrentDebit As Double, Optional msg As String)
bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   
  rs.Open "Select Balance From Balances Where BalanceID ='B1' ", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      CalBalance = (rs.Fields("Balance") + CurrentDebit) - CurrentCredit
      If rs.State = 1 Then rs.Close
   Else
      CalBalance = (CurrentDebit)
      If rs.State = 1 Then rs.Close
   End If
   If rs.State = 1 Then rs.Close
   

End Sub
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
sflag = False
For X = 0 To 6
Me.flxgDebtorPayment.TextMatrix(1, X) = ""
Next
Me.flxgDebtorPayment.Rows = 2
For X = 0 To 4
Me.flxgInvoices.TextMatrix(1, X) = ""
Next
Me.flxgInvoices.Rows = 2
Me.txtName.SetFocus
End Sub
