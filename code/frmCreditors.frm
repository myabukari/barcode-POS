VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmCreditors 
   BackColor       =   &H00C29E21&
   Caption         =   "Creditors Payments"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   Icon            =   "frmCreditors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   9495
   Begin VB.Frame Frame5 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   360
      TabIndex        =   26
      Top             =   7800
      Width           =   8775
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
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   1335
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
         Left            =   8760
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
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
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
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
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         Width           =   1335
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
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
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
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   7320
         TabIndex        =   37
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
         Image           =   "frmCreditors.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdSave 
         Height          =   375
         Left            =   120
         TabIndex        =   38
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
         Image           =   "frmCreditors.frx":075C
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdDelete 
         Height          =   375
         Left            =   1560
         TabIndex        =   39
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
         Image           =   "frmCreditors.frx":0BAE
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdFind 
         Height          =   375
         Left            =   3000
         TabIndex        =   40
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
         Image           =   "frmCreditors.frx":0D48
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdClear 
         Height          =   375
         Left            =   4440
         TabIndex        =   41
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
         Image           =   "frmCreditors.frx":119A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdReport 
         Height          =   375
         Left            =   5880
         TabIndex        =   42
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
         Image           =   "frmCreditors.frx":2A4B
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   3615
      Left            =   240
      TabIndex        =   13
      Top             =   4200
      Width           =   9015
      Begin VB.TextBox txtBalCD 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   34
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   2055
         Left            =   0
         TabIndex        =   24
         Top             =   1560
         Width           =   9015
         Begin MSComctlLib.ListView LstPayments 
            Height          =   1815
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3201
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483643
            BackColor       =   2499106
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
               Text            =   "Date"
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
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "AmountPaid"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Balance"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.TextBox txtChequeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   16
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox cboPaymentMode 
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCreditors.frx":2D65
         Left            =   1560
         List            =   "frmCreditors.frx":2D72
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtAmountPaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpPaymentDate 
         Height          =   315
         Left            =   6240
         TabIndex        =   17
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   21954563
         CurrentDate     =   39115
      End
      Begin MSComCtl2.DTPicker dtpChequeDueDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   22
         Top             =   1200
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   21954563
         CurrentDate     =   39115
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C29E21&
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5160
         TabIndex        =   35
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblChequeDue 
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   840
         Width           =   975
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   360
         Width           =   1095
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgSupplier 
         Height          =   1935
         Left            =   600
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3413
         _Version        =   393216
         BackColor       =   12754465
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorBkg    =   14402944
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   "<Supplier Name                                                     |<PhoneNumber                |<SupplierID"
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComctlLib.ListView LstInvoiceNo 
         Height          =   2415
         Left            =   4920
         TabIndex        =   5
         Top             =   120
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483643
         BackColor       =   2499106
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "InvoiceNumber"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TransactionDescription"
            Object.Width           =   3528
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
      End
      Begin VB.TextBox txtInvoiceNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtSupplierName 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   4800
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   4800
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   4800
         X2              =   4800
         Y1              =   120
         Y2              =   2880
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "InvoiceNumber:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "SupplierName:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   240
      TabIndex        =   32
      Top             =   7800
      Width           =   9015
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   9015
      Begin MSComCtl2.DTPicker dtpTransactionDate 
         Height          =   315
         Left            =   6240
         TabIndex        =   11
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   21954563
         CurrentDate     =   39118
      End
      Begin VB.ComboBox cboTransactionDescription 
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCreditors.frx":2D94
         Left            =   1560
         List            =   "frmCreditors.frx":2D9E
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtAmountDue 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5520
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Tansaction Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCreditors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Accountid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, CreditorAccid As String
Dim BalCD As Double, AccTransid As String, xflag As Boolean, AccountTransid As String, Supplierid As String
Dim CreditorAccountid As String

Private Sub cboPaymentMode_Click()
If Me.cboPaymentMode = "Cheque" Then
Me.txtChequeNo.Enabled = True
Me.dtpChequeDueDate.Enabled = True
Me.lblChequeDue.Enabled = True
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

Private Sub cboTransactionDescription_KeyPress(KeyAscii As Integer)
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
  Me.LstInvoiceNo.ListItems.Clear
  Me.LstPayments.ListItems.Clear
  Me.flxgSupplier.Visible = False
  Me.txtSupplierName.SetFocus
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
   
    cn.Execute "Delete From  CreditorAccount  Where SupplierID ='" & Supplierid & "' And CreditorAccID='" & CreditorAccountid & "'", Y
    If Y > 0 Then
    cn.Execute "Delete From  CreditorPayment  Where CreditorAccID ='" & CreditorAccountid & "'", Y
    End If
    If Y > 0 Then
    cn.CommitTrans
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     
     ClearCtrls
     Call ListPayments
   Else
      cn.RollbackTrans
      MsgBox "Sorry, Unable to Delete Products Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
    
   sflag = False
   Me.cmdDelete.Enabled = False
   If cn.State = 1 Then cn.Close
   Me.flxgSupplier.Visible = False
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
frmCreditorsRpt.Show
End Sub

Private Sub cmdSave_Click()

Me.txtChequeNo.Enabled = False
Me.dtpChequeDueDate.Enabled = False
Me.lblChequeDue.Enabled = False
Me.lblChequeNo.Enabled = False

If Trim(Me.txtSupplierName) = "" Then
   MsgBox "YOU MUST ENTER SUPPLIER NAME.", vbInformation, "SUPPLIER NAME"
   Me.txtSupplierName.SetFocus: Exit Sub
End If

If Trim(Me.txtInvoiceNo) = "" Then
   MsgBox "YOU MUST ENTER INVOICE NUMBER.", vbInformation, "ENTER INVOICE NUMBER"
   Me.txtInvoiceNo.SetFocus: Exit Sub
End If


If Trim(Me.cboTransactionDescription) = "" Then
   MsgBox "YOU MUST ENTER TRANSACTION DESCRIPTION.", vbInformation, "DESCRIPTION"
   Me.cboTransactionDescription.SetFocus: Exit Sub
End If


If Trim(Me.txtAmountDue) = "" Then
   MsgBox "YOU MUST ENTER AMOUNTDUE OF GOODS ON CREDIT.", vbInformation, "AMOUNTDUE"
   Me.txtAmountDue.SetFocus: Exit Sub
End If

 If Val(Me.txtBalCD) = 0 Then
   MsgBox "THE CREDITOR FOR THIS PARTICULAR INVOICE HAS BEEN CLEARED.", vbInformation, "CREDITOR CLEARED"
    Exit Sub
End If

 
 
  If Me.txtAmountPaid <> "" And (Val(Me.txtAmountPaid) > Val(Me.txtBalCD)) Then
     MsgBox "AMOUNT PAID SHOULD NOT BE MORE THAN BAL C/D", vbInformation, ""
     Me.txtAmountPaid = "": Me.txtAmountPaid.SetFocus: Exit Sub
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
    Call Generate_CreditorAccountID(CreditorAccid)
    Call AccountCal
   rs.Open "Select * From CreditorAccount Where InvoiceNo='" & Trim(Me.txtInvoiceNo) & "' And SupplierID='" & Supplierid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount = 0 Then
     
     cn.BeginTrans
     cn.Execute "Insert Into CreditorAccount  ([SupplierID],[InvoiceNo],[TransactionDescription],[AmountDue],[Date],[BalanceCD],[CreditorAccID]) select '" & (Supplierid) & "','" & Trim(Me.txtInvoiceNo.Text) & "','" & Trim(Me.cboTransactionDescription.Text) & "','" & CCur(Val(Me.txtAmountDue.Text)) & "','" & Trim(Me.dtpTransactionDate) & "','" & Val(Me.txtBalCD.Text) & "','" & (CreditorAccid) & "'", Y
   
     If Y > 0 Then
     If Me.cboPaymentMode <> "" And Me.txtAmountPaid <> "" Then
     cn.Execute "Insert Into CreditorPayment ([CreditorAccID],[PaymentMode],[ChequeNo],[AmountPaid],[PaymentDate],[ChequeDueDate]) select '" & (CreditorAccid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
     Else
     cn.Execute "Insert Into CreditorPayment ([CreditorAccID],[PaymentMode],[ChequeNo],[AmountPaid],[PaymentDate],[ChequeDueDate]) select '" & (CreditorAccid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
     End If
     End If
     If Y > 0 Then
     cn.CommitTrans
     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
       If rs.State = 1 Then rs.Close
       
      
       Call ListPayments
       Call ListInvoice
       Call ClearCtrls
     Me.flxgSupplier.Visible = False
     Me.txtSupplierName.SetFocus
     Else
     cn.RollbackTrans
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     If rs.State = 1 Then rs.Close
     Me.txtSupplierName.SetFocus
     End If
     
 Else
     'If Me.cboPaymentMode <> "" And Me.txtAmountPaid <> "" Then
     'cn.Execute "Insert Into CreditorPayment ([CreditorAccID],[PaymentMode],[ChequeNo],[AmountPaid],[PaymentDate],[ChequeDueDate]) select '" & (CreditorAccountid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
     'If Y > 0 Then
       'MsgBox "Saved Successfully!", vbInformation, "Save Successful"
       
       If rs.State = 1 Then rs.Close
       Call ListPayments
       Call ListInvoice
       Call ClearCtrls
       Me.flxgSupplier.Visible = False
       
       'End If
     'If rs.State = 1 Then rs.Close
     'End If
    If rs.State = 1 Then rs.Close
  End If
'Else



Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub

Private Sub flxgSupplier_Click()
Me.txtSupplierName = Me.flxgSupplier.TextMatrix(Me.flxgSupplier.Row, 0)
Supplierid = Me.flxgSupplier.TextMatrix(Me.flxgSupplier.Row, 2)

Me.flxgSupplier.Visible = False
Me.txtInvoiceNo = ""
Me.cmdSave.Enabled = True
sflag = False
oflag = False
Call ListInvoice
'Me.framAllInvoice.Caption = Me.txtName & " " & "Invoices"
Me.txtInvoiceNo.SetFocus

End Sub

Private Sub flxgSupplier_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtSupplierName = Me.flxgSupplier.TextMatrix(Me.flxgSupplier.Row, 0)
Supplierid = Me.flxgSupplier.TextMatrix(Me.flxgSupplier.Row, 2)

Me.flxgSupplier.Visible = False
Me.txtInvoiceNo = ""
Me.cmdSave.Enabled = True
sflag = False
oflag = False
Call ListInvoice
'Me.framAllInvoice.Caption = Me.txtName & " " & "Invoices"
Me.txtInvoiceNo.SetFocus
End If
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Form_Load()
Me.txtChequeNo.Enabled = False
Me.dtpChequeDueDate.Enabled = False
Me.lblChequeDue.Enabled = False
Me.lblChequeNo.Enabled = False
Me.dtpTransactionDate = Date
Me.dtpPaymentDate = Date
Me.Height = 9225
  Me.Width = 9615
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2

End Sub

Private Sub LstInvoiceNo_Click()

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
rs.Open "Select * From CreditorAccount Inner Join Suppliers On CreditorAccount.SupplierID=Suppliers.SupplierID", cn, adOpenForwardOnly, adLockReadOnly


    While Not rs.EOF
        If Me.LstInvoiceNo.SelectedItem.Text = rs("InvoiceNo") Then
                Me.txtInvoiceNo.Text = Trim(rs("InvoiceNo"))
                Me.cboTransactionDescription.Text = Trim(rs("TransactionDescription"))
                Me.txtAmountDue.Text = Trim(rs("AmountDue"))
                Me.dtpTransactionDate = Trim(rs("Date"))
                Me.txtBalCD.Text = Trim(rs("BalanceCD"))
                'Me.txtBalBD.Text = Trim(rs("BalanceCD"))
                Me.txtSupplierName.Text = Trim(rs("SupplierName"))
                'Me.txtPhoneNo.Text = Trim(rs("PhoneNumber"))
                'Me.txtCreditLimit.Text = Trim(rs("CreditLimit"))
                CreditorAccountid = Trim(rs("CreditorAccID"))
            rs.MoveNext
        Else
            rs.MoveNext
        End If
    Wend
    If rs.State = 1 Then rs.Close
    Call ListPayments
    Set rs = Nothing
    sflag = True
    Me.cmdDelete.Enabled = True
    Me.cmdSave.Enabled = True
    xflag = False
    
    If Val(Me.txtBalCD) = 0 Then
      MsgBox "THE CREDITOR FOR THIS PARTICULAR INVOICE HAS BEEN CLEARED.", vbInformation, "CREDITOR CLEARED"
      
    End If

    
    Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
      xflag = False
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "TRY AGAIN"
     Me.txtSupplierName.SetFocus
     Exit Sub
End Sub

Private Sub txtAmountDue_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789/|\;:.,()@#$%&*_-'"
If KeyAscii = vbKeyReturn Then
   Me.dtpTransactionDate.SetFocus
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

Private Sub txtChequeNo_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.dtpChequeDueDate.SetFocus
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
   Me.cboTransactionDescription.SetFocus
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

rs.Open "Select * From CreditorAccount Where InvoiceNo='" & (Me.txtInvoiceNo) & "' And SupplierID='" & Supplierid & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount = 0 Then
   Me.cboTransactionDescription = ""
   Me.txtAmountDue = ""
   Me.dtpTransactionDate = Date
   Me.cboPaymentMode = ""
   Me.txtChequeNo = ""
   Me.dtpPaymentDate = Date
   
   Me.txtBalCD = ""
   Me.LstPayments.ListItems.Clear
   If rs.State = 1 Then rs.Close
 Else
   MsgBox "THE INVOICE NUMBER" & " " & Me.txtInvoiceNo & " " & "ALREADY EXIST,SELECT IT FROM THE LIST", vbInformation, "SELECT INVOICE NUMBER FROM LIST OF INVOICES OR CHANGE INVOICE NUMBER"
   Me.txtInvoiceNo = ""
   Me.txtInvoiceNo.SetFocus
   If rs.State = 1 Then rs.Close: Exit Sub
End If
Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "Invoice Numbers Displaying", vbInformation, "Displaying"
     Exit Sub

End Sub

Private Sub txtSupplierName_Change()
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
   rs.Open "Select * From Suppliers  Where SupplierName Like '" & Trim(Me.txtSupplierName) & "%" & "' Order By SupplierName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   'flxgSupplier.Height = 950 + (285 * (rs.RecordCount - 1))
   
   'If flxgSupplier.Height >= 4455 Then
      'flxgSupplier.Height = 4455
   'End If
    flxgSupplier.Rows = rs.RecordCount + 1
   With flxgSupplier
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("SupplierName")
       .TextMatrix(X, 1) = rs.Fields("PhoneNo")
       .TextMatrix(X, 2) = rs.Fields("SupplierID")
       
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 2
      .RowSel = 1
   End With
   flxgSupplier.Visible = True
   If rs.State = 1 Then rs.Close
   Me.cmdSave.Enabled = True
Else
     flxgSupplier.Visible = False
     'Dquantity = "1"
      If cFlag = True Then
       If rs.State <> 0 Then
          rs.Close
       End If
       cFlag = False: Exit Sub
      End If
     
End If
Else
      flxgSupplier.Visible = False
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
    MsgBox "Creditors to Display", vbInformation, "Displaying"
     Exit Sub
End Sub

Private Sub txtsuppliername_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.flxgSupplier.Visible = True
   Me.flxgSupplier.SetFocus
End If
End Sub
Private Sub ListInvoice()
On Error GoTo SaveError
Me.LstInvoiceNo.ListItems.Clear
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


rs.Open "Select * From CreditorAccount Where SupplierID='" & (Supplierid) & "' Order By InvoiceNo", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.LstInvoiceNo.ListItems.Add(, , Trim(rs!InvoiceNo))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!TransactionDescription)
                List_Item.SubItems(2) = Trim(rs!AmountDue)
                List_Item.SubItems(3) = Trim(rs!Date)
                'List_Item.SubItems(4) = Trim(rs!Date)
                'List_Item.SubItems(5) = Trim(rs!AccID)
                 'List_Item.SubItems(6) = Trim(rs!BalanceCD)
                rs.MoveNext
            Loop
        End If
    Next i
    DoEvents
    
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    If cn.State = 1 Then cn.Close
    Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "SAVE FAILED"
     Exit Sub
    
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


rs.Open "Select CreditorPayment.* From CreditorPayment Inner Join CreditorAccount On CreditorPayment.CreditorAccID=CreditorAccount.CreditorAccID Where CreditorAccount.InvoiceNo='" & Trim(Me.txtInvoiceNo) & "' And SupplierID='" & (Supplierid) & "' Order By PaymentDate", cn, adOpenForwardOnly, adLockReadOnly
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
                'List_Item.SubItems(4) = Trim(rs!Arrears)
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
Private Function Generate_CreditorAccountID(Account_ID As String) As Boolean
Dim strg As String, strg1 As String, strg2 As String, bFlag As Boolean
On Error GoTo SaveError
rs.Open "Select CreditorAccID From CreditorAccount  order by CreditorAccID Desc", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
      strg1 = Trim(rs.Fields!CreditorAccid)

   strg1 = Trim(Str(Val(strg1) + 1))
   strg1 = String$(8 - Len(strg1), "0") & strg1
Else
   strg1 = "00000001"
End If
Account_ID = strg1

If rs.State = 1 Then rs.Close

Generate_CreditorAccountID = True

Exit Function
SaveError:
     If rs.State = 1 Then rs.Close
    Generate_CreditorAccountID = False
     Exit Function
     
End Function
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
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
       MsgBox "THE CREDITOR FOR THIS PARTICULAR INVOICE HAS BEEN CLEARED", vbInformation, "PAYMENT COMPLETE"
  End If


 rs.Open "Select * From CreditorAccount Where InvoiceNo='" & Trim(Me.txtInvoiceNo) & "' And SupplierID='" & (Supplierid) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
    BalCD = rs.Fields("BalanceCD") - Val(Me.txtAmountPaid)
    Me.txtBalCD = BalCD
  cn.BeginTrans
  If Me.cboPaymentMode <> "" And Me.txtAmountPaid <> "" Then
     cn.Execute "Insert Into CreditorPayment ([CreditorAccID],[PaymentMode],[ChequeNo],[AmountPaid],[PaymentDate],[ChequeDueDate]) select '" & (CreditorAccountid) & "','" & Trim(Me.cboPaymentMode.Text) & "','" & Trim(Me.txtChequeNo.Text) & "','" & Val(Me.txtAmountPaid.Text) & "','" & Trim(Me.dtpPaymentDate) & "','" & Trim(Me.dtpChequeDueDate) & "'", Y
       If rs.State = 1 Then rs.Close
  End If
       If rs.State = 1 Then rs.Close
     'End If
  
   If Y > 0 Then
    cn.Execute "Update CreditorAccount Set BalanceCD ='" & BalCD & "' Where InvoiceNo ='" & Trim(Me.txtInvoiceNo) & "' And SupplierID='" & (Supplierid) & "'", Y
    If rs.State = 1 Then rs.Close
   End If
   If Y > 0 Then
     cn.CommitTrans
     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
     'Call ListPayments
     'Call ListInvoice
     'Call ClearCtrls
     'Me.flxgSupplier.Visible = False
   Else
     cn.RollbackTrans
     MsgBox "Saved Failed!", vbInformation, "Try Again"
     End If
Else
  If Me.txtAmountDue <> "" Then
    BalCD = Val(Me.txtAmountDue) - Val(Me.txtAmountPaid)
    Me.txtBalCD = BalCD
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
