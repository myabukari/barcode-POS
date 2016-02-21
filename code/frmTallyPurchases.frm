VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C29E21&
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   14805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   8280
      Width           =   8655
      Begin lvButton_H.lvButtons_H cmdSave 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Ok"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         Image           =   "frmTallyPurchases.frx":0000
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdClear 
         Height          =   375
         Left            =   3480
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Clear"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         Image           =   "frmTallyPurchases.frx":0E82
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdPrint 
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Print"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         Image           =   "frmTallyPurchases.frx":2733
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   6720
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Exit"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         Image           =   "frmTallyPurchases.frx":2845
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdStock 
         Height          =   375
         Left            =   5040
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&View Stock"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         Image           =   "frmTallyPurchases.frx":2C97
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   6255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   8655
      Begin Crystal.CrystalReport CrystalInvoice 
         Left            =   360
         Top             =   5280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtReceipt 
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
         Left            =   6360
         TabIndex        =   44
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtParticulars 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   42
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox txtInitials 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   39
         Top             =   5760
         Width           =   2415
      End
      Begin VB.ComboBox cboRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTallyPurchases.frx":4501
         Left            =   5400
         List            =   "frmTallyPurchases.frx":450E
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   5760
         Width           =   2895
      End
      Begin VB.ComboBox cboClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTallyPurchases.frx":452D
         Left            =   1680
         List            =   "frmTallyPurchases.frx":452F
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   5160
         Width           =   2415
      End
      Begin VB.ComboBox cboPaymode 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmTallyPurchases.frx":4531
         Left            =   1680
         List            =   "frmTallyPurchases.frx":453B
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Cash"
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox txtBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox txtAmountPaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox txtNetCost 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1680
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3360
         Width           =   3735
      End
      Begin VB.OptionButton OptConvert 
         BackColor       =   &H00C29E21&
         Caption         =   "Old Currency"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   6360
         TabIndex        =   13
         Top             =   3240
         Width           =   1935
      End
      Begin VB.OptionButton OptConvert 
         BackColor       =   &H00C29E21&
         Caption         =   "New Currency"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   6360
         TabIndex        =   12
         Top             =   3600
         Width           =   1815
      End
      Begin Crystal.CrystalReport CrystalReceipt 
         Left            =   5640
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgCosts 
         Height          =   2415
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   720
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4260
         _Version        =   393216
         BackColor       =   15388531
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483628
         BackColorSel    =   -2147483647
         BackColorBkg    =   15388531
         GridColor       =   -2147483628
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
         FormatString    =   $"frmTallyPurchases.frx":454D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin lvButton_H.lvButtons_H cmdAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Add"
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
         Image           =   "frmTallyPurchases.frx":45EF
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdRemove 
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Remove"
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
         Image           =   "frmTallyPurchases.frx":5DB1
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdCompute 
         Height          =   375
         Left            =   3240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Compute"
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
         Image           =   "frmTallyPurchases.frx":5F4B
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   5400
         TabIndex        =   36
         Top             =   4560
         Width           =   2895
         _ExtentX        =   5106
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
         CalendarBackColor=   15718770
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   61014019
         CurrentDate     =   39091
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C29E21&
         Caption         =   "Particulars:"
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
         Left            =   4320
         TabIndex        =   43
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C29E21&
         Caption         =   "Initials:"
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
         Left            =   960
         TabIndex        =   41
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C29E21&
         Caption         =   "Remarks:"
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
         Left            =   4440
         TabIndex        =   40
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label8 
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
         Left            =   4800
         TabIndex        =   37
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C29E21&
         Caption         =   "Payment Mode:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   35
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C29E21&
         Caption         =   "Client:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C29E21&
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C29E21&
         Caption         =   "Amount Paid:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   24
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label10 
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
         Left            =   5280
         TabIndex        =   23
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C29E21&
         Caption         =   "Net Cost:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   3480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8655
      Begin VB.TextBox txtTotalCost 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cboQuantity 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTallyPurchases.frx":65C5
         Left            =   1680
         List            =   "frmTallyPurchases.frx":6626
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   6615
      End
      Begin VB.TextBox txtUnitPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Product Name:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "Discount(%):"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Total Cost:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Cartone Price:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000F&
         BorderStyle     =   6  'Inside Solid
         X1              =   0
         X2              =   11280
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000F&
         BorderStyle     =   6  'Inside Solid
         X1              =   0
         X2              =   11280
         Y1              =   2760
         Y2              =   2760
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgtry 
      Height          =   8055
      Left            =   8880
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   14208
      _Version        =   393216
      BackColor       =   15388531
      ForeColor       =   0
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   -2147483628
      BackColorSel    =   -2147483647
      BackColorBkg    =   15388531
      GridColor       =   -2147483639
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      FormatString    =   $"frmTallyPurchases.frx":66A0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, Productid As String, X As Integer, QuantityLeft As Integer
Dim sflag As Boolean, strg2 As String, strg3 As String, xx As Integer, xflag As Boolean, yflag As Boolean, eflag As Boolean, a As Integer
Dim Fval2 As Integer, val2 As Integer, Frow2 As Integer, val3 As Integer, i As Variant, Y As Variant, b As Integer, c As Integer, oflag As Boolean
Dim clearflag As Boolean, UnitPriceflag As Boolean, nameflag As Boolean, ReceiptNo As Integer
Dim NetCostflag As Boolean, ClientIDs() As Integer, Balance As Integer, ClientID As Integer
Private Sub cboClient_Click()
If Me.cboClient <> "NONE" Then
  ClientID = ClientIDs(Me.cboClient.ListIndex)
End If
End Sub

Private Sub cboClient_DropDown()
On Error GoTo SaveError

Me.cboClient.Clear
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

rs.Open "Select * From Debtors Order by DebtorName", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
   ReDim ClientIDs(rs.RecordCount - 1)
'   ReDim prefixs(rs.RecordCount - 1)
   For X = 0 To rs.RecordCount - 1
    Me.cboClient.AddItem rs.Fields("DebtorName")
    ClientIDs(X) = rs.Fields("DebtorID")
'    prefixs(X) = rs.Fields("CoursePrefix")
    rs.MoveNext
   Next
   
   If cboClient.ListCount > 1 Then
      Me.cboClient.AddItem "NONE"
   End If
   
Else
   MsgBox "No Clients has been Setup Yet - Setup Client(s).", vbInformation, ""
End If


If cn.State = 1 Then cn.Close
If rs.State = 1 Then rs.Close



Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Display Client(s):Please Try Again!", vbInformation, ""
     Exit Sub
End Sub

Private Sub cboQuantity_KeyPress(KeyAscii As Integer)
Dim strk1 As String

strk1 = "0123456789."
'If KeyAscii = vbKeyReturn Then
'   Me.txtTotalCost.SetFocus
'End If
If KeyAscii = 67 Or KeyAscii = 99 Then
  Call Clear
End If
If KeyAscii = 43 Or KeyAscii = vbKeyReturn Then
   Call cmdAdd_Click
End If

If KeyAscii = 92 Then
    If Me.OptConvert(0) Then
        Me.OptConvert(1).SetFocus
    ElseIf Me.OptConvert(1) Then
        Me.OptConvert(0).SetFocus
    Else
        Me.OptConvert(1).SetFocus
        Me.cboQuantity.SetFocus
    Exit Sub
    End If
        Me.cboQuantity.SetFocus
    
    If Me.txtNetCost <> "" Then
        NetCostflag = True
        txtNetCost = CurrencyConvertor(Trim(txtNetCost))
        NetCostflag = False
    End If
    If Me.txtAmountPaid <> "" Then
        txtAmountPaid = CurrencyConvertor(Trim(txtAmountPaid))
    End If
    
      MsgBox "Currency converted", vbInformation, "Converter"
      Me.cboQuantity.SetFocus
End If

If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub cmdAdd_Click()
Dim X As Integer, a As Integer
On Error GoTo OkError
'If eflag = False Then
'For xx = 1 To flxgCosts.Rows - 1
'           If flxgCosts.TextMatrix(xx, 0) = Trim(Me.txtName) And flxgCosts.TextMatrix(xx, 7) = Productid Then
'              MsgBox flxgCosts.TextMatrix(xx, 0) & " Has Already Been Added!: You Can Only Edit What is Entered"
'              UnitPriceflag = True
'              nameflag = True
'              Me.txtName = ""
'              Me.txtUnitPrice = ""
'              Me.txtCode = ""
'              Me.txtTotalCost = ""
'              Me.txtCode.SetFocus
'              UnitPriceflag = False
'              nameflag = False
'              Exit Sub
'           End If
'         Next
'End If
Call CalTotalPrice
If Trim(Me.txtname) = "" Then MsgBox "Please Select Product", vbInformation, "": Me.txtname.SetFocus: Exit Sub
    If Trim(Me.cboQuantity) = "" Then MsgBox "Please Enter Quantity of Product", vbInformation, "": Me.cboQuantity.SetFocus: Exit Sub
    If Val(Trim(Me.txtUnitPrice)) = 0 Then MsgBox "Please Enter Price of Product", vbInformation, "": Me.txtUnitPrice.SetFocus: Exit Sub
    If Trim(Me.txtTotalCost) = "" Then MsgBox "Please Compute TotalCost", vbInformation, "": Me.txtTotalCost.SetFocus: Exit Sub
    If Frow2 <= 0 Then
      If Fval2 > 0 Then
         For xx = 1 To flxgCosts.Rows - 1
           If flxgCosts.TextMatrix(xx, 0) = Trim(Me.txtname) And flxgCosts.TextMatrix(xx, 5) = Productid Then
              MsgBox flxgCosts.TextMatrix(xx, 0) & " Has Already Been Added!: You Can Only Edit What is Entered"
              Me.txtname = ""
'              Me.txtCode = ""
              Me.txtUnitPrice = ""
              Me.txtTotalCost = ""
              Me.txtname.SetFocus
              Exit Sub
           End If
         Next
       End If
      Fval2 = Fval2 + 1
      flxgCosts.Rows = Fval2 + 1
      flxgCosts.TextMatrix(Fval2, 0) = Me.txtname
      flxgCosts.TextMatrix(Fval2, 1) = Me.cboQuantity
'      If Me.cboSellCartone = "NO" Then
      flxgCosts.TextMatrix(Fval2, 2) = Me.txtUnitPrice
'      Else
'      flxgCosts.TextMatrix(Fval2, 2) = Me.txtCartonePrice
'      End If
      flxgCosts.TextMatrix(Fval2, 3) = Me.txtTotalCost
      flxgCosts.TextMatrix(Fval2, 4) = Me.txtDiscount
'      flxgCosts.TextMatrix(Fval2, 5) = Me.txtVat
''      flxgCosts.TextMatrix(Fval2, 5) = Me.txtPriceAfterDiscount
      flxgCosts.TextMatrix(Fval2, 5) = Trim(Productid)
'       flxgCosts.TextMatrix(Fval2, 8) = Trim(Me.cboSellCartone)
'       flxgCosts.TextMatrix(Fval2, 9) = Trim(Me.txtCartonePrice)
       
    Else
        flxgCosts.TextMatrix(Frow2, 0) = Me.txtname
        flxgCosts.TextMatrix(Frow2, 1) = Me.cboQuantity
'        If Me.cboSellCartone = "NO" Then
        flxgCosts.TextMatrix(Fval2, 2) = Me.txtUnitPrice
'        Else
'        flxgCosts.TextMatrix(Fval2, 2) = Me.txtCartonePrice
'        End If
        flxgCosts.TextMatrix(Frow2, 3) = Me.txtTotalCost
        flxgCosts.TextMatrix(Frow2, 4) = Me.txtDiscount
'        flxgCosts.TextMatrix(Frow2, 5) = Me.txtVat
'        flxgCosts.TextMatrix(Frow2, 5) = Me.txtPriceAfterDiscount
        flxgCosts.TextMatrix(Frow2, 5) = Trim(Productid)
'        flxgCosts.TextMatrix(Frow2, 8) = Trim(Me.cboSellCartone)
'        flxgCosts.TextMatrix(Frow2, 9) = Trim(Me.txtCartonePrice)
      Frow2 = 0
    
    End If
    UnitPriceflag = True
    'nameflag = True
    Me.txtname = ""
    Me.txtUnitPrice = ""
    Me.txtTotalCost = ""
    Me.txtDiscount = ""
'    Me.txtPriceAfterDiscount = ""
'    Me.txtVat = ""
'    Me.cboDepartment.Text = ""
'    Me.txtCode = ""
    Me.txtname.SetFocus
    Me.cmdSave.Enabled = True
    Me.cmdCompute.Enabled = True
    UnitPriceflag = False
    nameflag = False
    i = 0 'Reset the value i in cmdcompute to recompute netsales
    
    
    
    
    For X = 1 To Me.flxgCosts.Rows - 1
        If Trim(flxgCosts.TextMatrix(X, 3)) <> "" Then
            'a = Len(Trim(flxgCosts.TextMatrix(X, 3))) - 1
        Else
            Me.txtNetCost = "0.00"
            Exit Sub
        End If
        i = i + CDbl(flxgCosts.TextMatrix(X, 3))
    Next
    Me.txtNetCost.Text = Format$(i, "#,###.00")
    Me.cmdCompute.Enabled = True
    i = 0
    
    oflag = False
    If eflag = True Then
      eflag = False
    End If
'     Me.cboSellCartone.Text = "NO"
   
   Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
     Exit Sub
End Sub

Private Sub cmdPrint_Click()
Call cmdSave_Click
PrintReceipt
End Sub

Private Sub cmdRemove_Click()
Dim a As Integer, strgbal As String

On Error GoTo OkError

 If Frow2 = 0 Then Exit Sub
 
  If Frow2 = 1 Then
   If flxgCosts.TextMatrix(Frow2, 0) = "" Then
   Frow2 = 0
   Me.txtname.Text = ""
   Me.txtUnitPrice = ""
   Me.txtTotalCost = ""
   Me.txtname.SetFocus
   Exit Sub
   End If
 End If
 
 
If Trim(Me.txtNetCost) <> "" Then
    'a = Len(Trim(Me.txtNetCost)) - 1
    'b = Len(Trim(flxgCosts.TextMatrix(Frow2, 3))) - 1
    strgbal = CDbl(Me.txtNetCost) - CDbl(flxgCosts.TextMatrix(Frow2, 3))
    Me.txtNetCost.Text = Format$(strgbal, "#,###.00")
    Me.cmdCompute.Enabled = False
End If
   If Frow2 = flxgCosts.Rows - 1 Then
     If flxgCosts.Rows <> 2 Then
         flxgCosts.Rows = flxgCosts.Rows - 1
     Else
        For xx = 0 To 5
           flxgCosts.TextMatrix(Frow2, xx) = ""
        Next
        Me.txtname.Text = ""
        Me.txtUnitPrice = ""
        Me.txtTotalCost = ""
        Me.txtname.SetFocus
     End If
        Me.txtname.Text = ""
        Me.txtUnitPrice = ""
        Me.txtTotalCost = ""
        Me.txtname.SetFocus
   Else
         For xx = Frow2 To flxgCosts.Rows - 2
            flxgCosts.TextMatrix(xx, 0) = flxgCosts.TextMatrix(xx + 1, 0)
            flxgCosts.TextMatrix(xx, 1) = flxgCosts.TextMatrix(xx + 1, 1)
            flxgCosts.TextMatrix(xx, 2) = flxgCosts.TextMatrix(xx + 1, 2)
            flxgCosts.TextMatrix(xx, 3) = flxgCosts.TextMatrix(xx + 1, 3)
            flxgCosts.TextMatrix(xx, 4) = flxgCosts.TextMatrix(xx + 1, 4)
            flxgCosts.TextMatrix(xx, 5) = flxgCosts.TextMatrix(xx + 1, 5)
'            flxgCosts.TextMatrix(xx, 6) = flxgCosts.TextMatrix(xx + 1, 6)
'            flxgCosts.TextMatrix(xx, 7) = flxgCosts.TextMatrix(xx + 1, 7)
'            flxgCosts.TextMatrix(xx, 8) = flxgCosts.TextMatrix(xx + 1, 8)
'            flxgCosts.TextMatrix(xx, 9) = flxgCosts.TextMatrix(xx + 1, 9)
        Next
        Me.txtname.Text = ""
        Me.txtUnitPrice = ""
        Me.txtTotalCost = ""
        Me.txtname.SetFocus
        flxgCosts.Rows = flxgCosts.Rows - 1
   End If
  Fval2 = Fval2 - 1
  Frow2 = 0
  Me.cmdCompute.Enabled = True
  oflag = False
   Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "SORRY,UNABLE TO REMOVE,TRY AGAIN", vbInformation, "COMPUTATION"
     Exit Sub
End Sub

Private Sub cmdSave_Click()
'On Error GoTo OkError
'If Me.txtname = "" Then
'MsgBox "SELECT A PRODUCT", vbInformation
'Me.txtname.SetFocus: Exit Sub
'End If
For X = 1 To Me.flxgCosts.Rows - 1
   If Me.flxgCosts.TextMatrix(X, 0) = "" Then
     MsgBox "SELECT A PRODUCT", vbInformation
     Me.txtname.SetFocus: Exit Sub
   End If
Next

If Me.txtNetCost = "" Then
  MsgBox "COMPUTE NET COST", vbInformation
  Exit Sub
End If

If Me.txtAmountPaid = "" Then
  MsgBox "INPUT AMOUNTPAID", vbInformation
  Me.txtAmountPaid.SetFocus
  Exit Sub
End If

If Trim(Me.txtNetCost) <> "" Then
  'a = Len(Trim(Me.txtNetCost)) - 1
  If CDbl(Me.txtNetCost) > Me.txtAmountPaid Then
    MsgBox "AMOUNT PAID CANNOT BE LESS THAN NETCOST", vbInformation, " CHECK AMOUNT PAID"
    Me.txtAmountPaid.SetFocus: Exit Sub
  End If
End If

If Me.cboClient = "" Then
  MsgBox "PLEASE ENTER NAME OF CLIENT", vbInformation, ""
  Me.cboClient.SetFocus
  Exit Sub
End If

If Me.txtParticulars = "" Then
  MsgBox "PLEASE ENTER INVOICE NUMBER OF CLIENT", vbInformation, ""
  Randomize
  Me.txtParticulars = Round(10000000 * Rnd(), 0) & "/" & Round(1000 * Rnd(), 0)
  Me.txtParticulars.SetFocus
  Exit Sub
End If

Me.txtReceipt = ""
Me.txtReceipt = Me.txtParticulars

'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
    
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   'If rs.State = 1 Then rs.Close
   cn.BeginTrans
   For X = 1 To Me.flxgCosts.Rows - 1
   rs.Open "Select * From Products Inner Join Tally On Products.ProductID=Tally.ProductID Where Tally.ProductID= '" & Trim(Me.flxgCosts.TextMatrix(X, 5)) & "' order by IssueDate desc,IssueTime desc", cn, adOpenForwardOnly, adLockReadOnly
        'For i = 1 To rs.RecordCount
    If rs.RecordCount > 0 Then
'        If Trim(Me.flxgCosts.TextMatrix(X, 8)) = "NO" Then
          
          If rs.Fields("Balance") >= Val(Trim(Me.flxgCosts.TextMatrix(X, 1))) Then
            Call ComputeBalance(Trim(Me.flxgCosts.TextMatrix(X, 5)), Balance, , Trim(Me.flxgCosts.TextMatrix(X, 1)))
'           QuantityLeft = rs.Fields("Balance") - Val(Trim(Me.flxgCosts.TextMatrix(X, 1)))
'           cn.Execute "Update Tally Set Balance ='" & QuantityLeft & "' Where ProductID ='" & Trim(Me.flxgCosts.TextMatrix(X, 5)) & "'", Y
            cn.Execute "Insert Into Tally ([ProductID],[IssueDate],[Particulars],ReceivedIn,[IssueOut],[Balance],[Initials],[Remarks],IssueTime,Client) select '" & Trim(Me.flxgCosts.TextMatrix(X, 5)) & "','" & Trim(Me.dtpDate) & "','" & Trim(Me.txtParticulars.Text) & "','0','" & Trim(Me.flxgCosts.TextMatrix(X, 1)) & "','" & Balance & "','" & Trim(Me.txtInitials) & "','" & Trim(Me.cboRemarks.Text) & "','" & Format$(dtpDate, "Medium Date") & " " & Format$(Now, "Medium Time") & "','" & Trim(Me.cboClient) & "'", Y
            If Y < 0 Then
             MsgBox "UNABLE TO COMPUTE,PLEASE TRY AGAIN", vbiformation, "TRY AGAIN"
             Me.txtname.SetFocus: rs.Close: Exit Sub
            End If
          Else
            Y = 1
          
'         End If
          If rs.State = 1 Then rs.Close
       End If
          If rs.State = 1 Then rs.Close
          
    Else
        MsgBox "The Product Selected is not in Stock", vbInformation, ""
        If rs.State = 1 Then rs.Close
        Exit Sub
    End If
    'Next
    
   
     'For X = 1 To Me.flxgCosts.Rows - 1
'       If Trim(flxgCosts.TextMatrix(X, 3)) <> "" Then
'         a = Len(Trim(flxgCosts.TextMatrix(X, 3))) - 1
'       End If
     
    If Y > 0 Then
        rs.Open "Select * From ClientPurchases Where ProductID= '" & Trim(Me.flxgCosts.TextMatrix(X, 5)) & "' And PurchaseDate='" & Date & "' And DebtorID= '" & ClientID & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then 'And rs.Fields("UnitPrice") = Val(Trim(Me.flxgCosts.TextMatrix(X, 2))) Then
                AdjQty = rs.Fields("Quantity") + Val(Trim(Me.flxgCosts.TextMatrix(X, 1)))
'                AdjTotalCost = rs.Fields("TotalCost") + CDbl((Trim(Me.flxgCosts.TextMatrix(X, 3))))
            cn.Execute "Update ClientPurchases Set Quantity ='" & AdjQty & "' Where ProductID ='" & Trim(Me.flxgCosts.TextMatrix(X, 5)) & "'and PurchaseDate='" & Date & "' And DebtorID= '" & ClientID & "'", Y
                yflag = True
                If rs.State = 1 Then rs.Close
            
         Else
             cn.Execute "Insert Into ClientPurchases ([ProductID],[Quantity],[CartonePrice],InvoiceNo,PurchaseDate,DebtorID) select '" & Trim(Me.flxgCosts.TextMatrix(X, 5)) & "','" & Val(Trim(Me.flxgCosts.TextMatrix(X, 1))) & "','" & CDbl(Trim(Me.flxgCosts.TextMatrix(X, 2))) & "','" & Trim(Me.txtParticulars) & "','" & Me.dtpDate & "','" & ClientID & "'", Y
             yflag = True
              If rs.State = 1 Then rs.Close
        End If
          If rs.State = 1 Then rs.Close
     End If
     Next
     If Y > 0 Then

             cn.Execute "Insert Into ClientPayments ([InvoiceNo],[NetCost],[AmountPaid],Balance,PaymentDate,PaymentMode,DebtorID,ClientName) select '" & Trim(Me.txtParticulars) & "','" & CDbl(Trim(Me.txtNetCost)) & "','" & CDbl(Trim(Me.txtAmountPaid)) & "','" & CDbl(Trim(Me.txtBalance)) & "','" & Me.dtpDate & "','" & Trim(Me.cboPaymode) & "','" & ClientID & "','" & Trim(Me.cboClient) & "'", Y

     End If
     
   
   
      If Y > 0 Then
        cn.CommitTrans
        Fval2 = 0
      Else
       cn.RollbackTrans
       MsgBox "Save Failed,Please Retype Data and Save Again", vbInformation, "Try Again"
      End If
     
    
  


'   cn.BeginTrans
'  rs.Open "Select * From Payments ", cn, adOpenForwardOnly, adLockReadOnly
'
'            If rs.RecordCount > 0 Then
'               ReceiptNo = rs.Fields("ReceiptNo")
'               cn.Execute "Update Payments Set AmountPaid ='" & CDbl(Trim(Me.txtAmountPaid)) & "',PmtMode='Cash',Balance ='" & CDbl((Trim(Me.txtBalance))) & "',NetCost ='" & CDbl((Trim(Me.txtNetCost))) & "',Date='" & Date & "' Where ReceiptNo ='" & ReceiptNo & "'", Y
'               If rs.State = 1 Then rs.Close
'            Else
'              cn.Execute "Insert Into Payments ([AmountPaid],[PmtMode],[Balance],[NetCost],[Date],[UserID],ReceiptNo) select '" & CDbl(Trim(Me.txtAmountPaid)) & "','" & Trim(Me.cboPaymode) & "','" & CDbl((Trim(Me.txtBalance))) & "','" & CDbl((Trim(Me.txtNetCost))) & "','" & Date & "','111','1'", Y
'        '      If cn.State = 1 Then cn.Close
'              If rs.State = 1 Then rs.Close
'            End If
'                If rs.State = 1 Then rs.Close
'
'     If Y > 0 Then
'
'            rs.Open "Select * From [Receipt/Product] ", cn, adOpenForwardOnly, adLockReadOnly
'            If rs.RecordCount > 0 Then
'                cn.Execute "Delete From [Receipt/Product] Where ReceiptNo ='1'", Y
'                For X = 1 To Me.flxgCosts.Rows - 1
'               cn.Execute "Insert Into [Receipt/Product] ([ProductID],Quantity,[ReceiptNo],UnitPrice) select '" & Trim(Me.flxgCosts.TextMatrix(X, 7)) & "','" & Trim(Me.flxgCosts.TextMatrix(X, 1)) & "','1','" & CDbl(Me.flxgCosts.TextMatrix(X, 2)) & "'", Y
'                Next
'               If rs.State = 1 Then rs.Close
'            Else
'                For X = 1 To Me.flxgCosts.Rows - 1
'               cn.Execute "Insert Into [Receipt/Product] ([ProductID],Quantity,[ReceiptNo],UnitPrice) select '" & Trim(Me.flxgCosts.TextMatrix(X, 7)) & "','" & Trim(Me.flxgCosts.TextMatrix(X, 1)) & "','1','" & Trim(Me.flxgCosts.TextMatrix(X, 2)) & "'", Y
'                Next
'                If rs.State = 1 Then rs.Close
'            End If
'                If rs.State = 1 Then rs.Close
'
'    End If
    
'    If Y > 0 Then
'        cn.CommitTrans
'    Else
'        cn.RollbackTrans
'    End If
'
        
   If yflag = True Then
   ' i = Me.flxgCosts.Rows - 1
      For X = 1 To Me.flxgCosts.Rows - 1
            flxgCosts.TextMatrix(X, 0) = ""
            flxgCosts.TextMatrix(X, 1) = ""
            flxgCosts.TextMatrix(X, 2) = ""
            flxgCosts.TextMatrix(X, 3) = ""
            flxgCosts.TextMatrix(X, 4) = ""
            flxgCosts.TextMatrix(X, 5) = ""
'            flxgCosts.TextMatrix(X, 6) = ""
'            flxgCosts.TextMatrix(X, 7) = ""
'            flxgCosts.TextMatrix(X, 8) = ""
'            flxgCosts.TextMatrix(X, 9) = ""
       Next
    Fval2 = 0
    Me.flxgCosts.Rows = 2
    yflag = False
   End If
      
      
      Me.cmdSave.Enabled = False
      Me.cmdCompute.Enabled = True
'      Me.flxgInventory.Visible = False
      oflag = True
      Me.txtDiscount = "": Me.txtUnitPrice = ""
      Me.txtTotalCost = "": Me.cboQuantity = 1
      Me.txtAmountPaid = "": Me.txtNetCost = "": Me.txtBalance = ""
      
      Me.OptConvert(1).SetFocus
      Me.txtname.SetFocus
   Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "SORRY,THERE IS AN ERROR,TRY AGAIN", vbInformation, "COMPUTATION"
     Exit Sub
End Sub

Private Sub flxgCosts_Click()
If MsgBox("ARE YOU SURE  YOU WANT TO EDIT OR REPLACE THE PRODUCT CLICKED?", vbYesNo + vbQuestion, "CONFIRMATION") = vbYes Then
        xflag = True
        Me.txtname = flxgCosts.TextMatrix(flxgCosts.Row, 0)
        Me.cboQuantity = flxgCosts.TextMatrix(flxgCosts.Row, 1)
        Me.txtUnitPrice = flxgCosts.TextMatrix(flxgCosts.Row, 2)
        Me.txtTotalCost = flxgCosts.TextMatrix(flxgCosts.Row, 3)
        Me.txtDiscount = flxgCosts.TextMatrix(flxgCosts.Row, 4)
'        Me.txtVat = flxgCosts.TextMatrix(flxgCosts.Row, 5)
'        Me.txtPriceAfterDiscount = flxgCosts.TextMatrix(flxgCosts.Row, 6)
         'Trim(Productid) = flxgCosts.TextMatrix(flxgCosts.Row, 7)
'        Me.cboSellCartone = flxgCosts.TextMatrix(flxgCosts.Row, 8)
'        Me.txtCartonePrice = flxgCosts.TextMatrix(flxgCosts.Row, 9)
Frow2 = flxgCosts.Row
eflag = True
oflag = False
Me.cmdSave.Enabled = True
Me.cboQuantity.SetFocus
End If
End Sub

Private Sub flxgtry_Click()
'Productflag = True
Me.txtname = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 0)
'Productflag = False
Me.txtUnitPrice = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 1)
If Val(Trim(Me.txtUnitPrice)) = 0 Then
Me.txtUnitPrice = ""
Me.txtUnitPrice.SetFocus
End If
Me.txtDiscount = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 2)
'Me.txtCartonePrice = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 3)
' = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 4)
'Me.txtDiscount = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 5)
Productid = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 5)

'Me.flxgtry.Visible = False

Me.cmdSave.Enabled = True
sflag = True
oflag = False
Me.txtname.SetFocus

i = 0 'Reset the value i in cmdcompute to recompute netsales
If Val(Trim(Me.txtUnitPrice)) = 0 Then
Me.txtUnitPrice.SetFocus
Else
Me.cboQuantity.SetFocus
End If
End Sub

Private Sub flxgtry_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
'Productflag = True
Me.txtname = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 0)
'Productflag = False
Me.txtUnitPrice = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 1)
If Val(Trim(Me.txtUnitPrice)) = 0 Then
Me.txtUnitPrice = ""
Me.txtUnitPrice.SetFocus
End If
Me.txtDiscount = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 2)
'Me.txtCartonePrice = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 3)
' = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 4)
'Me.txtDiscount = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 5)
Productid = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 5)

'Me.flxgtry.Visible = False

Me.cmdSave.Enabled = True
sflag = True
oflag = False
Me.txtname.SetFocus

i = 0 'Reset the value i in cmdcompute to recompute netsales
If Val(Trim(Me.txtUnitPrice)) = 0 Then
Me.txtUnitPrice.SetFocus
Else
Me.cboQuantity.SetFocus
End If
End If
End Sub

Private Sub Form_Activate()
Me.OptConvert(1).SetFocus
Me.txtname.SetFocus
End Sub

Private Sub Form_Initialize()
'Me.OptConvert(1).SetFocus
'Me.txtname.SetFocus
End Sub

Private Sub Form_Load()
Me.dtpDate = Date
End Sub

Private Sub OptConvert_Click(Index As Integer)
If Me.txtNetCost <> "" Then
NetCostflag = True
txtNetCost = CurrencyConvertor(Trim(txtNetCost))
NetCostflag = False
End If
If Me.txtAmountPaid <> "" Then
txtAmountPaid = CurrencyConvertor(Trim(txtAmountPaid))
End If
End Sub

Private Sub txtAmountPaid_Change()
If Me.txtAmountPaid = "" Then
Me.txtBalance = ""
Exit Sub
End If
If Me.txtNetCost <> "" Then
    Dim a As Integer, strgbal As String
    If Trim(Me.txtNetCost) <> "" Then
'        a = Len(Trim(Me.txtNetCost)) - 1
    End If
    If CDbl((Me.txtAmountPaid)) >= CDbl(Me.txtNetCost) Then
        strgbal = CDbl((Me.txtAmountPaid)) - CDbl(Me.txtNetCost)
        Me.txtBalance.Text = Format$(strgbal, "#,###.00")
        
    Else
        Me.txtBalance = ""
    End If
End If
End Sub

Private Sub txtAmountPaid_GotFocus()
Randomize
 Me.txtParticulars = Round(10000000 * Rnd(), 0) & "/" & Round(1000 * Rnd(), 0)
End Sub

Private Sub txtAmountPaid_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   If Me.cmdSave.Enabled = True Then
    Me.cmdSave.SetFocus
   End If
End If
If KeyAscii = 67 Or KeyAscii = 99 Then
  Call Clear
End If
If KeyAscii = 80 Or KeyAscii = 112 Then
'  Call Receipt
'  Call Save
'  PrintReceipt
End If


If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub txtName_Change()
On Error GoTo OkError

If Me.txtname = "\" Then
'        MsgBox "Currency converted", vbInformation, "Converter"
Me.txtname = ""
Exit Sub
End If

'Open Connecttion to Server
'   Me.OptConvert(1).SetFocus
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   
   If nameflag = False Then
   If rs.State = 1 Then rs.Close
If xflag = False Then  ' when true u can view grid to add products,but when false u can edit with no grid shown
   rs.Open "Select Distinct  Products.*, ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where ProductName Like '" & Trim(Me.txtname) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
   'flxgInventory.Height = 950 + (285 * (rs.RecordCount - 1))
   
   'If flxgInventory.Height >= 4455 Then
      'flxgInventory.Height = 4455
   'End If
    flxgtry.Rows = rs.RecordCount + 1
   With flxgtry
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("ProductName")
       .TextMatrix(X, 1) = rs.Fields("UnitPrice")
       .TextMatrix(X, 2) = rs.Fields("Discount")
       .TextMatrix(X, 3) = rs.Fields("StockLevel")
       .TextMatrix(X, 4) = rs.Fields("ReorderLevel")
       .TextMatrix(X, 5) = rs.Fields("ProductID")
      
       
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 5
      .RowSel = 1
   End With
   flxgtry.Visible = True
   If rs.State = 1 Then rs.Close
   Me.cmdSave.Enabled = True
Else
     'flxgtry.Visible = False
     'Dquantity = "1"
      If cFlag = True Then
      If rs.State = 1 Then rs.Close
      cFlag = False: Exit Sub
      End If
     
End If
Else
      'flxgtry.Visible = False
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.cmdAdd.Enabled = True
      oflag = False
      xflag = False: Exit Sub
End If
'      If rs.State = 1 Then rs.Close


End If

Exit Sub
OkError:
    If rs.State = 1 Then rs.Close
    MsgBox "DISPLAYING ITEMS IN STOCK", vbInformation, "ITEMS IN STOCK"
     Exit Sub
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"

If KeyAscii = 27 Then
   Me.txtUnitPrice.SetFocus
End If
If KeyAscii = vbKeyReturn Then
   If Me.txtname = "" Then
     Me.txtAmountPaid.SetFocus
   Else
     Me.flxgtry.Visible = True
     Me.flxgtry.SetFocus
   End If
End If

If KeyAscii = 92 Then
    If Me.OptConvert(0) Then
        Me.OptConvert(1).SetFocus
    ElseIf Me.OptConvert(1) Then
        Me.OptConvert(0).SetFocus
    Else
        Me.OptConvert(1).SetFocus
        Me.txtname.SetFocus
    Exit Sub
    End If
        Me.txtname.SetFocus
    
    If Me.txtNetCost <> "" Then
        NetCostflag = True
        txtNetCost = CurrencyConvertor(Trim(txtNetCost))
        NetCostflag = False
    End If
    If Me.txtAmountPaid <> "" Then
        txtAmountPaid = CurrencyConvertor(Trim(txtAmountPaid))
    End If
    
      MsgBox "Currency converted", vbInformation, "Converter"
      Me.txtname.SetFocus
End If

If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Function CurrencyConvertor(Cedi As Double) As Variant
Dim Value As Double, OldFlag As Boolean, NewFlag As Boolean

If Me.OptConvert(0) Then
Value = CDbl(Cedi) * 10000
CurrencyConvertor = Format$(Value, "#,###.00")
End If


If Me.OptConvert(1) Then
Value = CDbl(Cedi) / 10000
CurrencyConvertor = Format$(Value, "#,###.00")
End If

End Function

Private Sub Clear()

On Error GoTo OkError
For X = 1 To Me.flxgCosts.Rows - 1
            flxgCosts.TextMatrix(X, 0) = ""
            flxgCosts.TextMatrix(X, 1) = ""
            flxgCosts.TextMatrix(X, 2) = ""
            flxgCosts.TextMatrix(X, 3) = ""
            flxgCosts.TextMatrix(X, 4) = ""
            flxgCosts.TextMatrix(X, 5) = ""
'            flxgCosts.TextMatrix(X, 6) = ""
'            flxgCosts.TextMatrix(X, 7) = ""
'            flxgCosts.TextMatrix(X, 8) = ""
'            flxgCosts.TextMatrix(X, 9) = ""
       Next
    Fval2 = 0
    Frow2 = 0
    Me.flxgCosts.Rows = 2
       clearflag = True
      Me.txtDiscount = "": Me.txtUnitPrice = ""
      Me.txtTotalCost = "": Me.txtVat = "": Me.cboQuantity = 1
      Me.cboPaymode = "": Me.txtAmountPaid = "": Me.txtBalance = "": Me.txtname = ""
      Me.cboDepartment.Text = ""
      Me.txtCartonePrice = ""
      Me.txtCode = ""
      Me.txtNetCost = ""
      Me.flxgInventory.Visible = False
      clearflag = False
      Me.OptConvert(1).SetFocus
      Me.txtname.SetFocus
   Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
     Exit Sub
End Sub


Private Sub CalTotalPrice()
'If Me.cboSellCartone = "NO" Then
strg2 = Val(Me.txtUnitPrice) * Val(Me.cboQuantity) * (100 - Val(Me.txtDiscount)) / 100
Me.txtTotalCost = Format$(strg2, "#,###.00")
'Else
'strg2 = Val(Me.txtCartonePrice) * Val(Me.cboQuantity) * (100 - Val(Me.txtDiscount)) / 100
'Me.txtTotalCost = Format$(strg2, "#,###.00")
'End If
End Sub

Private Sub txtNetCost_Change()
If Me.txtAmountPaid = "" Then
Me.txtBalance = ""
Exit Sub
End If
On Error GoTo OkError
If NetCostflag = False Then
If clearflag = False Then
    Dim a As Integer, strgbal As String
    If oflag = False Then
        If Trim(Me.txtNetCost) <> "" Then
'            a = Len(Trim(Me.txtNetCost)) - 1
        End If
        If CDbl((Me.txtAmountPaid)) >= CDbl(Me.txtNetCost) Then
            strgbal = CDbl((Me.txtAmountPaid)) - CDbl(Me.txtNetCost)
            Me.txtBalance.Text = Format$(strgbal, "#,###.00")
        Else
            Me.txtAmountPaid = ""
        End If
    End If
Else
    clearflag = False
    Exit Sub
End If

End If

Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "SORRY,PLEASE TRY AGAIN", vbInformation, "TRY AGAIN"
     Exit Sub
End Sub

Private Sub ComputeBalance(Productid As String, Balance As Integer, Optional ReceivedIn As Integer, Optional IssuedOut As Integer)

  
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
   
   If rs.State = 1 Then rs.Close
   rs.Open "Select * From Tally where ProductID='" & Productid & "' order by IssueDate desc,IssueTime desc ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
     rs.MoveFirst
'   If Me.cboReceivedIn <> "" Then
'   Balance = ReceivedIn + rs.Fields("Balance")
'   End If
'   If Me.cboIssuedOut <> "" Then
     Balance = rs.Fields("Balance") - IssuedOut
'   End If
      
     If rs.State = 1 Then rs.Close
   Else
     If rs.State = 1 Then rs.Close
   End If
     If rs.State = 1 Then rs.Close

Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "Stores could not display", , "Displaying"
     Exit Sub

End Sub
Private Sub PrintReceipt()
On Error GoTo OkError

   
   CrystalInvoice.ReportFileName = App.Path & "\rptTallyInvoice.rpt"
   CrystalInvoice.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"


   CrystalInvoice.SelectionFormula = "{ClientPurchases.InvoiceNo} ='" & Me.txtReceipt.Text & "'"

   CrystalInvoice.WindowState = crptMaximized
   CrystalInvoice.WindowShowRefreshBtn = True
   CrystalInvoice.WindowTitle = "INVOICE " & Format$(Date, "yyyy")
   CrystalInvoice.Action = 0
   
   Exit Sub
OkError:
           
    MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY RECEIPT,PLEASE TRY AGAIN", vbInformation, "RECEIPT"
    Exit Sub
End Sub
