VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmSellProducts 
   BackColor       =   &H00C29E21&
   Caption         =   "Daily Sales"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmSellProducts.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.CheckBox ChkFilterBarCode 
      BackColor       =   &H00C29E21&
      Caption         =   "Filter Out BarCode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   840
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   120
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox ChkSellCartone 
      BackColor       =   &H00C29E21&
      Caption         =   "Sell Cartone"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4440
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C29E21&
      Height          =   7575
      Left            =   0
      TabIndex        =   26
      Top             =   480
      Width           =   6135
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgtry 
         Height          =   7455
         Left            =   0
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   13150
         _Version        =   393216
         BackColor       =   15388531
         ForeColor       =   0
         Cols            =   7
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
         FormatString    =   $"frmSellProducts.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   5295
      Left            =   6120
      TabIndex        =   41
      Top             =   2880
      Width           =   8655
      Begin Crystal.CrystalReport CrystalReceipt 
         Left            =   6720
         Top             =   3840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   3840
         Width           =   3735
      End
      Begin VB.TextBox txtAmountPaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   4680
         Width           =   2415
      End
      Begin VB.TextBox txtBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2895
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgCosts 
         Height          =   2775
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4895
         _Version        =   393216
         BackColor       =   15388531
         ForeColor       =   0
         Cols            =   10
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483628
         BackColorSel    =   -2147483647
         BackColorBkg    =   15388531
         GridColor       =   -2147483628
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
         FormatString    =   $"frmSellProducts.frx":03C3
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
         _Band(0).Cols   =   10
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.ComboBox cboPaymode 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSellProducts.frx":04AA
         Left            =   6480
         List            =   "frmSellProducts.frx":04B4
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Cash"
         Top             =   1680
         Width           =   1695
      End
      Begin lvButton_H.lvButtons_H cmdAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   4
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
         Image           =   "frmSellProducts.frx":04C6
         cBack           =   -2147483633
      End
      Begin VB.CommandButton cmdAdd1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Add"
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
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin lvButton_H.lvButtons_H cmdRemove 
         Height          =   375
         Left            =   1680
         TabIndex        =   54
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
         Image           =   "frmSellProducts.frx":1C88
         cBack           =   -2147483633
      End
      Begin VB.CommandButton cmdRemove1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "R&emove"
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin lvButton_H.lvButtons_H cmdCompute 
         Height          =   375
         Left            =   3240
         TabIndex        =   55
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
         Image           =   "frmSellProducts.frx":1E22
         cBack           =   -2147483633
      End
      Begin VB.CommandButton cmdCompute1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Compute"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C29E21&
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   68
         Top             =   4680
         Width           =   375
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
         Left            =   480
         TabIndex        =   49
         Top             =   3960
         Width           =   975
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
         TabIndex        =   48
         Top             =   2640
         Width           =   1335
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
         Left            =   120
         TabIndex        =   47
         Top             =   4680
         Width           =   1335
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
         Left            =   4560
         TabIndex        =   46
         Top             =   4680
         Width           =   855
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
      Height          =   2415
      Left            =   6120
      TabIndex        =   15
      Top             =   480
      Width           =   8655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgNonBarcodeProducts 
         Height          =   1335
         Left            =   7560
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   -360
         Visible         =   0   'False
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   15388531
         ForeColor       =   0
         Cols            =   7
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
         FormatString    =   $"frmSellProducts.frx":249C
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
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox txtCode 
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
         Left            =   2040
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtCartonePrice 
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
         Left            =   6720
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox cboSellCartone 
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
         ItemData        =   "frmSellProducts.frx":2555
         Left            =   6240
         List            =   "frmSellProducts.frx":255F
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "NO"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtUnitPrice 
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
         Left            =   2040
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtname 
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
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   6255
      End
      Begin VB.ComboBox cboQuantity 
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
         ItemData        =   "frmSellProducts.frx":256C
         Left            =   2040
         List            =   "frmSellProducts.frx":25CD
         TabIndex        =   3
         Text            =   "1"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtVat 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2DEA2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtPriceAfterDiscount 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2DEA2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8280
         TabIndex        =   16
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtDiscount 
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
         Left            =   6480
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgInventory 
         Height          =   2535
         Left            =   -120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4471
         _Version        =   393216
         BackColor       =   -2147483625
         ForeColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   -2147483647
         BackColorBkg    =   15705705
         GridColor       =   -2147483628
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"frmSellProducts.frx":2647
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox txtTotalCost 
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
         Left            =   6480
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C29E21&
         Caption         =   "Esc"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   72
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "F1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   67
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C29E21&
         Caption         =   "Product Code:"
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
         TabIndex        =   63
         Top             =   840
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000F&
         BorderStyle     =   6  'Inside Solid
         X1              =   0
         X2              =   11280
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000F&
         BorderStyle     =   6  'Inside Solid
         X1              =   0
         X2              =   11280
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label7 
         BackColor       =   &H005B4311&
         Caption         =   "Vat:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   25
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H005B4311&
         Caption         =   "Adjusted Price After Discount:"
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
         TabIndex        =   24
         Top             =   2520
         Width           =   2655
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
         TabIndex        =   23
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblUnitPrice 
         BackColor       =   &H00C29E21&
         Caption         =   "Unit Price:"
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
         Left            =   480
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
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
         Left            =   5160
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
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
         TabIndex        =   20
         Top             =   840
         Width           =   1215
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
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblcartonePrice 
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
         TabIndex        =   66
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboDepartment 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2EBBF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmSellProducts.frx":26EA
      Left            =   1920
      List            =   "frmSellProducts.frx":26EC
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00C29E21&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C29E21&
      BorderStyle     =   0  'None
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
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   8040
      Width           =   14775
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   6240
         TabIndex        =   34
         Top             =   120
         Width           =   8175
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   120
            TabIndex        =   7
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
            Image           =   "frmSellProducts.frx":26EE
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   3480
            TabIndex        =   56
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
            Image           =   "frmSellProducts.frx":3570
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdSave1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Ok"
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
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdClear1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "C&lear"
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
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin lvButton_H.lvButtons_H cmdPrint 
            Height          =   375
            Left            =   1800
            TabIndex        =   57
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
            Image           =   "frmSellProducts.frx":4E21
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdPrint1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Print"
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
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   6720
            TabIndex        =   58
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
            Image           =   "frmSellProducts.frx":4F33
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdExit1 
            BackColor       =   &H00C0C0C0&
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
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   240
            Width           =   1335
         End
         Begin lvButton_H.lvButtons_H cmdStock 
            Height          =   375
            Left            =   5040
            TabIndex        =   59
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
            Image           =   "frmSellProducts.frx":5385
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdStock1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&View Stock"
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
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   3975
         Begin lvButton_H.lvButtons_H cmdpassword 
            Height          =   375
            Left            =   120
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   240
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   661
            Caption         =   "C&hange Password"
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
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdpassword1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "C&hange Password"
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
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   240
            Width           =   1815
         End
         Begin lvButton_H.lvButtons_H cmdSales 
            Height          =   375
            Left            =   2160
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Vie&w Sales"
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
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdSales1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Vie&w Sales"
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
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   1575
         End
      End
      Begin lvButton_H.lvButtons_H Command2 
         Height          =   615
         Left            =   4080
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         Caption         =   "Pro&ducts OutOf Stock"
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
         cBack           =   -2147483633
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00B08F1E&
         Caption         =   "Pro&ducts OutOf Stock"
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdProductOutStock 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pro&ducts OutOf Stock"
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   2040
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   2880
      Top             =   2400
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C29E21&
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   71
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C29E21&
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   70
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00C29E21&
      Caption         =   "User:"
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
      Left            =   7680
      TabIndex        =   51
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C29E21&
      Caption         =   "Date/Time:"
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
      Left            =   11040
      TabIndex        =   50
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSellProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, Productid As String, X As Integer, QuantityLeft As Integer
Dim sflag As Boolean, strg2 As String, strg3 As String, xx As Integer, xflag As Boolean, yflag As Boolean, eflag As Boolean, a As Integer
Dim Fval2 As Integer, val2 As Integer, Frow2 As Integer, val3 As Integer, i As Variant, Y As Variant, b As Integer, c As Integer, oflag As Boolean
Dim clearflag As Boolean, UnitPriceflag As Boolean, nameflag As Boolean, ReceiptNo As Integer, ShowCodeflag As Boolean
Dim NetCostflag As Boolean, Productflag As Boolean, Quantityflag As Boolean, Cartoneflag As Boolean
Private Sub cboDepartment_Click()
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
If xflag = False Then  ' when true u can view grid to add products,but when false u can edit with no grid shown
   rs.Open "Select Distinct  Products.*, ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where Department Like '" & Trim(Me.cboDepartment) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   
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
        .TextMatrix(X, 3) = rs.Fields("PricePerCartone")
       .TextMatrix(X, 4) = rs.Fields("StockLevel")
       .TextMatrix(X, 5) = rs.Fields("ReorderLevel")
       .TextMatrix(X, 6) = rs.Fields("ProductID")
      
       
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 6
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
     'MsgBox "'" & Me.txtname & "' is a Non-Stock Product!", vbInformation, "Non-Stock"
     'Me.txtname = "": Me.txtUnitPrice = "": Me.cboQuantity = ""
     'Me.txtTotalCost = "": Me.txtDiscount = "": Me.txtPriceAfterDiscount = "": Me.txtVat = ""
   
     'Frw1 = 1
     'If rs.State <> 0 Then
       'rs.Close
     'End If
     'txtname.SetFocus
     'Exit Sub
End If
Else
      'flxgtry.Visible = False
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.cmdAdd.Enabled = True
      oflag = False
      xflag = False: Exit Sub
End If
If rs.State = 1 Then rs.Close
Me.txtname = ""
Me.txtname.SetFocus
Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "DISPLAYING ITEMS IN STOCK", vbInformation, "ITEMS IN STOCK"
     Exit Sub
End Sub

Private Sub cboDepartment_DropDown()
'On Error GoTo OkError


'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

Me.cboDepartment.Clear

 rs.Open "Select Distinct Department from Products  Order By Department Asc", cn, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
   rs.MoveFirst
   For i = 1 To rs.RecordCount
   If rs.Fields("Department") <> "" Then
     Me.cboDepartment.AddItem rs.Fields!Department
   End If
     rs.MoveNext
   Next
   'If cboProducts.ListCount > 1 Then
      'Me.cboProducts.AddItem "All"
   'End If
End If
If rs.State = 1 Then rs.Close
Exit Sub
OkError:
       If rs.State = 1 Then rs.Close
       MsgBox "TRY AGAIN", vbInformation, "DEPARTMENTALISATION"
       Exit Sub



End Sub

Private Sub cboPaymode_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = ""

If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub cboQuantity_Change()
If Len(Me.cboQuantity) >= 4 Then
 Me.cboQuantity = 1
End If
End Sub

Private Sub cboQuantity_GotFocus()
Me.cboQuantity.BackColor = &HFFFF&
End Sub

Private Sub cboQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Me.txtCode.SetFocus
End If
If KeyCode = 113 Then
Me.txtAmountPaid.SetFocus
End If
If KeyCode = 38 Then
Me.txtUnitPrice.SetFocus
End If

If KeyCode = 27 Then
    Me.txtname.SetFocus
End If

If KeyCode = 115 Then
 If Me.ChkSellCartone = Checked Then
    Me.ChkSellCartone = Unchecked
    Else
    Me.ChkSellCartone = Checked
 End If
End If

If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If
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
   Call AddProduct
End If

If KeyAscii = 27 Then
''    If Me.ChkSellCartone = Checked Then
''    Me.ChkSellCartone = Unchecked
''    Else
''    Me.ChkSellCartone = Checked
''    End If
End If

If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub cboQuantity_LostFocus()
Me.cboQuantity.BackColor = &HF2EBBF
End Sub

Private Sub cboSellCartone_Click()
Me.cboQuantity.SetFocus
End Sub

Private Sub ChkFilterBarCode_Click()
Me.txtname.SetFocus
'If Me.ChkFilterBarCode = vbChecked Then
' Me.lblcartonePrice.Visible = True
' Me.lblUnitPrice.Visible = False
' If Me.txtname <> "" Then
'       Call GetCartonePrice
''       Me.cboQuantity.SetFocus
' Else
'      Me.txtname.SetFocus
' End If
'Else
' Me.lblcartonePrice.Visible = False
' Me.lblUnitPrice.Visible = True
' Me.txtname.SetFocus
' If Me.txtname <> "" Then
'       Call GetCartonePrice
''       Me.cboQuantity.SetFocus
' Else
'      Me.txtname.SetFocus
' End If
'End If
End Sub

Private Sub ChkSellCartone_Click()
If Me.ChkSellCartone = vbChecked Then
 Me.lblcartonePrice.Visible = True
 Me.lblUnitPrice.Visible = False
 If Me.txtname <> "" Then
       Call GetCartonePrice
'       Me.cboQuantity.SetFocus
 Else
      Me.txtname.SetFocus
 End If
Else
 Me.lblcartonePrice.Visible = False
 Me.lblUnitPrice.Visible = True
 Me.txtname.SetFocus
 If Me.txtname <> "" Then
       Call GetCartonePrice
'       Me.cboQuantity.SetFocus
 Else
      Me.txtname.SetFocus
 End If
End If

End Sub

Private Sub cmdAdd_Click()
Call AddProduct
End Sub

Private Sub cmdClear_Click()
Call Clear
End Sub

Private Sub cmdCompute_Click()
Dim X As Integer, a As Integer

On Error GoTo OkError

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
Me.cmdCompute.Enabled = False
Me.txtAmountPaid.SetFocus

   Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,UNABLE TO COMPUTE,TRY AGAIN", vbInformation, "COMPUTATION"
     Exit Sub
End Sub


Private Sub cmdDelete_Click()

End Sub

Private Sub cmdExit_Click()
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdFind_Click()

End Sub

Private Sub cmdPassword_Click()
frmChangePassword.Show
End Sub

Private Sub cmdPrint_Click()
'Call Receipt
If Me.txtNetCost = "" Then
  MsgBox "COMPUTE NET COST", vbInformation, "NET COST"
  Exit Sub
End If
If Me.txtAmountPaid = "" Then
  MsgBox "INPUT AMOUNT PAID", vbInformation, "AMOUNT PAID"
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

For X = 1 To Me.flxgCosts.Rows - 1
   If Me.flxgCosts.TextMatrix(X, 0) = "" Then
     MsgBox "SELECT A PRODUCT", vbInformation, ""
     Me.txtCode.SetFocus: Exit Sub
   End If
Next
Call Save
PrintReceipt
End Sub

Private Sub cmdProductOutStock_Click()
frmRestockRetail.Show
End Sub

Private Sub cmdRemove_Click()
Dim a As Integer, strgbal As String

On Error GoTo OkError

 If Frow2 = 0 Then Exit Sub
 
  If Frow2 = 1 Then
   If flxgCosts.TextMatrix(Frow2, 0) = "" Then
   Frow2 = 0
   Productflag = True
   Me.txtname.Text = ""
   Productflag = False
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
        For xx = 0 To 9
           flxgCosts.TextMatrix(Frow2, xx) = ""
        Next
        Productflag = True
        Me.txtname.Text = ""
        Productflag = False
        Me.txtUnitPrice = ""
        Me.txtTotalCost = ""
        Me.txtname.SetFocus
     End If
        Productflag = True
        Me.txtname.Text = ""
        Productflag = False
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
            flxgCosts.TextMatrix(xx, 6) = flxgCosts.TextMatrix(xx + 1, 6)
            flxgCosts.TextMatrix(xx, 7) = flxgCosts.TextMatrix(xx + 1, 7)
            flxgCosts.TextMatrix(xx, 8) = flxgCosts.TextMatrix(xx + 1, 8)
            flxgCosts.TextMatrix(xx, 9) = flxgCosts.TextMatrix(xx + 1, 9)
        Next
        Productflag = True
        Me.txtname.Text = ""
        Productflag = False
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
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,UNABLE TO REMOVE,TRY AGAIN", vbInformation, "COMPUTATION"
     Exit Sub
  
End Sub

Private Sub cmdSales_Click()
frmRevenueRpt.Show
End Sub

Private Sub cmdSave_Click()
Call Save

End Sub

Private Sub cmdStock_Click()
frmStockLevelRpt.Show
End Sub

Private Sub Command2_Click()
frmRestockRetail.Show
End Sub

Private Sub flxgCosts_Click()
If MsgBox("ARE YOU SURE  YOU WANT TO EDIT OR REPLACE THE PRODUCT CLICKED?", vbYesNo + vbQuestion, "CONFIRMATION") = vbYes Then
        xflag = True
        Productflag = True
        Me.txtname = flxgCosts.TextMatrix(flxgCosts.Row, 0)
        Productflag = False
        Me.cboQuantity = flxgCosts.TextMatrix(flxgCosts.Row, 1)
        Me.txtUnitPrice = flxgCosts.TextMatrix(flxgCosts.Row, 2)
        Me.txtTotalCost = flxgCosts.TextMatrix(flxgCosts.Row, 3)
        Me.txtDiscount = flxgCosts.TextMatrix(flxgCosts.Row, 4)
        Me.txtVat = flxgCosts.TextMatrix(flxgCosts.Row, 5)
        Me.txtPriceAfterDiscount = flxgCosts.TextMatrix(flxgCosts.Row, 6)
        Productid = flxgCosts.TextMatrix(flxgCosts.Row, 7)
        If flxgCosts.TextMatrix(flxgCosts.Row, 8) <> "" Then
         Me.cboSellCartone = flxgCosts.TextMatrix(flxgCosts.Row, 8)
        End If
        Me.txtCartonePrice = flxgCosts.TextMatrix(flxgCosts.Row, 9)
Frow2 = flxgCosts.Row
eflag = True
oflag = False
Me.cmdSave.Enabled = True
Me.cboQuantity.SetFocus
Else
       Me.txtname.SetFocus
End If
End Sub

Private Sub flxgCosts_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 40 Or KeyCode = 38 Then
 If Me.flxgtry.Visible = True Then
  Me.flxgtry.SetFocus
 Else
  Me.flxgCosts.SetFocus
 End If
End If

If KeyCode = 112 Then
 Me.txtCode.SetFocus
End If

If KeyCode = 115 Then
 If Me.ChkSellCartone = Checked Then
    Me.ChkSellCartone = Unchecked
 Else
    Me.ChkSellCartone = Checked
 End If
End If

If KeyCode = 27 Then
    Me.txtname.SetFocus
End If

If KeyCode = 113 Then
    Me.txtAmountPaid.SetFocus
End If
If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If

End Sub

Private Sub flxgCosts_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If MsgBox("ARE YOU SURE  YOU WANT TO EDIT OR REPLACE THE PRODUCT CLICKED?", vbYesNo + vbQuestion, "CONFIRMATION") = vbYes Then
            xflag = True
            Productflag = True
            Me.txtname = flxgCosts.TextMatrix(flxgCosts.Row, 0)
            Productflag = False
            Me.cboQuantity = flxgCosts.TextMatrix(flxgCosts.Row, 1)
            Me.txtUnitPrice = flxgCosts.TextMatrix(flxgCosts.Row, 2)
            Me.txtTotalCost = flxgCosts.TextMatrix(flxgCosts.Row, 3)
            Me.txtDiscount = flxgCosts.TextMatrix(flxgCosts.Row, 4)
            Me.txtVat = flxgCosts.TextMatrix(flxgCosts.Row, 5)
            Me.txtPriceAfterDiscount = flxgCosts.TextMatrix(flxgCosts.Row, 6)
            Productid = flxgCosts.TextMatrix(flxgCosts.Row, 7)
            If flxgCosts.TextMatrix(flxgCosts.Row, 8) <> "" Then
             Me.cboSellCartone = flxgCosts.TextMatrix(flxgCosts.Row, 8)
            End If
            Me.txtCartonePrice = flxgCosts.TextMatrix(flxgCosts.Row, 9)
    Frow2 = flxgCosts.Row
    eflag = True
    oflag = False
    Me.cmdSave.Enabled = True
    Me.cboQuantity.SetFocus
    Else
            Me.txtname.SetFocus
    End If
End If
End Sub

Private Sub flxgInventory_Click()
Me.txtname = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 0)
Me.txtUnitPrice = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 1)
Me.txtDiscount = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 2)
Me.txtVat = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 3)
' = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 4)
'Me.txtDiscount = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 5)
Productid = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 6)

Me.flxgInventory.Visible = False

Me.cmdSave.Enabled = True
sflag = True
oflag = False
Me.txtname.SetFocus

i = 0 'Reset the value i in cmdcompute to recompute netsales
Me.cboQuantity.SetFocus
End Sub

Private Sub flxgInventory_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtname = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 0)
Me.txtUnitPrice = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 1)
Me.txtDiscount = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 2)
Me.txtVat = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 3)
' = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 4)
'Me.txtDiscount = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 5)
Productid = Me.flxgInventory.TextMatrix(Me.flxgInventory.Row, 6)

Me.flxgInventory.Visible = False

Me.cmdSave.Enabled = True
sflag = True
Me.txtname.SetFocus
Me.cboQuantity.SetFocus
End If
End Sub

Private Sub flxgNonBarcodeProducts_Click()
Productflag = True
Me.txtname = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 0)
Productflag = False
Me.txtUnitPrice = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 1)
If Val(Trim(Me.txtUnitPrice)) = 0 Then
Me.txtUnitPrice = ""
Me.txtUnitPrice.SetFocus
End If
Me.txtDiscount = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 2)
Me.txtCartonePrice = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 3)
' = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 4)
'Me.txtDiscount = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 5)
Productid = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 6)

Me.flxgNonBarcodeProducts.Visible = False

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

Private Sub flxgNonBarcodeProducts_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Productflag = True
Me.txtname = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 0)
Productflag = False
Me.txtUnitPrice = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 1)
If Val(Trim(Me.txtUnitPrice)) = 0 Then
Me.txtUnitPrice = ""
Me.txtUnitPrice.SetFocus
End If
Me.txtDiscount = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 2)
Me.txtCartonePrice = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 3)
' = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 4)
'Me.txtDiscount = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 5)
Productid = Me.flxgNonBarcodeProducts.TextMatrix(Me.flxgtry.Row, 6)

Me.flxgNonBarcodeProducts.Visible = False

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

Private Sub flxgtry_Click()
Productflag = True
Me.txtname = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 0)
Productflag = False

If Me.ChkSellCartone = vbUnchecked Then
 Me.txtUnitPrice = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 1)
End If

If Val(Trim(Me.txtUnitPrice)) = 0 Then
 Me.txtUnitPrice = ""
 Me.txtUnitPrice.SetFocus
End If
 Me.txtDiscount = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 2)
 
If Me.ChkSellCartone = vbChecked Then
 Me.txtUnitPrice = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 3)
End If
Productid = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 6)
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
 Me.flxgtry.Visible = False
End Sub

Private Sub flxgtry_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Me.txtname.SetFocus
End If
End Sub

Private Sub flxgtry_KeyPress(KeyAscii As Integer)


If KeyAscii = vbKeyReturn Then
Productflag = True
Me.txtname = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 0)
Productflag = False

If Me.ChkSellCartone = vbUnchecked Then
 Me.txtUnitPrice = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 1)
End If

If Val(Trim(Me.txtUnitPrice)) = 0 Then
Me.txtUnitPrice = ""
Me.txtUnitPrice.SetFocus
End If
Me.txtDiscount = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 2)

If Me.ChkSellCartone = vbChecked Then
 Me.txtUnitPrice = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 3)
End If
Productid = Me.flxgtry.TextMatrix(Me.flxgtry.Row, 6)
Me.cmdSave.Enabled = True
sflag = True
'Me.txtname.SetFocus
If Val(Trim(Me.txtUnitPrice)) = 0 Then
   Me.txtUnitPrice.SetFocus
Else
   Me.cboQuantity.SetFocus
End If
   Me.flxgtry.Visible = False
End If
End Sub

Private Sub flxgtry_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 112 Then
 Me.txtCode.SetFocus
End If
If KeyCode = 113 Then
 Me.txtAmountPaid.SetFocus
End If

End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Me.txtCode.SetFocus
End If
If KeyCode = 113 Then
Me.txtAmountPaid.SetFocus
End If
If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If

If KeyCode = 27 Then
    Me.txtname.SetFocus
End If

End Sub



Private Sub Form_Load()
Me.cmdSave.Enabled = False

Me.txtDate = Now
CenterForm Me
Me.Height = 9600
Me.Width = 14985
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2

On Error GoTo SaveError

'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   MsgBox strg, vbInformation:
   Exit Sub
End If

rs.Open "Select UserName From Users Where UserNo =" & Val(Trim(frmMDI.txtU)), cn, adOpenKeyset, adLockOptimistic
 If rs.RecordCount > 0 Then
Me.txtUser.Text = rs.Fields("UserName")
 If rs.State = 1 Then rs.Close
End If
 If rs.State = 1 Then rs.Close

Exit Sub
SaveError:
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,THERE IS AN ERROR,TRY AGAIN", vbInformation, ""
     Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Fval2 = 0
Frow2 = 0
End Sub

Private Sub lvButtons_H1_Click()

End Sub

'Private Sub Timer1_Timer()

'Me.txtNetCost.Visible = False
'DoEvents
'End Sub

  
  
'Private Sub Timer2_Timer()

'Me.txtNetCost.Visible = True
'DoEvents
'End Sub

Private Sub txtAmountPaid_Change()

On Error GoTo OkError

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

     
     Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "Type 0.", vbInformation, "Amount Paid"
     Me.txtAmountPaid = ""
     Me.txtAmountPaid.SetFocus
     Exit Sub
End Sub

Private Sub txtAmountPaid_GotFocus()
Me.txtAmountPaid.BackColor = &HFFFF&
End Sub

Private Sub txtAmountPaid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Me.txtCode.SetFocus
End If
If KeyCode = 38 Then
Me.txtCode.SetFocus
End If

If KeyCode = 115 Then
 If Me.ChkSellCartone = Checked Then
    Me.ChkSellCartone = Unchecked
    Else
    Me.ChkSellCartone = Checked
 End If
End If

If KeyCode = 27 Then
    Me.txtname.SetFocus
End If

If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If
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
  Call Save
  PrintReceipt
End If


If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
'If KeyAscii = vbKeyReturn Then
'
'Dim a As Integer, strgbal As String
'If Trim(Me.txtNetCost) <> "" Then
'a = Len(Trim(Me.txtNetCost)) - 1
'End If
'
'If CDbl(Val(Me.txtAmountPaid)) >= CDbl((Mid(Me.txtNetCost, 2, a))) Then
'strgbal = CDbl(Val(Me.txtAmountPaid)) - CDbl((Mid(Me.txtNetCost, 2, a)))
'Me.txtBalance.Text = Format$(strgbal, "¢#,###.00")
'End If
'End If
End Sub

Private Sub txtAmountPaid_LostFocus()
Me.txtAmountPaid.BackColor = &HF2EBBF
End Sub



Private Sub txtBalance_GotFocus()
Me.txtBalance.BackColor = &HFFFF&
End Sub

Private Sub txtBalance_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
Me.txtAmountPaid.SetFocus
End If

If KeyCode = 115 Then
 If Me.ChkSellCartone = Checked Then
    Me.ChkSellCartone = Unchecked
    Else
    Me.ChkSellCartone = Checked
 End If
End If

If KeyCode = 27 Then
    Me.txtname.SetFocus
End If

If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If
End Sub

Private Sub txtBalance_LostFocus()
Me.txtBalance.BackColor = &HF2EBBF
End Sub

Private Sub txtCode_Change()

On Error GoTo OkError

If ShowCodeflag = False Then

If Me.txtCode = "" Then
 Exit Sub
End If

'Open Connecttion to Server

   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
''''   If rs.State = 1 Then rs.Close


  If Me.ChkSellCartone = vbChecked Then
    rs.Open "Select Distinct  Products.*, ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where ProductCode1 Like '" & Trim(Me.txtCode) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
    Cartoneflag = True
  Else
    rs.Open "Select Distinct  Products.*, ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where ProductCode1 Like '" & Trim(Me.txtCode) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
  End If
If rs.RecordCount > 0 Then
        If (rs.Fields("ProductCode1")) = (Me.txtCode) Then
            Me.txtDiscount = rs.Fields("Discount")
            Productid = rs.Fields("ProductID")
            UnitPriceflag = True
            If Cartoneflag = True Then
             Me.txtUnitPrice = rs.Fields("PricePerCartone")
            Else
             Me.txtUnitPrice = rs.Fields("UnitPrice")
            End If
             Cartoneflag = False
            Productflag = True
             Me.txtname = rs.Fields("ProductName")
            Productflag = False
            Me.cboQuantity.SetFocus
        Else
            Me.txtUnitPrice = ""
            Me.txtDiscount = ""
        End If

End If
      nameflag = False
      UnitPriceflag = False
      flxgtry.Visible = True
      If rs.State = 1 Then rs.Close
      Me.cmdSave.Enabled = True

End If
Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "DISPLAYING ITEMS IN STOCK", vbInformation, "ITEMS IN STOCK"
     Exit Sub
End Sub

Private Sub txtCode_GotFocus()
Me.txtCode.BackColor = &HFFFF&
End Sub



Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
Me.txtAmountPaid.SetFocus
End If
If KeyCode = 38 Then
Me.txtname.SetFocus
End If

If KeyCode = 115 Then
 If Me.ChkSellCartone = Checked Then
    Me.ChkSellCartone = Unchecked
    Else
    Me.ChkSellCartone = Checked
 End If
End If

If KeyCode = 27 Then
    Me.txtname.SetFocus
End If


If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
'strk1 = "0123456789/|;:.,()"
If KeyAscii = 27 Then
   Me.txtname.SetFocus
End If

If KeyAscii = vbKeyReturn Then
   Me.txtAmountPaid.SetFocus
End If

If KeyAscii = 92 Then

    
    

End If

'   Me.txtName.SetFocus
'End If

'If KeyAscii > 26 Then
'   If KeyAscii <> 32 Then
'      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
'         KeyAscii = 0
'      End If
'   End If
'End If
End Sub

Private Sub txtCode_LostFocus()
Me.txtCode.BackColor = &HF2EBBF
End Sub

Private Sub txtdiscount_GotFocus()
Me.txtDiscount.BackColor = &HFFFF&
End Sub

Private Sub txtdiscount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Me.txtCode.SetFocus
End If
If KeyCode = 113 Then
Me.txtAmountPaid.SetFocus
End If

If KeyCode = 115 Then
 If Me.ChkSellCartone = Checked Then
    Me.ChkSellCartone = Unchecked
    Else
    Me.ChkSellCartone = Checked
 End If
End If

If KeyCode = 27 Then
    Me.txtname.SetFocus
End If

If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
Dim strk1 As String

strk1 = "0123456789."
If KeyAscii = vbKeyReturn Then
   Me.txtPriceAfterDiscount.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtdiscount_LostFocus()
Me.txtDiscount.BackColor = &HF2EBBF
End Sub

Private Sub txtName_Change()
On Error GoTo OkError

If Me.txtname = "" Then
 Me.flxgtry.Visible = False
 Me.txtname.SetFocus
 Exit Sub
End If

If Productflag = False Then

'Open Connecttion to Server

   If rs.State = 1 Then rs.Close
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
''''''If xflag = False Then  ' when true u can view grid to add products,but when false u can edit with no grid shown
If Me.ChkFilterBarCode = vbChecked Then
   rs.Open "Select Distinct  Products.*, ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where ProductName Like '" & Trim(Me.txtname) & "%" & "'  and Products.ProductCode = '" & Null & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
Else
   rs.Open "Select Distinct  Products.*, ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where ProductName Like '" & Trim(Me.txtname) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
End If
   If rs.RecordCount > 0 Then
   'flxgInventory.Height = 950 + (285 * (rs.RecordCount - 1))
   
   'If flxgInventory.Height >= 4455 Then
      'flxgInventory.Height = 4455
   'End If
    flxgtry.Rows = rs.RecordCount + 1
   With flxgtry
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("ProductName")
       .TextMatrix(X, 1) = Format$(rs.Fields("UnitPrice"), "#,###.00")
       .TextMatrix(X, 2) = rs.Fields("Discount")
       .TextMatrix(X, 3) = rs.Fields("PricePerCartone")
       .TextMatrix(X, 4) = rs.Fields("StockLevel")
       .TextMatrix(X, 5) = rs.Fields("ReorderLevel")
       .TextMatrix(X, 6) = rs.Fields("ProductID")
      
       
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 6
      .RowSel = 1
   End With
   flxgtry.Visible = True
   If rs.State = 1 Then rs.Close
   Me.cmdSave.Enabled = True
Else
   If rs.State = 1 Then rs.Close
     
End If
  If rs.State = 1 Then rs.Close


End If

End If

Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "DISPLAYING ITEMS IN STOCK", vbInformation, "ITEMS IN STOCK"
     Exit Sub
End Sub

Private Sub txtname_GotFocus()
Me.txtname.BackColor = &HFFFF&
End Sub



Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 40 Or KeyCode = 38 Then
 If Me.flxgtry.Visible = True Then
  Me.flxgtry.SetFocus
 Else
  Me.flxgCosts.SetFocus
 End If
End If

If KeyCode = 112 Then
 Me.txtCode.SetFocus
End If

If KeyCode = 115 Then
 If Me.ChkSellCartone = Checked Then
    Me.ChkSellCartone = Unchecked
 Else
    Me.ChkSellCartone = Checked
 End If
End If


If KeyCode = 113 Then
    Me.txtAmountPaid.SetFocus
End If
If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If
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
'    If Me.ChkSellCartone = Checked Then
'     Me.ChkSellCartone = Unchecked
'    Else
'     Me.ChkSellCartone = Checked
'    End If
End If

If KeyAscii = vbKeyReturn And Me.txtname = "" Then
     Me.txtAmountPaid.SetFocus
End If

If KeyAscii = vbKeyReturn And Me.txtname <> "" And Len(Me.txtname) >= 4 Then
  Call GetProducts
End If

If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub txtname_LostFocus()
Me.txtname.BackColor = &HF2EBBF
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
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,PLEASE TRY AGAIN", vbInformation, "TRY AGAIN"
     Exit Sub
End Sub

Private Sub txtPriceAfterDiscount_KeyPress(KeyAscii As Integer)
strk1 = "0123456789."
If KeyAscii = vbKeyReturn Then
   Me.txtVat.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub txtPriceAfterDiscount_LostFocus()
Me.txtPriceAfterDiscount = Me.txtTotalCost
End Sub

Private Sub txtTotalCost_GotFocus()
Me.txtTotalCost.BackColor = &HFFFF&
End Sub

Private Sub txtTotalCost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Me.txtCode.SetFocus
End If
If KeyCode = 113 Then
Me.txtAmountPaid.SetFocus
End If

If KeyCode = 115 Then
 If Me.ChkSellCartone = Checked Then
    Me.ChkSellCartone = Unchecked
 Else
    Me.ChkSellCartone = Checked
 End If
End If

If KeyCode = 27 Then
    Me.txtname.SetFocus
End If

If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If
End Sub

Private Sub txtTotalCost_KeyPress(KeyAscii As Integer)
strk1 = "0123456789."
If KeyAscii = vbKeyReturn Then
   Me.cboQuantity.SetFocus
End If

If KeyAscii = 27 Then
''    If Me.ChkSellCartone = Checked Then
''    Me.ChkSellCartone = Unchecked
''    Else
''    Me.ChkSellCartone = Checked
''    End If
End If

If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub txtTotalCost_LostFocus()
Me.txtPriceAfterDiscount = Me.txtTotalCost
Me.txtTotalCost.BackColor = &HF2EBBF
End Sub



Private Sub txtUnitPrice_GotFocus()
Me.txtUnitPrice.BackColor = &HFFFF&
End Sub

Private Sub txtUnitPrice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Me.txtCode.SetFocus
End If
If KeyCode = 113 Then
Me.txtAmountPaid.SetFocus
End If
If KeyCode = 38 Then
Me.txtCode.SetFocus
End If

If KeyCode = 115 Then
 If Me.ChkSellCartone = Checked Then
    Me.ChkSellCartone = Unchecked
 Else
    Me.ChkSellCartone = Checked
 End If
End If
If KeyCode = 27 Then
    Me.txtname.SetFocus
End If
If KeyCode = 114 Then
    If Me.ChkFilterBarCode.Value = 1 Then
       Me.ChkFilterBarCode.Value = 0
       Me.txtname.SetFocus
    Else
       Me.ChkFilterBarCode.Value = 1
       Me.txtname.SetFocus
    End If
End If
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
Dim strk1 As String

strk1 = "0123456789."

If KeyAscii = vbKeyReturn Then
   Me.txtAmountPaid.SetFocus
End If


If KeyAscii = 27 Then
'    If Me.ChkSellCartone = Checked Then
'    Me.ChkSellCartone = Unchecked
''    If Me.txtname <> "" Then
''       Call GetCartonePrice
''       Me.cboQuantity.SetFocus
''     End If
'    Else
'    Me.ChkSellCartone = Checked
''     If Me.txtname <> "" Then
''       Call GetCartonePrice
''       Me.cboQuantity.SetFocus
''     End If
'    End If
    
    
End If



If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtUnitPrice_LostFocus()
Me.txtUnitPrice.BackColor = &HF2EBBF
End Sub

Private Sub txtVat_KeyPress(KeyAscii As Integer)
strk1 = "0123456789."
If KeyAscii = vbKeyReturn Then
   Me.cmdAdd.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub CalTotalPrice()
If Me.cboSellCartone = "NO" Then
 strg2 = Val(Me.txtUnitPrice) * Val(Me.cboQuantity) * (100 - Val(Me.txtDiscount)) / 100
 Me.txtTotalCost = Format$(strg2, "#,###.00")
Else
 strg2 = Val(Me.txtCartonePrice) * Val(Me.cboQuantity) * (100 - Val(Me.txtDiscount)) / 100
 Me.txtTotalCost = Format$(strg2, "#,###.00")
End If
End Sub
Private Sub Save()
On Error GoTo OkError
'If Me.txtname = "" Then
'MsgBox "SELECT A PRODUCT", vbInformation
'Me.txtname.SetFocus: Exit Sub
'End If
For X = 1 To Me.flxgCosts.Rows - 1
   If Me.flxgCosts.TextMatrix(X, 0) = "" Then
     MsgBox "SELECT A PRODUCT", vbInformation, "PRODUCT DESCRIPTION"
     Me.txtname.SetFocus: Exit Sub
   End If
Next
If Me.txtNetCost = "" Then
  MsgBox "COMPUTE NET COST", vbInformation, "NET COST"
  Exit Sub
End If
If Me.txtAmountPaid = "" Then
  MsgBox "INPUT AMOUNT PAID", vbInformation, "AMOUNT PAID"
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
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   cn.BeginTrans
   For X = 1 To Me.flxgCosts.Rows - 1
   rs.Open "Select ProductInventory.StockLevel,Products.PricePerCartone From ProductInventory Inner Join Products On ProductInventory.ProductID=Products.ProductID Where ProductInventory.ProductID= '" & Trim(Me.flxgCosts.TextMatrix(X, 7)) & "'", cn, adOpenForwardOnly, adLockReadOnly
        'For i = 1 To rs.RecordCount
    If rs.RecordCount > 0 Then
        If Trim(Me.flxgCosts.TextMatrix(X, 8)) = "NO" Then
        
          If rs.Fields("StockLevel") >= Val(Trim(Me.flxgCosts.TextMatrix(X, 1))) Then
           QuantityLeft = rs.Fields("StockLevel") - Val(Trim(Me.flxgCosts.TextMatrix(X, 1)))
           cn.Execute "Update ProductInventory Set StockLevel ='" & QuantityLeft & "' Where ProductID ='" & Trim(Me.flxgCosts.TextMatrix(X, 7)) & "'", Y
            If Y < 0 Then
             MsgBox "UNABLE TO COMPUTE,PLEASE TRY AGAIN", vbiformation, "TRY AGAIN"
             Me.txtname.SetFocus: rs.Close: Exit Sub
            End If
          Else
            Y = 1
          
         End If
          If rs.State = 1 Then rs.Close
       End If
          If rs.State = 1 Then rs.Close
    End If
    'Next
    
   
     'For X = 1 To Me.flxgCosts.Rows - 1
'       If Trim(flxgCosts.TextMatrix(X, 3)) <> "" Then
'         a = Len(Trim(flxgCosts.TextMatrix(X, 3))) - 1
'       End If
     
    If Y > 0 Then
        rs.Open "Select * From [Payments/Products] Where ProductID= '" & Trim(Me.flxgCosts.TextMatrix(X, 7)) & "' And Date='" & Date & "' And UnitPrice= '" & Val(Me.flxgCosts.TextMatrix(X, 2)) & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then 'And rs.Fields("UnitPrice") = Val(Trim(Me.flxgCosts.TextMatrix(X, 2))) Then
                AdjQty = rs.Fields("QtyBought") + Val(Trim(Me.flxgCosts.TextMatrix(X, 1)))
                AdjTotalCost = rs.Fields("TotalCost") + CDbl((Trim(Me.flxgCosts.TextMatrix(X, 3))))
            cn.Execute "Update [Payments/Products] Set QtyBought ='" & AdjQty & "',TotalCost='" & AdjTotalCost & "' Where ProductID ='" & Trim(Me.flxgCosts.TextMatrix(X, 7)) & "'and Date='" & Date & "' And UnitPrice= '" & Val(Me.flxgCosts.TextMatrix(X, 2)) & "'", Y
                yflag = True
                If rs.State = 1 Then rs.Close
            
         Else
             cn.Execute "Insert Into [Payments/Products] ([ProductID],[QtyBought],[TotalCost],[Date],[UnitPrice]) select '" & Trim(Me.flxgCosts.TextMatrix(X, 7)) & "','" & Val(Trim(Me.flxgCosts.TextMatrix(X, 1))) & "','" & CDbl(Trim(Me.flxgCosts.TextMatrix(X, 3))) & "','" & Date & "','" & Val(Trim(Me.flxgCosts.TextMatrix(X, 2))) & "'", Y
             yflag = True
              If rs.State = 1 Then rs.Close
        End If
          If rs.State = 1 Then rs.Close
     End If
   Next
      If Y > 0 Then
        cn.CommitTrans
        Fval2 = 0
      Else
       cn.RollbackTrans
       MsgBox "Save Failed,Please Retype Data and Save Again", vbInformation, "Try Again"
      End If
     
    
  


   cn.BeginTrans
  rs.Open "Select * From Payments ", cn, adOpenForwardOnly, adLockReadOnly
        
            If rs.RecordCount > 0 Then
               ReceiptNo = rs.Fields("ReceiptNo")
               cn.Execute "Update Payments Set AmountPaid ='" & CDbl(Trim(Me.txtAmountPaid)) & "',PmtMode='Cash',Balance ='" & CDbl((Trim(Me.txtBalance))) & "',NetCost ='" & CDbl((Trim(Me.txtNetCost))) & "',Date='" & Date & "' Where ReceiptNo ='" & ReceiptNo & "'", Y
               If rs.State = 1 Then rs.Close
            Else
              cn.Execute "Insert Into Payments ([AmountPaid],[PmtMode],[Balance],[NetCost],[Date],[UserID],ReceiptNo) select '" & CDbl(Trim(Me.txtAmountPaid)) & "','" & Trim(Me.cboPaymode) & "','" & CDbl((Trim(Me.txtBalance))) & "','" & CDbl((Trim(Me.txtNetCost))) & "','" & Date & "','111','1'", Y
        '      If cn.State = 1 Then cn.Close
              If rs.State = 1 Then rs.Close
            End If
                If rs.State = 1 Then rs.Close
        
     If Y > 0 Then
  
            rs.Open "Select * From [Receipt/Product] ", cn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount > 0 Then
                cn.Execute "Delete From [Receipt/Product] Where ReceiptNo ='1'", Y
                For X = 1 To Me.flxgCosts.Rows - 1
               cn.Execute "Insert Into [Receipt/Product] ([ProductID],Quantity,[ReceiptNo],UnitPrice) select '" & Trim(Me.flxgCosts.TextMatrix(X, 7)) & "','" & Trim(Me.flxgCosts.TextMatrix(X, 1)) & "','1','" & CDbl(Me.flxgCosts.TextMatrix(X, 2)) & "'", Y
                Next
               If rs.State = 1 Then rs.Close
            Else
                For X = 1 To Me.flxgCosts.Rows - 1
               cn.Execute "Insert Into [Receipt/Product] ([ProductID],Quantity,[ReceiptNo],UnitPrice) select '" & Trim(Me.flxgCosts.TextMatrix(X, 7)) & "','" & Trim(Me.flxgCosts.TextMatrix(X, 1)) & "','1','" & Trim(Me.flxgCosts.TextMatrix(X, 2)) & "'", Y
                Next
                If rs.State = 1 Then rs.Close
            End If
                If rs.State = 1 Then rs.Close
    
    End If
    
    If Y > 0 Then
        cn.CommitTrans
    Else
        cn.RollbackTrans
    End If
      
        
   If yflag = True Then
   ' i = Me.flxgCosts.Rows - 1
      For X = 1 To Me.flxgCosts.Rows - 1
            flxgCosts.TextMatrix(X, 0) = ""
            flxgCosts.TextMatrix(X, 1) = ""
            flxgCosts.TextMatrix(X, 2) = ""
            flxgCosts.TextMatrix(X, 3) = ""
            flxgCosts.TextMatrix(X, 4) = ""
            flxgCosts.TextMatrix(X, 5) = ""
            flxgCosts.TextMatrix(X, 6) = ""
            flxgCosts.TextMatrix(X, 7) = ""
            flxgCosts.TextMatrix(X, 8) = ""
            flxgCosts.TextMatrix(X, 9) = ""
       Next
    Fval2 = 0
    Me.flxgCosts.Rows = 2
    yflag = False
   End If
      
      
      Me.cmdSave.Enabled = False
      Me.cmdCompute.Enabled = True
      Me.flxgInventory.Visible = False
      oflag = True
      Me.txtDiscount = "": Me.txtPriceAfterDiscount = "": Me.txtUnitPrice = ""
      Me.txtTotalCost = "": Me.txtVat = "": Me.cboQuantity = 1
      Me.txtAmountPaid = "": Me.txtNetCost = "": Me.txtBalance = ""
      
      
      Me.txtname.SetFocus
   Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,THERE IS AN ERROR,TRY AGAIN", vbInformation, "COMPUTATION"
     Exit Sub
End Sub
Private Sub AddProduct()
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
           If flxgCosts.TextMatrix(xx, 0) = Trim(Me.txtname) And flxgCosts.TextMatrix(xx, 7) = Productid Then
              MsgBox flxgCosts.TextMatrix(xx, 0) & " Has Already Been Added!: You Can Only Edit What is Entered", vbInformation, ""
              Productflag = True
              Me.txtname = ""
              Productflag = False
              Me.txtCode = ""
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
       If Me.cboSellCartone = "NO" Then
        flxgCosts.TextMatrix(Fval2, 2) = Me.txtUnitPrice
       Else
        flxgCosts.TextMatrix(Fval2, 2) = Me.txtCartonePrice
       End If
       flxgCosts.TextMatrix(Fval2, 3) = Me.txtTotalCost
       flxgCosts.TextMatrix(Fval2, 4) = Me.txtDiscount
       flxgCosts.TextMatrix(Fval2, 5) = Me.txtVat
       flxgCosts.TextMatrix(Fval2, 6) = Me.txtPriceAfterDiscount
       flxgCosts.TextMatrix(Fval2, 7) = Productid
       flxgCosts.TextMatrix(Fval2, 8) = Trim(Me.cboSellCartone)
       flxgCosts.TextMatrix(Fval2, 9) = Trim(Me.txtCartonePrice)
       
    Else
        flxgCosts.TextMatrix(Frow2, 0) = Me.txtname
        flxgCosts.TextMatrix(Frow2, 1) = Me.cboQuantity
        If Me.cboSellCartone = "NO" Then
         flxgCosts.TextMatrix(Frow2, 2) = Me.txtUnitPrice
        Else
         flxgCosts.TextMatrix(Frow2, 2) = Me.txtCartonePrice
        End If
        flxgCosts.TextMatrix(Frow2, 3) = Me.txtTotalCost
        flxgCosts.TextMatrix(Frow2, 4) = Me.txtDiscount
        flxgCosts.TextMatrix(Frow2, 5) = Me.txtVat
        flxgCosts.TextMatrix(Frow2, 6) = Me.txtPriceAfterDiscount
        flxgCosts.TextMatrix(Frow2, 7) = Productid
        flxgCosts.TextMatrix(Frow2, 8) = Trim(Me.cboSellCartone)
        flxgCosts.TextMatrix(Frow2, 9) = Trim(Me.txtCartonePrice)
      Frow2 = 0
    
    End If
    UnitPriceflag = True
    'nameflag = True
    Productflag = True
    Me.txtname = ""
    Productflag = False
    Me.txtUnitPrice = ""
    Me.txtTotalCost = ""
    Me.txtDiscount = ""
    Me.txtPriceAfterDiscount = ""
    Me.txtVat = ""
    Me.cboDepartment.Text = ""
    ShowCodeflag = True
    Me.txtCode = ""
    ShowCodeflag = False
    Me.txtname.SetFocus
    Me.cmdSave.Enabled = True
    Me.cmdCompute.Enabled = True
    UnitPriceflag = False
    nameflag = False
    i = 0 'Reset the value i in cmdcompute to recompute netsales
    
    If Me.ChkSellCartone = vbChecked Then
       Me.ChkSellCartone = vbUnchecked
       Me.lblcartonePrice.Visible = False
       Me.lblUnitPrice.Visible = True
    End If

    
    
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
   Me.cboSellCartone.Text = "NO"
   
   Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
     Exit Sub
End Sub
Private Sub Clear()
On Error GoTo OkError
For X = 1 To Me.flxgCosts.Rows - 1
            flxgCosts.TextMatrix(X, 0) = ""
            flxgCosts.TextMatrix(X, 1) = ""
            flxgCosts.TextMatrix(X, 2) = ""
            flxgCosts.TextMatrix(X, 3) = ""
            flxgCosts.TextMatrix(X, 4) = ""
            flxgCosts.TextMatrix(X, 5) = ""
            flxgCosts.TextMatrix(X, 6) = ""
            flxgCosts.TextMatrix(X, 7) = ""
            flxgCosts.TextMatrix(X, 8) = ""
            flxgCosts.TextMatrix(X, 9) = ""
       Next
    Fval2 = 0
    Frow2 = 0
    Me.flxgCosts.Rows = 2
       clearflag = True
      Me.txtDiscount = "": Me.txtPriceAfterDiscount = "": Me.txtUnitPrice = ""
      Me.txtTotalCost = "": Me.txtVat = "": Me.cboQuantity = 1
      Me.cboPaymode = "": Me.txtAmountPaid = "": Me.txtBalance = ""
      Productflag = True
      Me.txtname = ""
      Productflag = False
      Me.cboDepartment.Text = ""
      Me.txtCartonePrice = ""
      ShowCodeflag = True
      Me.txtCode = ""
      ShowCodeflag = False
      Me.txtNetCost = ""
      Me.flxgInventory.Visible = False
      clearflag = False
      
      If Me.ChkSellCartone = vbChecked Then
       Me.ChkSellCartone = vbUnchecked
       Me.lblcartonePrice.Visible = False
       Me.lblUnitPrice.Visible = True
      End If
      
      Me.flxgtry.Visible = False
      
      Me.txtname.SetFocus
   Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
     Exit Sub
End Sub
Private Sub Receipt()
For X = 1 To Me.flxgCosts.Rows - 1
 If flxgCosts.TextMatrix(X, 0) = "" Then
 MsgBox "NOTHING TO BE PRINTED", vbInformation, "BLANK PRINT"
 Exit Sub
 End If
Next
If Me.txtAmountPaid = "" Then
 MsgBox "AMOUNT PAID NOT ENTERED", vbInformation, "BLANK PRINT"
 Me.txtAmountPaid.SetFocus: Exit Sub
End If
a = Len(Format$(Me.txtAmountPaid, "#,###.00"))
b = Len(Me.txtNetCost)
c = Len(Me.txtBalance)

On Error GoTo SaveError

Printer.FontSize = 7
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.Print Tab(15); "Abu MULTIMART"
Printer.FontBold = False
Printer.Print ""
Printer.FontSize = 6
Printer.Print Tab(2); "Date:"; Space(2); Date
Printer.Print Tab(2); "Time:"; Space(2); Format(Now, "hh:mm:ss ampm")
Printer.Print Tab(2); "Vat Reg #:"
Printer.Print ""
Printer.FontSize = 7
Printer.Print Tab(2); "PRODUCT"; Space(20); "QTY"; Space(15); "PRICE"
Printer.Print Tab(2); "----------------------------------------------------------------------"
Printer.Print Tab(2); "----------------------------------------------------------------------"
For X = 1 To Me.flxgCosts.Rows - 1
d = Len(flxgCosts.TextMatrix(X, 3))
Printer.Print Tab(2); Left(flxgCosts.TextMatrix(X, 0), 14); Tab(25); flxgCosts.TextMatrix(X, 1); Tab(35 + (a - d)); flxgCosts.TextMatrix(X, 3)
Next
Printer.Print Tab(2); "----------------------------------------------------------------------"
Printer.Print Tab(2); "----------------------------------------------------------------------"
Printer.Print Tab(2); "Total Amount:"; Tab(35 + (a - b)); Me.txtNetCost
Printer.Print Tab(2); "Amount Paid:"; Tab(35); Format$(Me.txtAmountPaid, "#,###.00")
Printer.Print Tab(2); "Balance:"; Tab(35 + (a - c)); Me.txtBalance
Printer.Print ""
Printer.FontSize = 6
Printer.Print Tab(2); "VAT Inclusive"
Printer.Print Tab(20); "THANK YOU"
Printer.Print Tab(2); "Developed by Abu Yakubu"
Printer.Print Tab(2); "MTX SoftWare Systems"
Printer.Print Tab(2); "Tel:0242324486"
Printer.EndDoc

 Exit Sub
SaveError:
 MsgBox "TRY AGAIN", vbInformation, "PRINT FAILED"
End Sub
Private Sub PrintReceipt()
On Error GoTo OkError

   
CrystalReceipt.ReportFileName = App.Path & "\rptReceipt.rpt"
CrystalReceipt.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"

'If Me.cboProducts <> "All" Then
'CrystalProducts.SelectionFormula = "{Products.ProductName} ='" & Me.cboProducts.Text & "'"
'Else
'CrystalProducts.SelectionFormula = ""
'End If
'   CrystalReceipt.WindowState = crptMaximized
'   CrystalReceipt.WindowShowRefreshBtn = True
'   CrystalReceipt.WindowTitle = "Products Prices " & Format$(Date, "yyyy")
   CrystalReceipt.Action = 0
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY RECEIPT,PLEASE TRY AGAIN", vbInformation, "RECEIPT"
       Exit Sub
End Sub
'Private Function CurrencyConvertor(Cedi As Double) As Variant
'Dim Value As Double, OldFlag As Boolean, NewFlag As Boolean
'
'If Me.OptConvert(0) Then
'Value = CDbl(Cedi) * 10000
'CurrencyConvertor = Format$(Value, "#,###.00")
'End If
'
'
'If Me.OptConvert(1) Then
'Value = CDbl(Cedi) / 10000
'CurrencyConvertor = Format$(Value, "#,###.00")
'End If
'
'End Function

Private Sub GetCartonePrice()

On Error GoTo OkError

   If rs.State = 1 Then rs.Close
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   If Me.ChkSellCartone = vbChecked Then
    Cartoneflag = True
   Else
    Cartoneflag = False
   End If
   
   rs.Open "Select Distinct  Products.*, ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where ProductName = '" & Trim(Me.txtname) & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then

            Me.txtDiscount = rs.Fields("Discount")
            Productid = rs.Fields("ProductID")
            UnitPriceflag = True
            'nameflag = True
            If Cartoneflag = True Then
             Me.txtUnitPrice = rs.Fields("PricePerCartone")
            Else
             Me.txtUnitPrice = rs.Fields("UnitPrice")
            End If
             Cartoneflag = False
            nameflag = True
            Me.txtname = rs.Fields("ProductName")
            nameflag = False
            Me.cboQuantity.SetFocus
            Me.flxgNonBarcodeProducts.Visible = False
    Else
            Me.txtUnitPrice = ""
            Me.txtDiscount = ""
            Me.txtname.SetFocus
    End If
    If rs.State = 1 Then rs.Close
   

Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "DISPLAYING ITEMS IN STOCK", vbInformation, "ITEMS IN STOCK"
     Exit Sub

End Sub

Private Sub GetProducts()

'On Error GoTo OkError

'Open Connecttion to Server

   If rs.State = 1 Then rs.Close
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
      
If rs.State = 1 Then rs.Close

   If Me.ChkSellCartone = vbChecked Then
    rs.Open "Select Distinct  Products.*, ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where ProductCode = '" & Trim(Me.txtname) & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
    Cartoneflag = True
   Else
   rs.Open "Select Distinct  Products.*, ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where ProductCode = '" & Trim(Me.txtname) & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   End If
   If rs.RecordCount > 0 Then
'       For X = 1 To rs.RecordCount
       If Len(rs.Fields("ProductCode")) = Len(Me.txtname) Then
            Me.txtDiscount = rs.Fields("Discount")
            Productid = rs.Fields("ProductID")
            UnitPriceflag = True
            If Cartoneflag = True Then
             Me.txtUnitPrice = rs.Fields("PricePerCartone")
            Else
             Me.txtUnitPrice = rs.Fields("UnitPrice")
            End If
             Cartoneflag = False
             
            Productflag = True
            Me.txtname = rs.Fields("ProductName")
            Productflag = False
            Me.cboQuantity.SetFocus
        Else
            Me.txtUnitPrice = ""
            Me.txtDiscount = ""
        End If
'        rs.MoveNext
'        Next

      
         flxgtry.Visible = True
         If rs.State = 1 Then rs.Close
         Me.cmdSave.Enabled = True
Else
        MsgBox "The Product Scanned is not is the System", , ""
        Me.txtname = ""
        Me.txtname.SetFocus
        If rs.State = 1 Then rs.Close
End If

Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "DISPLAYING ITEMS IN STOCK", vbInformation, "ITEMS IN STOCK"
     Exit Sub

End Sub

