VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmPricing 
   BackColor       =   &H00C29E21&
   Caption         =   "Pricing"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   Icon            =   "frmPricing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   10905
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   7560
      Width           =   10575
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   1080
         TabIndex        =   25
         Top             =   120
         Width           =   8175
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   120
            TabIndex        =   29
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
            Image           =   "frmPricing.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdDelete 
            Height          =   375
            Left            =   1800
            TabIndex        =   30
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
            Image           =   "frmPricing.frx":075C
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdFind 
            Height          =   375
            Left            =   3480
            TabIndex        =   31
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
            Image           =   "frmPricing.frx":08F6
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
            TabIndex        =   10
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
            TabIndex        =   11
            Top             =   240
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
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   6720
            TabIndex        =   32
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
            Image           =   "frmPricing.frx":0D48
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   5160
            TabIndex        =   33
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
            Image           =   "frmPricing.frx":119A
            cBack           =   -2147483633
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
            TabIndex        =   13
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
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   7455
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   10575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProducts 
         Height          =   2055
         Left            =   1680
         TabIndex        =   26
         Top             =   4680
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3625
         _Version        =   393216
         BackColor       =   -2147483625
         ForeColor       =   -2147483643
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorBkg    =   14996075
         GridColor       =   -2147483634
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   "<Product Description                                   |<Number Per Cartone  |<ProductID  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.TextBox txtWholeSalePrice 
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
         Left            =   8160
         TabIndex        =   8
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ComboBox cboWholeSalePercent 
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
         ItemData        =   "frmPricing.frx":2A4B
         Left            =   8160
         List            =   "frmPricing.frx":2A6D
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   3855
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   10335
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgPricing 
            Height          =   3495
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   6165
            _Version        =   393216
            BackColor       =   0
            ForeColor       =   -2147483639
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorBkg    =   14996075
            GridColor       =   -2147483628
            AllowBigSelection=   0   'False
            FocusRect       =   2
            SelectionMode   =   1
            FormatString    =   $"frmPricing.frx":2A9C
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
            _Band(0).Cols   =   6
         End
      End
      Begin VB.ComboBox cboHandlingPrice 
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
         ItemData        =   "frmPricing.frx":2B29
         Left            =   2280
         List            =   "frmPricing.frx":2B45
         TabIndex        =   1
         Top             =   840
         Width           =   3255
      End
      Begin VB.ComboBox cboTnT 
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
         ItemData        =   "frmPricing.frx":2B7A
         Left            =   2280
         List            =   "frmPricing.frx":2B9F
         TabIndex        =   0
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtProductName 
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
         Left            =   2040
         TabIndex        =   4
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtPricePerCartone 
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
         Left            =   2040
         TabIndex        =   5
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtNoPerCartone 
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
         Left            =   2040
         TabIndex        =   6
         Top             =   2520
         Width           =   3495
      End
      Begin VB.ComboBox cboSalesPercent 
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
         ItemData        =   "frmPricing.frx":2BE9
         Left            =   8160
         List            =   "frmPricing.frx":2C0B
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtUnitPrice 
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
         Left            =   8160
         TabIndex        =   7
         Top             =   1560
         Width           =   2055
      End
      Begin lvButton_H.lvButtons_H cmdAdd 
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
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
         Image           =   "frmPricing.frx":2C3A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdRemove 
         Height          =   375
         Left            =   8040
         TabIndex        =   28
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
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
         Image           =   "frmPricing.frx":43FC
         cBack           =   -2147483633
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C29E21&
         Caption         =   "Selling Price (WholeSale):"
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
         Left            =   5880
         TabIndex        =   35
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C29E21&
         Caption         =   "Sales % (WholeSale):"
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
         Left            =   6240
         TabIndex        =   34
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C29E21&
         Caption         =   "Selling Price (Retail):"
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
         Left            =   6240
         TabIndex        =   27
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C29E21&
         Caption         =   "Sales % (Retail):"
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
         Left            =   6600
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Product Description:"
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
         TabIndex        =   20
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "C/P Per Cartone:"
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
         TabIndex        =   19
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Number Per Cartone:"
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
         TabIndex        =   18
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Handling Cost:"
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
         Left            =   840
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "T n T Cost Per Cartone:"
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
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, Frow2 As Integer, Fval2 As Integer
Dim xflag As Boolean, yflag As Boolean

Private Sub cboHandlingPrice_Change()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
 Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If
End Sub

Private Sub cboHandlingPrice_Click()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
 Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If
End Sub

Private Sub cboHandlingPrice_KeyPress(KeyAscii As Integer)
Dim strk1 As String
'   Me.cboTnT.SetFocus
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.cboTnT.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub cboHandlingPrice_LostFocus()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
 Call CalWholeSalePrice
End If

End Sub

Private Sub cboSalesPercent_Change()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
' Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If

End Sub

Private Sub cboSalesPercent_Click()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
' Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If
End Sub

Private Sub cboSalesPercent_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
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



Private Sub cboTnT_Change()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
 Call CalWholeSalePrice
 Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If

End Sub

Private Sub cboTnT_Click()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
 Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If
End Sub

Private Sub cboTnT_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.cboSalesPercent.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub cboWholeSalePercent_Change()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
' Call CalRetailPrice
 Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If

End Sub

Private Sub cboWholeSalePercent_Click()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
' Call CalRetailPrice
 Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If
End Sub

Private Sub cmdAdd_Click()
Dim X As Integer, a As Integer
On Error GoTo OkError
'xflag = True
'If eflag = False Then
'For xx = 1 To flxgPricing.Rows - 1
'           If flxgPricing.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgPricing.TextMatrix(xx, 4) = Productid Then
'              MsgBox flxgPricing.TextMatrix(xx, 0) & " Has Already Been Added!: You Can Only Edit What is Entered"
'              Me.txtProductName.SetFocus
'              Exit Sub
'           End If
'         Next
'End If
'Call CalPrice
If Trim(Me.cboTnT) = "" Then MsgBox "Please Enter Transportation Cost", vbInformation: Me.cboTnT.SetFocus: Exit Sub
If Trim(Me.cboSalesPercent) = "" And Trim(Me.cboWholeSalePercent) = "" Then MsgBox "Please Enter Some Percentage", vbInformation, "": Me.cboSalesPercent.SetFocus: Exit Sub
If Trim(Me.txtProductName) = "" Then MsgBox "Please Enter Product Description", vbInformation, "": Me.txtProductName.SetFocus: Exit Sub
If Val(Trim(Me.txtPricePerCartone)) = 0 Then MsgBox "Please Enter Cost Price", vbInformation, "": Me.txtPricePerCartone.SetFocus: Exit Sub
If Val(Trim(Me.txtNoPerCartone)) = 0 Then MsgBox "Please Enter Number Per Cartone", vbInformation, "": Me.txtNoPerCartone.SetFocus: Exit Sub

  Call CalRetailPrice
  Call CalWholeSalePrice
    If Trim(Me.txtUnitPrice) = "" Then MsgBox "Please Compute S/P (Retail)", vbInformation, "": Me.txtUnitPrice.SetFocus: Exit Sub
    If Trim(Me.txtWholeSalePrice) = "" Then MsgBox "Please Compute S/P (WholeSale)", vbInformation, "": Me.txtUnitPrice.SetFocus: Exit Sub
    
    If Frow2 <= 0 Then
      If Fval2 > 1 Then
         For xx = 1 To flxgPricing.Rows - 1
           If flxgPricing.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgPricing.TextMatrix(xx, 4) = Productid Then
              MsgBox Me.txtProductName & " Has Already Been Added!: You Can Only Edit What is Entered", vbInformation, ""
              Me.txtProductName = ""
              Exit Sub
           End If
         Next
       End If
      Fval2 = Fval2 + 1
      flxgPricing.Rows = Fval2 + 1
      flxgPricing.TextMatrix(Fval2, 0) = Me.txtProductName
      flxgPricing.TextMatrix(Fval2, 1) = Format$(Me.txtPricePerCartone, "#,###.00")
      flxgPricing.TextMatrix(Fval2, 2) = Format$(Me.txtUnitPrice, "#,###.00")
      flxgPricing.TextMatrix(Fval2, 3) = Format$(Me.txtWholeSalePrice, "#,###.00")
      flxgPricing.TextMatrix(Fval2, 4) = Me.txtNoPerCartone
      flxgPricing.TextMatrix(Fval2, 5) = Productid
      'flxgPricing.TextMatrix(Fval2, 5) = Me.txtVat
      'flxgPricing.TextMatrix(Fval2, 6) = Me.txtPriceAfterDiscount
       'flxgPricing.TextMatrix(Fval2, 7) = Trim(Productid)Format$(Me.txtWholeSalePrice, "#,###.00")
       
    Else
       flxgPricing.TextMatrix(Frow2, 0) = Me.txtProductName
       flxgPricing.TextMatrix(Frow2, 1) = Format$(Me.txtPricePerCartone, "#,###.00")
       flxgPricing.TextMatrix(Frow2, 2) = Format$(Me.txtUnitPrice, "#,###.00")
       flxgPricing.TextMatrix(Frow2, 3) = Format$(Me.txtWholeSalePrice, "#,###.00")
       flxgPricing.TextMatrix(Frow2, 4) = Me.txtNoPerCartone
       flxgPricing.TextMatrix(Frow2, 5) = Productid
        'flxgPricing.TextMatrix(Frow2, 6) = Me.txtPriceAfterDiscount
        'flxgPricing.TextMatrix(Frow2, 7) = Trim(Productid)
      Frow2 = 0
    
    End If
    Me.txtProductName = ""
    Me.txtPricePerCartone = ""
    Me.txtNoPerCartone = ""
    Me.txtUnitPrice = ""
    Me.txtWholeSalePrice = ""
    Productid = ""
    
    Me.txtProductName.SetFocus
    Me.cmdSave.Enabled = True
    
    
    
    oflag = False
    If eflag = True Then
     eflag = False
     End If
   
   Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "SORRY,TRY AGAIN", vbInformation, "TRY AGAIN"
     Exit Sub
End Sub

Private Sub cmdClear_Click()
For X = 1 To Me.flxgPricing.Rows - 1
            flxgPricing.TextMatrix(X, 0) = ""
            flxgPricing.TextMatrix(X, 1) = ""
            flxgPricing.TextMatrix(X, 2) = ""
            flxgPricing.TextMatrix(X, 3) = ""
            flxgPricing.TextMatrix(X, 4) = ""
            flxgPricing.TextMatrix(X, 5) = ""
            'flxgPricing.TextMatrix(X, 6) = ""
            'flxgPricing.TextMatrix(X, 7) = ""
       Next
    Fval2 = 0
    Frow2 = 0
    Productid = ""
    Me.flxgPricing.Rows = 2
Call ClearCtrls

Me.flxgProducts.Visible = False

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRemove_Click()
Dim a As Integer, strgbal As String

On Error GoTo OkError

 If Frow2 = 0 Then Exit Sub
 
  If Frow2 = 1 Then
   If flxgPricing.TextMatrix(Frow2, 0) = "" Then
   Frow2 = 0
   Me.txtProductName.Text = ""
   Me.flxgProducts.Visible = False
   Me.txtPricePerCartone = ""
   Me.txtNoPerCartone = ""
   Productid = ""
   Me.txtProductName.SetFocus
   Exit Sub
   End If
 End If
 
 
'If Trim(Me.txtNetCost) <> "" Then
'a = Len(Trim(Me.txtNetCost)) - 1
'b = Len(Trim(flxgPricing.TextMatrix(Frow2, 3))) - 1
'strgbal = CDbl((Mid(Me.txtNetCost, 2, a))) - CDbl((Mid(flxgPricing.TextMatrix(Frow2, 3), 2, b)))
'Me.txtNetCost.Text = Format$(strgbal, "¢#,###.00")
'Me.cmdCompute.Enabled = False
'End If
   If Frow2 = flxgPricing.Rows - 1 Then
     If flxgPricing.Rows <> 2 Then
         flxgPricing.Rows = flxgPricing.Rows - 1
     Else
        For xx = 0 To 5
           flxgPricing.TextMatrix(Frow2, xx) = ""
        Next
        Me.txtProductName.Text = ""
        Me.flxgProducts.Visible = False
        Me.txtPricePerCartone = ""
        Me.txtNoPerCartone = ""
        Productid = ""
        Me.txtProductName.SetFocus
     End If
        Me.txtProductName.Text = ""
        Me.flxgProducts.Visible = False
        Me.txtPricePerCartone = ""
        Me.txtNoPerCartone = ""
        Productid = ""
        Me.txtProductName.SetFocus
   Else
         For xx = Frow2 To flxgPricing.Rows - 2
            flxgPricing.TextMatrix(xx, 0) = flxgPricing.TextMatrix(xx + 1, 0)
            flxgPricing.TextMatrix(xx, 1) = flxgPricing.TextMatrix(xx + 1, 1)
            flxgPricing.TextMatrix(xx, 2) = flxgPricing.TextMatrix(xx + 1, 2)
            flxgPricing.TextMatrix(xx, 3) = flxgPricing.TextMatrix(xx + 1, 3)
            flxgPricing.TextMatrix(xx, 4) = flxgPricing.TextMatrix(xx + 1, 4)
            flxgPricing.TextMatrix(xx, 5) = flxgPricing.TextMatrix(xx + 1, 5)
            
        Next
        Me.txtProductName.Text = ""
        Me.flxgProducts.Visible = False
        Me.txtPricePerCartone = ""
        Me.txtNoPerCartone = ""
        Productid = ""
        Me.txtProductName.SetFocus
        flxgPricing.Rows = flxgPricing.Rows - 1
   End If
  Fval2 = Fval2 - 1
  Frow2 = 0
  
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

If Me.flxgPricing.TextMatrix(1, 0) = "" Then
   MsgBox "YOU MUST ENTER PRODUCT NAME.", vbInformation, "PRODUCT NAME"
   Me.txtProductName.SetFocus: Exit Sub
End If




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

   For X = 1 To flxgPricing.Rows - 1
   cn.Execute "Update Products Set PricePerCartone ='" & Trim(Me.flxgPricing.TextMatrix(X, 1)) & "',BaseUnit='" & Trim(Me.flxgPricing.TextMatrix(X, 2)) & "',UnitPrice='" & Trim(Me.flxgPricing.TextMatrix(X, 3)) & "' Where ProductID ='" & Trim(Me.flxgPricing.TextMatrix(X, 4)) & "'", Y
   Next
   If Y > 0 Then
   MsgBox "SAVED SUCCESSFULLY", vbInformation, "SAVED"
   Me.txtProductName.SetFocus
   yflag = True
   Else
   MsgBox " UNABLE TO SAVED ", vbInformation, "TRY AGAIN"
   End If
   If yflag = True Then
    
      For X = 1 To Me.flxgPricing.Rows - 1
            flxgPricing.TextMatrix(X, 0) = ""
            flxgPricing.TextMatrix(X, 1) = ""
            flxgPricing.TextMatrix(X, 2) = ""
            flxgPricing.TextMatrix(X, 3) = ""
            flxgPricing.TextMatrix(X, 4) = ""
            
       Next
    Fval2 = 0
    Me.flxgPricing.Rows = 2
    yflag = False
   End If
  
   Me.txtProductName.SetFocus
   
   
Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub

Private Sub flxgPricing_Click()
 xflag = True
        Me.txtProductName = Me.flxgPricing.TextMatrix(Me.flxgPricing.Row, 0)
        Me.txtPricePerCartone = Me.flxgPricing.TextMatrix(Me.flxgPricing.Row, 1)
        Me.txtUnitPrice = Me.flxgPricing.TextMatrix(Me.flxgPricing.Row, 2)
        Me.txtWholeSalePrice = Me.flxgPricing.TextMatrix(Me.flxgPricing.Row, 3)
        Me.txtNoPerCartone = Me.flxgPricing.TextMatrix(Me.flxgPricing.Row, 4)
        Productid = Me.flxgPricing.TextMatrix(Me.flxgPricing.Row, 5)
        
Frow2 = flxgPricing.Row
eflag = True
'oflag = False
Me.cmdSave.Enabled = True
Me.txtPricePerCartone.SetFocus
End Sub

Private Sub flxgProducts_Click()
    Me.txtProductName = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 0)
    Me.txtNoPerCartone = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 1)
    Productid = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 2)
    Me.flxgProducts.Visible = False
    Me.cmdSave.Enabled = True
    sflag = False
    oflag = False
    Me.txtPricePerCartone.SetFocus
End Sub

Private Sub flxgProducts_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.txtProductName = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 0)
    Me.txtNoPerCartone = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 1)
    Productid = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 2)
    Me.flxgProducts.Visible = False
    Me.cmdSave.Enabled = True
    sflag = False
    oflag = False
    Me.txtPricePerCartone.SetFocus
End If
End Sub

Private Sub Form_Load()
  Me.Height = 9330
  Me.Width = 11025
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Fval2 = 0
Frow2 = 0
End Sub

Private Sub txtNoPerCartone_Change()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
 Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If

End Sub

Private Sub txtNoPerCartone_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.cboHandlingPrice.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtNoPerCartone_LostFocus()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
 Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If

End Sub

Private Sub txtPricePerCartone_Change()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
 Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If

End Sub

Private Sub txtPricePerCartone_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.txtNoPerCartone.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtPricePerCartone_LostFocus()
If Val(Me.txtPricePerCartone) <> 0 And Val(Me.txtNoPerCartone) <> O Then
 Call CalRetailPrice
 Call CalWholeSalePrice
Else
 Me.txtUnitPrice = ""
 Me.txtWholeSalePrice = ""
End If

End Sub

Private Sub txtProductName_Change()

If Me.txtProductName = "" Then
 Me.flxgProducts.Visible = False
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

   rs.Open "Select * From Products  Where ProductName Like '" & Trim(Me.txtProductName) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
        flxgProducts.Height = 950 + (285 * (rs.RecordCount - 1))
        If flxgProducts.Height >= 4455 Then
           flxgProducts.Height = 4455
        End If
           flxgProducts.Rows = rs.RecordCount + 1
        With flxgProducts
           For X = 1 To rs.RecordCount
            .TextMatrix(X, 0) = rs.Fields("ProductName")
            .TextMatrix(X, 1) = rs.Fields("BaseUnit")
            .TextMatrix(X, 2) = rs.Fields("ProductID")
             rs.MoveNext
           Next
           .Col = 0
           .Row = 1
           .ColSel = 2
           .RowSel = 1
        End With
        flxgProducts.Visible = True
        If rs.State = 1 Then rs.Close
        Me.cmdSave.Enabled = True
  Else
        flxgProducts.Visible = False
        If rs.State = 1 Then rs.Close
  End If
        If rs.State = 1 Then rs.Close

Exit Sub
OkError:
       If rs.State = 1 Then rs.Close
       MsgBox "Items In Stock", , "Displaying"
       Exit Sub
End Sub

Private Sub txtProductName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()@#$%&*_-'"
If KeyAscii = vbKeyReturn Then
Me.flxgProducts.Visible = True
   Me.flxgProducts.SetFocus
End If

If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub
Private Sub CalRetailPrice()
Dim X As Double, Y As Double, Z As Double
If Me.cboSalesPercent <> "" Then
    X = Val(Me.txtPricePerCartone) + Val(Me.cboHandlingPrice) + Val(Me.cboTnT)
    If Val(Me.txtNoPerCartone) <> 0 Then
    Y = (X / Val(Me.txtNoPerCartone))
    'Else
    'Y = (Val(Me.txtNoPerCartone))
    End If
    Z = Y + ((Val(Me.cboSalesPercent) * Y)) / 100
    Me.txtUnitPrice = FormatNumber(Z, 2)
Else
    Me.txtUnitPrice = ""
End If
End Sub
Private Sub CalWholeSalePrice()
Dim X As Double, Y As Double, Z As Double
If Me.cboWholeSalePercent <> "" Then
    X = Val(Me.txtPricePerCartone) + Val(Me.cboHandlingPrice) + Val(Me.cboTnT)
        If Val(Me.txtNoPerCartone) <> 0 Then
        Y = (X / Val(Me.txtNoPerCartone))
    '    Else
    '    Y = (Val(Me.txtNoPerCartone))
        End If
    Z = Y + ((Val(Me.cboWholeSalePercent) * Y)) / 100
    Me.txtWholeSalePrice = FormatNumber(Z, 2)
Else
    Me.txtWholeSalePrice = ""
End If
End Sub
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
   ctrl = ""
   End If
Next
End Sub

