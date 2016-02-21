VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmReceivals 
   Appearance      =   0  'Flat
   BackColor       =   &H00C29E21&
   Caption         =   "Goods Receivals"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   9840
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   8535
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9615
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   7320
         Width           =   9375
         Begin VB.Frame Frame5 
            BackColor       =   &H00C29E21&
            Height          =   735
            Left            =   720
            TabIndex        =   18
            Top             =   120
            Width           =   7575
            Begin lvButton_H.lvButtons_H cmdExit 
               Height          =   375
               Left            =   5160
               TabIndex        =   23
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
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
               Image           =   "Form1.frx":030A
               cBack           =   -2147483633
            End
            Begin lvButton_H.lvButtons_H cmdSave 
               Height          =   375
               Left            =   240
               TabIndex        =   24
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
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
               Image           =   "Form1.frx":075C
               cBack           =   -2147483633
            End
            Begin lvButton_H.lvButtons_H cmdClear 
               Height          =   375
               Left            =   2760
               TabIndex        =   25
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
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
               Image           =   "Form1.frx":0BAE
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
               Left            =   7560
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   360
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
               Left            =   7560
               Style           =   1  'Graphical
               TabIndex        =   19
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
               Left            =   7560
               MaskColor       =   &H0000FF00&
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   360
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C29E21&
         Height          =   7335
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   9375
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgSupplier 
            Height          =   1215
            Left            =   4560
            TabIndex        =   26
            Top             =   720
            Visible         =   0   'False
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   2143
            _Version        =   393216
            BackColor       =   16117969
            FixedCols       =   0
            BackColorFixed  =   8421504
            ForeColorFixed  =   -2147483634
            BackColorBkg    =   12754465
            AllowBigSelection=   0   'False
            FocusRect       =   2
            SelectionMode   =   1
            FormatString    =   "<SupplierName                                                      |<SupplierID"
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgStoreName 
            Height          =   1215
            Left            =   360
            TabIndex        =   31
            Top             =   720
            Visible         =   0   'False
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   2143
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
            FormatString    =   "<StoreName                                                   |<StoreID             "
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
         Begin VB.TextBox txtStoreName 
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
            Left            =   1680
            TabIndex        =   0
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txtReceiptNo 
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
            Left            =   1680
            TabIndex        =   2
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtSupplier 
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
            TabIndex        =   1
            Top             =   360
            Width           =   2655
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProducts 
            Height          =   2295
            Left            =   2040
            TabIndex        =   21
            Top             =   2520
            Visible         =   0   'False
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   4048
            _Version        =   393216
            BackColor       =   16117969
            FixedCols       =   0
            BackColorFixed  =   8421504
            ForeColorFixed  =   -2147483634
            BackColorBkg    =   12754465
            AllowBigSelection=   0   'False
            FocusRect       =   2
            SelectionMode   =   1
            FormatString    =   "<ProductName                                                                             |>ProductID      "
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgAddProducts 
            Height          =   3255
            Left            =   240
            TabIndex        =   22
            Top             =   3840
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   5741
            _Version        =   393216
            BackColor       =   16117969
            Cols            =   4
            FixedCols       =   0
            BackColorFixed  =   8421504
            ForeColorFixed  =   -2147483634
            BackColorBkg    =   12754465
            AllowBigSelection=   0   'False
            FocusRect       =   2
            SelectionMode   =   1
            FormatString    =   "<ProductName                            |<Store Name                      |<# Of Cartones  |<ProductID    "
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
            Top             =   2160
            Width           =   6855
         End
         Begin VB.TextBox txtTotalQuantity 
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
            TabIndex        =   7
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox txtNumberPerPackage 
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
            Top             =   3360
            Width           =   1815
         End
         Begin VB.TextBox txtNumberOfPackages 
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
            Top             =   2760
            Width           =   1815
         End
         Begin lvButton_H.lvButtons_H cmdAdd 
            Height          =   375
            Left            =   7200
            TabIndex        =   8
            Top             =   3840
            Width           =   1815
            _ExtentX        =   3201
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
            Image           =   "Form1.frx":245F
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdRemove 
            Height          =   375
            Left            =   7200
            TabIndex        =   9
            Top             =   4320
            Width           =   1815
            _ExtentX        =   3201
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
            Image           =   "Form1.frx":3C21
            cBack           =   -2147483633
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   315
            Left            =   6240
            TabIndex        =   3
            Top             =   960
            Width           =   2655
            _ExtentX        =   4683
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
            CalendarBackColor=   15454586
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   57409539
            CurrentDate     =   39091
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000F&
            X1              =   0
            X2              =   10560
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C29E21&
            Caption         =   "Date of Receival:"
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
            Left            =   4680
            TabIndex        =   30
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C29E21&
            Caption         =   "Receiving Store:"
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
            TabIndex        =   29
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C29E21&
            Caption         =   "Receipt Number:"
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
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C29E21&
            Caption         =   "Supplier:"
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
            TabIndex        =   27
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C29E21&
            Caption         =   "Total Quantity:"
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
            Left            =   4800
            TabIndex        =   16
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label6 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C29E21&
            Caption         =   "Number Of Cartones:"
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
            TabIndex        =   14
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C29E21&
            Caption         =   "ProductName:"
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
            TabIndex        =   13
            Top             =   2280
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmReceivals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String, AdjNoOfPackages As Integer, AdjNoPerPackage As Integer, AdjTotal As Integer
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, xflag As Boolean
Dim Fval2 As Integer, Frow2 As Integer, xx As Integer, eflag As Boolean, Receiptid As String, X As Integer
Dim Supplierid As String, yflag As Boolean, ShowProductflag As Boolean, Storeid As Integer

Private Sub cmdAdd_Click()
Dim X As Integer, a As Integer
'On Error GoTo OkError
'If eflag = False Then
'For xx = 1 To flxgAddProducts.Rows - 1
'           If flxgAddProducts.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgAddProducts.TextMatrix(xx, 4) = Productid Then
'              MsgBox flxgAddProducts.TextMatrix(xx, 0) & " Has Already Been Added!: You Can Only Edit What is Entered"
'              Me.txtProductName.SetFocus
'              Exit Sub
'           End If
'         Next
'End If
Me.txtTotalQuantity = Val(Trim(Me.txtNumberOfPackages)) * Val(Trim(Me.txtNumberPerPackage))
'strg2 = Val(Me.txtUnitPrice) * Val(Me.cboQuantity) * (100 - Val(Me.txtDiscount)) / 100
'Me.txtTotalCost = Format$(strg2, "¢#,###.00")
If Trim(Me.txtProductName) = "" Then MsgBox "Please Select Product", vbInformation, "": Me.txtProductName.SetFocus: Exit Sub
If Trim(Me.txtNumberOfPackages) = "" Then MsgBox "Please Select Number of Cartones", vbInformation, "": Me.txtNumberOfPackages.SetFocus: Exit Sub
'If Trim(Me.txtNumberPerPackage) = "" Then MsgBox "Please Select Number Per Cartone", vbInformation, "": Me.txtNumberPerPackage.SetFocus: Exit Sub
'    If Trim(Me.txtTotalQuantity) = "" Then MsgBox "Please Enter Quantity of Product", vbInformation, "": Me.txtTotalQuantity.SetFocus: Exit Sub
    'If Trim(Me.txtTotalCost) = "" Then MsgBox "Please Compute TotalCost", vbInformation: Me.txtTotalCost.SetFocus: Exit Sub
    If Frow2 <= 0 Then
      If Fval2 > 0 Then
         For xx = 1 To flxgAddProducts.Rows - 1
           If flxgAddProducts.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgAddProducts.TextMatrix(xx, 3) = Productid Then
              MsgBox Me.txtProductName & " Has Already Been Added!: You Can Only Edit What is Entered"
              Me.txtProductName = ""
              Exit Sub
           End If
         Next
       End If
      Fval2 = Fval2 + 1
      flxgAddProducts.Rows = Fval2 + 1
      flxgAddProducts.TextMatrix(Fval2, 0) = Me.txtProductName
      flxgAddProducts.TextMatrix(Fval2, 1) = Me.txtStoreName
      flxgAddProducts.TextMatrix(Fval2, 2) = Me.txtNumberOfPackages
      flxgAddProducts.TextMatrix(Fval2, 3) = Productid
'      flxgAddProducts.TextMatrix(Fval2, 4) = Me.txtTotalQuantity
'      flxgAddProducts.TextMatrix(Fval2, 5) = Productid
      'flxgAddProducts.TextMatrix(Fval2, 5) = Me.txtVat
      'flxgAddProducts.TextMatrix(Fval2, 6) = Me.txtPriceAfterDiscount
       'flxgAddProducts.TextMatrix(Fval2, 7) = Trim(Productid)
       
    Else
       flxgAddProducts.TextMatrix(Frow2, 0) = Me.txtProductName
       flxgAddProducts.TextMatrix(Frow2, 1) = Me.txtStoreName
       flxgAddProducts.TextMatrix(Frow2, 2) = Me.txtNumberOfPackages
       flxgAddProducts.TextMatrix(Frow2, 3) = Productid
'       flxgAddProducts.TextMatrix(Frow2, 4) = Me.txtTotalQuantity
'       flxgAddProducts.TextMatrix(Frow2, 5) = Productid
        'flxgAddProducts.TextMatrix(Frow2, 5) = Me.txtVat
        'flxgAddProducts.TextMatrix(Frow2, 6) = Me.txtPriceAfterDiscount
        'flxgAddProducts.TextMatrix(Frow2, 7) = Trim(Productid)
      Frow2 = 0
    
    End If
    Me.txtProductName = ""
      
    Me.txtNumberOfPackages = ""
    Me.txtNumberPerPackage = ""
    Me.txtTotalQuantity = ""
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
    MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
     Exit Sub
End Sub


Private Sub cmdClear_Click()

For X = 1 To Me.flxgAddProducts.Rows - 1
            flxgAddProducts.TextMatrix(X, 0) = ""
            flxgAddProducts.TextMatrix(X, 1) = ""
            flxgAddProducts.TextMatrix(X, 2) = ""
            flxgAddProducts.TextMatrix(X, 3) = ""
            flxgAddProducts.TextMatrix(X, 4) = ""
            flxgAddProducts.TextMatrix(X, 5) = ""
            'flxgAddProducts.TextMatrix(X, 6) = ""
            'flxgAddProducts.TextMatrix(X, 7) = ""
       Next
    Fval2 = 0
    Frow2 = 0
    Me.flxgAddProducts.Rows = 2
Call ClearCtrls

Me.flxgProducts.Visible = False
Me.flxgSupplier.Visible = False
End Sub

Private Sub cmdExit_Click()
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdRemove_Click()
Dim a As Integer, strgbal As String

On Error GoTo OkError

 If Frow2 = 0 Then Exit Sub
 
  If Frow2 = 1 Then
   If flxgAddProducts.TextMatrix(Frow2, 0) = "" Then
   Frow2 = 0
   ShowProductflag = True
   Me.txtProductName.Text = ""
   ShowProductflag = False
   Me.flxgProducts.Visible = False
    Me.txtNumberOfPackages = ""
    Me.txtNumberPerPackage = ""
    Me.txtTotalQuantity = ""
    Me.txtProductName.SetFocus
   Exit Sub
   End If
 End If
 
   If Frow2 = flxgAddProducts.Rows - 1 Then
     If flxgAddProducts.Rows <> 2 Then
         flxgAddProducts.Rows = flxgAddProducts.Rows - 1
     Else
        For xx = 0 To 3
           flxgAddProducts.TextMatrix(Frow2, xx) = ""
        Next
        ShowProductflag = True
        Me.txtProductName.Text = ""
        ShowProductflag = False
        Me.flxgProducts.Visible = False
        Me.txtNumberOfPackages = ""
        Me.txtNumberPerPackage = ""
        Me.txtTotalQuantity = ""
        Me.txtProductName.SetFocus
     End If
        ShowProductflag = True
        Me.txtProductName.Text = ""
        ShowProductflag = False
        Me.flxgProducts.Visible = False
        Me.txtNumberOfPackages = ""
        Me.txtNumberPerPackage = ""
        Me.txtTotalQuantity = ""
        Me.txtProductName.SetFocus
   Else
         For xx = Frow2 To flxgAddProducts.Rows - 2
            flxgAddProducts.TextMatrix(xx, 0) = flxgAddProducts.TextMatrix(xx + 1, 0)
            flxgAddProducts.TextMatrix(xx, 1) = flxgAddProducts.TextMatrix(xx + 1, 1)
            flxgAddProducts.TextMatrix(xx, 2) = flxgAddProducts.TextMatrix(xx + 1, 2)
            flxgAddProducts.TextMatrix(xx, 3) = flxgAddProducts.TextMatrix(xx + 1, 3)
'            flxgAddProducts.TextMatrix(xx, 4) = flxgAddProducts.TextMatrix(xx + 1, 4)
'            flxgAddProducts.TextMatrix(xx, 5) = flxgAddProducts.TextMatrix(xx + 1, 5)
        Next
        ShowProductflag = True
        Me.txtProductName.Text = ""
        ShowProductflag = False
        Me.flxgProducts.Visible = False
        Me.txtNumberOfPackages = ""
        Me.txtNumberPerPackage = ""
        Me.txtTotalQuantity = ""
        Me.txtProductName.SetFocus
        flxgAddProducts.Rows = flxgAddProducts.Rows - 1
   End If
  Fval2 = Fval2 - 1
  Frow2 = 0
 
  'oflag = False
   Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "SORRY,UNABLE TO REMOVE,TRY AGAIN", vbInformation, "COMPUTATION"
     Exit Sub
  
End Sub

Private Sub cmdSave_Click()

If Trim(Me.txtStoreName) = "" Then
   MsgBox "YOU MUST ENTER RECEIVING STORE'S NAME.", vbInformation, "STORE'S NAME"
   Me.txtStoreName.SetFocus: Exit Sub
End If

 
If Trim(Me.txtSupplier) = "" Then
   MsgBox "YOU MUST ENTER SUPPLIERS NAME.", vbInformation, "SUPPLIERS NAME"
   Me.txtSupplier.SetFocus: Exit Sub
End If
 
 
If Trim(Me.dtpDate) = "" Then
   MsgBox "YOU MUST ENTER DATE OF RECEIVAL.", vbInformation, "DATE OF RECEIVAL"
   Me.dtpDate.SetFocus: Exit Sub
End If


If Me.flxgAddProducts.TextMatrix(1, 0) = "" Then
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
'Me.txtStock = Trim(Me.txtTotalQuantity)
Call Generate_ReceiptID(Receiptid)

cn.BeginTrans
   'cn.Execute "Insert Into MainReceipts ([ReceiptNo],[Amonutpaid],[DateReceived],[MainReceiptsID],[Receipiant],[SupplierID]) select '" & Trim(Me.txtReceiptNo) & "','" & Val(Trim(Me.txtAmount)) & "','" & Trim(Me.dtpDate) & "','" & (Receiptid) & "','" & Trim(Me.txtReceiver) & "','" & Supplierid & "'", Y
   'If Y > 0 Then
   For X = 1 To Me.flxgAddProducts.Rows - 1
   cn.Execute "Insert Into Receivals ([NoOfPackages],[ProductID],[Date],StoreID) select '" & Val(Trim(Me.flxgAddProducts.TextMatrix(X, 2))) & "','" & Trim(Me.flxgAddProducts.TextMatrix(X, 3)) & "','" & Trim(Me.dtpDate) & "','" & Storeid & "'", Y
   If Y > 0 Then
   
    rs.Open "Select * From WholeSaleInventory Where ProductID='" & Trim(Me.flxgAddProducts.TextMatrix(X, 3)) & "' and StoreID='" & Storeid & "'", cn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then

     cn.Execute "Update WholeSaleInventory Set StockLevel ='" & rs.Fields("StockLevel") + Val(Trim(Me.flxgAddProducts.TextMatrix(X, 2))) & "' where ProductID ='" & Trim(Me.flxgAddProducts.TextMatrix(X, 3)) & "' and StoreID='" & Storeid & "'", Y
     If rs.State = 1 Then rs.Close
    Else

     cn.Execute "Insert Into WholeSaleInventory ([StoreID],[ProductID],[StockLevel]) select'" & Storeid & "','" & Trim(Me.flxgAddProducts.TextMatrix(X, 3)) & "','" & Val(Trim(Me.flxgAddProducts.TextMatrix(X, 2))) & "'", Y
     If rs.State = 1 Then rs.Close
    End If
    'cn.Execute "Insert Into [MainReceipts/MainStoreProducts] ([ProductID],[MainReceiptsID]) select '" & Trim(Me.flxgAddProducts.TextMatrix(X, 4)) & "','" & Receiptid & "'", Y
    End If
    Next
   'End If
   If Y > 0 Then
    cn.CommitTrans
    MsgBox "SAVED SUCCESSFULLY", vbInformation, "SAVED"
    yflag = True
   Else
    MsgBox "SAVE FAILED", vbInformation, "TRY AGAIN"
    yflag = True
   End If

   If yflag = True Then
    
      For X = 1 To Me.flxgAddProducts.Rows - 1
            flxgAddProducts.TextMatrix(X, 0) = ""
            flxgAddProducts.TextMatrix(X, 1) = ""
            flxgAddProducts.TextMatrix(X, 2) = ""
            flxgAddProducts.TextMatrix(X, 3) = ""
'            flxgAddProducts.TextMatrix(X, 4) = ""
'            flxgAddProducts.TextMatrix(X, 5) = ""
       Next
    Fval2 = 0
    Me.flxgAddProducts.Rows = 2
    yflag = False
   End If
   Call ClearCtrls
   Me.flxgSupplier.Visible = False
   Me.flxgProducts.Visible = False
   Me.txtReceiptNo.SetFocus
Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub




Private Sub dtpDate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
'   Me.txtAmount.SetFocus
End If
End Sub

Private Sub TetxtReceiver_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

End Sub

Private Sub flxgAddProducts_Click()
' xflag = True
        ShowProductflag = True
        Me.txtProductName = flxgAddProducts.TextMatrix(flxgAddProducts.Row, 0)
        ShowProductflag = False
        Me.txtStoreName = flxgAddProducts.TextMatrix(flxgAddProducts.Row, 1)
        Me.txtNumberOfPackages = flxgAddProducts.TextMatrix(flxgAddProducts.Row, 2)
        Productid = flxgAddProducts.TextMatrix(flxgAddProducts.Row, 3)
'        Me.txtTotalQuantity = flxgAddProducts.TextMatrix(flxgAddProducts.Row, 4)
'        Productid = flxgAddProducts.TextMatrix(flxgAddProducts.Row, 5)
        
Frow2 = flxgAddProducts.Row
eflag = True
'oflag = False
Me.cmdSave.Enabled = True
Me.txtNumberOfPackages.SetFocus
End Sub

Private Sub flxgProducts_Click()
Me.txtProductName = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 0)
Productid = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 1)
Me.flxgProducts.Visible = False
Me.txtNumberOfPackages = ""
'Me.txtNumberPerPackage = ""
'Me.txtTotalQuantity = ""
Me.cmdSave.Enabled = True
sflag = False
oflag = False
Me.txtNumberOfPackages.SetFocus
End Sub

Private Sub flxgProducts_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtProductName = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 0)

Productid = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 1)


Me.flxgProducts.Visible = False

Me.cmdSave.Enabled = True
sflag = False
oflag = False
Me.txtNumberOfPackages.SetFocus
End If
End Sub

Private Sub flxgStoreName_Click()
Me.txtStoreName = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 0)

Storeid = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 1)
'Call FindStockLevel
'Call FindRetailStockLevel
Me.flxgStoreName.Visible = False
Me.txtSupplier.SetFocus
End Sub

Private Sub flxgSupplier_Click()
        Me.txtSupplier = flxgSupplier.TextMatrix(flxgSupplier.Row, 0)
        Supplierid = flxgSupplier.TextMatrix(flxgSupplier.Row, 1)
        flxgSupplier.Visible = False
        Me.txtReceiptNo.SetFocus
End Sub

Private Sub Form_Load()
Me.dtpDate = Date

CenterForm Me
Me.Height = 9315
Me.Width = 9960
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
Frow2 = 0
Fval2 = 0
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
'   Me.txtReceiver.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtNumberOfPackages_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.txtNumberPerPackage.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtNumberPerPackage_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.txtTotalQuantity.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtNumberPerPackage_LostFocus()
Me.txtTotalQuantity = Val(Trim(Me.txtNumberOfPackages)) * Val(Trim(Me.txtNumberPerPackage))
End Sub

Private Sub txtProductName_Change()
If Me.txtProductName = "" Then
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
   
If ShowProductflag = False Then  ' when true u can view grid to add products,but when false u can edit with no grid shown
   rs.Open "Select Distinct  ProductName,ProductID From Products  Where ProductName Like '" & Trim(Me.txtProductName) & "%" & "' Order By ProductName", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   flxgProducts.Height = 950 + (285 * (rs.RecordCount - 1))
   
   If flxgProducts.Height >= 4455 Then
      flxgProducts.Height = 4455
   End If
    flxgProducts.Rows = rs.RecordCount + 1
   With flxgProducts
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("ProductName")
       ' .TextMatrix(X, 3) = rs.Fields("VAT")
       .TextMatrix(X, 1) = rs.Fields("ProductID")
             
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 1
      .RowSel = 1
   End With
   flxgProducts.Visible = True
   If rs.State = 1 Then rs.Close
   Me.cmdSave.Enabled = True
Else
     flxgProducts.Visible = False
     'Dquantity = "1"
      If cFlag = True Then
       If rs.State <> 0 Then
          rs.Close
       End If
       cFlag = False: Exit Sub
      End If
     
End If
Else
      flxgProducts.Visible = False
      'If cn.State = 1 Then cn.Close
      'If rs.State = 1 Then rs.Close
      'oflag = False
      xflag = False ': Exit Sub
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

Private Sub txtProductName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If KeyAscii = vbKeyReturn Then
    Me.flxgProducts.Visible = True
   Me.flxgProducts.SetFocus
End If
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

End Sub

Private Sub txtReceiptNo_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If KeyAscii = vbKeyReturn Then
   Me.dtpDate.SetFocus
End If
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

End Sub

Private Sub txtReceiver_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If KeyAscii = vbKeyReturn Then
   Me.txtProductName.SetFocus
End If
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
End Sub

Private Sub txtStoreName_Change()
If Me.txtStoreName = "" Then
Me.flxgStoreName.Visible = False
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
   
   If ShowStoreflag = False Then
   
   rs.Open "Select * From Stores  Where StoreName Like '" & Trim(Me.txtStoreName) & "%" & "' Order By StoreName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   flxgStoreName.Height = 950 + (285 * (rs.RecordCount - 1))
   
   If flxgStoreName.Height >= 4455 Then
      flxgStoreName.Height = 4455
   End If
    flxgStoreName.Rows = rs.RecordCount + 1
   With flxgStoreName
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("StoreName")
       .TextMatrix(X, 1) = rs.Fields("StoreID")
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 1
      .RowSel = 1
   End With
   flxgStoreName.Visible = True
   If rs.State = 1 Then rs.Close
'   Me.cmdSave.Enabled = True
 Else
     flxgStoreName.Visible = False
     If rs.State = 1 Then rs.Close
      
     
End If

 




If rs.State = 1 Then rs.Close
End If
Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
    MsgBox "Stores could not display", , "Displaying"
     Exit Sub
End Sub

Private Sub txtStoreName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtSupplier.SetFocus
End If

End Sub

Private Sub txtSupplier_Change()
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
'If xflag = False Then  ' when true u can view grid to add products,but when false u can edit with no grid shown
   rs.Open "Select *  From Suppliers  Where SupplierName Like '" & Trim(Me.txtSupplier) & "%" & "' Order By SupplierName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   'flxgProducts.Height = 950 + (285 * (rs.RecordCount - 1))
   
   'If flxgProducts.Height >= 4455 Then
      'flxgProducts.Height = 4455
   'End If
    flxgSupplier.Rows = rs.RecordCount + 1
   With flxgSupplier
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("SupplierName")
       ' .TextMatrix(X, 3) = rs.Fields("VAT")
       .TextMatrix(X, 1) = rs.Fields("SupplierID")
             
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 1
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

Private Sub txtSupplier_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtReceiptNo.SetFocus
End If

End Sub

Private Sub txtTotalQuantity_GotFocus()
Me.txtTotalQuantity = Val(Trim(Me.txtNumberOfPackages)) * Val(Trim(Me.txtNumberPerPackage))
End Sub

Private Sub txtTotalQuantity_KeyPress(KeyAscii As Integer)
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
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
End Sub
Private Function Generate_ReceiptID(Receipt_ID As String) As Boolean
Dim strg As String, strg1 As String, strg2 As String, bFlag As Boolean
On Error GoTo SaveError
rs.Open "Select MainReceiptsID From MainReceipts  order by MainReceiptsID Desc", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
      strg1 = Trim(rs.Fields!MainReceiptsID)

   strg1 = Trim(Str(Val(strg1) + 1))
   strg1 = String$(8 - Len(strg1), "0") & strg1
Else
   strg1 = "00000001"
End If
Receipt_ID = strg1

If rs.State = 1 Then rs.Close

Generate_ReceiptID = True

Exit Function
SaveError:
     If rs.State = 1 Then rs.Close
    Generate_ReceiptID = False
     Exit Function
     
End Function
