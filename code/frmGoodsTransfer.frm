VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmGoodsTransfer 
   BackColor       =   &H00C29E21&
   Caption         =   "GoodsTransfer"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   Icon            =   "frmGoodsTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   9780
   Begin VB.Frame Frame5 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   7560
      Width           =   9495
      Begin lvButton_H.lvButtons_H cmdSave 
         Height          =   375
         Left            =   120
         TabIndex        =   20
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
         Image           =   "frmGoodsTransfer.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdWholeSaleStock 
         Height          =   375
         Left            =   6720
         TabIndex        =   21
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Caption         =   "&Adjust WholeSale Stock"
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
         Image           =   "frmGoodsTransfer.frx":075C
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdClear 
         Height          =   375
         Left            =   1800
         TabIndex        =   22
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
         Image           =   "frmGoodsTransfer.frx":0A76
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   5040
         TabIndex        =   23
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
         Image           =   "frmGoodsTransfer.frx":2327
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdReport 
         Height          =   375
         Left            =   3480
         TabIndex        =   26
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
         Image           =   "frmGoodsTransfer.frx":2779
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   7215
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   9495
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProductSearch 
         Height          =   1695
         Left            =   1920
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   16117969
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   "<Product Description                                              |>Cartone Price              |<ProductID    |<NoPerCartone   "
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
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProducts 
         Height          =   2895
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5106
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
         FormatString    =   "<Product Description                           |<Quantity  |>Rate            |>Value                       |<ProductID    "
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
         _Band(0).Cols   =   5
      End
      Begin VB.ComboBox cboQuantity 
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
         ItemData        =   "frmGoodsTransfer.frx":402A
         Left            =   1560
         List            =   "frmGoodsTransfer.frx":40A6
         TabIndex        =   4
         Top             =   2760
         Width           =   2895
      End
      Begin lvButton_H.lvButtons_H cmdAdd 
         Height          =   375
         Left            =   6000
         TabIndex        =   7
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "frmGoodsTransfer.frx":4141
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdRemove 
         Height          =   375
         Left            =   7680
         TabIndex        =   12
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "frmGoodsTransfer.frx":5903
         cBack           =   -2147483633
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgDestinationStore 
         Height          =   1095
         Left            =   6240
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1931
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
         FormatString    =   "<StoreName                                |<StoreID"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgSourceStore 
         Height          =   1095
         Left            =   1560
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1931
         _Version        =   393216
         BackColor       =   16117969
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         FillStyle       =   1
         SelectionMode   =   1
         FormatString    =   "<StoreName                               |<StoreID             "
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
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.DTPicker dtpIsueingDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.TextBox txtDestination 
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
         Left            =   6240
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtSource 
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
      Begin VB.TextBox txtCartonePrice 
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
         TabIndex        =   6
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtProductName 
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
         TabIndex        =   3
         Top             =   2280
         Width           =   7695
      End
      Begin VB.TextBox txtNoPerCartone 
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
         TabIndex        =   5
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C29E21&
         Caption         =   "Rate:"
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
         TabIndex        =   25
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C29E21&
         Caption         =   "No Per Cartone:"
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
         TabIndex        =   24
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000F&
         X1              =   0
         X2              =   9480
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "Source Store:"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Destination Store:"
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
         Left            =   4560
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Issueing Date:"
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
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Quantity in (Cartones):"
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
         TabIndex        =   11
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Product Name:"
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
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmGoodsTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Storeid As Integer
Dim sflag As Boolean, ListProductID As String, ctrl As Control, Destinationid As Integer, Sourceid As Integer
Dim Productid As String, Fval2 As Integer, Frow2 As Integer, xx As Integer, ShowProductsflag As Boolean, ShowStoreflag As Boolean, ShowDestinationflag As Boolean

Private Sub cboQuantity_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,-+"
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

Private Sub cmdAdd_Click()
Dim X As Integer, a As Integer
'On Error GoTo OkError
'If eflag = False Then
'For xx = 1 To flxgProducts.Rows - 1
'           If flxgProducts.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgProducts.TextMatrix(xx, 6) = Productid And flxgProducts.TextMatrix(xx, 3) = Trim(Me.txtSource) And flxgProducts.TextMatrix(xx, 4) = Trim(Me.txtDestination) Then
'              MsgBox Me.txtProductName & " Has Already Been Added!: You Can Only Edit What is Entered", vbInformation, ""
'              ShowProductsflag = True
'              Me.txtProductName = ""
'              Me.cboQuantity = ""
'              ShowProductsflag = False
'              Me.txtProductName.SetFocus
'              Exit Sub
'           End If
'         Next
'End If
'Me.txtTotalQuantity = Val(Trim(Me.txtNoPackages)) * Val(Trim(Me.txtNoPerPackage))
''strg2 = Val(Me.txtUnitPrice) * Val(Me.cboQuantity) * (100 - Val(Me.txtDiscount)) / 100
'Me.txtTotalCost = Format$(strg2, "¢#,###.00")
If Trim(Me.txtSource) = "" Then MsgBox "From which Store are you issuing goods", vbInformation, "": Me.txtSource.SetFocus: Exit Sub
If Trim(Me.txtDestination) = "" Then MsgBox "Please select destination of Goods", vbInformation, "": Me.txtDestination.SetFocus: Exit Sub
If Trim(Me.txtProductName) = "" Then MsgBox "Please Select Product", vbInformation, "ProductName": Me.txtProductName.SetFocus: Exit Sub
    If Trim(Me.cboQuantity) = "" Or Trim(Me.cboQuantity) = 0 Then MsgBox "Please Enter Quantity of Product", vbInformation, "Quantity": Me.cboQuantity.SetFocus: Exit Sub
    
    
'    If MsgBox("ARE YOU SURE  ABOUT THE ISSUING DATE ENTERED?", vbYesNo + vbQuestion, "CONFIRM ISSUING DATE") = vbNo Then
'    Me.dtpIsueingDate.SetFocus
'    Exit Sub
'    End If
    
    
    'If Trim(Me.txtTotalCost) = "" Then MsgBox "Please Compute TotalCost", vbInformation: Me.txtTotalCost.SetFocus: Exit Sub
    If Frow2 <= 0 Then
      If Fval2 > 0 Then
         For xx = 1 To flxgProducts.Rows - 1
           If flxgProducts.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgProducts.TextMatrix(xx, 4) = Productid Then
              MsgBox Me.txtProductName & " Has Already Been Added!: You Can Only Edit What is Entered", vbInformation, ""
              ShowProductsflag = True
              Me.txtProductName = ""
              Me.cboQuantity = ""
              ShowProductsflag = False
              Me.txtProductName.SetFocus
              Exit Sub
           End If
         Next
       End If
      Fval2 = Fval2 + 1
      flxgProducts.Rows = Fval2 + 1
      flxgProducts.TextMatrix(Fval2, 0) = Me.txtProductName
      flxgProducts.TextMatrix(Fval2, 1) = Me.cboQuantity
      flxgProducts.TextMatrix(Fval2, 2) = Format$(Val(Me.txtCartonePrice), "#,###.00")
      flxgProducts.TextMatrix(Fval2, 3) = Format$(Val(Me.txtCartonePrice) * Val(Me.cboQuantity), "#,###.00")
      flxgProducts.TextMatrix(Fval2, 4) = Productid
'      flxgProducts.TextMatrix(Fval2, 5) = Sourceid'Format$(Val(Me.txtCartonePrice) * val(Me.cboQuantity), "#,###.00")
'      flxgProducts.TextMatrix(Fval2, 6) = Productid
'      flxgProducts.TextMatrix(Fval2, 7) = Destinationid
       'flxgProducts.TextMatrix(Fval2, 7) = Trim(Productid)
       
    Else
       flxgProducts.TextMatrix(Frow2, 0) = Me.txtProductName
       flxgProducts.TextMatrix(Frow2, 1) = Me.cboQuantity
       flxgProducts.TextMatrix(Frow2, 2) = Format$(Val(Me.txtCartonePrice), "#,###.00")
       flxgProducts.TextMatrix(Frow2, 3) = Format$(Val(Me.txtCartonePrice) * Val(Me.cboQuantity), "#,###.00")
       flxgProducts.TextMatrix(Fval2, 4) = Productid
'        flxgProducts.TextMatrix(Frow2, 5) = Sourceid
'        flxgProducts.TextMatrix(Frow2, 6) = Productid
'       flxgProducts.TextMatrix(Frow2, 7) = Destinationid
        'flxgProducts.TextMatrix(Frow2, 7) = Trim(Productid)
      Frow2 = 0
    
    End If
    Me.cboQuantity = ""
    Me.txtCartonePrice = ""
    ShowProductsflag = True
    Me.txtProductName = ""
    ShowProductsflag = False
    Me.txtProductName.SetFocus
'    Me.cmdSave.Enabled = True
    
    
    
    
    
'    If eflag = True Then
'    eflag = False
'   End If
   
   Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
    MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
     Exit Sub
End Sub

Private Sub cmdClear_Click()
Call ClearCtrls
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRemove_Click()
Dim a As Integer, strgbal As String

On Error GoTo OkError

 If Frow2 = 0 Then Exit Sub
 
 
 If Frow2 = 1 Then
   If flxgProducts.TextMatrix(Frow2, 0) = "" Then
   Frow2 = 0
    ShowProductsflag = True
    Me.txtProductName.Text = ""
    ShowProductsflag = False
    
    ShowStoreflag = True
    Me.txtSource = ""
    ShowStoreflag = False
    
    ShowDestinationflag = True
    Me.txtDestination = ""
    ShowDestinationflag = False
    
    Me.txtCartonePrice = ""
    Me.cboQuantity = ""
    Me.txtSource.SetFocus
   Exit Sub
   End If
 End If

   If Frow2 = flxgProducts.Rows - 1 Then
     If flxgProducts.Rows <> 2 Then
         flxgProducts.Rows = flxgProducts.Rows - 1
     Else
        For xx = 0 To 4
           flxgProducts.TextMatrix(Frow2, xx) = ""
        Next
            ShowProductsflag = True
            Me.txtProductName.Text = ""
            ShowProductsflag = False
            
            ShowStoreflag = True
            Me.txtSource = ""
            ShowStoreflag = False
            
            ShowDestinationflag = True
            Me.txtDestination = ""
            ShowDestinationflag = False
            
            Me.txtCartonePrice = ""
            Me.cboQuantity = ""
            Me.txtSource.SetFocus
     End If
            ShowProductsflag = True
            Me.txtProductName.Text = ""
            ShowProductsflag = False
            
            ShowStoreflag = True
            Me.txtSource = ""
            ShowStoreflag = False
            
            ShowDestinationflag = True
            Me.txtDestination = ""
            ShowDestinationflag = False
            
            Me.txtCartonePrice = ""
            Me.cboQuantity = ""
            Me.txtSource.SetFocus
   Else
         For xx = Frow2 To flxgProducts.Rows - 2
            flxgProducts.TextMatrix(xx, 0) = flxgProducts.TextMatrix(xx + 1, 0)
            flxgProducts.TextMatrix(xx, 1) = flxgProducts.TextMatrix(xx + 1, 1)
            flxgProducts.TextMatrix(xx, 2) = flxgProducts.TextMatrix(xx + 1, 2)
            flxgProducts.TextMatrix(xx, 3) = flxgProducts.TextMatrix(xx + 1, 3)
            flxgProducts.TextMatrix(xx, 4) = flxgProducts.TextMatrix(xx + 1, 4)
'            flxgProducts.TextMatrix(xx, 5) = flxgProducts.TextMatrix(xx + 1, 5)
'            flxgProducts.TextMatrix(xx, 6) = flxgProducts.TextMatrix(xx + 1, 6)
'            flxgProducts.TextMatrix(xx, 7) = flxgProducts.TextMatrix(xx + 1, 7)
        Next
            ShowProductsflag = True
            Me.txtProductName.Text = ""
            ShowProductsflag = False
            
            ShowStoreflag = True
            Me.txtSource = ""
            ShowStoreflag = False
            
            ShowDestinationflag = True
            Me.txtDestination = ""
            ShowDestinationflag = False
            
            Me.txtCartonePrice = ""
            Me.cboQuantity = ""
            Me.txtSource.SetFocus
        flxgProducts.Rows = flxgProducts.Rows - 1
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

Private Sub cmdReport_Click()
frmGoodsTransferRpt.Show
End Sub

Private Sub cmdSave_Click()
'On Error GoTo OkError
If Me.txtSource = "" Then
    MsgBox "Select Store Name", vbInformation, "Store Name"
    Me.txtSource.SetFocus: Exit Sub
End If

If Me.txtDestination = "" Then
  MsgBox "Select Destination Name", vbInformation, "Destination Name"
  Me.txtDestination.SetFocus: Exit Sub
End If
'If Me.txtProductName = "" Then
'     MsgBox "Select Products to be issued", vbInformation, "Select Products"
'     Me.txtProductName.SetFocus: Exit Sub
'End If

'If Me.cboQuantity = "" Then
'     MsgBox "Select Quantity of Products to be issued", vbInformation, "Select Quantity"
'     Me.cboQuantity.SetFocus: Exit Sub
'End If

For X = 1 To Me.flxgProducts.Rows - 1
   If Me.flxgProducts.TextMatrix(X, 0) = "" Then
     MsgBox "Add Products to the grid below", vbInformation, "Add Products"
     Me.txtProductName.SetFocus: Exit Sub
   End If
Next



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
   For X = 1 To Me.flxgProducts.Rows - 1
   
    cn.Execute "Insert Into GoodsTransfer (ProductID,Storeid,Quantity,Destination,TransferDate,TransferTime,Source,Rate) select '" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "','" & Sourceid & "','" & Trim(Me.flxgProducts.TextMatrix(X, 1)) & "','" & Trim(Me.txtDestination) & "','" & Me.dtpIsueingDate & "','" & Format$(Me.dtpIsueingDate, "General Date") & "','" & Trim(Me.txtSource) & "','" & Trim(Me.flxgProducts.TextMatrix(X, 2)) & "'", Y
     Next
   If Y > 0 Then
'   X = 0
   For X = 1 To Me.flxgProducts.Rows - 1
   If Trim(Me.txtSource) <> "RETAIL" Then
        rs.Open "Select * From WholeSaleInventory  Where Storeid= '" & Sourceid & "' and ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", cn, adOpenForwardOnly, adLockReadOnly
         If rs.RecordCount > 0 Then
             If rs.Fields("StockLevel") >= Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) Then
              cn.Execute "Update WholeSaleInventory Set StockLevel ='" & rs.Fields("StockLevel") - Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) & "' Where Storeid ='" & Sourceid & "' and ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", Y
             Else
              Y = 1
             End If
             If rs.State = 1 Then rs.Close
         Else
'             If Trim(Me.flxgProducts.TextMatrix(X, 5)) <> "" Then
'             cn.Execute "Insert Into WholeSaleInventory (ProductID,Sourceid,StockLevel) select '" & Trim(Me.flxgProducts.TextMatrix(X, 6)) & "','" & Trim(Me.flxgProducts.TextMatrix(X, 5)) & "','" & Trim(Me.flxgProducts.TextMatrix(X, 1)) & "'", Y
'             End If
'             If rs.State = 1 Then rs.Close
         End If
             If rs.State = 1 Then rs.Close
    End If
   If Trim(Me.txtSource) = "RETAIL" Then
         
         rs.Open "Select * From Products inner Join ProductInventory on Products.ProductID=ProductInventory.ProductID where  Products.ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", cn, adOpenForwardOnly, adLockReadOnly
         If rs.RecordCount > 0 Then
            If rs.Fields("StockLevel") >= Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) Then
              cn.Execute "Update ProductInventory Set StockLevel ='" & rs.Fields("StockLevel") - Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) & "' Where ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", Y
            Else
              Y = 1
            End If
            If rs.State = 1 Then rs.Close
          End If
    End If
'    If Trim(Me.flxgProducts.TextMatrix(X, 5)) = "" Then
''              Y = 1
         
''    End If
            If rs.State = 1 Then rs.Close
    Next
   End If
   
   
   If Y > 0 Then
'   X = 0
      For X = 1 To Me.flxgProducts.Rows - 1
          If Trim(Me.txtDestination) <> "RETAIL" Then
                rs.Open "Select * From WholeSaleInventory  Where Storeid= '" & Destinationid & "' and ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", cn, adOpenForwardOnly, adLockReadOnly
                 If rs.RecordCount > 0 Then
'                        If rs.Fields("StockLevel") >= Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) Then
                         cn.Execute "Update WholeSaleInventory Set StockLevel ='" & rs.Fields("StockLevel") + Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) & "' Where Sourceid ='" & Destinationid & "' and ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", Y
'                        Else
'                         Y = 1
'                        End If
                        If rs.State = 1 Then rs.Close
                Else
'                    If Trim(Me.flxgProducts.TextMatrix(X, 7)) <> "" Then
                       cn.Execute "Insert Into WholeSaleInventory (ProductID,Storeid,StockLevel) select '" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "','" & Destinationid & "','" & Trim(Me.flxgProducts.TextMatrix(X, 1)) & "'", Y
'                    End If
                    If rs.State = 1 Then rs.Close
                End If
                    If rs.State = 1 Then rs.Close
         End If
        If Trim(Me.txtDestination) = "RETAIL" Then
              
              rs.Open "Select * From Products inner Join ProductInventory on Products.ProductID=ProductInventory.ProductID where  Products.ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", cn, adOpenForwardOnly, adLockReadOnly
              If rs.RecordCount > 0 Then
'                 If rs.Fields("StockLevel") >= Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) Then
                   cn.Execute "Update ProductInventory Set StockLevel ='" & rs.Fields("StockLevel") + Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) & "' Where ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", Y
'                 Else
'                   Y = 1
'                 End If
                 If rs.State = 1 Then rs.Close
              End If
                 If rs.State = 1 Then rs.Close
         End If
'         If Destinationid = "" Then
'              Y = 1
'         End If
         Next
   End If
      
            If Y > 0 Then
              cn.CommitTrans
              MsgBox "Saved successfully", vbInformation, "Saved successfully"
              Fval2 = 0
              Call ClearCtrls
            Else
              cn.RollbackTrans
              MsgBox "Save Failed,Please Retype Data and Save Again", vbInformation, "Try Again"
            End If
            
            
            
   Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,THERE IS AN ERROR,TRY AGAIN", vbInformation, "COMPUTATION"
     Exit Sub
End Sub

Private Sub cmdWholeSaleStock_Click()
frmAdjustWholeSaleStock.Show
frmAdjustWholeSaleStock.txtStoreName = Me.txtSource
frmAdjustWholeSaleStock.flxgStoreName.Visible = False
frmAdjustWholeSaleStock.txtProductName.SetFocus
End Sub

Private Sub flxgDestinationStore_Click()
Me.txtDestination = Me.flxgDestinationStore.TextMatrix(Me.flxgDestinationStore.Row, 0)
If Me.txtSource <> "" And (Me.txtDestination = Me.txtSource) Then
    MsgBox "You cannot issue goods to the same store", vbInformation, "Source and Destination should not be the same"
    Me.txtDestination = "": Me.txtDestination.SetFocus: Me.flxgDestinationStore.Visible = False: Exit Sub
End If
Destinationid = Me.flxgDestinationStore.TextMatrix(Me.flxgDestinationStore.Row, 1)
Me.flxgDestinationStore.Visible = False
Me.dtpIsueingDate.SetFocus
End Sub

Private Sub flxgDestinationStore_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then

Me.txtDestination = Me.flxgDestinationStore.TextMatrix(Me.flxgDestinationStore.Row, 0)
If Me.txtSource <> "" And (Me.txtDestination = Me.txtSource) Then
    MsgBox "You cannot issue goods to the same store", vbInformation, "Source and Destination should not be the same"
    Me.txtDestination = "": Me.txtDestination.SetFocus: Me.flxgDestinationStore.Visible = False: Exit Sub
End If
Destinationid = Me.flxgDestinationStore.TextMatrix(Me.flxgDestinationStore.Row, 1)
Me.flxgDestinationStore.Visible = False
Me.dtpIsueingDate.SetFocus

End If
End Sub

Private Sub flxgProducts_Click()
        ShowProductsflag = True
        Me.txtProductName = flxgProducts.TextMatrix(flxgProducts.Row, 0)
        ShowProductsflag = False
        Me.cboQuantity = flxgProducts.TextMatrix(flxgProducts.Row, 1)
        Me.txtCartonePrice = flxgProducts.TextMatrix(flxgProducts.Row, 2)
        Productid = flxgProducts.TextMatrix(flxgProducts.Row, 4)
        
        Frow2 = flxgProducts.Row
        Me.cboQuantity.SetFocus
End Sub

Private Sub flxgProductSearch_Click()
Me.txtProductName = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 0)
Me.txtCartonePrice = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 1)
Productid = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 2)
Me.txtNoPerCartone = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 3)
Me.flxgProductSearch.Visible = False
Me.cboQuantity.SetFocus
End Sub

Private Sub flxgProductSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then

Me.txtProductName = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 0)
Me.txtCartonePrice = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 1)
Productid = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 2)
Me.txtNoPerCartone = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 3)
Me.flxgProductSearch.Visible = False
Me.cboQuantity.SetFocus

End If
End Sub

Private Sub flxgSourceStore_Click()
Me.txtSource = Me.flxgSourceStore.TextMatrix(Me.flxgSourceStore.Row, 0)
If Me.txtDestination <> "" And (Me.txtDestination = Me.txtSource) Then
    MsgBox "You cannot issue goods to the same store", vbInformation, "Source and Destination should not be the same"
    Me.txtSource = "": Me.txtSource.SetFocus: Me.flxgSourceStore.Visible = False: Exit Sub
End If
Sourceid = Me.flxgSourceStore.TextMatrix(Me.flxgSourceStore.Row, 1)
Me.flxgSourceStore.Visible = False
Me.txtDestination.SetFocus
End Sub

Private Sub flxgSourceStore_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then

Me.txtSource = Me.flxgSourceStore.TextMatrix(Me.flxgSourceStore.Row, 0)
If Me.txtDestination <> "" And (Me.txtDestination = Me.txtSource) Then
    MsgBox "You cannot issue goods to the same store", vbInformation, "Source and Destination should not be the same"
    Me.txtSource = "": Me.txtSource.SetFocus: Me.flxgSourceStore.Visible = False: Exit Sub
End If
Sourceid = Me.flxgSourceStore.TextMatrix(Me.flxgSourceStore.Row, 1)
Me.flxgSourceStore.Visible = False
Me.txtDestination.SetFocus

End If
End Sub

Private Sub Form_Load()
Me.dtpIsueingDate = Date
Me.Height = 9045
  Me.Width = 9900
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub txtCartonePrice_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,-+"
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

Private Sub txtDestination_Change()
If Me.txtDestination = "" Then
 Me.flxgDestinationStore.Visible = False
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

   If ShowDestinationflag = False Then

   rs.Open "Select * From Stores  Where StoreName Like '" & Trim(Me.txtDestination) & "%" & "' Order By StoreName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   flxgDestinationStore.Height = 950 + (285 * (rs.RecordCount - 1))
   
   If flxgDestinationStore.Height >= 4455 Then
      flxgDestinationStore.Height = 4455
   End If
    flxgDestinationStore.Rows = rs.RecordCount + 1
   With flxgDestinationStore
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("StoreName")
       .TextMatrix(X, 1) = rs.Fields("Storeid")
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 1
      .RowSel = 1
   End With
      flxgDestinationStore.Visible = True
      If rs.State = 1 Then rs.Close
 Else
     flxgDestinationStore.Visible = False
     If rs.State = 1 Then rs.Close
     MsgBox "'" & Me.txtDestination & "' is a Not in the System !," & " " & " To SetUp a new Store goto Set-Ups---->SetUp Stores", vbInformation, "What are you typing"
     Me.txtDestination = ""
     Me.txtDestination.SetFocus
      
     
End If
If rs.State = 1 Then rs.Close
End If
Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "Stores could not display", , "Displaying"
     Exit Sub
End Sub

Private Sub txtDestination_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
 Me.flxgDestinationStore.SetFocus
End If
End Sub

Private Sub txtDestination_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.dtpIsueingDate.SetFocus
End If
End Sub

Private Sub txtNoPerCartone_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,-+"
If KeyAscii = vbKeyReturn Then
   Me.txtCartonePrice.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtProductName_Change()

If Me.txtProductName = "" Then
 Me.flxgProductSearch.Visible = False
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

  If ShowProductsflag = False Then

   rs.Open "Select * From Products  Where ProductName Like '" & Trim(Me.txtProductName) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   flxgProductSearch.Height = 950 + (285 * (rs.RecordCount - 1))
   
   If flxgProductSearch.Height >= 4455 Then
      flxgProductSearch.Height = 4455
   End If
    flxgProductSearch.Rows = rs.RecordCount + 1
   With flxgProductSearch
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("ProductName")
       .TextMatrix(X, 1) = rs.Fields("PricePerCartone")
       .TextMatrix(X, 2) = rs.Fields("ProductID")
       .TextMatrix(X, 3) = rs.Fields("BaseUnit")
       
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 3
      .RowSel = 1
   End With
   flxgProductSearch.Visible = True
   If rs.State = 1 Then rs.Close
'   Me.cmdSave.Enabled = True
 Else
     flxgProductSearch.Visible = False
     If rs.State = 1 Then rs.Close
     MsgBox "'" & Me.txtProductName & "' is a Not in the System !," & " " & " To SetUp a new Product goto Set-Ups---->SetUp Products", vbInformation, "What are you typing"
     Me.txtProductName = ""
     Me.txtProductName.SetFocus
     
End If
If rs.State = 1 Then rs.Close
End If
Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
    MsgBox "Stores could not display", , "Displaying"
     Exit Sub
End Sub

Private Sub txtProductName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
 Me.flxgProductSearch.SetFocus
End If
End Sub

Private Sub txtProductName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.cboQuantity.SetFocus
End If
End Sub

Private Sub txtSource_Change()
If Me.txtSource = "" Then
Me.flxgSourceStore.Visible = False
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
   
   rs.Open "Select * From Stores  Where StoreName Like '" & Trim(Me.txtSource) & "%" & "' Order By StoreName ", cn, adOpenForwardOnly, adLockReadOnly
   
If rs.RecordCount > 0 Then
   flxgSourceStore.Height = 950 + (285 * (rs.RecordCount - 1))
   
   If flxgSourceStore.Height >= 4455 Then
      flxgSourceStore.Height = 4455
   End If
    flxgSourceStore.Rows = rs.RecordCount + 1
   With flxgSourceStore
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("StoreName")
       .TextMatrix(X, 1) = rs.Fields("Storeid")
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 1
      .RowSel = 1
   End With
   flxgSourceStore.Visible = True
   If rs.State = 1 Then rs.Close
'   Me.cmdSave.Enabled = True
 Else
     flxgSourceStore.Visible = False
     If rs.State = 1 Then rs.Close
     MsgBox "'" & Me.txtSource & "' is a Not in the System !," & " " & " To SetUp a new Store goto Set-Ups---->SetUp Stores", vbInformation, "What are you typing"
     Me.txtSource = ""
     Me.txtSource.SetFocus
     
End If
If rs.State = 1 Then rs.Close
End If
Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
    MsgBox "Stores could not display", , "Displaying"
     Exit Sub
End Sub

Private Sub txtSource_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
 Me.flxgSourceStore.SetFocus
End If
End Sub

Private Sub txtSource_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtDestination.SetFocus
End If
End Sub
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
   ctrl = ""
   End If
Next
Fval2 = 0
Frow2 = 0
sflag = False
For X = 0 To 4
Me.flxgProducts.TextMatrix(1, X) = ""
Next
Me.flxgProducts.Rows = 2
Me.flxgSourceStore.Visible = False
Me.flxgDestinationStore.Visible = False
Productid = ""
'Sourceid = ""
'Destinationid = ""
Me.txtSource.SetFocus
End Sub

