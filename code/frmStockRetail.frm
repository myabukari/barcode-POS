VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmStockRetail 
   BackColor       =   &H00C29E21&
   Caption         =   "Supply to Retail"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   Icon            =   "frmStockRetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   10560
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   10335
      Begin VB.Frame Frame5 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   10095
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   7800
            TabIndex        =   18
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
            Image           =   "frmStockRetail.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   2760
            TabIndex        =   19
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
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
            Image           =   "frmStockRetail.frx":075C
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
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
            Image           =   "frmStockRetail.frx":200D
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdReport 
            Height          =   375
            Left            =   5400
            TabIndex        =   26
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            Image           =   "frmStockRetail.frx":245F
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   7095
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   10335
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgStoreName 
         Height          =   1215
         Left            =   2040
         TabIndex        =   23
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
         ForeColorSel    =   -2147483641
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   "<StoreName                                                   |<StoreID             "
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
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProductName 
         Height          =   2295
         Left            =   2040
         TabIndex        =   24
         Top             =   2160
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4048
         _Version        =   393216
         BackColor       =   16117969
         ForeColor       =   -2147483630
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12615680
         ForeColorSel    =   -2147483641
         BackColorBkg    =   12754465
         GridColor       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"frmStockRetail.frx":3D10
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
         _Band(0).Cols   =   4
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
         TabIndex        =   3
         Top             =   1440
         Width           =   7575
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
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   7575
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
         Left            =   7440
         TabIndex        =   6
         Top             =   1920
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   22085635
         CurrentDate     =   39092
      End
      Begin VB.TextBox txtNoPerPackage 
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
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtNoPackages 
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
         Top             =   1920
         Width           =   2175
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgStockRetail 
         Height          =   3255
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5741
         _Version        =   393216
         BackColor       =   16117969
         ForeColor       =   -2147483630
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         GridColor       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"frmStockRetail.frx":3D9B
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
         _Band(0).Cols   =   8
      End
      Begin lvButton_H.lvButtons_H cmdAdd 
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   3120
         Width           =   1695
         _ExtentX        =   2990
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
         Image           =   "frmStockRetail.frx":3E45
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdRemove 
         Height          =   375
         Left            =   7800
         TabIndex        =   22
         Top             =   3120
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
         Image           =   "frmStockRetail.frx":5607
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtCartonePrice 
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
         Left            =   7440
         TabIndex        =   7
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtReferenceNo 
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
         Left            =   7440
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtWholeSaleStockLevel 
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
         TabIndex        =   28
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C29E21&
         Caption         =   "Ref Number:"
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
         Left            =   6240
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "Cartone Price:"
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
         Left            =   6000
         TabIndex        =   25
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "From Store:"
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
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C29E21&
         Caption         =   "Number Per Package:"
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
         TabIndex        =   14
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label5 
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
         Left            =   6000
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Stocking Date:"
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
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmStockRetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String, AdjNoOfPackages As Integer, AdjNoPerPackage As Integer, AdjTotal As Integer
Dim sflag As Boolean, ListProductID As String, ctrl As Control, Storeid As Integer, xflag As Boolean
Dim Fval2 As Integer, Frow2 As Integer, xx As Integer, eflag As Boolean, yflag As Boolean, X As Integer
Dim ShowStoreGridFlag As Boolean, WholeSaleStock As Integer
Private Sub cmdAdd_Click()
Dim X As Integer, a As Integer
On Error GoTo OkError
'If eflag = False Then
'For xx = 1 To flxgStockRetail.Rows - 1
'           If flxgStockRetail.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgStockRetail.TextMatrix(xx, 5) = Productid Then
'              MsgBox flxgStockRetail.TextMatrix(xx, 0) & " HAS ALREADY BEEN ADDED!: YOU CAN ONLY EDIT WHAT IS ENTERED", vbInformation, ""
'              Me.txtProductName.SetFocus
'              Exit Sub
'           End If
'         Next
'End If

If Me.flxgProductName.Visible = True Then
   MsgBox "Please Select Product.", vbInformation, "Product Description"
   Me.txtProductName.SetFocus: Exit Sub
End If

Me.txtTotalQuantity = Val(Trim(Me.txtNoPackages)) * Val(Trim(Me.txtNoPerPackage))
'strg2 = Val(Me.txtUnitPrice) * Val(Me.cboQuantity) * (100 - Val(Me.txtDiscount)) / 100
'Me.txtTotalCost = Format$(strg2, "¢#,###.00")
If Trim(Me.txtProductName) = "" Then MsgBox "Please Select Product", vbInformation, "": Me.txtProductName.SetFocus: Exit Sub
If Trim(Me.txtNoPackages) = "" Or Val(Trim(Me.txtNoPackages)) = 0 Then MsgBox "Please Enter Number of Cartones", vbInformation, "": txtNoPackages = "": Me.txtNoPackages.SetFocus: Exit Sub
If Trim(Me.txtNoPerPackage) = "" Or Val(Trim(Me.txtNoPerPackage)) = 0 Then MsgBox "Please Enter Number Per Cartone", vbInformation, "": txtNoPerPackage = "": Me.txtNoPerPackage.SetFocus: Exit Sub
    If Trim(Me.txtTotalQuantity) = "" Or Val(Trim(Me.txtTotalQuantity)) = 0 Then MsgBox "Please Enter Quantity of Product", vbInformation, "": txtTotalQuantity = "": Me.txtTotalQuantity.SetFocus: Exit Sub
    If Trim(Me.txtCartonePrice) = "" Or Val(Trim(Me.txtCartonePrice)) = 0 Then MsgBox "Please Enter Cartone Price", vbInformation: Me.txtCartonePrice.SetFocus: Exit Sub
    If Frow2 <= 0 Then
      If Fval2 > 0 Then
         For xx = 1 To flxgStockRetail.Rows - 1
           If flxgStockRetail.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgStockRetail.TextMatrix(xx, 4) = Productid Then
              MsgBox Me.txtProductName & " Has Already Been Added!: You Can Only Edit What is Entered", vbInformation, ""
              Me.txtProductName = ""
              Exit Sub
           End If
         Next
       End If
      Fval2 = Fval2 + 1
      flxgStockRetail.Rows = Fval2 + 1
      flxgStockRetail.TextMatrix(Fval2, 0) = Me.txtProductName
      flxgStockRetail.TextMatrix(Fval2, 1) = Me.txtNoPackages
      flxgStockRetail.TextMatrix(Fval2, 2) = Me.txtNoPerPackage
      flxgStockRetail.TextMatrix(Fval2, 3) = Format$(Val(Me.txtCartonePrice), "#,###.00")
      flxgStockRetail.TextMatrix(Fval2, 4) = Format(Val(Trim(Me.txtCartonePrice)) * Val(Trim(Me.txtNoPackages)), "#,###.00")
      flxgStockRetail.TextMatrix(Fval2, 5) = Me.txtTotalQuantity
      flxgStockRetail.TextMatrix(Fval2, 6) = Me.dtpDate
      flxgStockRetail.TextMatrix(Fval2, 7) = Productid
'      flxgStockRetail.TextMatrix(Fval2, 6) = Me.txtVat Format$(Me.txtCartonePrice, "#,###.00")
      'flxgStockRetail.TextMatrix(Fval2, 6) = Me.txtPriceAfterDiscount
       'flxgStockRetail.TextMatrix(Fval2, 7) = Trim(Productid)
       
    Else
       flxgStockRetail.TextMatrix(Frow2, 0) = Me.txtProductName
       flxgStockRetail.TextMatrix(Frow2, 1) = Me.txtNoPackages
       flxgStockRetail.TextMatrix(Frow2, 2) = Me.txtNoPerPackage
       flxgStockRetail.TextMatrix(Frow2, 3) = Format$(Val(Me.txtCartonePrice), "#,###.00")
       flxgStockRetail.TextMatrix(Frow2, 4) = Format(Val(Trim(Me.txtCartonePrice)) * Val(Trim(Me.txtNoPackages)), "#,###.00")
       flxgStockRetail.TextMatrix(Frow2, 5) = Me.txtTotalQuantity
       flxgStockRetail.TextMatrix(Frow2, 6) = Me.dtpDate
       flxgStockRetail.TextMatrix(Frow2, 7) = Productid
        'flxgStockRetail.TextMatrix(Frow2, 5) = Me.txtVat
        'flxgStockRetail.TextMatrix(Frow2, 6) = Me.txtPriceAfterDiscount
        'flxgStockRetail.TextMatrix(Frow2, 7) = Trim(Productid)
      Frow2 = 0
    
    End If
    Me.txtProductName = ""
    Me.txtNoPackages = ""
    Me.txtNoPerPackage = ""
    Me.txtTotalQuantity = ""
    Me.txtCartonePrice = ""
    Me.txtWholeSaleStockLevel = ""
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
On Error GoTo OkError
      For X = 1 To Me.flxgStockRetail.Rows - 1
            flxgStockRetail.TextMatrix(X, 0) = ""
            flxgStockRetail.TextMatrix(X, 1) = ""
            flxgStockRetail.TextMatrix(X, 2) = ""
            flxgStockRetail.TextMatrix(X, 3) = ""
            flxgStockRetail.TextMatrix(X, 4) = ""
            flxgStockRetail.TextMatrix(X, 5) = ""
            'flxgStockRetail.TextMatrix(X, 6) = ""
            'flxgStockRetail.TextMatrix(X, 7) = ""
       Next
    Fval2 = 0
    Frow2 = 0
    Me.flxgStockRetail.Rows = 2
    
    Call ClearCtrls
    Me.txtProductName.SetFocus
    Me.flxgProductName.Visible = False
       
   Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
     MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
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
'Dim strg1 As String
'xflag = True
'Me.flxgProductName.Visible = False
'On Error GoTo SaveError

'strg1 = InputBox("Enter The ProductName Or The First Few or All of the Characters of The ProductName.", "ProductName")
'If Trim(strg1) <> "" Then

   
   'Me.cmdFind.Enabled = False
          
   'Open Connecttion to Server
   
   'bFlag = OpenConnection(cn, strg)
   
   'If bFlag = False Then
      'If cn.State = 1 Then cn.Close
      'If rs.State = 1 Then rs.Close
      'Me.MousePointer = vbDefault
      'Me.cmdFind.Enabled = True
       'MsgBox strg, vbInformation:
      'Exit Sub
   'End If
   'strg1 = strg1 & "%"
   'rs.Open "Select * From MainStoreProducts Inner Join StockingRetail On MainStoreProducts.ProductID=StockingRetail.ProductID Where MainStoreProducts.ProductName Like '" & strg1 & "' and Date=#" & Date & "#  Order By MainStoreProducts.ProductName", cn, adOpenForwardOnly, adLockReadOnly
   'If rs.RecordCount <= 0 Then
      'rs.Close: cn.Close
      'MsgBox "THERE IS NO PRODUCT TRANSFERED TO RETAIL TODAY WITH THE NAME ENTERED.", vbInformation, "SEARCH FAILED"
      'Me.MousePointer = vbDefault: Me.cmdFind.Enabled = True: Me.cmdFind.SetFocus: Exit Sub
   'Else
      'If rs.RecordCount = 1 Then
         'sflag = True
         'Me.txtProductName = rs.Fields("ProductName")
         'Me.txtNoPackages = rs.Fields("NoOfPackages")
         'Me.txtNoPerPackage = rs.Fields("NoPerPackage")
         'Me.txtTotalQuantity = rs.Fields("TotalQuantity")
         'Me.dtpDate = rs.Fields("Date")
         ''''''''''''Productid = rs.Fields("MainStoreProducts.ProductID")
         'rs.Close
        
         'Me.cmdDelete.Enabled = True
         'sflag = True
      'Else
         'rs.MoveFirst
         'Me.flxgfind.Rows = rs.RecordCount + 1
         'For X = 1 To rs.RecordCount
           'Me.flxgfind.TextMatrix(X, 0) = rs.Fields("ProductName")
           'Me.flxgfind.TextMatrix(X, 1) = rs.Fields("NoOfPackages")
           'Me.flxgfind.TextMatrix(X, 2) = rs.Fields("NoPerPackage")
           'Me.flxgfind.TextMatrix(X, 3) = rs.Fields("TotalQuantity")
           'Me.flxgfind.TextMatrix(X, 4) = rs.Fields("Date")
           'Me.flxgfind.TextMatrix(X, 5) = rs.Fields("MainStoreProducts.ProductID")
           
          ' rs.MoveNext
        ' Next
        ' Me.flxgfind.Visible = True
        ' Me.flxgfind.SetFocus
        ' rs.Close
    ''  End If
  ' End If
   

'End If

'If cn.State = 1 Then cn.Close
'If rs.State = 1 Then rs.Close

'Me.MousePointer = vbDefault
'Me.cmdFind.Enabled = True
'Me.cmdSave.Enabled = True
'Exit Sub
'SaveError:
     'If cn.State = 1 Then cn.Close
   '  If rs.State = 1 Then rs.Close
     'Me.MousePointer = vbDefault
    ' Me.cmdFind.Enabled = True
    ' MsgBox "Sorry, Unable to Find Products Details:Please Try Again!", vbInformation, "Search Failed"
     'Exit Sub
End Sub

Private Sub cmdOrder_Click()
frmReOderProducts.Show
End Sub

Private Sub cmdRemove_Click()
Dim a As Integer, strgbal As String

On Error GoTo OkError

 If Frow2 = 0 Then Exit Sub
 
 
 If Frow2 = 1 Then
   If flxgStockRetail.TextMatrix(Frow2, 0) = "" Then
   Frow2 = 0
   Me.txtProductName.Text = ""
   Me.flxgProductName.Visible = False
    Me.txtNoPackages = ""
    Me.txtNoPerPackage = ""
    Me.txtProductName.SetFocus
   Exit Sub
   End If
 End If

   If Frow2 = flxgStockRetail.Rows - 1 Then
     If flxgStockRetail.Rows <> 2 Then
         flxgStockRetail.Rows = flxgStockRetail.Rows - 1
     Else
        For xx = 0 To 7
           flxgStockRetail.TextMatrix(Frow2, xx) = ""
        Next
        Me.txtProductName.Text = ""
        Me.flxgProductName.Visible = False
        Me.txtProductName.SetFocus
     End If
        Me.txtProductName.Text = ""
        Me.flxgProductName.Visible = False
        Me.txtProductName.SetFocus
   Else
         For xx = Frow2 To flxgStockRetail.Rows - 2
            flxgStockRetail.TextMatrix(xx, 0) = flxgStockRetail.TextMatrix(xx + 1, 0)
            flxgStockRetail.TextMatrix(xx, 1) = flxgStockRetail.TextMatrix(xx + 1, 1)
            flxgStockRetail.TextMatrix(xx, 2) = flxgStockRetail.TextMatrix(xx + 1, 2)
            flxgStockRetail.TextMatrix(xx, 3) = flxgStockRetail.TextMatrix(xx + 1, 3)
            flxgStockRetail.TextMatrix(xx, 4) = flxgStockRetail.TextMatrix(xx + 1, 4)
            flxgStockRetail.TextMatrix(xx, 5) = flxgStockRetail.TextMatrix(xx + 1, 5)
            flxgStockRetail.TextMatrix(xx, 6) = flxgStockRetail.TextMatrix(xx + 1, 6)
            flxgStockRetail.TextMatrix(xx, 7) = flxgStockRetail.TextMatrix(xx + 1, 7)
        Next
        Me.txtProductName.Text = ""
        Me.flxgProductName.Visible = False
        Me.txtProductName.SetFocus
        flxgStockRetail.Rows = flxgStockRetail.Rows - 1
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
frmStockingRetailRpt.Show
End Sub

Private Sub cmdSave_Click()
If flxgStoreName.Visible = True Then
   MsgBox "YOU MUST SELECT STORE.", vbInformation, "STORE NAME"
   Me.flxgStoreName.SetFocus: Exit Sub
End If


If Trim(Me.txtStoreName) = "" Then
   MsgBox "YOU MUST ENTER STORE NAME.", vbInformation, "STORE NAME"
   Me.txtStoreName.SetFocus: Exit Sub
End If


If Me.flxgStockRetail.TextMatrix(1, 0) = "" Then
   MsgBox "YOU MUST ENTER PRODUCT NAME.", vbInformation, "PRODUCT NAME"
   Me.txtProductName.SetFocus: Exit Sub
End If


If Trim(Me.txtReferenceNo) = "" Then
   MsgBox "YOU MUST ENTER REFERENCE NUMBER.", vbInformation, "REFERENCE NUMBER"
   Me.txtReferenceNo.SetFocus: Exit Sub
End If

'On Error GoTo SaveError
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


   'save part
   cn.BeginTrans
   For X = 1 To Me.flxgStockRetail.Rows - 1
   cn.Execute "Insert Into StockingRetail ([ProductID],[NoOfPackages],[NoPerPackage],[TotalQuantity],[Date],StoreID,CartonePrice,RefNumber) select '" & Trim(Me.flxgStockRetail.TextMatrix(X, 7)) & "','" & Val(Trim(Me.flxgStockRetail.TextMatrix(X, 1))) & "','" & Val(Trim(Me.flxgStockRetail.TextMatrix(X, 2))) & "','" & Val(Trim(Me.flxgStockRetail.TextMatrix(X, 5))) & "','" & Trim(Me.dtpDate) & "','" & Storeid & "','" & Val(Trim(Me.flxgStockRetail.TextMatrix(X, 3))) & "','" & Trim(Me.txtReferenceNo) & "'", Y
   
   If Y > 0 Then
   
   rs.Open "Select * From WholeSaleInventory Where ProductID='" & Trim(Me.flxgStockRetail.TextMatrix(X, 7)) & "' and StoreID='" & Storeid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
     If rs.Fields("StockLevel") >= Val(Trim(Me.flxgStockRetail.TextMatrix(X, 1))) Then
      AdjNoOfPackages = rs.Fields("StockLevel") - Val(Trim(Me.flxgStockRetail.TextMatrix(X, 1)))
     End If
     
     cn.Execute "Update WholeSaleInventory Set StockLevel ='" & AdjNoOfPackages & "' Where ProductID ='" & Trim(Me.flxgStockRetail.TextMatrix(X, 7)) & "' and StoreID='" & Storeid & "'", Y
     If rs.State = 1 Then rs.Close
   End If
     If rs.State = 1 Then rs.Close
   End If
   
   If Y > 0 Then
      rs.Open "Select * From ProductInventory Where ProductID='" & Trim(Me.flxgStockRetail.TextMatrix(X, 7)) & "'", cn, adOpenForwardOnly, adLockReadOnly
     If rs.RecordCount > 0 Then
       AdjStockLevel = rs.Fields("StockLevel") + Val(Trim(Me.flxgStockRetail.TextMatrix(X, 5)))
'     Else
'       AdjStockLevel = Val(Trim(Me.flxgStockRetail.TextMatrix(X, 3)))
     End If
     cn.Execute "Update ProductInventory Set StockLevel ='" & AdjStockLevel & "' Where ProductID ='" & Trim(Me.flxgStockRetail.TextMatrix(X, 7)) & "'", Y
     If rs.State = 1 Then rs.Close
   End If
   
   Next
   
   If Y > 0 Then
        cn.CommitTrans
        MsgBox "Retail Updated Successfully", vbInformation, "SAVED"
        yflag = True
        Fval2 = 0
   Else
        cn.RollbackTrans
        MsgBox "Retail Update failed", vbInformation, "Please Try Again"
        yflag = True
   End If

    If yflag = True Then
    
      For X = 1 To Me.flxgStockRetail.Rows - 1
            flxgStockRetail.TextMatrix(X, 0) = ""
            flxgStockRetail.TextMatrix(X, 1) = ""
            flxgStockRetail.TextMatrix(X, 2) = ""
            flxgStockRetail.TextMatrix(X, 3) = ""
            flxgStockRetail.TextMatrix(X, 4) = ""
            flxgStockRetail.TextMatrix(X, 5) = ""
            flxgStockRetail.TextMatrix(X, 6) = ""
            flxgStockRetail.TextMatrix(X, 7) = ""
       Next
    Fval2 = 0
    Me.flxgStockRetail.Rows = 2
    yflag = False
   End If
   Me.flxgProductName.Visible = False
   Call ClearCtrls
   Me.flxgProductName.Visible = False

Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   Me.txtReferenceNo.SetFocus
End If
End Sub

Private Sub flxgfind_Click()
'Me.txtProductName = Me.flxgfind.TextMatrix(Me.flxgfind.Row, 0)
'Me.txtNoPackages = Me.flxgfind.TextMatrix(Me.flxgfind.Row, 1)
'Me.txtNoPerPackage = Me.flxgfind.TextMatrix(Me.flxgfind.Row, 2)
'Me.txtTotalQuantity = Me.flxgfind.TextMatrix(Me.flxgfind.Row, 3)
'Me.dtpDate = Me.flxgfind.TextMatrix(Me.flxgfind.Row, 4)
'Productid = Me.flxgfind.TextMatrix(Me.flxgfind.Row, 5)
'
'Me.flxgfind.Visible = False
'
'Me.cmdSave.Enabled = True
'sflag = True
'oflag = False

End Sub

Private Sub flxgProductName_Click()
WholeSaleStock = 0
Me.txtProductName = Me.flxgProductName.TextMatrix(Me.flxgProductName.Row, 0)
Productid = Me.flxgProductName.TextMatrix(Me.flxgProductName.Row, 1)
Me.txtNoPerPackage = Me.flxgProductName.TextMatrix(Me.flxgProductName.Row, 2)
Me.txtCartonePrice = Me.flxgProductName.TextMatrix(Me.flxgProductName.Row, 3)
Me.flxgProductName.Visible = False
Call FindWholeSaleStock(Productid)
Me.txtWholeSaleStockLevel = WholeSaleStock & " " & "Cartone(s)"
Me.cmdSave.Enabled = True
sflag = False
oflag = False
Me.txtNoPackages.SetFocus

End Sub

Private Sub flxgProductName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
WholeSaleStock = 0
Me.txtProductName = Me.flxgProductName.TextMatrix(Me.flxgProductName.Row, 0)
Productid = Me.flxgProductName.TextMatrix(Me.flxgProductName.Row, 1)
Me.txtNoPerPackage = Me.flxgProductName.TextMatrix(Me.flxgProductName.Row, 2)
Me.txtCartonePrice = Me.flxgProductName.TextMatrix(Me.flxgProductName.Row, 3)
Me.flxgProductName.Visible = False
Call FindWholeSaleStock(Productid)
Me.txtWholeSaleStockLevel = WholeSaleStock & " " & "Cartone(s)"
Me.cmdSave.Enabled = True
sflag = False
    oflag = False
Me.txtNoPackages.SetFocus

End If
End Sub

Private Sub flxgStockRetail_Click()
xflag = True
        Me.txtProductName = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 0)
        Me.txtNoPackages = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 1)
        Me.txtNoPerPackage = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 2)
        Me.txtCartonePrice = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 3)
        Me.txtTotalQuantity = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 5)
        If Me.flxgStockRetail.Rows > 2 Then
        Me.dtpDate = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 6)
        End If
        Productid = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 7)
        'Me.txtVat = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 5)
        'Me.txtPriceAfterDiscount = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 6)
         'Trim(Productid) = flxgStockRetail.TextMatrix(flxgStockRetail.Row, 7)
Frow2 = flxgStockRetail.Row
eflag = True
'oflag = False
Me.cmdSave.Enabled = True
Me.txtNoPackages.SetFocus
End Sub

Private Sub flxgStoreName_Click()
If Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 0) = "RETAIL" Then
MsgBox "You cannot stock retail from retail ", vbInformation, "Choose store other than retail"
Me.flxgStoreName.Visible = False
Me.txtStoreName = ""
Me.txtStoreName.SetFocus
Exit Sub
End If

Me.txtStoreName = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 0)

Storeid = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 1)
'Call FindStockLevel
'Call FindRetailStockLevel
Me.flxgStoreName.Visible = False
Me.dtpDate.SetFocus

End Sub

Private Sub flxgStoreName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 0) = "RETAIL" Then
MsgBox "You cannot stock retail from retail ", vbInformation, "Choose store other than retail"
Me.flxgStoreName.Visible = False
Me.txtStoreName = ""
Me.txtStoreName.SetFocus
Exit Sub
End If

Me.txtStoreName = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 0)

Storeid = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 1)
'Call FindStockLevel
'Call FindRetailStockLevel
Me.flxgStoreName.Visible = False
Me.dtpDate.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.dtpDate = Date
'CenterForm Me
Me.Height = 9210
Me.Width = 10680

Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Frow2 = 0
Fval2 = 0

End Sub

Private Sub txtCartonePrice_KeyPress(KeyAscii As Integer)
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

Private Sub txtNoPackages_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.txtNoPerPackage.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtNoPerPackage_KeyPress(KeyAscii As Integer)
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

Private Sub txtNoPerPackage_LostFocus()
Me.txtTotalQuantity = Val(Trim(Me.txtNoPackages)) * Val(Trim(Me.txtNoPerPackage))
End Sub

Private Sub txtProductName_Change()

If Me.txtProductName = "" Then
Me.flxgProductName.Visible = False
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
   rs.Open "Select Distinct  ProductName,ProductID,BaseUnit,PricePerCartone From Products  Where ProductName Like '" & Trim(Me.txtProductName) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   flxgProductName.Height = 950 + (285 * (rs.RecordCount - 1))

   If flxgProductName.Height >= 4455 Then
      flxgProductName.Height = 4455
   End If
    flxgProductName.Rows = rs.RecordCount + 1
   With flxgProductName
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("ProductName")
       ' .TextMatrix(X, 3) = rs.Fields("VAT")
       .TextMatrix(X, 1) = rs.Fields("ProductID")
       .TextMatrix(X, 2) = rs.Fields("BaseUnit")
        .TextMatrix(X, 3) = rs.Fields("PricePerCartone")
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 3
      .RowSel = 1
   End With
   flxgProductName.Visible = True
   If rs.State = 1 Then rs.Close
   Me.cmdSave.Enabled = True
Else
     flxgProductName.Visible = False
     'Dquantity = "1"
      If cFlag = True Then
       If rs.State = 1 Then rs.Close
       cFlag = False: Exit Sub
      End If
     MsgBox "'" & Me.txtProductName & "' is a Not in the System !", vbInformation, "Non-Stock"
'     Me.txtname = "": Me.txtUnitPrice = "": Me.cboQuantity = ""
     'Me.txtTotalCost = "": Me.txtDiscount = "": Me.txtPriceAfterDiscount = "": Me.txtVat = ""
   
     'Frw1 = 1
     'If rs.State <> 0 Then
       'rs.Close
     'End If
     'txtname.SetFocus
     'Exit Sub
End If
Else
      flxgProductName.Visible = False
      
      xflag = False ': Exit Sub
End If
      If rs.State = 1 Then rs.Close

Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "Items In Stock", , "Displaying"
     Exit Sub
End Sub

Private Sub txtProductName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
 Me.flxgProductName.SetFocus
End If
End Sub

Private Sub txtProductName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If KeyAscii = vbKeyReturn Then
   Me.txtNoPackages.SetFocus
End If
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
End Sub

Private Sub txtReferenceNo_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   Me.txtProductName.SetFocus
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
'  Me.cmdSave.Enabled = True
 Else
     flxgStoreName.Visible = False
     If rs.State = 1 Then rs.Close
     MsgBox "'" & Me.txtStoreName & "' is a Not in the System !,", vbInformation, "What are you typing"
     Me.txtStoreName = ""
     Me.txtStoreName.SetFocus
End If

    If rs.State = 1 Then rs.Close
End If

    Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "Stores could not display", , "Displaying"
     Exit Sub
End Sub

Private Sub txtStoreName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
 Me.flxgStoreName.SetFocus
End If
End Sub

Private Sub txtStoreName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.dtpDate.SetFocus
End If
End Sub

Private Sub txtStoreName_LostFocus()
'    If ShowStoreGridFlag = False Then
'        If flxgStoreName.Visible = True Then
'            MsgBox "Please Select Store from the list below", vbInformation, ""
'            Me.txtStoreName.SetFocus
'        End If
'    End If
End Sub

Private Sub txtTotalQuantity_GotFocus()
Me.txtTotalQuantity = Val(Trim(Me.txtNoPackages)) * Val(Trim(Me.txtNoPerPackage))
End Sub

Private Sub txtTotalQuantity_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
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
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
End Sub

Private Sub GetStores()

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
   
   If ShowStoreflag = False Then
   
   rs.Open "Select * From Stores  order by StoreName ", cn, adOpenForwardOnly, adLockReadOnly
   
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
'     flxgStoreName.Visible = False
'     If rs.State = 1 Then rs.Close
'     MsgBox "'" & Me.txtStoreName & "' is a Not in the System !", vbInformation, "What are you typing"
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

Private Sub GetProductList()

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

   rs.Open "Select Distinct  ProductName,ProductID,BaseUnit From Products  Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
   flxgProductName.Height = 950 + (285 * (rs.RecordCount - 1))

   If flxgProductName.Height >= 4455 Then
      flxgProductName.Height = 4455
   End If
      flxgProductName.Rows = rs.RecordCount + 1
   With flxgProductName
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("ProductName")
       ' .TextMatrix(X, 3) = rs.Fields("VAT")
       .TextMatrix(X, 1) = rs.Fields("ProductID")
       .TextMatrix(X, 2) = rs.Fields("BaseUnit")
             
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 1
      .RowSel = 1
   End With
   flxgProductName.Visible = True
   If rs.State = 1 Then rs.Close
   Me.cmdSave.Enabled = True
Else
      flxgProductName.Visible = False
     'Dquantity = "1"
      If rs.State = 1 Then rs.Close
'      MsgBox "'" & Me.txtProductName & "' is a Not in the System !", vbInformation, "Non-Stock"

End If
      If rs.State = 1 Then rs.Close

Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "Items In Stock", , "Displaying"
     Exit Sub

End Sub

Private Sub FindWholeSaleStock(ProductNo As String)
 
 On Error GoTo OkError
 
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If

   rs.Open "Select * From WholeSaleInventory Where ProductID='" & Productid & "' and StoreID='" & Storeid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
     WholeSaleStock = rs.Fields("StockLevel")
     If rs.State = 1 Then rs.Close
   End If
     If rs.State = 1 Then rs.Close
   
   
   Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "Product Save Pending", vbInformation, "Pending"
     Exit Sub
   
End Sub

