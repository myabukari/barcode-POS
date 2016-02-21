VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmAdjustWholeSaleStock 
   BackColor       =   &H00C29E21&
   Caption         =   "Adjust WholeSale Stock"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   Icon            =   "frmWholeSaleStockLevel.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   9660
   Begin VB.Frame Frame5 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   6720
      Width           =   9375
      Begin lvButton_H.lvButtons_H cmdSave 
         Height          =   375
         Left            =   240
         TabIndex        =   5
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
         Image           =   "frmWholeSaleStockLevel.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdClear 
         Height          =   375
         Left            =   3720
         TabIndex        =   6
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
         Image           =   "frmWholeSaleStockLevel.frx":075C
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   6960
         TabIndex        =   7
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
         Image           =   "frmWholeSaleStockLevel.frx":200D
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   6375
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   9375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgStoreName 
         Height          =   1095
         Left            =   1560
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProductSearch 
         Height          =   1695
         Left            =   1560
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   16117969
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   "<Product Description                                             |>StockLevel              |<ProductID"
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
      Begin VB.ComboBox cboStockLevel 
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
         ItemData        =   "frmWholeSaleStockLevel.frx":245F
         Left            =   1560
         List            =   "frmWholeSaleStockLevel.frx":24DB
         TabIndex        =   2
         ToolTipText     =   "Whatever value you input here will be added to current stocklevel."
         Top             =   1440
         Width           =   3255
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
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtStoreName 
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
         Top             =   360
         Width           =   3255
      End
      Begin lvButton_H.lvButtons_H cmdAdd 
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1920
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
         Image           =   "frmWholeSaleStockLevel.frx":2576
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdRemove 
         Height          =   375
         Left            =   7560
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1920
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
         Image           =   "frmWholeSaleStockLevel.frx":3D38
         cBack           =   -2147483633
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProducts 
         Height          =   3615
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6376
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
         FormatString    =   $"frmWholeSaleStockLevel.frx":3ED2
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
      Begin VB.TextBox txtCurrentStockLevel 
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
         Left            =   7080
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Current StockLevel:"
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
         Left            =   5280
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Store Name:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Adjust Stock to (Cartones):"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Top             =   960
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAdjustWholeSaleStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Storeid As Integer
Dim sflag As Boolean, ListProductID As String, ctrl As Control, Destinationid As Integer, Sourceid As Integer
Dim Productid As String, Fval2 As Integer, Frow2 As Integer, xx As Integer, ShowProductsflag As Boolean, ShowStoreflag As Boolean, ShowDestinationflag As Boolean
Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.cboStockLevel.SetFocus
End If

End Sub

Private Sub cboStockLevel_KeyPress(KeyAscii As Integer)
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
'If Trim(Me.txtSource) = "" Then MsgBox "From which Store are you issuing goods", vbInformation, "": Me.txtSource.SetFocus: Exit Sub

If Trim(Me.txtProductName) = "" Then MsgBox "Please Select Product", vbInformation, "ProductName": Me.txtProductName.SetFocus: Exit Sub
If Trim(Me.txtStoreName) = "" Then MsgBox "Please select Store's Name", vbInformation, "": Me.txtStoreName.SetFocus: Exit Sub

    If Trim(Me.cboStockLevel) = "" Then MsgBox "Please Enter Quantity of Product to bo adjusted", vbInformation, "Quantity": Me.cboStockLevel.SetFocus: Exit Sub
    
    
'    If MsgBox("ARE YOU SURE  ABOUT THE ISSUING DATE ENTERED?", vbYesNo + vbQuestion, "CONFIRM ISSUING DATE") = vbNo Then
'    Me.dtpIsueingDate.SetFocus
'    Exit Sub
'    End If
    
    
    'If Trim(Me.txtTotalCost) = "" Then MsgBox "Please Compute TotalCost", vbInformation: Me.txtTotalCost.SetFocus: Exit Sub
    If Frow2 <= 0 Then
      If Fval2 > 0 Then
         For xx = 1 To flxgProducts.Rows - 1
           If flxgProducts.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgProducts.TextMatrix(xx, 4) = Productid And flxgProducts.TextMatrix(xx, 3) = Storeid Then
              MsgBox Me.txtProductName & " Has Already Been Added!: You Can Only Edit What is Entered", vbInformation, ""
              ShowProductsflag = True
              Me.txtProductName = ""
              Me.cboStockLevel = ""
              ShowProductsflag = False
              Me.txtProductName.SetFocus
              Exit Sub
           End If
         Next
       End If
      Fval2 = Fval2 + 1
      flxgProducts.Rows = Fval2 + 1
      flxgProducts.TextMatrix(Fval2, 0) = Me.txtProductName
      flxgProducts.TextMatrix(Fval2, 1) = Me.cboStockLevel
      flxgProducts.TextMatrix(Fval2, 2) = Me.txtStoreName
      flxgProducts.TextMatrix(Fval2, 3) = Storeid
      flxgProducts.TextMatrix(Fval2, 4) = Productid
'      flxgProducts.TextMatrix(Fval2, 5) = Sourceid
'      flxgProducts.TextMatrix(Fval2, 6) = Productid
'      flxgProducts.TextMatrix(Fval2, 7) = Destinationid
       'flxgProducts.TextMatrix(Fval2, 7) = Trim(Productid)
       
    Else
       flxgProducts.TextMatrix(Frow2, 0) = Me.txtProductName
       flxgProducts.TextMatrix(Frow2, 1) = Me.cboStockLevel
       flxgProducts.TextMatrix(Frow2, 2) = Me.txtStoreName
       flxgProducts.TextMatrix(Frow2, 3) = Storeid
       flxgProducts.TextMatrix(Frow2, 4) = Productid
'        flxgProducts.TextMatrix(Frow2, 5) = Sourceid
'        flxgProducts.TextMatrix(Frow2, 6) = Productid
'       flxgProducts.TextMatrix(Frow2, 7) = Destinationid
        'flxgProducts.TextMatrix(Frow2, 7) = Trim(Productid)
      Frow2 = 0
    
    End If
    Me.cboStockLevel = ""
    
    ShowProductsflag = True
    Me.txtProductName = ""
    ShowProductsflag = False
    Me.txtCurrentStockLevel = ""
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

'On Error GoTo OkError

 If Frow2 = 0 Then Exit Sub
 
 
 If Frow2 = 1 Then
   If flxgProducts.TextMatrix(Frow2, 0) = "" Then
   Frow2 = 0
        ShowProductsflag = True
        Me.txtProductName.Text = ""
        ShowProductsflag = False
        
'        ShowStoreflag = True
'        Me.txtStoreName = ""
'        ShowStoreflag = False
        
        Me.cboStockLevel = ""
        Me.txtCurrentStockLevel = ""
        Productid = ""
        Me.txtProductName.SetFocus
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
            
'            ShowStoreflag = True
'            Me.txtStoreName = ""
'            ShowStoreflag = False
'
            Me.cboStockLevel = ""
            Me.txtCurrentStockLevel = ""
            Productid = ""
            Me.txtProductName.SetFocus
     End If
            ShowProductsflag = True
            Me.txtProductName.Text = ""
            ShowProductsflag = False
            
'            ShowStoreflag = True
'            Me.txtStoreName = ""
'            ShowStoreflag = False

            Me.cboStockLevel = ""
            Me.txtCurrentStockLevel = ""
            Productid = ""
            Me.txtProductName.SetFocus
      Else
         For xx = Frow2 To flxgProducts.Rows - 2
            flxgProducts.TextMatrix(xx, 0) = flxgProducts.TextMatrix(xx + 1, 0)
            flxgProducts.TextMatrix(xx, 1) = flxgProducts.TextMatrix(xx + 1, 1)
            flxgProducts.TextMatrix(xx, 2) = flxgProducts.TextMatrix(xx + 1, 2)
            flxgProducts.TextMatrix(xx, 3) = flxgProducts.TextMatrix(xx + 1, 3)
            flxgProducts.TextMatrix(xx, 4) = flxgProducts.TextMatrix(xx + 1, 4)
'            flxgProducts.TextMatrix(xx, 5) = flxgProducts.TextMatrix(xx + 1, 5)
'            flxgProducts.TextMatrix(xx, 6) = flxgProducts.TextMatrix(xx + 1, 6)
            'flxgProducts.TextMatrix(xx, 7) = flxgProducts.TextMatrix(xx + 1, 7)
        Next
            ShowProductsflag = True
            Me.txtProductName.Text = ""
            ShowProductsflag = False
            
'            ShowStoreflag = True
'            Me.txtStoreName = ""
'            ShowStoreflag = False
            
            Me.cboStockLevel = ""
            Me.txtCurrentStockLevel = ""
            Productid = ""
            Me.txtProductName.SetFocus
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

Private Sub cmdSave_Click()
'On Error GoTo OkError
If Me.txtStoreName = "" Then
    MsgBox "Select Store Name", vbInformation, "Store Name"
    Me.txtStoreName.SetFocus: Exit Sub
End If



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
   
'   cn.BeginTrans
  For X = 1 To Me.flxgProducts.Rows - 1
  If Trim(Me.flxgProducts.TextMatrix(X, 2)) <> "RETAIL" Then
   rs.Open "Select * From WholeSaleInventory  Where StoreID= '" & Trim(Me.flxgProducts.TextMatrix(X, 3)) & "' and ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
         cn.Execute "Update WholeSaleInventory Set StockLevel ='" & Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) & "' Where StoreID ='" & Trim(Me.flxgProducts.TextMatrix(X, 3)) & "' and ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", Y
         If rs.State = 1 Then rs.Close
        Else
         If Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) >= 0 Then
            cn.Execute "Insert Into WholeSaleInventory (ProductID,StoreID,StockLevel) select '" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "','" & Trim(Me.flxgProducts.TextMatrix(X, 3)) & "','" & Trim(Me.flxgProducts.TextMatrix(X, 1)) & "'", Y
            If rs.State = 1 Then rs.Close
         End If
            If rs.State = 1 Then rs.Close
        End If
         If rs.State = 1 Then rs.Close
   ElseIf Trim(Me.flxgProducts.TextMatrix(X, 2)) = "RETAIL" Then
     rs.Open "Select * From ProductInventory  Where ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
         cn.Execute "Update ProductInventory Set StockLevel ='" & rs.Fields("StockLevel") + Val(Trim(Me.flxgProducts.TextMatrix(X, 1))) & "' Where ProductID='" & Trim(Me.flxgProducts.TextMatrix(X, 4)) & "'", Y
         If rs.State = 1 Then rs.Close
        
        End If
         If rs.State = 1 Then rs.Close
   End If
   Next
   
    If Y > 0 Then
      MsgBox "Saved successfully", vbInformation, "Saved successfully"
      Fval2 = 0
      Call ClearCtrls
    Else
     MsgBox "Save Failed,Please Retype Data and Save Again", vbInformation, "Try Again"
    End If
            
      Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "Save Failed,Please Try Again", vbInformation, "Save Failed"
     Exit Sub

End Sub

Private Sub flxgProducts_Click()
        ShowProductsflag = True
        Me.txtProductName = flxgProducts.TextMatrix(flxgProducts.Row, 0)
        ShowProductsflag = False
        Me.cboStockLevel = flxgProducts.TextMatrix(flxgProducts.Row, 1)
        
        ShowStoreflag = True
        Me.txtStoreName = flxgProducts.TextMatrix(flxgProducts.Row, 2)
        ShowStoreflag = False
        
        If Me.flxgProducts.Rows > 2 Then
        Storeid = flxgProducts.TextMatrix(flxgProducts.Row, 3)
        End If
        Productid = flxgProducts.TextMatrix(flxgProducts.Row, 4)
        Frow2 = flxgProducts.Row
        Me.cboStockLevel.SetFocus
End Sub

Private Sub flxgProducts_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        ShowProductsflag = True
        Me.txtProductName = flxgProducts.TextMatrix(flxgProducts.Row, 0)
        ShowProductsflag = False
        Me.cboStockLevel = flxgProducts.TextMatrix(flxgProducts.Row, 1)
        
        ShowStoreflag = True
        Me.txtStoreName = flxgProducts.TextMatrix(flxgProducts.Row, 2)
        ShowStoreflag = False
        
        If Me.flxgProducts.Rows > 2 Then
        Storeid = flxgProducts.TextMatrix(flxgProducts.Row, 3)
        End If
        Productid = flxgProducts.TextMatrix(flxgProducts.Row, 4)
        Frow2 = flxgProducts.Row
        Me.cboStockLevel.SetFocus
  End If
End Sub

Private Sub flxgProductSearch_Click()
Me.txtProductName = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 0)
'Me.txtCartonePrice = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 1)
Productid = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 2)
Me.flxgProductSearch.Visible = False
Call FindStockLevel
Call FindRetailStockLevel
Me.cboStockLevel.SetFocus
End Sub

Private Sub flxgProductSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtProductName = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 0)
'Me.txtCartonePrice = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 1)
Productid = Me.flxgProductSearch.TextMatrix(Me.flxgProductSearch.Row, 2)
Me.flxgProductSearch.Visible = False
Call FindStockLevel
Call FindRetailStockLevel
Me.cboStockLevel.SetFocus

End If
End Sub

Private Sub flxgStoreName_Click()
Me.txtStoreName = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 0)

Storeid = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 1)
Call FindStockLevel
Call FindRetailStockLevel
Me.flxgStoreName.Visible = False
Me.txtProductName.SetFocus
End Sub

Private Sub flxgStoreName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.txtStoreName = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 0)
    Storeid = Me.flxgStoreName.TextMatrix(Me.flxgStoreName.Row, 1)
    Call FindStockLevel
    Call FindRetailStockLevel
    Me.flxgStoreName.Visible = False
    Me.txtProductName.SetFocus
End If

End Sub

Private Sub Form_Load()
Me.Height = 8310
Me.Width = 9780
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub txtProductName_Change()
If Me.txtProductName = "" Then
 Me.flxgProductSearch.Visible = False
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
'       .TextMatrix(X, 1) = rs.Fields("PricePerCartone")
       .TextMatrix(X, 2) = rs.Fields("ProductID")
       
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 2
      .RowSel = 1
   End With
   flxgProductSearch.Visible = True
   If rs.State = 1 Then rs.Close
'   Me.cmdSave.Enabled = True
 Else
     flxgProductSearch.Visible = False
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

Private Sub txtProductName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
  If Me.flxgProductSearch.Visible = True Then
   Me.flxgProductSearch.SetFocus
  Else
   Me.flxgProducts.SetFocus
  End If
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
   Me.cboStockLevel.SetFocus
End If

End Sub

Private Sub txtStoreName_Change()
If Me.txtStoreName = "" Then
Me.flxgStoreName.Visible = False
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
Private Sub FindStockLevel()
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      
       MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   rs.Open "Select * From WholeSaleInventory  Where StoreID= '" & Storeid & "' and ProductID= '" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
    Me.txtCurrentStockLevel = rs.Fields("StockLevel") & " " & "Cartone(s)"
    If rs.State = 1 Then rs.Close
   Else
    Me.txtCurrentStockLevel = 0
   End If
    If rs.State = 1 Then rs.Close
   
End Sub

Private Sub txtStoreName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
  If Me.flxgStoreName.Visible = True Then
   Me.flxgStoreName.SetFocus
  End If
End If
'If KeyCode = 40 And Me.flxgItems.Visible = False Then
' Me.txtUnitPrice.SetFocus
'End If
End Sub

Private Sub txtStoreName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtProductName.SetFocus
End If

End Sub
Private Sub FindRetailStockLevel()
If Me.txtStoreName = "RETAIL" And Me.txtProductName <> "" Then
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
       MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   rs.Open "Select * From ProductInventory where ProductID= '" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
    Me.txtCurrentStockLevel = rs.Fields("StockLevel")
    If rs.State = 1 Then rs.Close
   Else
    Me.txtCurrentStockLevel = 0
   End If
    If rs.State = 1 Then rs.Close
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
Me.txtStoreName.SetFocus
End Sub


