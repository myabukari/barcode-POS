VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmWholeSaleStockRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "WholeSale StockLevel "
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   Icon            =   "frmWholeSaleStockLevelRpt.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   8160
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   7575
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   720
         TabIndex        =   4
         Top             =   0
         Width           =   5895
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   3480
            TabIndex        =   5
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
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
            Image           =   "frmWholeSaleStockLevelRpt.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdOk 
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            Caption         =   "&Ok"
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
            Image           =   "frmWholeSaleStockLevelRpt.frx":075C
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Caption         =   "SPECIFY STORE AND PRODUCT TO VIEW REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      Begin Crystal.CrystalReport CrystalWholeSaleStock 
         Left            =   6840
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cboProducts 
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
         TabIndex        =   7
         Top             =   1200
         Width           =   4575
      End
      Begin VB.ComboBox cboStore 
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
         TabIndex        =   1
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "Products's Name:"
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
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Store's Name:"
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
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmWholeSaleStockRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer
Private Sub cboProducts_DropDown()
On Error GoTo OkError


'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

Me.cboProducts.Clear

 rs.Open "Select Distinct ProductName from Products  Order By ProductName Asc", cn, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
   rs.MoveFirst
   For i = 1 To rs.RecordCount
     Me.cboProducts.AddItem rs.Fields!ProductName
     
     rs.MoveNext
   Next
   If cboProducts.ListCount > 1 Then
      Me.cboProducts.AddItem "All"
   End If
End If
If rs.State = 1 Then
   rs.Close
End If
Exit Sub
OkError:
       If rs.State = 1 Then
          rs.Close
       End If
       MsgBox "Try Again", vbInformation, "Could not list Products"
       Exit Sub
End Sub

Private Sub cboStore_DropDown()
On Error GoTo OkError


'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

Me.cboStore.Clear

 rs.Open "Select Distinct StoreName from Stores  Order By StoreName Asc", cn, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
   rs.MoveFirst
   For i = 1 To rs.RecordCount
     Me.cboStore.AddItem rs.Fields!StoreName
     
     rs.MoveNext
   Next
   If cboStore.ListCount > 1 Then
      Me.cboStore.AddItem "All"
   End If
End If
If rs.State = 1 Then
   rs.Close
End If
Exit Sub
OkError:
       If rs.State = 1 Then
          rs.Close
       End If
       MsgBox "Sorry cannot display Store Names,Please try again ", vbInformation, "Store Names"
       Exit Sub

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError

If Me.cboStore = "" Then
    MsgBox "Please Select a Store or All to view report ", vbInformation, "Store's Name"
    Me.cboStore.SetFocus: Exit Sub
End If

If Me.cboProducts = "" Then
    MsgBox "Please Select a Product or All to view report ", vbInformation, "Product's Name"
    Me.cboProducts.SetFocus: Exit Sub
End If

'CrystalWholeSaleStock.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalWholeSaleStock.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalWholeSaleStock.ReportFileName = App.Path & "\rptWholeSaleStockLevel.rpt"
CrystalWholeSaleStock.Connect = "DSN=nxomen;UID=sa;PWD=abu;DSQ=ZuksData"



bFlag = OpenConnection(cn, strg)
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
 
   MsgBox strg, vbInformation:
   Exit Sub
End If

If Me.cboStore <> "All" Then
rs.Open "Select  * From Stores inner Join WholeSaleInventory on Stores.StoreID=WholeSaleInventory.StoreID  where StoreName='" & Me.cboStore & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED STORE", vbInformation, "STORE NOT IN REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close



If Me.cboProducts <> "All" Then
rs.Open "Select  * From Products inner Join WholeSaleInventory on Products.ProductID=WholeSaleInventory.ProductID  where ProductName='" & Me.cboProducts & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED PRODUCT", vbInformation, "PRODUCT NOT IN REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close

If Me.cboProducts <> "All" And Me.cboStore <> "All" Then
rs.Open "Select  * From (Products inner Join WholeSaleInventory on Products.ProductID=WholeSaleInventory.ProductID) inner Join Stores on Stores.StoreID=WholeSaleInventory.StoreID  where ProductName='" & Me.cboProducts & "' and StoreName='" & Me.cboStore & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED FIELDS", vbInformation, "WHOLASALE STOCKLEVEL REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close


If Me.cboStore = "All" And Me.cboProducts = "All" Then
CrystalWholeSaleStock.SelectionFormula = ""
End If

If Me.cboStore = "All" And Me.cboProducts <> "All" Then
CrystalWholeSaleStock.SelectionFormula = "{Products.ProductName}='" & Trim(Me.cboProducts) & "'"
End If

If Me.cboStore <> "All" And Me.cboProducts = "All" Then
CrystalWholeSaleStock.SelectionFormula = "{Stores.StoreName}='" & Trim(Me.cboStore) & "'"
End If

If Me.cboStore <> "All" And Me.cboProducts <> "All" Then
CrystalWholeSaleStock.SelectionFormula = "{Stores.StoreName}='" & Trim(Me.cboStore) & "' and {Products.ProductName}='" & Trim(Me.cboProducts) & "'"
End If


   CrystalWholeSaleStock.WindowState = crptMaximized
   CrystalWholeSaleStock.WindowShowRefreshBtn = True
   Me.CrystalWholeSaleStock.WindowTitle = "WHOLESALE STOCKLEVEL " & Format$(Date, "yyyy")
   
   CrystalWholeSaleStock.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "DEBTOR'S PAYMENTS"
       Exit Sub


End Sub

Private Sub Form_Load()
Me.cboStore = "All"
Me.cboProducts = "All"

Me.Height = 4785
Me.Width = 8280
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub
