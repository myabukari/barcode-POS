VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmStockingRetailRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "StockingRetail From WholeSale Report"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmStockingRetailRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   7770
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgDates 
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   16117969
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   12754465
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      FormatString    =   "<Date                            |<Store                      |<Reference Number                 "
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00C29E21&
      Caption         =   "SELECT STORE PRODUCT TO VIEW REPORT"
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
      Height          =   1695
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   7215
      Begin VB.ComboBox cboProducts 
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
         Left            =   1800
         TabIndex        =   1
         Text            =   "All"
         Top             =   1080
         Width           =   4095
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
         Left            =   1800
         TabIndex        =   0
         Text            =   "All"
         Top             =   480
         Width           =   4095
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   6240
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   480
         TabIndex        =   9
         Top             =   0
         Width           =   6375
         Begin lvButton_H.lvButtons_H cmdOk 
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
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
            Image           =   "frmStockingRetailRpt.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   4080
            TabIndex        =   11
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            Caption         =   "&Exit"
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
            Image           =   "frmStockingRetailRpt.frx":118C
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Caption         =   "SPECIFY DATE RANGE"
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
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   7215
      Begin Crystal.CrystalReport CrystalStockingRetail 
         Left            =   3120
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   16515075
         CurrentDate     =   39091
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   16515075
         CurrentDate     =   39091
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "To:"
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
         Left            =   3960
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "From:"
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
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmStockingRetailRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer
Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

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
       MsgBox "Sorry cannot display Product Names,Please try again", vbInformation, "Product Names"
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
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError


If Me.cboStore = "" Then
MsgBox "Please specify store's name", vbInformation, "store's name"
Me.cboStore.SetFocus
Exit Sub
End If


If Me.cboProducts = "" Then
MsgBox "Please specify Product's name", vbInformation, "Product's name"
Me.cboProducts.SetFocus
Exit Sub
End If


If Me.dtpfrom > Me.dtpto Then
MsgBox "WRONG DATE RANGE ENTERED", vbInformation, "DATE RANGE"
Exit Sub
End If


If Me.dtpfrom <= Me.dtpto Then
bFlag = OpenConnection(cn, strg)
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
 
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select  * From StockingRetail where Date >='" & Me.dtpfrom & "' and Date <='" & Me.dtpto & "'  order by Date desc", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
           
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED DATE", vbInformation, "DATE RANGE"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close

If Me.cboProducts <> "All" Then
rs.Open "Select  * From StockingRetail inner Join Products on StockingRetail.ProductID=Products.ProductID  where ProductName='" & Me.cboProducts & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED PRODUCT", vbInformation, "PRODUCT NOT IN REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close


If Me.cboStore <> "All" Then
rs.Open "Select  * From Stores inner Join StockingRetail on Stores.StoreID=StockingRetail.StoreID  where StoreName='" & Me.cboStore & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED STORE", vbInformation, "STORE NOT IN REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close


If Me.cboProducts <> "All" And Me.cboStore <> "All" Then
rs.Open "Select  * From (Products inner Join StockingRetail on Products.ProductID=StockingRetail.ProductID) inner Join Stores on Stores.StoreID=StockingRetail.StoreID  where ProductName='" & Me.cboProducts & "' and StoreName='" & Me.cboStore & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED FIELDS", vbInformation, "WHOLASALE STOCKLEVEL REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close





'CrystalStockingRetail.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalStockingRetail.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalStockingRetail.ReportFileName = App.Path & "\StockingRetail.rpt"
CrystalStockingRetail.Connect = "DSN=nxomen;UID=sa;PWD=abu;DSQ=ZuksData"


'If Me.cboReport = "Today" And Me.cboProducts = "All" Then 'Me.dtpfrom = Me.dtpto Then
'CrystalStockingRetail.SelectionFormula = "{Payments/Products.Date} =#" & Date & "#"
'End If

'If Me.cboReport = "Today" And Me.cboProducts <> "All" Then 'Me.dtpfrom = Me.dtpto Then
'CrystalStockingRetail.SelectionFormula = "{Payments/Products.Date} =#" & Date & "#and {Products.ProductName}='" & Trim(Me.cboProducts) & "'"
'End If


If Me.cboStore = "All" And Me.cboProducts <> "All" Then
CrystalStockingRetail.SelectionFormula = "{StockingRetail.Date} >=#" & Me.dtpfrom & "# and {StockingRetail.Date} <=#" & Me.dtpto & "# and {Products.ProductName} ='" & Me.cboProducts.Text & "'"
End If

If Me.cboStore <> "All" And Me.cboProducts = "All" Then
CrystalStockingRetail.SelectionFormula = "{StockingRetail.Date} >=#" & Me.dtpfrom & "# and {StockingRetail.Date} <=#" & Me.dtpto & "# and {Stores.StoreName} ='" & Me.cboStore & "'"
End If

If Me.cboStore <> "All" And Me.cboProducts <> "All" Then
CrystalStockingRetail.SelectionFormula = "{StockingRetail.Date} >=#" & Me.dtpfrom & "# and {StockingRetail.Date} <=#" & Me.dtpto & "# and {Stores.StoreName} ='" & Me.cboStore & "' and {Products.ProductName} ='" & Me.cboProducts.Text & "'"
End If

If Me.cboProducts = "All" And Me.cboStore = "All" Then
CrystalStockingRetail.SelectionFormula = "{StockingRetail.Date} >=#" & Me.dtpfrom & "# and {StockingRetail.Date} <=#" & Me.dtpto & "#"
End If





   CrystalStockingRetail.WindowState = crptMaximized
   CrystalStockingRetail.WindowShowRefreshBtn = True
   Me.CrystalStockingRetail.WindowTitle = "STOCKING RETAIL REPORT " & Format$(Date, "yyyy")
   
   CrystalStockingRetail.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "STOCKING RETAIL REPORT"
       Exit Sub
End Sub

Private Sub flxgDates_Click()

On Error GoTo OkError

'CrystalStockingRetail.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalStockingRetail.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalStockingRetail.ReportFileName = App.Path & "\StockingRetail.rpt"
CrystalStockingRetail.Connect = "DSN=nxomen;UID=sa;PWD=abu;DSQ=ZuksData"


CrystalStockingRetail.SelectionFormula = "{StockingRetail.Date} =#" & Me.flxgDates.TextMatrix(flxgDates.Row, 0) & "# and {Stores.StoreName} ='" & Me.flxgDates.TextMatrix(flxgDates.Row, 1) & "' and {StockingRetail.RefNumber} ='" & Me.flxgDates.TextMatrix(flxgDates.Row, 2) & "'"


   CrystalStockingRetail.WindowState = crptMaximized
   CrystalStockingRetail.WindowShowRefreshBtn = True
   Me.CrystalStockingRetail.WindowTitle = "STOCKING RETAIL REPORT " & Format$(Date, "yyyy")
   CrystalStockingRetail.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "STOCKING RETAIL REPORT"
       Exit Sub
End Sub

Private Sub flxgDates_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
On Error GoTo OkError

'CrystalStockingRetail.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalStockingRetail.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalStockingRetail.ReportFileName = App.Path & "\StockingRetail.rpt"
CrystalStockingRetail.Connect = "DSN=nxomen;UID=sa;PWD=abu;DSQ=ZuksData"


CrystalStockingRetail.SelectionFormula = "{StockingRetail.Date} =#" & Me.flxgDates.TextMatrix(flxgDates.Row, 0) & "# and {Stores.StoreName} ='" & Me.flxgDates.TextMatrix(flxgDates.Row, 1) & "'"


   CrystalStockingRetail.WindowState = crptMaximized
   CrystalStockingRetail.WindowShowRefreshBtn = True
   Me.CrystalStockingRetail.WindowTitle = "STOCKING RETAIL REPORT " & Format$(Date, "yyyy")
   CrystalStockingRetail.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "STOCKING RETAIL REPORT"
       Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.dtpfrom = Date
Me.dtpto = Date
Call GetDates
'CenterForm Me
Me.Height = 7785
Me.Width = 7890
Me.cboProducts.Text = "All"
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub GetDates()

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

 rs.Open "Select Distinct Date,RefNumber,StoreName from Stores inner Join StockingRetail on Stores.StoreID=StockingRetail.StoreID  Order By Date", cn, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
   rs.MoveFirst
 Me.flxgDates.Rows = rs.RecordCount + 1
  
  For X = 1 To rs.RecordCount
   Me.flxgDates.TextMatrix(X, 0) = rs.Fields("Date")
   Me.flxgDates.TextMatrix(X, 1) = rs.Fields("StoreName")
   If rs.Fields("RefNumber") <> "" Then
   Me.flxgDates.TextMatrix(X, 2) = rs.Fields("RefNumber")
   End If
'   Me.flxgDates.TextMatrix(X, 3) = rs.Fields("Mode")
'   Me.flxgExpenditure.TextMatrix(X, 4) = rs.Fields("ChequeNo")
'   Me.flxgExpenditure.TextMatrix(X, 5) = rs.Fields("ExpenditureID")
'   Me.flxgExpenditure.TextMatrix(X, 5) = rs.Fields("AurthorsName")
'   Me.flxgExpenditure.TextMatrix(X, 6) = rs.Fields("BookID")
'   Me.flxgExpenditure.TextMatrix(X, 7) = rs.Fields("BookBorrowID")
   rs.MoveNext
  Next
    If rs.State = 1 Then rs.Close
  

Else
  For X = 0 To 2
    Me.flxgDates.TextMatrix(1, X) = ""
  Next
    Me.flxgDates.Rows = 2
End If
If rs.State = 1 Then rs.Close

Exit Sub
OkError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Find Dates:Please Try Again!", vbInformation, "Please Try Again!"
     Exit Sub

End Sub
