VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmGoodsTransferRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "Goods Transfer Report"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   Icon            =   "frmGoodsTransferRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   7710
   Begin VB.Frame Frame4 
      BackColor       =   &H00C29E21&
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
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   7215
      Begin VB.ComboBox cboProducts 
         BackColor       =   &H00F2DEA2&
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
         Left            =   1920
         TabIndex        =   0
         Text            =   "All"
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label3 
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
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   5880
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   855
         Left            =   1320
         TabIndex        =   10
         Top             =   0
         Width           =   4335
         Begin lvButton_H.lvButtons_H cmdOk 
            Height          =   375
            Left            =   480
            TabIndex        =   2
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Image           =   "frmGoodsTransferRpt.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   2280
            TabIndex        =   3
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Image           =   "frmGoodsTransferRpt.frx":118C
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   7215
      Begin Crystal.CrystalReport CrystalGoodsTransfer 
         Left            =   5400
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   4320
         TabIndex        =   5
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   15718770
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16515075
         CurrentDate     =   39087
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   15718770
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16515075
         CurrentDate     =   39087
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
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
         Left            =   3840
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgDates 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5318
      _Version        =   393216
      BackColor       =   16117969
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   12754465
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      FormatString    =   "<Date                       |<Issueing  Store                      |<Receiving Store                    "
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
End
Attribute VB_Name = "frmGoodsTransferRpt"
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

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError

If Me.cboProducts = "" Then
    MsgBox "Please Select a Product or All to view report ", vbInformation, "Product's Name"
    Me.cboProducts.SetFocus: Exit Sub
End If



'CrystalGoodsTransfer.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalGoodsTransfer.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalGoodsTransfer.ReportFileName = App.Path & "\rptGoodsTransfer.rpt"
CrystalGoodsTransfer.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"


If Me.dtpfrom <= Me.dtpto Then
bFlag = OpenConnection(cn, strg)
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
 
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select  * From GoodsTransfer inner Join Products on GoodsTransfer.ProductID=Products.ProductID  where TransferDate >='" & Me.dtpfrom & "' and TransferDate <='" & Me.dtpto & "'  order by TransferDate desc", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
           
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED DATE", vbInformation, "DATE RANGE"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close

If Me.cboProducts <> "All" Then
rs.Open "Select  * From GoodsTransfer inner Join Products on GoodsTransfer.ProductID=Products.ProductID  where ProductName='" & Me.cboProducts & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED PRODUCT", vbInformation, "PRODUCT NOT IN REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close

If Me.dtpfrom <= Me.dtpto And Me.cboProducts = "All" Then
CrystalGoodsTransfer.SelectionFormula = "{GoodsTransfer.TransferDate} >=#" & Me.dtpfrom & "# and {GoodsTransfer.TransferDate} <=#" & Me.dtpto & "#"
End If

If Me.dtpfrom <= Me.dtpto And Me.cboProducts <> "All" Then
CrystalGoodsTransfer.SelectionFormula = "{GoodsTransfer.TransferDate} >=#" & Me.dtpfrom & "# and {GoodsTransfer.TransferDate} <=#" & Me.dtpto & "# and {Products.ProductName}='" & Trim(Me.cboProducts) & "'"
End If

If Me.dtpfrom > Me.dtpto Then
MsgBox "WRONG DATE RANGE ENTERED", vbInformation, "DATE RANGE"
Exit Sub
End If

   CrystalGoodsTransfer.WindowState = crptMaximized
   CrystalGoodsTransfer.WindowShowRefreshBtn = True
   Me.CrystalGoodsTransfer.WindowTitle = "GOODS TRANSFER " & Format$(Date, "yyyy")
   
   CrystalGoodsTransfer.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "GOODS TRANSFER"
       Exit Sub


End Sub

Private Sub flxgDates_Click()
On Error GoTo OkError

'CrystalStockingRetail.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalStockingRetail.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalGoodsTransfer.ReportFileName = App.Path & "\rptGoodsTransfer.rpt"
CrystalGoodsTransfer.Connect = "DSN=nxomen;UID=sa;PWD=abu;DSQ=ZuksData"


CrystalGoodsTransfer.SelectionFormula = "{GoodsTransfer.TransferDate} =#" & Me.flxgDates.TextMatrix(flxgDates.Row, 0) & "# and {GoodsTransfer.Source} ='" & Me.flxgDates.TextMatrix(flxgDates.Row, 1) & "' and {GoodsTransfer.Destination} ='" & Me.flxgDates.TextMatrix(flxgDates.Row, 2) & "'"


   CrystalGoodsTransfer.WindowState = crptMaximized
   CrystalGoodsTransfer.WindowShowRefreshBtn = True
   Me.CrystalGoodsTransfer.WindowTitle = "GOODS MOVEMENT REPORT " & Format$(Date, "yyyy")
   CrystalGoodsTransfer.Action = 1
   
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
CrystalGoodsTransfer.ReportFileName = App.Path & "\rptGoodsTransfer.rpt"
CrystalGoodsTransfer.Connect = "DSN=nxomen;UID=sa;PWD=abu;DSQ=ZuksData"


CrystalGoodsTransfer.SelectionFormula = "{GoodsTransfer.TransferDate} =#" & Me.flxgDates.TextMatrix(flxgDates.Row, 0) & "# and {GoodsTransfer.Source} ='" & Me.flxgDates.TextMatrix(flxgDates.Row, 1) & "' and {GoodsTransfer.Destination} ='" & Me.flxgDates.TextMatrix(flxgDates.Row, 2) & "'"


   CrystalGoodsTransfer.WindowState = crptMaximized
   CrystalGoodsTransfer.WindowShowRefreshBtn = True
   Me.CrystalGoodsTransfer.WindowTitle = "GOODS MOVEMENT REPORT " & Format$(Date, "yyyy")
   CrystalGoodsTransfer.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "STOCKING RETAIL REPORT"
       Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.dtpfrom = Date
Me.dtpto = Date
Me.cboProducts = "All"
Me.Height = 7410
Me.Width = 7830
Call GetDates

Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub GetDates()

'On Error GoTo OkError

'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

Me.cboProducts.Clear

 rs.Open "Select Distinct TransferDate,Destination,Source from GoodsTransfer Order By TransferDate", cn, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
   rs.MoveFirst
 Me.flxgDates.Rows = rs.RecordCount + 1
  
  For X = 1 To rs.RecordCount
   Me.flxgDates.TextMatrix(X, 0) = rs.Fields("TransferDate")
   Me.flxgDates.TextMatrix(X, 1) = rs.Fields("Source")
'   If rs.Fields("RefNumber") <> "" Then
   Me.flxgDates.TextMatrix(X, 2) = rs.Fields("Destination")
'   End If
'   Me.flxgDates.TextMatrix(X, 3) = rs.Fields("TransferID")
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

