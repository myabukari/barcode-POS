VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmStockLevelRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "Retail StockLevel"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   Icon            =   "frmStockLevel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   6885
   Begin Crystal.CrystalReport CrystalStockLevel 
      Left            =   6000
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   6375
      Begin lvButton_H.lvButtons_H cmdOk 
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
         Image           =   "frmStockLevel.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
         Image           =   "frmStockLevel.frx":118C
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Caption         =   "SELECT A PRODUCT TO VIEW STOCKLEVEL"
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
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6375
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
         Left            =   1440
         TabIndex        =   2
         Text            =   "All"
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmStockLevelRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection, rs As New ADODB.Recordset, i As Integer
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
       MsgBox "TRY AGAIN", vbInformation, "RECIEVALS REPORTS"
       Exit Sub


End Sub

Private Sub cmdExit_Click()
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError

If Me.cboProducts = "" Then
    MsgBox "Please select Product to view report", vbInformation, "Product name"
    Me.cboProducts.SetFocus
    Exit Sub
End If

'CrystalStockLevel.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
  ' CrystalStockLevel.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalStockLevel.ReportFileName = App.Path & "\StockLevel.rpt"
CrystalStockLevel.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"

If Me.cboProducts <> "All" Then
 CrystalStockLevel.SelectionFormula = "{Products.ProductName} ='" & Me.cboProducts.Text & "'"
Else
 CrystalStockLevel.SelectionFormula = ""
End If
   CrystalStockLevel.WindowState = crptMaximized
   CrystalStockLevel.WindowShowRefreshBtn = True
   Me.CrystalStockLevel.WindowTitle = "LAST RECIEVALS " & Format$(Date, "yyyy")
   CrystalStockLevel.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "RECIEVALS"
       Exit Sub

End Sub

Private Sub Form_Load()
'CenterForm Me
Me.Height = 3585
Me.Width = 7005
Me.cboProducts = "All"
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

