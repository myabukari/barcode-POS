VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmProductsRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "Products Report"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   Icon            =   "frmProductsRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   7095
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6495
      Begin Crystal.CrystalReport CrystalProducts 
         Left            =   5520
         Top             =   120
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
         Left            =   1440
         TabIndex        =   2
         Text            =   "All"
         Top             =   480
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
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   6495
      Begin lvButton_H.lvButtons_H cmdOk 
         Height          =   375
         Left            =   600
         TabIndex        =   4
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
         Image           =   "frmProductsRpt.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   3720
         TabIndex        =   5
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
         Image           =   "frmProductsRpt.frx":118C
         cBack           =   -2147483633
      End
   End
End
Attribute VB_Name = "frmProductsRpt"
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
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError

   'CrystalProducts.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalProducts.ReportFileName = App.Path & "\Products.rpt"
CrystalProducts.Connect = "DSN=nxomen;UID=sa;PWD=abu;DSQ=ZuksData"

If Me.cboProducts <> "All" Then
CrystalProducts.SelectionFormula = "{Products.ProductName} ='" & Me.cboProducts.Text & "'"
Else
CrystalProducts.SelectionFormula = ""
End If
   CrystalProducts.WindowState = crptMaximized
   CrystalProducts.WindowShowRefreshBtn = True
   Me.CrystalProducts.WindowTitle = "Products Report" & Format$(Date, "yyyy")
   CrystalProducts.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "RECIEVALS"
       Exit Sub
End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub Form_Load()
Me.Height = 3690
Me.Width = 7215
Me.cboProducts = "All"
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub
