VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmMainStoreStockLevelRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "WholeSale StockLevel Reports"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   Icon            =   "frmMainStoreStockLevelReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   6630
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   6135
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   615
         Left            =   1080
         TabIndex        =   4
         Top             =   120
         Width           =   4215
         Begin lvButton_H.lvButtons_H cmdOk 
            Height          =   375
            Left            =   480
            TabIndex        =   7
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
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
            Image           =   "frmMainStoreStockLevelReport.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   2280
            TabIndex        =   8
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
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
            Image           =   "frmMainStoreStockLevelReport.frx":118C
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdOk1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Ok"
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
            Left            =   480
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdExit1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Exit"
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
            Left            =   2280
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      Begin Crystal.CrystalReport CrystaMainStockLevel 
         Left            =   5400
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
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
         Left            =   1920
         TabIndex        =   1
         Text            =   "All"
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
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
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMainStoreStockLevelRpt"
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

 rs.Open "Select Distinct ProductName from MainStoreProducts  Order By ProductName Asc", cn, adOpenForwardOnly, adLockReadOnly

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

'CrystaMainStockLevel.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
   'CrystaMainStockLevel.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystaMainStockLevel.ReportFileName = App.Path & "\MainStoreStockLevel.rpt"
CrystaMainStockLevel.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"



If Me.cboProducts <> "All" Then
CrystaMainStockLevel.SelectionFormula = "{MainStoreProducts.ProductName} ='" & Me.cboProducts.Text & "'"
Else
CrystaMainStockLevel.SelectionFormula = ""
End If
   CrystaMainStockLevel.WindowState = crptMaximized
   CrystaMainStockLevel.WindowShowRefreshBtn = True
   Me.CrystaMainStockLevel.WindowTitle = "WHOLESALE STOCKLEVEL " & Format$(Date, "yyyy")
   CrystaMainStockLevel.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "RECIEVALS"
       Exit Sub

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Height = 3930
Me.Width = 6750
Me.cboProducts.Text = "All"
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

