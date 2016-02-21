VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmWholeSaleWorth 
   BackColor       =   &H00C29E21&
   Caption         =   "WholeSaleWorth"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   Icon            =   "frmWholeSaleWorth.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   7230
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   6615
      Begin lvButton_H.lvButtons_H cmdOk 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
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
         Image           =   "frmWholeSaleWorth.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
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
         Image           =   "frmWholeSaleWorth.frx":118C
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6615
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
         Left            =   1560
         TabIndex        =   6
         Text            =   "All"
         Top             =   600
         Width           =   4455
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
         Left            =   1560
         TabIndex        =   1
         Text            =   "All"
         Top             =   1320
         Width           =   4455
      End
      Begin Crystal.CrystalReport CrystalWholeSaleWorth 
         Left            =   6000
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   1095
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
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmWholeSaleWorth"
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
 MsgBox "Select Store's name to view report", vbInformation, "Store Name"
 Me.cboStore.SetFocus: Exit Sub
End If

If Me.cboProducts = "" Then
 MsgBox "Select Product Name to view Report", vbInformation, "Product Name"
 Me.cboProducts.SetFocus: Exit Sub
End If

   'CrystalWholeSaleWorth.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalWholeSaleWorth.ReportFileName = App.Path & "\WholeSaleWorth.rpt"
CrystalWholeSaleWorth.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"

If Me.cboProducts <> "All" And Me.cboStore = "All" Then
CrystalWholeSaleWorth.SelectionFormula = "{Products.ProductName} ='" & Trim(Me.cboProducts.Text) & "'"

ElseIf Me.cboProducts = "All" And Me.cboStore <> "All" Then
CrystalWholeSaleWorth.SelectionFormula = "{Stores.StoreName} ='" & Trim(Me.cboStore) & "'"

ElseIf Me.cboProducts = "All" And Me.cboStore = "All" Then
CrystalWholeSaleWorth.SelectionFormula = ""

ElseIf Me.cboProducts <> "All" And Me.cboStore <> "All" Then
CrystalWholeSaleWorth.SelectionFormula = "{Products.ProductName} ='" & Trim(Me.cboProducts.Text) & "' and {Stores.StoreName} ='" & Trim(Me.cboStore) & "'"

End If

   CrystalWholeSaleWorth.WindowState = crptMaximized
   CrystalWholeSaleWorth.WindowShowRefreshBtn = True
   Me.CrystalWholeSaleWorth.WindowTitle = "WholeSaleWorth Report " & Format$(Date, "yyyy")
   CrystalWholeSaleWorth.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "WholeSaleWorth"
       Exit Sub
End Sub

Private Sub Form_Load()
Me.Height = 4215
Me.Width = 7350
Me.cboProducts = "All"
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub
