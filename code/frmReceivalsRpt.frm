VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmMainReceivalsRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "WholeSale Receivals Report"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   Icon            =   "frmReceivalsRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   7920
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   7695
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Width           =   6975
         Begin lvButton_H.lvButtons_H cmdOk 
            Height          =   375
            Left            =   1800
            TabIndex        =   11
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
            Image           =   "frmReceivalsRpt.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   3840
            TabIndex        =   12
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
            Image           =   "frmReceivalsRpt.frx":118C
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
            Left            =   1800
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   10
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
            Left            =   3840
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7695
      Begin Crystal.CrystalReport CrystalReceivals 
         Left            =   6960
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   5760
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   61014019
         CurrentDate     =   39090
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
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   5760
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   61014019
         CurrentDate     =   39090
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
         Left            =   5160
         TabIndex        =   6
         Top             =   1200
         Width           =   495
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
         Left            =   5040
         TabIndex        =   5
         Top             =   600
         Width           =   615
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
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMainReceivalsRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer

Private Sub cboProductName_DropDown()

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
If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
End If
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError
If Me.cboProducts = "" Then
MsgBox "Select Product to View Report", vbInformation, "Select Product"
Me.cboProducts.SetFocus: Exit Sub
End If
'If Me.cboReport = "" Then
'MsgBox "SPECIFY PERIOD RANGE", vbInformation, ""
'Me.cboReport.SetFocus
'Exit Sub
'End If
'CrystalReceivals.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalReceivals.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalReceivals.ReportFileName = App.Path & "\Receivals.rpt"
CrystalReceivals.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"


'If Me.cboReport = "Today" And Me.cboProducts = "All" Then 'Me.dtpfrom = Me.dtpto Then
'CrystalReceivals.SelectionFormula = "{Payments/Products.Date} =#" & Date & "#"
'End If

'If Me.cboReport = "Today" And Me.cboProducts <> "All" Then 'Me.dtpfrom = Me.dtpto Then
'CrystalReceivals.SelectionFormula = "{Payments/Products.Date} =#" & Date & "#and {Products.ProductName}='" & Trim(Me.cboProducts) & "'"
'End If

If Me.dtpfrom <= Me.dtpto And Me.cboProducts = "All" Then
CrystalReceivals.SelectionFormula = "{Receivals.Date} >=#" & Me.dtpfrom & "# and {Receivals.Date} <=#" & Me.dtpto & "#"
End If

If Me.dtpfrom <= Me.dtpto And Me.cboProducts <> "All" Then
CrystalReceivals.SelectionFormula = "{Receivals.Date} >=#" & Me.dtpfrom & "# and {Receivals.Date} <=#" & Me.dtpto & "# and {MainStoreProducts.ProductName}='" & Trim(Me.cboProducts) & "'"
End If

If Me.dtpfrom > Me.dtpto Then
MsgBox "WRONG DATE RANGE ENTERED", vbInformation, "DATE RANGE"
Exit Sub
End If

   CrystalReceivals.WindowState = crptMaximized
   CrystalReceivals.WindowShowRefreshBtn = True
   Me.CrystalReceivals.WindowTitle = "WHOLESALE RECEIVALS REPORT " & Format$(Date, "yyyy")
   
   CrystalReceivals.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "SALES REVENUE"
       Exit Sub

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Height = 4305
Me.Width = 8040
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2

End Sub
