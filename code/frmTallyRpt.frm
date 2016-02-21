VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmTallyRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "Tally Report  "
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   Icon            =   "frmTallyRpt.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4755
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   8295
      Begin VB.CheckBox ChkDate 
         BackColor       =   &H00C29E21&
         Caption         =   "Specify Date Range"
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
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
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
         Left            =   1680
         TabIndex        =   11
         Text            =   "All"
         Top             =   480
         Width           =   6135
      End
      Begin Crystal.CrystalReport CrystalTallyReport 
         Left            =   7440
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   8295
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
         Left            =   9000
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
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
         Left            =   9000
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin lvButton_H.lvButtons_H cmdOk 
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
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
         Image           =   "frmTallyRpt.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
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
         Image           =   "frmTallyRpt.frx":118C
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Caption         =   "Date Range"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   8295
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   4920
         TabIndex        =   1
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
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
         Format          =   61014019
         CurrentDate     =   39087
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
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
         Format          =   61014019
         CurrentDate     =   39087
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
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   375
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmTallyRpt"
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
       MsgBox "TRY AGAIN", vbInformation, "RECIEVALS REPORTS"
       Exit Sub



End Sub

Private Sub ChkDate_Click()
If Me.ChkDate = vbChecked Then
Me.dtpfrom.Enabled = True
Me.dtpto.Enabled = True
Else
Me.dtpfrom.Enabled = False
Me.dtpto.Enabled = False
End If
End Sub

Private Sub cmdExit_Click()
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError

If Me.dtpfrom > Me.dtpto Then
MsgBox "WRONG DATE RANGE ENTERED", vbInformation, "DATE RANGE"
Exit Sub
End If

'If Me.cboReport = "" Then
'MsgBox "SPECIFY PERIOD RANGE", vbInformation, ""
'Me.cboReport.SetFocus
'Exit Sub
'End If
'CrystalRevenue.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalRevenue.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalTallyReport.ReportFileName = App.Path & "\rtpTallyCard.rpt"
CrystalTallyReport.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"


'If Me.cboReport = "Today" And Me.cboProducts = "All" Then 'Me.dtpfrom = Me.dtpto Then
'CrystalRevenue.SelectionFormula = "{Payments/Products.Date} =#" & Date & "#"
'End If
'
'If Me.cboReport = "Today" And Me.cboProducts <> "All" Then 'Me.dtpfrom = Me.dtpto Then
'CrystalRevenue.SelectionFormula = "{Payments/Products.Date} =#" & Date & "#and {Products.ProductName}='" & Trim(Me.cboProducts) & "'"
'End If


If Me.cboProducts <> "All" And Me.ChkDate = vbUnchecked Then
CrystalTallyReport.SelectionFormula = "{Products.ProductName}='" & Trim(Me.cboProducts) & "'"
End If

If Me.cboProducts <> "All" And Me.ChkDate = vbChecked Then
CrystalTallyReport.SelectionFormula = "{Tally.IssueDate} =#" & Date & "# and {Products.ProductName}='" & Trim(Me.cboProducts) & "'"
End If

'If Me.cboProducts = "All" And Me.ChkDate = vbChecked Then
'CrystalRevenue.SelectionFormula = "{Payments/Products.Date} =#" & Date & "#"
'End If

If Me.cboProducts = "All" And Me.ChkDate = vbUnchecked Then
CrystalTallyReport.SelectionFormula = ""
End If


If Me.dtpfrom <= Me.dtpto And Me.cboProducts = "All" And Me.ChkDate = vbChecked Then
CrystalTallyReport.SelectionFormula = "{Tally.IssueDate} >=#" & Me.dtpfrom & "# and {Tally.IssueDate} <=#" & Me.dtpto & "#"
End If

'And Me.cboReport = "Specified Time Period" Then

If Me.dtpfrom <= Me.dtpto And Me.cboProducts <> "All" And Me.ChkDate = vbChecked Then
CrystalTallyReport.SelectionFormula = "{Tally.IssueDate} >=#" & Me.dtpfrom & "# and {Tally.IssueDate} <=#" & Me.dtpto & "# and {Products.ProductName}='" & Trim(Me.cboProducts) & "'"
End If



   CrystalTallyReport.WindowState = crptMaximized
   CrystalTallyReport.WindowShowRefreshBtn = True
   Me.CrystalTallyReport.WindowTitle = "TALLY/BIN CARD " & Format$(Date, "yyyy")
   
   CrystalTallyReport.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "SALES REVENUE"
       Exit Sub

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub Form_Load()
Me.dtpfrom = Date
Me.dtpto = Date
Me.dtpfrom.Enabled = False
Me.dtpto.Enabled = False
'CenterForm Me

Me.Height = 5055
Me.Width = 8955
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub


