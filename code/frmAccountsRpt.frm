VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmAccountsRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "Debtors Payments"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   Icon            =   "frmAccountsRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   8265
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   7575
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   1800
         TabIndex        =   6
         Top             =   0
         Width           =   4095
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   2400
            TabIndex        =   9
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            Image           =   "frmAccountsRpt.frx":030A
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdExit1 
            BackColor       =   &H00C0C0C0&
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
            Left            =   2400
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin lvButton_H.lvButtons_H cmdOk 
            Height          =   375
            Left            =   480
            TabIndex        =   10
            Top             =   240
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
            Image           =   "frmAccountsRpt.frx":075C
            cBack           =   -2147483633
         End
         Begin VB.CommandButton cmdOk1 
            BackColor       =   &H00C0C0C0&
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
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      Begin Crystal.CrystalReport CrystalAccounts 
         Left            =   5520
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cboInvoiceNo 
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
         TabIndex        =   2
         Top             =   1080
         Width           =   4575
      End
      Begin VB.ComboBox cboName 
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
         Caption         =   "Invoice Number:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Debtors Name:"
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
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAccountsRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer

Private Sub cboInvoiceNo_DropDown()
On Error GoTo OkError


'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

Me.cboInvoiceNo.Clear

 rs.Open "Select InvoiceNo from AccountHolders Inner Join AccountTransaction On AccountHolders.AccID=AccountTransaction.AccID Where AccountHolders.Name='" & Trim(Me.cboName) & "' Order By InvoiceNo Asc", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
   For i = 1 To rs.RecordCount
     Me.cboInvoiceNo.AddItem rs.Fields!InvoiceNo
     rs.MoveNext
   Next
   If cboInvoiceNo.ListCount > 1 Then
      Me.cboInvoiceNo.AddItem "All"
   End If
 End If
If rs.State = 1 Then
   rs.Close
End If


If Me.cboName = "All" Then
 rs.Open "Select InvoiceNo from AccountHolders Inner Join AccountTransaction On AccountHolders.AccID=AccountTransaction.AccID Order By InvoiceNo Asc", cn, adOpenForwardOnly, adLockReadOnly
 If rs.RecordCount > 0 Then
   rs.MoveFirst
   For i = 1 To rs.RecordCount
     Me.cboInvoiceNo.AddItem rs.Fields!InvoiceNo
     rs.MoveNext
   Next
   If cboInvoiceNo.ListCount > 1 Then
      Me.cboInvoiceNo.AddItem "All"
   End If
 End If
If rs.State = 1 Then
   rs.Close
End If
End If

If cn.State = 1 Then cn.Close
If rs.State = 1 Then rs.Close
Exit Sub
OkError:
       If rs.State = 1 Then
          rs.Close
       End If
       MsgBox "TRY AGAIN", vbInformation, "ACCOUNTS REPORTS"
       Exit Sub



End Sub

Private Sub cboName_DropDown()
On Error GoTo OkError


'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

Me.cboName.Clear

 rs.Open "Select Distinct Name from AccountHolders  Order By Name Asc", cn, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
   rs.MoveFirst
   For i = 1 To rs.RecordCount
     Me.cboName.AddItem rs.Fields!Name
     
     rs.MoveNext
   Next
   If cboName.ListCount > 1 Then
      Me.cboName.AddItem "All"
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
       MsgBox "TRY AGAIN", vbInformation, "ACCOUNTS REPORTS"
       Exit Sub
End Sub

Private Sub cmdExit_Click()
If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
End If
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError
If Me.cboName = "" Then
MsgBox "SELECT ACCOUNT NAME", vbInformation, ""
Me.cboName.SetFocus
Exit Sub
End If

If Me.cboInvoiceNo = "" Then
MsgBox "SELECT INVOICE NUMBER", vbInformation, ""
Me.cboInvoiceNo.SetFocus
Exit Sub
End If

'CrystalAccounts.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalAccounts.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalAccounts.ReportFileName = App.Path & "\Accounts.rpt"
CrystalAccounts.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"

If Me.cboName <> "All" And Me.cboInvoiceNo = "All" Then
CrystalAccounts.SelectionFormula = "{AccountHolders.Name} ='" & Trim(Me.cboName) & "'"
End If

If Me.cboName <> "All" And Me.cboInvoiceNo <> "All" Then
CrystalAccounts.SelectionFormula = "{AccountHolders.Name} ='" & Trim(Me.cboName) & "'and {AccountTransaction.InvoiceNo} ='" & Trim(Me.cboInvoiceNo) & "'"
End If

If Me.cboName = "All" And Me.cboInvoiceNo = "All" Then
CrystalAccounts.SelectionFormula = ""
End If

   CrystalAccounts.WindowState = crptMaximized
   CrystalAccounts.WindowShowRefreshBtn = True
   Me.CrystalAccounts.WindowTitle = "DEBTOR'S ACCOUNT " & Format$(Date, "yyyy")
   
   CrystalAccounts.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "DEBTOR'S ACCOUNT "
       Exit Sub

End Sub

Private Sub Form_Load()
Me.Width = 8385
Me.Height = 4470
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub
