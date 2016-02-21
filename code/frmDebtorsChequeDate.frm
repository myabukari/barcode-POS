VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmDebtorsChequeDate 
   BackColor       =   &H00C29E21&
   Caption         =   "ChequeDue Date"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13740
   Icon            =   "frmDebtorsChequeDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   13740
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Caption         =   "CREDITOR'S CHEQUE DUE DATES"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   13455
      Begin MSComctlLib.ListView LstCreditor 
         Height          =   1935
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483643
         BackColor       =   -2147483625
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payment Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Payment Mode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Creditor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Invoice Number"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount Paid"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cheque Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ChequeDue Date"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Caption         =   "DEBTOR'S CHEQUE DUE DATES"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   13455
      Begin MSComctlLib.ListView LstChequeDueDate 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483643
         BackColor       =   0
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payment Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Payment Mode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Debtor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Invoice Number"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount Paid"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cheque Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cheque Due Date"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin lvButton_H.lvButtons_H cmdeXIT 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   661
      Caption         =   "E&xit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Image           =   "frmDebtorsChequeDate.frx":030A
      cBack           =   -2147483633
   End
   Begin VB.CommandButton cmdeXIT1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   9780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C29E21&
      Caption         =   " CHEQUE DUE DATES FOR TODAY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   13455
   End
End
Attribute VB_Name = "frmDebtorsChequeDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Accountid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer
Dim BalCD As Double, AccTransid As String, xflag As Boolean, AccountTransid As String

Private Sub cmdExit_Click()
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub Form_Load()
Call ListCheque
Call listCreditorCheque
 Me.Height = 7170
  Me.Width = 13860
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub ListCheque()
On Error GoTo SaveError
Me.LstChequeDueDate.ListItems.Clear
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select * From AccountHolders Inner Join(AccountTransaction Inner Join TransPayments On AccountTransaction.AccTransID=TransPayments.AccTransID ) On AccountHolders.AccID=AccountTransaction.AccID Where ChequeDueDate= ' " & Date & " ' And  PaymentMode='Cheque'", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.LstChequeDueDate.ListItems.Add(, , Trim(rs!PaymentDate))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!PaymentMode)
                List_Item.SubItems(2) = Trim(rs!Name)
                List_Item.SubItems(3) = Trim(rs!InvoiceNo)
                List_Item.SubItems(4) = Trim(rs!AmountPaid)
               List_Item.SubItems(5) = Trim(rs!ChequeNo)
                List_Item.SubItems(6) = Trim(rs!ChequeDueDate)
                 'List_Item.SubItems(6) = Trim(rs!BalanceCD)
                rs.MoveNext
            Loop
        End If
    Next i
    DoEvents
    
  rs.Close
    Set rs = Nothing
    If cn.State = 1 Then cn.Close
    Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "TRY AGAIN"
     Exit Sub
    
End Sub

Private Sub listCreditorCheque()
On Error GoTo SaveError
Me.LstCreditor.ListItems.Clear
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select * From Suppliers Inner Join(CreditorAccount Inner Join CreditorPayment On CreditorAccount.CreditorAccID=CreditorPayment.CreditorAccID ) On Suppliers.SupplierID=CreditorAccount.SupplierID Where ChequeDueDate= ' " & Date & " ' And  PaymentMode='Cheque'", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.LstCreditor.ListItems.Add(, , Trim(rs!PaymentDate))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!PaymentMode)
                List_Item.SubItems(2) = Trim(rs!SupplierName)
                List_Item.SubItems(3) = Trim(rs!InvoiceNo)
                List_Item.SubItems(4) = Trim(rs!AmountPaid)
               List_Item.SubItems(5) = Trim(rs!ChequeNo)
                List_Item.SubItems(6) = Trim(rs!ChequeDueDate)
                 'List_Item.SubItems(6) = Trim(rs!BalanceCD)
                rs.MoveNext
            Loop
        End If
    Next i
    DoEvents
    
  rs.Close
    Set rs = Nothing
    If cn.State = 1 Then cn.Close
    Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "TRY AGAIN"
     Exit Sub
    
End Sub

