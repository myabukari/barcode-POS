VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmIndividualDebtorPaymentsRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "Specified Debtor's Report"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   Icon            =   "frmIndividualDebtorPaymentsRpt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   8115
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   7575
      Begin Crystal.CrystalReport CrystalIndividualDebtorsPayments 
         Left            =   6960
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cboDebtor 
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
         TabIndex        =   5
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Debtor's Name:"
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
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   7575
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   1800
         TabIndex        =   1
         Top             =   0
         Width           =   4095
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   2400
            TabIndex        =   2
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
            Image           =   "frmIndividualDebtorPaymentsRpt.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdOk 
            Height          =   375
            Left            =   480
            TabIndex        =   3
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
            Image           =   "frmIndividualDebtorPaymentsRpt.frx":075C
            cBack           =   -2147483633
         End
      End
   End
End
Attribute VB_Name = "frmIndividualDebtorPaymentsRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer

Private Sub cboDebtor_DropDown()
On Error GoTo OkError


'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

Me.cboDebtor.Clear

 rs.Open "Select Distinct DebtorName from Debtors  Order By DebtorName Asc", cn, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
   rs.MoveFirst
   For i = 1 To rs.RecordCount
     Me.cboDebtor.AddItem rs.Fields!DebtorName
     
     rs.MoveNext
   Next
   If cboDebtor.ListCount > 1 Then
      Me.cboDebtor.AddItem "All"
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
       MsgBox "Sorry cannot display Debtor's Name,Please try again ", vbInformation, "Debtor's Name"
       Exit Sub

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError

If Me.cboDebtor = "" Then
    MsgBox "Please Select a Debtor or All to view report ", vbInformation, "Debtor's Name"
    Me.cboDebtor.SetFocus: Exit Sub
End If

'CrystalIndividualDebtorsPayments.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalIndividualDebtorsPayments.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalIndividualDebtorsPayments.ReportFileName = App.Path & "\rptIndividualDebtorsPayments.rpt"
CrystalIndividualDebtorsPayments.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"


If Me.cboDebtor <> "All" Then
bFlag = OpenConnection(cn, strg)
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
 
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select  * From Debtors inner Join DebtorsTransaction on Debtors.DebtorID=DebtorsTransaction.DebtorID  where DebtorName='" & Me.cboDebtor & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED DEBTOR", vbInformation, "DEBTOR NOT IN REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close

If Me.cboDebtor = "All" Then
CrystalIndividualDebtorsPayments.SelectionFormula = ""
End If

If Me.cboDebtor <> "All" Then
CrystalIndividualDebtorsPayments.SelectionFormula = "{Debtors.DebtorName}='" & Trim(Me.cboDebtor) & "'"
End If


   CrystalIndividualDebtorsPayments.WindowState = crptMaximized
   CrystalIndividualDebtorsPayments.WindowShowRefreshBtn = True
   Me.CrystalIndividualDebtorsPayments.WindowTitle = "DEBTOR'S PAYMENTS " & Format$(Date, "yyyy")
   
   CrystalIndividualDebtorsPayments.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "DEBTOR'S PAYMENTS"
       Exit Sub

End Sub

Private Sub Form_Load()
Me.cboDebtor = "All"
  Me.Height = 4440
  Me.Width = 8235
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

