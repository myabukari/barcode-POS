VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmDebtor 
   BackColor       =   &H00C29E21&
   Caption         =   "Debtors Information"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   Icon            =   "frmDebtor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   8625
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   8295
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   7335
         Begin VB.CommandButton cmdSave1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Save"
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
            Left            =   8160
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdExit1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton cmdClear1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Clear"
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
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdFind1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Find"
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
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelete1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "D&elete"
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
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Save"
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
            Image           =   "frmDebtor.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdDelete 
            Height          =   375
            Left            =   1560
            TabIndex        =   20
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "D&elete"
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
            Image           =   "frmDebtor.frx":075C
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdFind 
            Height          =   375
            Left            =   3000
            TabIndex        =   21
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Find"
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
            Image           =   "frmDebtor.frx":08F6
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   4440
            TabIndex        =   22
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Clear"
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
            Image           =   "frmDebtor.frx":0D48
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   5880
            TabIndex        =   23
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
            Image           =   "frmDebtor.frx":25F9
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   8295
      Begin MSComctlLib.ListView LstDebtors 
         Height          =   2295
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   2499106
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PhoneNumber"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Address"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Account Opening Date"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   8295
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgAccounts 
         Height          =   1695
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   16117969
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"frmDebtor.frx":2A4B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
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
         TabIndex        =   0
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtPhoneNumber 
         Appearance      =   0  'Flat
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
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
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
         Height          =   675
         Left            =   5640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   57147395
         CurrentDate     =   39115
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Name:"
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
         Left            =   720
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Address:"
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
         Left            =   4680
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C29E21&
         Caption         =   "Accont Opening Date:"
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
         Height          =   615
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "PhoneNumber:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmDebtor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Debtorid As Integer
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, AccTransid As String

Private Sub cmdClear_Click()
Call ClearCtrls
End Sub

Private Sub cmdDelete_Click()
If Me.txtname = "" Then
MsgBox "Enter Debtor's name to be deleted", vbInformation, "DELETE WHAT"
Me.txtname.SetFocus: Exit Sub
End If
If MsgBox("ARE YOU SURE  YOU WANT TO DELETE THIS DEBTOR'S DETAILS?", vbYesNo + vbQuestion, "CONFIRM DELETE") = vbYes Then
   
   Me.cmdDelete.Enabled = False
   
'   On Error GoTo SaveError
       
   'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      Me.cmdDelete.Enabled = True
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   cn.BeginTrans
   cn.Execute "Delete From Debtors Where DebtorID ='" & Debtorid & "'", Y
    If Y > 0 Then
    
    rs.Open "Select * From DebtorsTransaction  Where DebtorID = '" & Debtorid & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            MsgBox "This debtor has transactions with you and deleting will not be allowed." & vbCr & "If deleted, your balances will distort ", vbInformation, "Please you cannot proceed"
'            cn.Execute "Delete From DebtorsTransactions Where DebtorID ='" & Debtorid & "'", Y
            If rs.State = 1 Then rs.Close
            cn.RollbackTrans
            Exit Sub
        Else
            Y = 1
            If rs.State = 1 Then rs.Close
        End If
    End If
    
'    If Y > 0 Then
'
'     rs.Open "Select * From DebtorsPayments inner Join DebtorsTransactions on  DebtorsPayments.InvoiceNo=DebtorsTransactions.InvoiceNo  Where DebtorID = '" & Debtorid & "'", cn, adOpenForwardOnly, adLockReadOnly
'        If rs.RecordCount > 0 Then
'            rs.MoveFirst
'            For X = 1 To rs.RecordCount
'            cn.Execute "Delete From DebtorsPayments Where InvoiceNo ='" & rs.Fields("InvoiceNo") & "'", Y
'            rs.MoveNext
'            Next
'            If rs.State = 1 Then rs.Close
'        Else
'            Y = 1
'            If rs.State = 1 Then rs.Close
'        End If
'
'    End If
    
'    If Y > 0 Then
'
'
'    rs.Open "Select * From StockingRetail  Where StoreID Like '" & Storeid & "'", cn, adOpenForwardOnly, adLockReadOnly
'        If rs.RecordCount > 0 Then
'            cn.Execute "Delete From StockingRetail Where StoreID ='" & Storeid & "'", Y
'            If rs.State = 1 Then rs.Close
'        Else
'            Y = 1
'            If rs.State = 1 Then rs.Close
'        End If
'
'    End If
    
    If Y > 0 Then
    cn.CommitTrans
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     'Clear ctrls and setfocus to supplier name ctrl
      ClearCtrls
      Call ListDebtors
      Me.txtname.SetFocus
   Else
      cn.RollbackTrans
      MsgBox "Sorry, Unable to Delete Store's Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
    Me.txtname.SetFocus
   sflag = False
   Me.cmdDelete.Enabled = False
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   
   
   Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Delete Debtor's Details:Please Try Again!", vbInformation, "Delete Failed"
        Exit Sub
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim strg1 As String
'On Error GoTo SaveError

strg1 = InputBox("ENTER DEBTOR'S NAME OR FEW CHARACTERS OF THE NAME.", "ACCOUNTHOLDER'S NAME")
If Trim(strg1) <> "" Then

   
   Me.cmdFind.Enabled = False
          
   'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      Me.cmdFind.Enabled = True
       MsgBox strg, vbInformation:
      Exit Sub
   End If
   strg1 = strg1 & "%"
   rs.Open "Select * From Debtors  Where DebtorName Like '" & strg1 & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount <= 0 Then
      rs.Close: cn.Close
      MsgBox "THERE IS NO DEBTOR WITH THE NAME ENTERED.", vbInformation, "SEARCH FAILED"
      Me.MousePointer = vbDefault: Me.cmdFind.Enabled = True: Me.cmdFind.SetFocus: Exit Sub
   Else
      If rs.RecordCount = 1 Then
         sflag = True
         Me.txtname = rs.Fields("DebtorName")
         Me.txtAddress = rs.Fields("Address")
         Me.txtPhoneNumber = rs.Fields("PhoneNo")
         Me.dtpDate = rs.Fields("AccountOpenDate")
         Debtorid = rs.Fields("DebtorID")
         
         If rs.State = 1 Then rs.Close
         
         Me.txtname.SetFocus
         Me.cmdDelete.Enabled = True
        
      Else
         rs.MoveFirst
         Me.flxgAccounts.Rows = rs.RecordCount + 1
         For X = 1 To rs.RecordCount
           Me.flxgAccounts.TextMatrix(X, 0) = rs.Fields("DebtorName")
           Me.flxgAccounts.TextMatrix(X, 1) = rs.Fields("PhoneNo")
           Me.flxgAccounts.TextMatrix(X, 2) = rs.Fields("AccountOpenDate")
           Me.flxgAccounts.TextMatrix(X, 3) = rs.Fields("Address")
           Me.flxgAccounts.TextMatrix(X, 4) = rs.Fields("DebtorID")
           rs.MoveNext
         Next
         Me.flxgAccounts.Visible = True
         Me.flxgAccounts.SetFocus
         If rs.State = 1 Then rs.Close
      End If
   End If
   

End If

If cn.State = 1 Then cn.Close
If rs.State = 1 Then rs.Close

Me.MousePointer = vbDefault
Me.cmdFind.Enabled = True
Me.cmdSave.Enabled = True
Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     Me.MousePointer = vbDefault
     Me.cmdFind.Enabled = True
     MsgBox "Sorry, Unable to Find Debtors Details:Please Try Again!", vbInformation, "Search Failed"
     Exit Sub
End Sub

Private Sub cmdSave_Click()
If Trim(Me.txtname) = "" Then
   MsgBox "YOU MUST ENTER DEBTOR'S NAME.", vbInformation, "DEBTOR'S NAME"
   Me.txtname.SetFocus: Exit Sub
End If



'On Error GoTo SaveError
Me.cmdSave.Enabled = False

'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If

If sflag = False Then
   'save part
rs.Open "Select DebtorName From Debtors Where DebtorName ='" & Trim(Me.txtname) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "A DEBTOR HAS ALREADY BEEN SETUP WITH THE NAME ENTERED.", vbInformation, "NAME ALREADY EXIST"
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtname.SetFocus: Exit Sub
   
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
End If
If rs.State = 1 Then rs.Close

   'Call Generate_AccountID(Accountid)
   
   cn.Execute "Insert Into Debtors ([DebtorName],[Address],[PhoneNo],[AccountOpenDate]) select '" & Trim(Me.txtname.Text) & "','" & Trim(Me.txtAddress.Text) & "','" & Trim(Me.txtPhoneNumber.Text) & "','" & Trim(Me.dtpDate) & "'", Y
   If Y > 0 Then
     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
       Call ClearCtrls
       Call ListDebtors
       Me.txtname.SetFocus
   Else
      MsgBox "Sorry, Unable to Save Debtor's Details:Please Try Again!", vbInformation, "Save Failed"
      Me.txtname.SetFocus
   End If
 
   
Else
'edit part

   rs.Open "Select DebtorName From Debtors Where DebtorName ='" & Trim(Me.txtname) & "' and DebtorID<>'" & Debtorid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.Close: cn.Close
      MsgBox "A DEBTOR HAS ALREADY BEEN SETUP WITH THE NAME ENTERED.", vbInformation, "NAME ALREADY EXIST"
      Me.txtname.SetFocus: Exit Sub
   End If
   If rs.State = 1 Then rs.Close
   
   
   
   
   cn.Execute "Update Debtors Set DebtorName ='" & Trim(Me.txtname.Text) & "',Address='" & Trim(Me.txtAddress.Text) & "',PhoneNo='" & Trim(Me.txtPhoneNumber.Text) & "' Where DebtorID ='" & Debtorid & "'", Y
      
   If Y > 0 Then
      MsgBox "Edit Successful!", vbInformation, "Edit Successful"
      Call ClearCtrls
      Call ListDebtors
      
      Me.txtname.SetFocus
   Else
      MsgBox "Sorry, Unable to Edit Debtor's Details:Please Try Again!", vbInformation, "Edit Failed"
   End If

 End If

 sflag = False
 If cn.State = 1 Then cn.Close
 If rs.State = 1 Then rs.Close

 Me.MousePointer = vbDefault
 Me.cmdSave.Enabled = True


 Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub

Private Sub flxgAccounts_Click()
Me.txtname = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 0)
Me.txtPhoneNumber = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 1)
Me.dtpDate = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 2)
Me.txtAddress = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 3)
Debtorid = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 4)

Me.flxgAccounts.Visible = False
Me.cmdDelete.Enabled = True
Me.cmdSave.Enabled = True
sflag = True
Me.txtname.SetFocus
End Sub

Private Sub Form_Load()
Call ListDebtors
Me.dtpDate = Date
Me.Height = 7560
  Me.Width = 8745
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub LstDebtors_Click()
On Error GoTo SaveError
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If
rs.Open "Select * From Debtors", cn, adOpenForwardOnly, adLockReadOnly


    While Not rs.EOF
        If Me.LstDebtors.SelectedItem.Text = rs("DebtorName") Then
            Me.txtname.Text = Trim(rs("DebtorName"))
            
                Debtorid = rs("DebtorID")
                Me.txtAddress.Text = Trim(rs("Address"))
                Me.txtPhoneNumber.Text = Trim(rs("PhoneNo"))
                Me.dtpDate = Trim(rs("AccountOpenDate"))
                
            rs.MoveNext
        Else
            rs.MoveNext
        End If
    Wend
    rs.Close
    Set rs = Nothing
    sflag = True
    Me.cmdDelete.Enabled = True
    Me.cmdSave.Enabled = True
    
    Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "SAVE FAILED"
     Exit Sub
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()@#$%&*_-'"
If KeyAscii = vbKeyReturn Then
   
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtPhoneNumber.SetFocus
End If
End Sub
Private Sub ListDebtors()
On Error GoTo SaveError
Me.LstDebtors.ListItems.Clear
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select * From Debtors Order By DebtorName", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.LstDebtors.ListItems.Add(, , Trim(rs!DebtorName))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!PhoneNo)
                List_Item.SubItems(2) = Trim(rs!Address)
                List_Item.SubItems(3) = Trim(rs!AccountOpenDate)
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
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "SAVE FAILED"
     Exit Sub
    
End Sub
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
sflag = False
End Sub

