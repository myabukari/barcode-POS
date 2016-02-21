VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmAccountsInfo 
   BackColor       =   &H00C29E21&
   Caption         =   "Set Debtors"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   Icon            =   "frmAccountsInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   9705
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   9375
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   480
         TabIndex        =   16
         Top             =   120
         Width           =   8175
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
            TabIndex        =   21
            Top             =   240
            Width           =   1215
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
            TabIndex        =   20
            Top             =   480
            Width           =   1335
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
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   120
            TabIndex        =   26
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
            Image           =   "frmAccountsInfo.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdDelete 
            Height          =   375
            Left            =   1800
            TabIndex        =   27
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
            Image           =   "frmAccountsInfo.frx":075C
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdFind 
            Height          =   375
            Left            =   3480
            TabIndex        =   28
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
            Image           =   "frmAccountsInfo.frx":08F6
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   5160
            TabIndex        =   29
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
            Image           =   "frmAccountsInfo.frx":0D48
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   6720
            TabIndex        =   30
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
            Image           =   "frmAccountsInfo.frx":25F9
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   3495
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   9375
      Begin MSComctlLib.ListView LstAccounts 
         Height          =   3255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5741
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
         NumItems        =   7
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
            Text            =   "Email"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fax"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "CreditLimit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgAccounts 
         Height          =   1695
         Left            =   1080
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   12754465
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"frmAccountsInfo.frx":2A4B
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
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   1920
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
         Format          =   57081859
         CurrentDate     =   39115
      End
      Begin VB.TextBox txtEmail 
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
         TabIndex        =   12
         Top             =   2520
         Width           =   7095
      End
      Begin VB.TextBox txtCreditLimit 
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
         Left            =   6240
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtFax 
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
         TabIndex        =   7
         Top             =   1440
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
         Left            =   6240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   960
         Width           =   2415
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
         TabIndex        =   1
         Top             =   480
         Width           =   2415
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
         TabIndex        =   5
         Top             =   960
         Width           =   1215
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
         Left            =   720
         TabIndex        =   23
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C29E21&
         Caption         =   "Email:"
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
         Left            =   840
         TabIndex        =   11
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Credit Limit:"
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
         Left            =   5040
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Fax:"
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
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   495
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
         Left            =   5280
         TabIndex        =   6
         Top             =   480
         Width           =   855
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
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C29E21&
         Height          =   2895
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   9135
      End
   End
End
Attribute VB_Name = "frmAccountsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Accountid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, AccTransid As String

Private Sub cmdClear_Click()
 Call ClearCtrls
 Me.txtName.SetFocus
 Me.flxgAccounts.Visible = False
 sflag = False
End Sub

Private Sub cmdDelete_Click()
Dim X As Integer, rec As Integer, AccTransid() As String
If Me.txtName = "" Then
MsgBox "There is Nothing to be Deleted", vbInformation, "Select Debtor to Delete"
Exit Sub
End If

If MsgBox("ARE YOU SURE  YOU WANT TO DELETE THE DEBTOR'S DETAILS?", vbYesNo + vbQuestion, "CONFIRM DELETE") = vbYes Then
   
   Me.cmdDelete.Enabled = False
   
   On Error GoTo SaveError
       
   'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      Me.MousePointer = vbDefault
      Me.cmdDelete.Enabled = True
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   rs.Open "Select AccTransID From AccountTransaction  Where AccID='" & Accountid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
   rec = rs.RecordCount - 1
   ReDim AccTransid(rs.RecordCount - 1)
   For X = 0 To rs.RecordCount - 1
   AccTransid(X) = rs.Fields("AccTransID")
   rs.MoveNext
   Next
   If rs.State = 1 Then rs.Close
   Else
    ReDim AccTransid(0)
    If rs.State = 1 Then rs.Close
   End If
   If rs.State = 1 Then rs.Close

   
   cn.BeginTrans
   cn.Execute "Delete From AccountHolders Where AccID ='" & Accountid & "'", Y
    If Y > 0 Then
    
    rs.Open "Select * From AccountTransaction  Where AccID='" & Accountid & "'", cn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
     cn.Execute "Delete From  AccountTransaction  Where AccID ='" & Accountid & "'", Y
     If rs.State = 1 Then rs.Close
    End If
     If rs.State = 1 Then rs.Close
    
    End If
    If Y > 0 Then
    For X = 0 To rec
    rs.Open "Select * From TransPayments  Where AccTransID='" & AccTransid(X) & "'", cn, adOpenForwardOnly, adLockReadOnly
      If rs.RecordCount > 0 Then
        cn.Execute "Delete From  TransPayments  Where AccTransID ='" & AccTransid(X) & "'", Y
       If rs.State = 1 Then rs.Close
      End If
       If rs.State = 1 Then rs.Close
    Next
    End If
    If Y > 0 Then
      cn.CommitTrans
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     
      ClearCtrls
      Me.txtName.SetFocus
   Else
      cn.RollbackTrans
      MsgBox "Sorry, Unable to Delete Products Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
    Me.txtName.SetFocus
   sflag = False
   Me.cmdDelete.Enabled = False
   If cn.State = 1 Then cn.Close
   
   Me.MousePointer = vbDefault
   
  ' Call ListProducts
   Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        If rs.State = 1 Then rs.Close
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Delete Products  Details:Please Try Again!", vbInformation, "Delete Failed"
        Exit Sub
End If
End Sub

Private Sub cmdExit_Click()
If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
End If
End Sub

Private Sub cmdFind_Click()
Dim strg1 As String
On Error GoTo SaveError

strg1 = InputBox("ENTER ACCOUNTHOLDER'S NAME OR FEW CHARASTERS OF THE NAME.", "ACCOUNTHOLDER'S NAME")
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
   rs.Open "Select * From AccountHolders  Where Name Like '" & strg1 & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount <= 0 Then
      rs.Close: cn.Close
      MsgBox "THERE IS NO ACCOUNTHOLDER WITH THE NAME ENTERED.", vbInformation, "SEARCH FAILED"
      Me.MousePointer = vbDefault: Me.cmdFind.Enabled = True: Me.cmdFind.SetFocus: Exit Sub
   Else
      If rs.RecordCount = 1 Then
         sflag = True
         Me.txtName = rs.Fields("Name")
         Me.txtAddress = rs.Fields("Address")
         Me.txtPhoneNumber = rs.Fields("PhoneNumber")
         Me.txtFax = rs.Fields("Fax")
         Me.txtEmail = rs.Fields("Email")
         Me.txtCreditLimit = rs.Fields("CreditLimit")
         Me.dtpDate = rs.Fields("AccDate")
         Accountid = rs.Fields("AccID")
         
         rs.Close
         
         Me.txtName.SetFocus
         Me.cmdDelete.Enabled = True
        
      Else
         rs.MoveFirst
         Me.flxgAccounts.Rows = rs.RecordCount + 1
         For X = 1 To rs.RecordCount
           Me.flxgAccounts.TextMatrix(X, 0) = rs.Fields("Name")
           Me.flxgAccounts.TextMatrix(X, 1) = rs.Fields("PhoneNumber")
           Me.flxgAccounts.TextMatrix(X, 2) = rs.Fields("Fax")
           Me.flxgAccounts.TextMatrix(X, 3) = rs.Fields("AccDate")
           Me.flxgAccounts.TextMatrix(X, 4) = rs.Fields("Address")
           Me.flxgAccounts.TextMatrix(X, 5) = rs.Fields("CreditLimit")
           Me.flxgAccounts.TextMatrix(X, 6) = rs.Fields("Email")
           Me.flxgAccounts.TextMatrix(X, 7) = rs.Fields("AccID")
           rs.MoveNext
         Next
         Me.flxgAccounts.Visible = True
         Me.flxgAccounts.SetFocus
         rs.Close
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
     MsgBox "Sorry, Unable to Find Products Details:Please Try Again!", vbInformation, "Search Failed"
     Exit Sub
End Sub

Private Sub cmdSave_Click()
If Trim(Me.txtName) = "" Then
   MsgBox "YOU MUST ENTER ACCOUNTHOLDER'S NAME.", vbInformation, "ACCOUNTHOLDER'S NAME"
   Me.txtName.SetFocus: Exit Sub
End If



On Error GoTo SaveError
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
rs.Open "Select Name From AccountHolders Where Name ='" & Trim(Me.txtName) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "AN ACCOUNTHOLDER HAS ALREADY BEEN SETUP WITH THE NAME ENTERED.", vbInformation, "NAME ALREADY EXIST"
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtName.SetFocus: Exit Sub
   
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
End If
rs.Close

   Call Generate_AccountID(Accountid)
   
   cn.Execute "Insert Into AccountHolders ([AccID],[Name],[Address],[PhoneNumber],[Fax],[Email],[CreditLimit],[AccDate]) select '" & Trim(Accountid) & "','" & Trim(Me.txtName.Text) & "','" & Trim(Me.txtAddress.Text) & "','" & Trim(Me.txtPhoneNumber.Text) & "','" & Trim(Me.txtFax.Text) & "','" & Trim(Me.txtEmail.Text) & "','" & Val(Trim(Me.txtCreditLimit.Text)) & "','" & Trim(Me.dtpDate) & "'", Y
   If Y > 0 Then
     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
       Call ClearCtrls
       Call ListAccounts
       Me.txtName.SetFocus
   Else
      MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
      Me.txtName.SetFocus
   End If
 
   
Else
'edit part

   rs.Open "Select Name From AccountHolders Where Name ='" & Trim(Me.txtName) & "' and AccID<>'" & Accountid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.Close: cn.Close
      MsgBox "AN ACCOUNTHOLDER HAS ALREADY BEEN SETUP WITH THE NAME ENTERED.", vbInformation, "NAME ALREADY EXIST"
      Me.txtName.SetFocus: Exit Sub
   End If
   rs.Close
   
   
   
   
   cn.Execute "Update AccountHolders Set Name ='" & Trim(Me.txtName.Text) & "',Address='" & Trim(Me.txtAddress.Text) & "',PhoneNumber='" & Trim(Me.txtPhoneNumber.Text) & "',Fax='" & Trim(Me.txtFax.Text) & "',Email='" & Trim(Me.txtEmail.Text) & "',CreditLimit='" & Trim(Me.txtCreditLimit.Text) & "' Where AccID ='" & Accountid & "'", Y
      
   If Y > 0 Then
      MsgBox "Edit Successful!", vbInformation, "Edit Successful"
      Call ClearCtrls
      Call ListAccounts
      
      Me.txtName.SetFocus
   Else
      MsgBox "Sorry, Unable to Edit Product Details:Please Try Again!", vbInformation, "Edit Failed"
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
Me.txtName = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 0)
Me.txtPhoneNumber = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 1)
Me.txtFax = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 2)
Me.dtpDate = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 3)
Me.txtAddress = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 4)
Me.txtCreditLimit = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 5)
Me.txtEmail = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 6)
Accountid = Me.flxgAccounts.TextMatrix(Me.flxgAccounts.Row, 7)

Me.flxgAccounts.Visible = False
Me.cmdDelete.Enabled = True
Me.cmdSave.Enabled = True
sflag = True
Me.txtName.SetFocus
End Sub

Private Sub Form_Load()
Call ListAccounts
Me.Width = 9825
Me.Height = 8925
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub LstAccounts_Click()
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
rs.Open "Select * From AccountHolders", cn, adOpenForwardOnly, adLockReadOnly


    While Not rs.EOF
        If Me.LstAccounts.SelectedItem.Text = rs("Name") Then
            Me.txtName.Text = Trim(rs("Name"))
            
                Accountid = rs("AccID")
                Me.txtAddress.Text = Trim(rs("Address"))
                Me.txtPhoneNumber.Text = Trim(rs("PhoneNumber"))
                Me.txtFax.Text = Trim(rs("Fax"))
                Me.txtEmail.Text = Trim(rs("Email"))
                Me.txtCreditLimit = Trim(rs("CreditLimit"))
                Me.dtpDate = Trim(rs("AccDate"))
                
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
   Me.txtCreditLimit.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtCreditLimit_KeyPress(KeyAscii As Integer)
Dim strk1 As String



strk1 = "0123456789."
If KeyAscii = vbKeyReturn Then
   Me.cmdSave.SetFocus
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
Private Function Generate_AccountID(Account_ID As String) As Boolean
Dim strg As String, strg1 As String, strg2 As String, bFlag As Boolean
On Error GoTo SaveError
rs.Open "Select AccID From AccountHolders  order by AccID Desc", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
      strg1 = Trim(rs.Fields!AccID)

   strg1 = Trim(Str(Val(strg1) + 1))
   strg1 = String$(5 - Len(strg1), "0") & strg1
Else
   strg1 = "00001"
End If
Account_ID = strg1

If rs.State = 1 Then rs.Close

Generate_AccountID = True

Exit Function
SaveError:
     If rs.State = 1 Then rs.Close
    Generate_AccountID = False
     Exit Function
     
End Function
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
End Sub
Private Sub ListAccounts()
On Error GoTo SaveError
Me.LstAccounts.ListItems.Clear
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


rs.Open "Select * From AccountHolders Order By Name", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.LstAccounts.ListItems.Add(, , Trim(rs!Name))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!PhoneNumber)
                List_Item.SubItems(2) = Trim(rs!Address)
                List_Item.SubItems(3) = Trim(rs!Email)
                List_Item.SubItems(4) = Trim(rs!Fax)
                List_Item.SubItems(5) = Trim(rs!CreditLimit)
                List_Item.SubItems(6) = Trim(rs!AccDate)
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

