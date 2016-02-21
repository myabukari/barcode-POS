VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmsuppliers 
   BackColor       =   &H00C29E21&
   Caption         =   "SUPPLIERS"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   Icon            =   "frmsuppliers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9555
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   9255
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   480
         TabIndex        =   17
         Top             =   120
         Width           =   8175
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   6480
            TabIndex        =   20
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
            Image           =   "frmsuppliers.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   240
            TabIndex        =   21
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
            Image           =   "frmsuppliers.frx":075C
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdDelete 
            Height          =   375
            Left            =   1800
            TabIndex        =   22
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Delete"
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
            Image           =   "frmsuppliers.frx":0BAE
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdFind 
            Height          =   375
            Left            =   3360
            TabIndex        =   23
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
            Image           =   "frmsuppliers.frx":0D48
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   4920
            TabIndex        =   24
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
            Image           =   "frmsuppliers.frx":119A
            cBack           =   -2147483633
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
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
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
            Height          =   375
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
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
            TabIndex        =   9
            Top             =   240
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
            TabIndex        =   8
            Top             =   240
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
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   9255
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgSuppliers 
         Height          =   2055
         Left            =   720
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3625
         _Version        =   393216
         BackColor       =   16117969
         ForeColor       =   -2147483641
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         GridColor       =   -2147483634
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"frmsuppliers.frx":2A4B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComctlLib.ListView lstSuppliers 
         Height          =   2175
         Left            =   360
         TabIndex        =   18
         Top             =   1920
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483628
         BackColor       =   0
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Supplier Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Phone #"
            Object.Width           =   3616
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Postal  Address"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Email"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fax Number"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtphonenumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtsuppliername 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   330
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Postal  Address:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4680
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C29E21&
         Caption         =   "Phone Number:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Fax Number:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "Supplier Name:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmsuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset, ctrl As Control, CreditorAccid As String
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Supplierid As String, sflag As Boolean

Private Sub cmdClear_Click()
     Me.txtSupplierName = "": Me.txtAddress = "": Me.txtEmail = "": Me.txtFax = ""
     Me.txtPhoneNumber = ""
     Me.txtSupplierName.SetFocus
     sflag = False
End Sub

Private Sub cmdDelete_Click()
Dim X As Integer, CreditorAccid() As String, rec As Integer

If Me.txtSupplierName = "" Then
MsgBox "Please Specify Supplier to be deleted", vbInformation, ""
Me.txtSupplierName.SetFocus
Exit Sub
End If
If MsgBox("ARE YOU SURE  YOU WANT TO DELETE THE SUPPLIERS DETAILS?", vbYesNo + vbQuestion, "CONFIRM DELETE") = vbYes Then
   
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
   
   rs.Open "Select CreditorAccID From CreditorAccount  Where SupplierID='" & Supplierid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
   rec = rs.RecordCount - 1
   ReDim CreditorAccid(rs.RecordCount - 1)
   For X = 0 To rs.RecordCount - 1
   CreditorAccid(X) = rs.Fields("CreditorAccID")
   rs.MoveNext
   Next
   If rs.State = 1 Then rs.Close
   Else
    ReDim CreditorAccid(0)
    If rs.State = 1 Then rs.Close
   End If
   If rs.State = 1 Then rs.Close

   
   cn.BeginTrans
   cn.Execute "Delete From Suppliers Where SupplierID ='" & Supplierid & "'", Y
   If Y > 0 Then
     rs.Open "Select * From CreditorAccount  Where SupplierID='" & Supplierid & "'", cn, adOpenForwardOnly, adLockReadOnly
      If rs.RecordCount > 0 Then
       cn.Execute "Delete From CreditorAccount Where SupplierID ='" & Supplierid & "'", Y
       If rs.State = 1 Then rs.Close
      End If
       If rs.State = 1 Then rs.Close
   End If
   If Y > 0 Then
     For X = 0 To rec
      rs.Open "Select * From CreditorPayment  Where CreditorAccID='" & CreditorAccid(X) & "'", cn, adOpenForwardOnly, adLockReadOnly
       If rs.RecordCount > 0 Then
        cn.Execute "Delete From CreditorPayment Where CreditorAccID ='" & CreditorAccid(X) & "'", Y
        If rs.State = 1 Then rs.Close
       End If
        If rs.State = 1 Then rs.Close
      Next
   End If
   If Y > 0 Then
   cn.CommitTrans
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     
     ClearCtrls
     Me.txtSupplierName.SetFocus
   Else
   cn.RollbackTrans
      MsgBox "Sorry, Unable to Delete Products Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
   Me.txtSupplierName.SetFocus
   sflag = False
   Me.cmdDelete.Enabled = False
   If cn.State = 1 Then cn.Close
   Me.MousePointer = vbDefault
   Call ListSuppliers
   
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
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdFind_Click()
Dim strg1 As String
On Error GoTo SaveError

strg1 = InputBox("Enter The SupplierName Or The First Few or All of the Characters of The SupplierName.", "SupplierName")
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
   rs.Open "Select Suppliers.* From Suppliers  Where SupplierName Like '" & strg1 & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount <= 0 Then
      rs.Close: cn.Close
      MsgBox "There is No Suppliers with the Name Entered.", vbInformation, "Search Failed"
      Me.cmdFind.Enabled = True: Me.cmdFind.SetFocus: Exit Sub
   Else
      If rs.RecordCount = 1 Then
         sflag = True
         Me.txtSupplierName = rs.Fields("SupplierName")
         Me.txtPhoneNumber = rs.Fields("PhoneNo")
         Me.txtEmail = rs.Fields("Email")
         Me.txtAddress = rs.Fields("Address")
         Me.txtFax = rs.Fields("FaxNo")
         Supplierid = rs.Fields("SupplierID")
         
         rs.Close
         Me.txtSupplierName.SetFocus
         Me.cmdDelete.Enabled = True
         
    Else
         rs.MoveFirst
         Me.flxgSuppliers.Rows = rs.RecordCount + 1
         For X = 1 To rs.RecordCount
           Me.flxgSuppliers.TextMatrix(X, 0) = rs.Fields("SupplierName")
           Me.flxgSuppliers.TextMatrix(X, 1) = rs.Fields("PhoneNo")
           Me.flxgSuppliers.TextMatrix(X, 2) = rs.Fields("Email")
           Me.flxgSuppliers.TextMatrix(X, 3) = rs.Fields("Address")
           Me.flxgSuppliers.TextMatrix(X, 4) = rs.Fields("FaxNo")
           Me.flxgSuppliers.TextMatrix(X, 5) = rs.Fields("SupplierID")
           
           rs.MoveNext
         Next
         Me.flxgSuppliers.Visible = True
         Me.flxgSuppliers.SetFocus
         rs.Close
      End If
   End If
End If

If cn.State = 1 Then cn.Close
If rs.State = 1 Then rs.Close


Me.cmdFind.Enabled = True
Me.cmdSave.Enabled = True
Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     
     Me.cmdFind.Enabled = True
     MsgBox "Sorry, Unable to Find Suppliers Details:Please Try Again!", vbInformation, "Search Failed"
     Exit Sub
End Sub


Private Sub cmdSave_Click()
 
If Trim(Me.txtSupplierName) = "" Then
   MsgBox "You Must Enter Product Name.", vbInformation, "Product Name"
   Me.txtSupplierName.SetFocus: Exit Sub
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
rs.Open "Select SupplierName From Suppliers Where SupplierName ='" & Trim(Me.txtSupplierName) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "A Supplier Has Already Been Setup with the Name.", vbInformation
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtSupplierName.SetFocus: Exit Sub
   
End If
rs.Close

   Call Generate_SupplierID(Supplierid)
   cn.Execute "Insert Into Suppliers ([SupplierID],[SupplierName],[PhoneNo],[Email],[Address],[FaxNo]) select '" & Trim(Supplierid) & "','" & Trim(Me.txtSupplierName.Text) & "','" & Trim(Me.txtPhoneNumber.Text) & "','" & Trim(Me.txtEmail.Text) & "','" & Trim(Me.txtAddress.Text) & "','" & Trim(Me.txtFax.Text) & "'", Y
   If Y > 0 Then

     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
       'Call ClearCtrls
       Call ListSuppliers
     Me.txtSupplierName = "": Me.txtAddress = "": Me.txtEmail = "": Me.txtFax = ""
     Me.txtPhoneNumber = ""
     Me.txtSupplierName.SetFocus
   Else
 
   MsgBox "Sorry, Unable to Save Supplier Details:Please Try Again!", vbInformation, "Save Failed"
   Me.txtSupplierName.SetFocus
   End If
   Else
'edit part

   rs.Open "Select SupplierName From Suppliers Where SupplierName ='" & Trim(Me.txtSupplierName) & "' and SupplierID<>'" & Supplierid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.Close: cn.Close
      MsgBox "A Supplier Has Already Been Setup with the Name.", vbInformation
      Me.txtSupplierName.SetFocus: Exit Sub
   End If
   rs.Close
   cn.Execute "Update Suppliers Set SupplierName ='" & Trim(Me.txtSupplierName.Text) & "',PhoneNo='" & Trim(Me.txtPhoneNumber.Text) & "',Email='" & Trim(Me.txtEmail.Text) & "',Address='" & Trim(Me.txtAddress.Text) & "',FaxNo='" & Trim(Me.txtFax.Text) & "' Where SupplierID ='" & Supplierid & "'", Y
   
   If Y > 0 Then
     
      MsgBox "Edit Successful!", vbInformation, "Edit Successful"
     'Clear ctrls and setfocus to Clinic ctrl
      'Call ClearCtrls
      Call ListSuppliers
      Me.txtSupplierName = "": Me.txtAddress = "": Me.txtEmail = "": Me.txtFax = ""
     Me.txtPhoneNumber = ""
     Me.txtSupplierName.SetFocus
   Else
      MsgBox "Sorry, Unable to Edit Suppliers Details:Please Try Again!", vbInformation, "Edit Failed"
      Me.txtSupplierName.SetFocus
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
     MsgBox "Sorry, Unable to Save Suppliers Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub


End Sub



Private Sub flxgSuppliers_Click()
Me.txtSupplierName = Me.flxgSuppliers.TextMatrix(Me.flxgSuppliers.Row, 0)
Me.txtPhoneNumber = Me.flxgSuppliers.TextMatrix(Me.flxgSuppliers.Row, 1)
Me.txtEmail = Me.flxgSuppliers.TextMatrix(Me.flxgSuppliers.Row, 2)
Me.txtFax = Me.flxgSuppliers.TextMatrix(Me.flxgSuppliers.Row, 3)
Me.txtAddress = Me.flxgSuppliers.TextMatrix(Me.flxgSuppliers.Row, 4)
Supplierid = Me.flxgSuppliers.TextMatrix(Me.flxgSuppliers.Row, 5)
Me.flxgSuppliers.Visible = False
Me.cmdDelete.Enabled = True
Me.cmdSave.Enabled = True
sflag = True
Me.txtSupplierName.SetFocus
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Width = 9675
Me.Height = 6270
Call ListSuppliers
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub lstSuppliers_Click()
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
rs.Open "Select Suppliers.* From Suppliers Order By SupplierName", cn, adOpenForwardOnly, adLockReadOnly

    While Not rs.EOF
        If Me.lstSuppliers.SelectedItem.Text = rs("SupplierName") Then
               Me.txtSupplierName.Text = Trim(rs("SupplierName"))
            
                Supplierid = rs("SupplierID")
                Me.txtPhoneNumber.Text = Trim(rs("PhoneNo"))
                Me.txtEmail.Text = Trim(rs("Email"))
                Me.txtAddress.Text = Trim(rs("Address"))
                Me.txtFax.Text = Trim(rs("FaxNo"))
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
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Display Suppliers  Details:Please Try Again!", vbInformation, "Suppliers Details"
        Exit Sub

End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
KeyAscii = KeyAscii - 32
Exit Sub
End If
strk1 = "0123456789/|\;:.,@%$()"
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

Private Sub txtCreditLimit_KeyPress(KeyAscii As Integer)
Dim strk1 As String

strk1 = "0123456789/|\;:.,@%$()"
If KeyAscii = vbKeyReturn Then
   Me.txtAddress.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
'Dim strk1 As String
'If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then
'KeyAscii = KeyAscii - 32
'Exit Sub
 'Else
'If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then Exit Sub
'End If
'strk1 = "0123456789/|\;:.,@%$()"
'If KeyAscii = vbKeyReturn Then
   'Me.txtFax.SetFocus
'End If
'If KeyAscii > 26 Then
   'If KeyAscii <> 32 Then
      'If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         'KeyAscii = 0
      'End If
   'End If
'End If
End Sub

Private Sub txtphonenumber_KeyPress(KeyAscii As Integer)
Dim strk1 As String

strk1 = "0123456789/|\;:.,_-()"
If KeyAscii = vbKeyReturn Then
   Me.txtAddress.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtsuppliername_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()-_"
If KeyAscii = vbKeyReturn Then
   Me.txtPhoneNumber.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub
Private Function Generate_SupplierID(Supplier_ID As String) As Boolean
Dim strg As String, strg1 As String, strg2 As String, bFlag As Boolean
On Error GoTo SaveError
rs.Open "Select SupplierID From Suppliers  order by SupplierID Desc", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
      strg1 = Trim(rs.Fields!Supplierid)

   strg1 = Trim(Str(Val(strg1) + 1))
   strg1 = String$(5 - Len(strg1), "0") & strg1
Else
   strg1 = "00001"
End If
Supplier_ID = strg1

If rs.State = 1 Then rs.Close

Generate_SupplierID = True

Exit Function
SaveError:
     If rs.State = 1 Then rs.Close
    Generate_SupplierID = False
     Exit Function
     
End Function
Private Sub ListSuppliers()
On Error GoTo SaveError
Me.lstSuppliers.ListItems.Clear


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


rs.Open "Select Suppliers.* From Suppliers Order By SupplierName", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.lstSuppliers.ListItems.Add(, , Trim(rs!SupplierName))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
                'List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!PhoneNo)
                List_Item.SubItems(2) = Trim(rs!Address)
                List_Item.SubItems(3) = Trim(rs!Email)
                List_Item.SubItems(4) = Trim(rs!FaxNo)
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
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Display Suppliers  Details:Please Try Again!", vbInformation, "Suppliers Details"
        Exit Sub
End Sub
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
End Sub

