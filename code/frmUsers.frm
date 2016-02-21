VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmUsers 
   BackColor       =   &H00C29E21&
   Caption         =   "USERS"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   9435
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   240
      TabIndex        =   18
      Top             =   2520
      Width           =   8895
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   8895
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   7440
            TabIndex        =   22
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
            Image           =   "frmUsers.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   120
            TabIndex        =   23
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
            Image           =   "frmUsers.frx":075C
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdDelete 
            Height          =   375
            Left            =   1560
            TabIndex        =   24
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
            Image           =   "frmUsers.frx":0BAE
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdFind 
            Height          =   375
            Left            =   3120
            TabIndex        =   25
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
            Image           =   "frmUsers.frx":0D48
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   4560
            TabIndex        =   26
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
            Image           =   "frmUsers.frx":119A
            cBack           =   -2147483633
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
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   7
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
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1215
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
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1215
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
            Left            =   8880
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1215
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
            Left            =   8880
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin lvButton_H.lvButtons_H Command1 
            Height          =   375
            Left            =   6000
            TabIndex        =   27
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Users Info"
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
            Image           =   "frmUsers.frx":2A4B
            cBack           =   -2147483633
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Users Info"
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
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   8895
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgUsers 
         Height          =   1455
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2566
         _Version        =   393216
         BackColor       =   16117969
         ForeColor       =   -2147483630
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorSel    =   -2147483647
         BackColorBkg    =   12754465
         GridColor       =   16777215
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   "<UserName                     |<UserID            |<AccessLevel             |<Password       |<Status Of Account    |UserNo"
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
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.ComboBox cboAccessLevel 
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmUsers.frx":42FC
         Left            =   1560
         List            =   "frmUsers.frx":430F
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cboStatusAccount 
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmUsers.frx":4343
         Left            =   6480
         List            =   "frmUsers.frx":434D
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtUserID 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtConfirmPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6480
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F0D1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6480
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C29E21&
         Caption         =   "Access Level:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Status Of Accounts:"
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
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Confirm Password:"
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
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Password:"
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
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "UserID:"
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
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "UserName:"
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
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String
Dim sflag As Boolean, UserNo As Integer, ctrl As Control
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub cboAccessLevel_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
If KeyAscii = vbKeyReturn Then
   Me.txtPassword.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub cboStatusAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   Me.cmdSave.SetFocus
End If
End Sub

Private Sub cmdClear_Click()
 Call ClearCtrls
 Me.txtUserName.SetFocus
 sflag = False
 
End Sub

Private Sub cmdDelete_Click()
If MsgBox("ARE YOU SURE  YOU WANT TO DELETE THE USERS DETAILS?", vbYesNo + vbQuestion, "CONFIRM DELETE") = vbYes Then
   If Me.txtUserID = "" Then
     MsgBox "ENTER USER ID TO BE DELETED", vbInformation, "DELETE"
     Exit Sub
   End If
   
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
   
   
   rs.Open "Select Users.* From Users ", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount = 1 And rs.Fields("AccessLevel") = "MANAGER" Then
     MsgBox "CREATE NEW ACCOUNT BEFORE DELETING", vbInformation, "DELETE"
     rs.Close: Exit Sub
   End If
     
    
   cn.Execute "Delete From Users Where UserNo =" & UserNo, Y
    If Y > 0 Then
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     'Clear ctrls and setfocus to supplier name ctrl
     Call ClearCtrls
     Me.txtUserName.SetFocus
   Else
      MsgBox "Sorry, Unable to Delete Users Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
    Me.txtUserName.SetFocus
   sflag = False
   Me.cmdDelete.Enabled = False
   If cn.State = 1 Then cn.Close
   
  
    Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Delete Users  Details:Please Try Again!", vbInformation, "Delete Failed"
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

strg1 = InputBox("Enter The ProductName Or The First Few or All of the Characters of The ProductName.", "ProductName")
If Trim(strg1) <> "" Then
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
   rs.Open "Select Users.* From Users Where UserName Like '" & strg1 & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount <= 0 Then
      rs.Close: cn.Close
      MsgBox "There is no User with the Name Entered.", vbInformation, "Search Failed"
      Me.cmdFind.Enabled = True: Me.cmdFind.SetFocus: Exit Sub

   Else
      If rs.RecordCount = 1 Then
         sflag = True
         Me.txtUserName = rs.Fields("UserName")
         Me.txtUserID = rs.Fields("UserID")
         Me.txtPassword = rs.Fields("UserPassword")
         Me.cboStatusAccount = rs.Fields("StatusOfAccount")
         UserNo = rs.Fields("UserNo")
         Me.cboAccessLevel = rs.Fields("AccessLevel")
         
         
         
         
         rs.Close
        ' rs.Open "Select Description From Clinics Where ClinicID =" & ClinicID, cn, adOpenForwardOnly, adLockReadOnly
         'If rs.RecordCount > 0 Then
            'Me.cboClinic = rs.Fields("Description")
         'End If
         
         Me.txtUserName.SetFocus
         Me.cmdDelete.Enabled = True
         
Else
         rs.MoveFirst
         Me.flxgUsers.Rows = rs.RecordCount + 1
         For X = 1 To rs.RecordCount
           Me.flxgUsers.TextMatrix(X, 0) = rs.Fields("UserName")
           Me.flxgUsers.TextMatrix(X, 1) = rs.Fields("UserID")
           Me.flxgUsers.TextMatrix(X, 2) = rs.Fields("AccessLevel")
           Me.flxgUsers.TextMatrix(X, 3) = rs.Fields("UserPassword")
           Me.flxgUsers.TextMatrix(X, 4) = rs.Fields("StatusOfAccount")
           Me.flxgUsers.TextMatrix(X, 5) = rs.Fields("UserNo")
           
           rs.MoveNext
         Next
         Me.flxgUsers.Visible = True
         Me.flxgUsers.SetFocus
         rs.Close
      End If
   End If
End If
If cn.State = 1 Then cn.Close
If rs.State = 1 Then rs.Close
Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     Me.cmdFind.Enabled = True
     MsgBox "Sorry,  Try Again!", vbInformation, "Search Failed"
     Exit Sub
End Sub

Private Sub cmdSave_Click()
If Me.txtUserName.Text = "" Then
  MsgBox "You must Enter UserName", vbInformation, "UserName"
  Me.txtUserName.SetFocus: Exit Sub
End If
If Me.txtUserID.Text = "" Then
  MsgBox "You must Enter UserID", vbInformation, "UserID"
  Me.txtUserID.SetFocus: Exit Sub
End If
  If Me.txtPassword.Text = "" Then
  MsgBox "You must Enter Password", vbInformation, "Password"
  Me.txtPassword.SetFocus: Exit Sub
End If
If sflag = False Then
If Me.txtConfirmPassword.Text = "" Then
  MsgBox "You must ConfirmPassword", vbInformation, "ConfirmPassword"
  Me.txtConfirmPassword.SetFocus: Exit Sub

End If
If Me.txtConfirmPassword.Text <> Me.txtPassword.Text Then
  MsgBox "ConfirmPassword Does not match Password", vbInformation, "ConfirmPassword"
  Me.txtConfirmPassword.SetFocus: Exit Sub
End If
End If
If Me.cboStatusAccount.Text = "" Then
  MsgBox "You must Enter StatusOfAccount", vbInformation, "StatusOfAccount"
  Me.cboStatusAccount.SetFocus: Exit Sub
End If
If Me.cboAccessLevel.Text = "" Then
  MsgBox "You must Enter AccessLevel", vbInformation, "AccessLevel"
  Me.cboAccessLevel.SetFocus: Exit Sub
End If

  
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

If sflag = False Then
   'save part
rs.Open "Select UserID From Users Where UserID ='" & Trim(Me.txtUserID) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "A User Has Already Been Setup with the UserID.", vbInformation
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtUserName.SetFocus: Exit Sub
   
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
End If
rs.Close

cn.Execute "Insert Into Users ([UserName],[UserID],[AccessLevel],[UserPassword],[StatusOfAccount],[Date]) select '" & Trim(Me.txtUserName.Text) & "','" & Trim(Me.txtUserID.Text) & "','" & Trim(Me.cboAccessLevel.Text) & "','" & Trim(Me.txtPassword.Text) & "','" & Trim(Me.cboStatusAccount.Text) & "','" & Date & "'", Y
If Y > 0 Then
     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
     Call ClearCtrls
Else
   MsgBox "Sorry, Unable to Save User Details:Please Try Again!", vbInformation, "Save Failed"
   Me.txtUserName.SetFocus
End If

Else
'edit part

   rs.Open "Select UserName From Users Where UserID ='" & Trim(Me.txtUserID) & "' and UserNo<>'" & UserNo & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.Close: cn.Close
      MsgBox "A USER HAS ALREADY BEEN SETUP WITH THE USERID.", vbInformation
      Me.txtUserID.SetFocus: Exit Sub
   End If
   rs.Close
   
   
   cn.Execute "Update Users Set UserName ='" & Trim(Me.txtUserName.Text) & "',UserID='" & Trim(Me.txtUserID.Text) & "',AccessLevel='" & Trim(Me.cboAccessLevel.Text) & "',StatusOfAccount='" & Trim(Me.cboStatusAccount.Text) & "',UserPassword='" & Trim(Me.txtPassword.Text) & "' Where UserNo =" & UserNo, Y
   If Y > 0 Then
      MsgBox "Edit Successful!", vbInformation, "Edit Successful"
     'Clear ctrls and setfocus to Clinic ctrl
      Call ClearCtrls
      Me.txtUserName.SetFocus
   Else
      MsgBox "SORRY, UNABLE TO EDIT USERS DETAILS:PLEASE TRY AGAIN!", vbInformation, "EDIT FAILED"
   End If
End If

sflag = False
If cn.State = 1 Then cn.Close
If rs.State = 1 Then rs.Close

Me.MousePointer = vbDefault
Me.cmdSave.Enabled = True
Me.txtUserName.SetFocus

Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY, UNABLE TO SAVE USERS DETAILS:PLEASE TRY AGAIN!", vbInformation, "SAVE FAILED"
     Exit Sub


End Sub

Private Sub Command1_Click()
frmUsersInfo.Show

End Sub

Private Sub flxgUsers_Click()
Me.txtUserName = Me.flxgUsers.TextMatrix(Me.flxgUsers.Row, 0)
Me.txtUserID = Me.flxgUsers.TextMatrix(Me.flxgUsers.Row, 1)
Me.cboAccessLevel = Me.flxgUsers.TextMatrix(Me.flxgUsers.Row, 2)
Me.txtPassword = Me.flxgUsers.TextMatrix(Me.flxgUsers.Row, 3)
Me.cboStatusAccount = Me.flxgUsers.TextMatrix(Me.flxgUsers.Row, 4)
UserNo = Me.flxgUsers.TextMatrix(Me.flxgUsers.Row, 5)


Me.flxgUsers.Visible = False
Me.cmdDelete.Enabled = True
Me.cmdSave.Enabled = True
sflag = True
Me.txtUserName.SetFocus
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Width = 9555
Me.Height = 4350
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub txtConfirmPassword_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
If KeyAscii = vbKeyReturn Then
   Me.cboStatusAccount.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
If KeyAscii = vbKeyReturn Then
   Me.txtConfirmPassword.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
If KeyAscii = vbKeyReturn Then
   Me.cboAccessLevel.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
If KeyAscii = vbKeyReturn Then
   Me.txtUserID.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
sflag = False
End Sub

