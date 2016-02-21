VERSION 5.00
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmChangePassword 
   BackColor       =   &H00C29E21&
   Caption         =   "Change UserID orPassword "
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8640
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   8295
      Begin lvButton_H.lvButtons_H cmdSave 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Caption         =   "&Save Changes"
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
         Image           =   "frmChangePassword.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   5520
         TabIndex        =   15
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
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
         Image           =   "frmChangePassword.frx":075C
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8295
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Caption         =   "Input Changes Below"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   0
         TabIndex        =   8
         Top             =   1680
         Width           =   8295
         Begin VB.TextBox txtConfirmChangePassword 
            Appearance      =   0  'Flat
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
            IMEMode         =   3  'DISABLE
            Left            =   5520
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtChangePassword 
            Appearance      =   0  'Flat
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
            IMEMode         =   3  'DISABLE
            Left            =   5520
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtChangeUserID 
            Appearance      =   0  'Flat
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
            Left            =   960
            TabIndex        =   2
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lblConfirmPassword 
            BackColor       =   &H00C29E21&
            Caption         =   "Confirm Password:"
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
            Left            =   3480
            TabIndex        =   11
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblPassword 
            BackColor       =   &H00C29E21&
            Caption         =   "Password:"
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
            Left            =   4320
            TabIndex        =   10
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblUserID 
            BackColor       =   &H00C29E21&
            Caption         =   "UserID:"
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
            TabIndex        =   9
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.TextBox txtUserID 
         Appearance      =   0  'Flat
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
         Left            =   960
         TabIndex        =   0
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   5400
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
      Begin lvButton_H.lvButtons_H cmdOk 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "&Authenticate UserID and Password"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         Image           =   "frmChangePassword.frx":0BAE
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "UserID:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Password:"
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
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, UserNo As Integer

Private Sub cmdExit_Click()
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdOk_Click()
If Trim(Me.txtUserID) = "" Then
   MsgBox "You Must Enter UserID.", vbInformation, "UserID"
   Me.txtUserID.SetFocus: Exit Sub
End If
If Trim(Me.txtPassword) = "" Then
   MsgBox "You Must Enter Password.", vbInformation, "Password"
   Me.txtPassword.SetFocus: Exit Sub
End If
On Error GoTo SaveError


'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
  
   MsgBox strg, vbInformation:
   Exit Sub
End If

rs.Open "Select * from Users Where UserID ='" & Me.txtUserID & "' And UserPassword ='" & txtPassword & "'", cn, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
  If rs.Fields("StatusOfAccount") = "UNBLOCKED" Then
     Me.txtChangeUserID.Enabled = True
     Me.txtChangePassword.Enabled = True
     Me.txtConfirmChangePassword.Enabled = True
     Me.lblUserID.Enabled = True
     Me.lblPassword.Enabled = True
     Me.lblConfirmPassword.Enabled = True
     UserNo = rs.Fields("UserNo")
     Me.txtChangePassword.SetFocus
     
     
  ElseIf rs.Fields("StatusOfAccount") = "BLOCKED" Then
   MsgBox "Sorry, Your Accounts is Currently Blocked - You Cannot Access the System.", vbInformation, "System Access Denied"
   Me.txtUserID.SetFocus
  End If
  
'  Else
'    MsgBox "Incorrect User Name And/Or Password", vbInformation, "Enter The Correct User Name And/Or Password"
'    Me.txtUserID.SetFocus
'  End If
Else
   MsgBox "Incorrect User Name And/Or Password", vbInformation, "Enter The Correct User Name And/Or Password"
    Me.txtUserID.SetFocus
End If
If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   
   Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        MsgBox "SORRY,UNABLE TO LOGIN :PLEASE TRY AGAIN!", vbInformation, "TRY AGAIN"
        Exit Sub
End Sub

Private Sub cmdSave_Click()
If Me.txtChangePassword.Text = "" Then
  MsgBox "YOU MUST ENTER PASSWORD", vbInformation, "ENTER PASSWORD"
  Exit Sub
End If
If Me.txtConfirmChangePassword.Text = "" Then
  MsgBox "YOU MUST CONFIRM PASSWORD", vbInformation, "CONFIRM PASSWORD"
  Me.txtConfirmChangePassword.SetFocus: Exit Sub
End If
  
If Me.txtConfirmChangePassword.Text <> Me.txtChangePassword.Text Then
  MsgBox "CONFIRMED PASSWORD DOES NOT MATCH PASSWORD", vbInformation, "CONFIRM PASSWORD"
  Me.txtChangePassword.SetFocus: Exit Sub
End If
'If Me.cboStatusAccount.Text = "" Then
  'MsgBox "You must Enter StatusOfAccount", vbInformation, "StatusOfAccount"
  'Me.cboStatusAccount.SetFocus: Exit Sub
'End If
  
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


   'save part
rs.Open "Select UserID From Users Where UserID ='" & Trim(Me.txtChangeUserID) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "A USER HAS ALREADY BEEN SETUP WITH THE USERID.", vbInformation
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtChangeUserID.SetFocus: Exit Sub
   
End If
 If rs.State = 1 Then rs.Close

If Trim(Me.txtChangeUserID) <> "" Then

cn.Execute "Update Users Set UserID='" & Trim(Me.txtChangeUserID.Text) & "',UserPassword='" & Trim(Me.txtChangePassword.Text) & "' Where UserNo =" & UserNo, Y
Else
cn.Execute "Update Users Set UserPassword='" & Trim(Me.txtChangePassword.Text) & "' Where UserNo =" & UserNo, Y
End If
   If Y > 0 Then
      MsgBox "EDIT SUCCESSFUL!", vbInformation, "EDIT SUCCESSFUL"
     'Clear ctrls and setfocus to Clinic ctrl
      Call ClearCtrls
       Me.txtChangeUserID.Enabled = False
       Me.txtChangePassword.Enabled = False
       Me.txtConfirmChangePassword.Enabled = False
       Me.lblUserID.Enabled = False
       Me.lblPassword.Enabled = False
       Me.lblConfirmPassword.Enabled = False
        Me.txtUserID.SetFocus
   Else
      MsgBox "SORRY, UNABLE TO EDIT USERS DETAILS:PLEASE TRY AGAIN!", vbInformation, "EDIT FAILED"
   End If
If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   
   Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        MsgBox "SORRY,UNABLE TO EDIT PASSWORD :PLEASE TRY AGAIN!", vbInformation, "TRY AGAIN"
        Exit Sub

End Sub

Private Sub Form_Load()
   Me.txtChangeUserID.Enabled = False
   Me.txtChangePassword.Enabled = False
   Me.txtConfirmChangePassword.Enabled = False
   Me.lblUserID.Enabled = False
   Me.lblPassword.Enabled = False
   Me.lblConfirmPassword.Enabled = False
   
  CenterForm Me
  Me.Height = 5505
  Me.Width = 8760
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2

  
End Sub

Private Sub txtChangePassword_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
If KeyAscii = vbKeyReturn Then
   Me.txtConfirmChangePassword.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtChangeUserID_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
If KeyAscii = vbKeyReturn Then
   Me.txtChangePassword.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtConfirmChangePassword_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
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

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
If KeyAscii = vbKeyReturn Then
   Me.cmdOk.SetFocus
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
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
End Sub
