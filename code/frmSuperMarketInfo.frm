VERSION 5.00
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form2"
   ScaleHeight     =   5190
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   4080
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin lvButton_H.lvButtons_H cmdSave 
         Height          =   375
         Left            =   120
         TabIndex        =   17
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
         Image           =   "frmSuperMarketInfo.frx":0000
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdDelete 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
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
         Image           =   "frmSuperMarketInfo.frx":0452
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdFind 
         Height          =   375
         Left            =   3000
         TabIndex        =   19
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
         Image           =   "frmSuperMarketInfo.frx":05EC
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdClear 
         Height          =   375
         Left            =   4440
         TabIndex        =   20
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
         Image           =   "frmSuperMarketInfo.frx":0A3E
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   5880
         TabIndex        =   21
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
         Image           =   "frmSuperMarketInfo.frx":22EF
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   7335
      Begin VB.TextBox txtCellPhone 
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
         Left            =   2160
         TabIndex        =   2
         Top             =   1440
         Width           =   3255
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
         Left            =   2160
         TabIndex        =   4
         Top             =   2640
         Width           =   4695
      End
      Begin VB.TextBox txtPhoneNo 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   3255
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
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   4695
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
         Height          =   555
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "CellPhone Number:"
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
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Email Address:"
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
         Left            =   600
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Postal Address:"
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
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "SuperMarket's Name:"
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
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "Phone Number:"
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
         Left            =   600
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset, ctrl As Control, CreditorAccid As String
Dim bFlag As Boolean, strg As String, List_Item As ListItem, SuperMarketid As Integer, sflag As Boolean

Private Sub Text1_Change()

End Sub

Private Sub cmdClear_Click()
Call ClearCtrls
End Sub

Private Sub cmdDelete_Click()
If Me.txtName = "" Then
MsgBox "ENTER NAME TO BE DELETED", vbInformation, "DELETE WHAT"
Me.txtName.SetFocus: Exit Sub
End If
If MsgBox("ARE YOU SURE  YOU WANT TO DELETE THE SUPERMARKET'S DETAILS?", vbYesNo + vbQuestion, "CONFIRM DELETE") = vbYes Then
   
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
   
    cn.Execute "Delete From SuperMarketInfo Where SuperMarketID ='" & SuperMarketid & "'", Y
    
    If Y > 0 Then
    
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     
     
     ClearCtrls
     Me.txtName.SetFocus
   Else
      
      MsgBox "Sorry, Unable to Delete SuperMarket's Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
    Me.txtName.SetFocus
   sflag = False
   Me.cmdDelete.Enabled = False
   If cn.State = 1 Then cn.Close
   
   Me.MousePointer = vbDefault
   
   
   Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Delete Products  Details:Please Try Again!", vbInformation, "Delete Failed"
        Exit Sub
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim strg1 As String
On Error GoTo SaveError

strg1 = InputBox("Enter SuperMarket's name Or The First Few or All of the Characters of The SuperMarket's name.", "SuperMarket's name")
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
   rs.Open "Select * From SuperMarketInfo  Where SuperMarketName Like '" & strg1 & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount <= 0 Then
      rs.Close: cn.Close
      MsgBox "There is no information with the Name Entered.", vbInformation, "Search Failed"
      Me.cmdFind.Enabled = True: Me.cmdFind.SetFocus: Exit Sub
   Else
      If rs.RecordCount = 1 Then
         sflag = True
         Me.txtName = rs.Fields("SuperMarketName")
         Me.txtPhoneNo = rs.Fields("PhoneNo")
         Me.txtCellPhone = rs.Fields("CellPhoneNo")
         Me.txtAddress = rs.Fields("PostalAddress")
         Me.txtEmail = rs.Fields("Email")
         SuperMarketid = rs.Fields("SuperMarketID")
         
         rs.Close
         Me.txtName.SetFocus
         Me.cmdDelete.Enabled = True

'    Else
'         rs.MoveFirst
'         Me.flxgSuppliers.Rows = rs.RecordCount + 1
'         For X = 1 To rs.RecordCount
'           Me.flxgSuppliers.TextMatrix(X, 0) = rs.Fields("SupplierName")
'           Me.flxgSuppliers.TextMatrix(X, 1) = rs.Fields("PhoneNo")
'           Me.flxgSuppliers.TextMatrix(X, 2) = rs.Fields("Email")
'           Me.flxgSuppliers.TextMatrix(X, 3) = rs.Fields("Address")
'           Me.flxgSuppliers.TextMatrix(X, 4) = rs.Fields("FaxNo")
'           Me.flxgSuppliers.TextMatrix(X, 5) = rs.Fields("SupplierID")
'
'           rs.MoveNext
'         Next
'         Me.flxgSuppliers.Visible = True
'         Me.flxgSuppliers.SetFocus
'         rs.Close
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
     MsgBox "Sorry, Unable to Find SuperMarket's Details:Please Try Again!", vbInformation, "Search Failed"
     Exit Sub
End Sub

Private Sub cmdSave_Click()
 
If Trim(Me.txtName) = "" Then
   MsgBox "You Must Enter Name of SuperMarket.", vbInformation, "SuperMarket's Name"
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
rs.Open "Select * From SuperMarketInfo ", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "Name Has Already Been Setup for the SuperMarket,You can only change name.", vbInformation, ""
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtName.SetFocus: Exit Sub
   
End If
rs.Close

   
   cn.Execute "Insert Into SuperMarketInfo ([SuperMarketName],[PhoneNo],[CellPhoneNo],[PostalAddress],[Email]) select '" & Trim(Me.txtName) & "','" & Trim(Me.txtPhoneNo) & "','" & Trim(Me.txtCellPhone) & "','" & Trim(Me.txtAddress) & "','" & Trim(Me.txtEmail) & "'", Y
   If Y > 0 Then

     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
     Me.txtName = "": Me.txtPhoneNo = "": Me.txtCellPhone = "": Me.txtAddress = ""
     Me.txtEmail = ""
     Me.txtName.SetFocus
   Else
 
   MsgBox "Sorry, Unable to Save SuperMarket's Details:Please Try Again!", vbInformation, "Save Failed"
   Me.txtName.SetFocus
   End If
   Else
'edit part

   rs.Open "Select SuperMarketName From SuperMarketInfo Where SuperMarketName ='" & Trim(Me.txtName) & "' and SuperMarketID<>'" & SuperMarketid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.Close: cn.Close
      MsgBox "Name Has Already Been Setup for the SuperMarket,You can only change name.", vbInformation, ""
      Me.txtName.SetFocus: Exit Sub
   End If
   rs.Close
   cn.Execute "Update SuperMarketInfo Set SuperMarketName ='" & Trim(Me.txtName) & "',PhoneNo='" & Trim(Me.txtPhoneNo) & "',CellPhoneNo='" & Trim(Me.txtCellPhone) & "',PostalAddress='" & Trim(Me.txtAddress) & "',Email='" & Trim(Me.txtEmail) & "' Where SuperMarketID ='" & SuperMarketid & "'", Y
   
   If Y > 0 Then
     
     MsgBox "Edit Successful!", vbInformation, "Edit Successful"
     Me.txtName = "": Me.txtPhoneNo = "": Me.txtCellPhone = "": Me.txtAddress = ""
     Me.txtEmail = ""
     Me.txtName.SetFocus
   Else
      MsgBox "Sorry, Unable to Edit SuperMarket's Details:Please Try Again!", vbInformation, "Edit Failed"
      Me.txtName.SetFocus
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
     MsgBox "Sorry, Unable to Save SuperMarket's Details:Please Try Again!", vbInformation, "Save Failed"
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
   Me.txtEmail.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtCellPhone_KeyPress(KeyAscii As Integer)
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

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtPhoneNo.SetFocus
End If
End Sub

Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
Dim strk1 As String

strk1 = "0123456789/|\;:.,_-()"
If KeyAscii = vbKeyReturn Then
   Me.txtCellPhone.SetFocus
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



