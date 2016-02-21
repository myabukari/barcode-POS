VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmStores 
   BackColor       =   &H00C29E21&
   Caption         =   "Stores"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   8010
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   7575
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   7335
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   360
            Width           =   1335
         End
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   120
            TabIndex        =   12
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
            Image           =   "frmStores.frx":0000
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdDelete 
            Height          =   375
            Left            =   1560
            TabIndex        =   13
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
            Image           =   "frmStores.frx":0452
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdFind 
            Height          =   375
            Left            =   3000
            TabIndex        =   14
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
            Image           =   "frmStores.frx":05EC
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   4440
            TabIndex        =   15
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
            Image           =   "frmStores.frx":0A3E
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   5880
            TabIndex        =   16
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
            Image           =   "frmStores.frx":22EF
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgStores 
         Height          =   1215
         Left            =   600
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2143
         _Version        =   393216
         BackColor       =   16117969
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   "<StoreNmae                              |<StoreLocation                                                      |<StoreID"
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
         _Band(0).Cols   =   3
      End
      Begin VB.TextBox txtStoreLocation 
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
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   960
         Width           =   4815
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
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "Store Location:"
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
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "StoreName:"
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
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmStores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Storeid As Integer
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, AccTransid As String

Private Sub cmdClear_Click()
Call ClearCtrls


End Sub

Private Sub cmdDelete_Click()
If Me.txtname = "" Then
MsgBox "ENTER STORE TO BE DELETED", vbInformation, "DELETE WHAT"
Me.txtname.SetFocus: Exit Sub
End If
If MsgBox("ARE YOU SURE  YOU WANT TO DELETE THE STORE'S DETAILS?", vbYesNo + vbQuestion, "CONFIRM DELETE") = vbYes Then
   
   Me.cmdDelete.Enabled = False
   
   On Error GoTo SaveError
       
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
   cn.Execute "Delete From Stores Where StoreID ='" & Storeid & "'", Y
    If Y > 0 Then
    
    rs.Open "Select * From WholeSaleInventory  Where StoreID = '" & Storeid & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            cn.Execute "Delete From WholeSaleInventory Where StoreID ='" & Storeid & "'", Y
            If rs.State = 1 Then rs.Close
        Else
            Y = 1
            If rs.State = 1 Then rs.Close
        End If
    End If
    
    If Y > 0 Then
    
     rs.Open "Select * From Receivals  Where StoreID = '" & Storeid & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            cn.Execute "Delete From Receivals Where StoreID ='" & Storeid & "'", Y
            If rs.State = 1 Then rs.Close
        Else
            Y = 1
            If rs.State = 1 Then rs.Close
        End If
            
    End If
    
    If Y > 0 Then
    
    
    rs.Open "Select * From StockingRetail  Where StoreID = '" & Storeid & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            cn.Execute "Delete From StockingRetail Where StoreID ='" & Storeid & "'", Y
            If rs.State = 1 Then rs.Close
        Else
            Y = 1
            If rs.State = 1 Then rs.Close
        End If
    
    End If
    
    If Y > 0 Then
    cn.CommitTrans
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     'Clear ctrls and setfocus to supplier name ctrl
     ClearCtrls
     Me.txtname.SetFocus
   Else
      cn.RollbackTrans
      MsgBox "Sorry, Unable to Delete Store's Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
    Me.txtname.SetFocus
   sflag = False
   Me.cmdDelete.Enabled = False
   If cn.State = 1 Then cn.Close
   
   Me.MousePointer = vbDefault
   
   
   Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Delete Store's Details:Please Try Again!", vbInformation, "Delete Failed"
        Exit Sub
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim strg1 As String
'On Error GoTo SaveError

strg1 = InputBox("Enter Store's name or a few characters of the name.", "Store's name")
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
   rs.Open "Select * From Stores  Where StoreName Like '" & strg1 & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount <= 0 Then
      rs.Close: cn.Close
      MsgBox "THERE IS NO STORE WITH THE NAME ENTERED.", vbInformation, "SEARCH FAILED"
      Me.MousePointer = vbDefault: Me.cmdFind.Enabled = True: Me.cmdFind.SetFocus: Exit Sub
   Else
      If rs.RecordCount = 1 Then
         sflag = True
         Me.txtname = rs.Fields("StoreName")
         Me.txtStoreLocation = rs.Fields("StoreLocation")
         Storeid = rs.Fields("StoreID")
'         Me.dtpDate = rs.Fields("AccountOpenDate")
'         Debtorid = rs.Fields("DebtorID")
         
         If rs.State = 1 Then rs.Close
         
         Me.txtname.SetFocus
         Me.cmdDelete.Enabled = True
        
      Else
         rs.MoveFirst
         Me.flxgStores.Rows = rs.RecordCount + 1
         For X = 1 To rs.RecordCount
           Me.flxgStores.TextMatrix(X, 0) = rs.Fields("StoreName")
           Me.flxgStores.TextMatrix(X, 1) = rs.Fields("StoreLocation")
           Me.flxgStores.TextMatrix(X, 2) = rs.Fields("StoreID")
'           Me.flxgStores.TextMatrix(X, 3) = rs.Fields("Address")
'           Me.flxgStores.TextMatrix(X, 4) = rs.Fields("DebtorID")
           rs.MoveNext
         Next
         Me.flxgStores.Visible = True
         Me.flxgStores.SetFocus
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
     MsgBox "Sorry, Unable to Find Stores Details:Please Try Again!", vbInformation, "Search Failed"
     Exit Sub
End Sub

Private Sub cmdSave_Click()

If Trim(Me.txtname) = "" Then
   MsgBox "YOU MUST ENTER STORE'S NAME.", vbInformation, "STORE'S NAME"
   Me.txtname.SetFocus: Exit Sub
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
rs.Open "Select StoreName From Stores Where StoreName ='" & Trim(Me.txtname) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "A STORE HAS ALREADY BEEN SETUP WITH THE NAME ENTERED.", vbInformation, "NAME ALREADY EXIST"
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtname.SetFocus: Exit Sub
   
End If
   If rs.State = 1 Then rs.Close

   'Call Generate_AccountID(Accountid)
   
   cn.Execute "Insert Into Stores ([StoreName],[StoreLocation]) select '" & Trim(Me.txtname.Text) & "','" & Trim(Me.txtStoreLocation) & "'", Y
   If Y > 0 Then
     MsgBox "Saved Successfully!", vbInformation, "Save Successful"
       Call ClearCtrls
       
       Me.txtname.SetFocus
   Else
       MsgBox "Sorry, Unable to Save Store's Details:Please Try Again!", vbInformation, "Save Failed"
       Me.txtname.SetFocus
   End If
 
   
Else
'edit part

   rs.Open "Select StoreName From Stores Where StoreName ='" & Trim(Me.txtname) & "' and StoreID<>'" & Storeid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.Close: cn.Close
      MsgBox "A STORE HAS ALREADY BEEN SETUP WITH THE NAME ENTERED.", vbInformation, "NAME ALREADY EXIST"
      Me.txtname.SetFocus: Exit Sub
   End If
      If rs.State = 1 Then rs.Close
   
   
   
   
   cn.Execute "Update Stores Set StoreName ='" & Trim(Me.txtname.Text) & "',StoreLocation='" & Trim(Me.txtStoreLocation.Text) & "' Where StoreID ='" & Storeid & "'", Y
      
   If Y > 0 Then
      MsgBox "Edit Successful!", vbInformation, "Edit Successful"
      Call ClearCtrls
      
      
      Me.txtname.SetFocus
   Else
      MsgBox "Sorry, Unable to Edit Store's Details:Please Try Again!", vbInformation, "Edit Failed"
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
     MsgBox "Sorry, Unable to Save Store's Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub
End Sub

Private Sub flxgStores_Click()
Me.txtname = Me.flxgStores.TextMatrix(Me.flxgStores.Row, 0)
Me.txtStoreLocation = Me.flxgStores.TextMatrix(Me.flxgStores.Row, 1)
Storeid = Me.flxgStores.TextMatrix(Me.flxgStores.Row, 2)
'Me.txtAddress = Me.flxgStores.TextMatrix(Me.flxgStores.Row, 3)
'Debtorid = Me.flxgStores.TextMatrix(Me.flxgStores.Row, 4)

Me.flxgStores.Visible = False
Me.cmdDelete.Enabled = True
Me.cmdSave.Enabled = True
sflag = True
Me.txtname.SetFocus
End Sub

Private Sub Form_Load()
Me.Height = 4410
Me.Width = 8130

Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.txtStoreLocation.SetFocus
End If
End Sub

Private Sub txtStoreLocation_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   
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

