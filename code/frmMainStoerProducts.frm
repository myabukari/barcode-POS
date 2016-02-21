VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmMainStoreProducts 
   BackColor       =   &H00C29E21&
   Caption         =   "MainStoreProducts"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   Icon            =   "frmMainStoerProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   6240
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   5895
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   7320
         TabIndex        =   12
         Top             =   120
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
         Image           =   "frmMainStoerProducts.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdSave 
         Height          =   375
         Left            =   120
         TabIndex        =   13
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
         Image           =   "frmMainStoerProducts.frx":075C
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdDelete 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
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
         Image           =   "frmMainStoerProducts.frx":0BAE
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdFind 
         Height          =   375
         Left            =   3000
         TabIndex        =   15
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
         Image           =   "frmMainStoerProducts.frx":0D48
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
         Left            =   5880
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelete1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Delete"
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
         Left            =   5880
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
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
         Height          =   495
         Left            =   5880
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit12 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5880
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin lvButton_H.lvButtons_H cmdExit1 
         Height          =   375
         Left            =   4440
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
         Image           =   "frmMainStoerProducts.frx":119A
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   4575
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   5655
      Begin MSComctlLib.ListView LstMainProducts 
         Height          =   4095
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   0
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ProductName"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ReOrderLevel"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ProductID"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin VB.TextBox txtReOrderLevel 
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
         Left            =   1800
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtProductName 
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
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "ReOrderLevel:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "ProductName:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMainStoreProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer

Private Sub cmdDelete_Click()
If Me.txtProductName = "" Then
MsgBox "ENTER PRODUCTNAME TO DELETE", vbInformation, "DELETE"
Me.txtProductName.SetFocus: Exit Sub
End If
If MsgBox("ARE YOU SURE  YOU WANT TO DELETE THE PRODUCTS DETAILS?", vbYesNo + vbQuestion, "CONFIRM DELETE") = vbYes Then
   
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
   cn.BeginTrans
   cn.Execute "Delete From MainStoreProducts Where ProductID ='" & Productid & "'", Y
   
    If Y > 0 Then
    rs.Open "Select * From MainInventory Where ProductID ='" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
     cn.Execute "Delete From MainInventory Where ProductID ='" & Productid & "'", Y
     If rs.State = 1 Then rs.Close
    End If
     If rs.State = 1 Then rs.Close
    End If
    
    If Y > 0 Then
     rs.Open "Select * From Receivals Where ProductID ='" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
     If rs.RecordCount > 0 Then
     cn.Execute "Delete From Receivals Where ProductID ='" & Productid & "'", Y
     If rs.State = 1 Then rs.Close
    End If
     If rs.State = 1 Then rs.Close
    End If
    
    If Y > 0 Then
     rs.Open "Select * From StockingRetail Where ProductID ='" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
     If rs.RecordCount > 0 Then
     cn.Execute "Delete From StockingRetail Where ProductID ='" & Productid & "'", Y
     If rs.State = 1 Then rs.Close
    End If
     If rs.State = 1 Then rs.Close
    End If
     
    If Y > 0 Then
     rs.Open "Select * From [MainReceipts/MainStoreProducts] Where ProductID ='" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
     If rs.RecordCount > 0 Then
     cn.Execute "Delete From [MainReceipts/MainStoreProducts] Where ProductID ='" & Productid & "'", Y
     If rs.State = 1 Then rs.Close
    End If
     If rs.State = 1 Then rs.Close
    End If
    
    If Y > 0 Then
    cn.CommitTrans
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     'Clear ctrls and setfocus to supplier name ctrl
     ClearCtrls
     Me.txtProductName.SetFocus
   Else
      
      MsgBox "Sorry, Unable to Delete Products Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
    Me.txtProductName.SetFocus
   sflag = False
   Me.cmdDelete.Enabled = False
   If cn.State = 1 Then cn.Close
   
   Me.MousePointer = vbDefault
   
   Call ListProducts
   Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Delete Products  Details:Please Try Again!", vbInformation, "Delete Failed"
        Exit Sub
End If
End Sub

Private Sub cmdExit1_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If Trim(Me.txtProductName) = "" Then
   MsgBox "YOU MUST ENTER PRODUCT NAME.", vbInformation, "PRODUCT NAME"
   Me.txtProductName.SetFocus: Exit Sub
End If



On Error GoTo SaveError
'Me.cmdSave.Enabled = False

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
rs.Open "Select ProductName From MainStoreProducts Where ProductName ='" & Trim(Me.txtProductName) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "A ProductName Has Already Been Setup with the Name.", vbInformation
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtProductName.SetFocus: Exit Sub
   
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
End If
rs.Close

   Call Generate_ProductID(Productid)
   cn.Execute "Insert Into MainStoreProducts ([ProductID],[ProductName],[ReOrderLevel]) select '" & Trim(Productid) & "','" & Trim(Me.txtProductName.Text) & "','" & Val(Me.txtReOrderLevel.Text) & "'", Y
   If Y > 0 Then
   MsgBox "Saved Successfully!", vbInformation, "Save Successful"
       Call ClearCtrls
       Call ListProducts
     
     Me.txtProductName.SetFocus
   Else
   
   MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
   Me.txtProductName.SetFocus
   End If
   Else
'edit part

   rs.Open "Select ProductName From MainStoreProducts Where ProductName ='" & Trim(Me.txtProductName) & "' and ProductID<>'" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.Close: cn.Close
      MsgBox "A Product Has Already Been Setup with the Name.", vbInformation
      Me.txtProductName.SetFocus: Exit Sub
   End If
   rs.Close
   
   
   
   
   cn.Execute "Update MainStoreProducts Set ProductName ='" & Trim(Me.txtProductName.Text) & "',ReOrderLevel='" & Val(Trim(Me.txtReOrderLevel.Text)) & "' Where ProductID ='" & Productid & "'", Y
      
   If Y > 0 Then
     
      MsgBox "Edit Successful!", vbInformation, "Edit Successful"
     'Clear ctrls and setfocus to Clinic ctrl
      Call ListProducts
      Call ClearCtrls
      Me.txtProductName.SetFocus
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

Private Sub Form_Load()
Call ListProducts
CenterForm Me
Me.Height = 7635
Me.Width = 6360
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2

End Sub

Private Sub LstMainProducts_Click()
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
rs.Open "Select MainStoreProducts.* From MainStoreProducts  Order By MainStoreProducts.ProductName", cn, adOpenForwardOnly, adLockReadOnly


    While Not rs.EOF
        If Me.LstMainProducts.SelectedItem.Text = rs("ProductName") Then
            Me.txtProductName.Text = Trim(rs("ProductName"))
            
               Productid = rs("ProductID")
                
                Me.txtReOrderLevel.Text = Trim(rs("ReOrderLevel"))
                

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

Private Sub txtProductName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If KeyAscii = vbKeyReturn Then
   Me.cmdSave.SetFocus
End If
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
End Sub
Private Function Generate_ProductID(Product_ID As String) As Boolean
Dim strg As String, strg1 As String, strg2 As String, bFlag As Boolean
On Error GoTo SaveError
rs.Open "Select ProductID From MainStoreProducts  order by ProductID Desc", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
      strg1 = Trim(rs.Fields!Productid)

   strg1 = Trim(Str(Val(strg1) + 1))
   strg1 = String$(5 - Len(strg1), "0") & strg1
Else
   strg1 = "00001"
End If
Product_ID = strg1

If rs.State = 1 Then rs.Close

Generate_ProductID = True

Exit Function
SaveError:
     If rs.State = 1 Then rs.Close
    Generate_ProductID = False
     Exit Function
End Function
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
End Sub
Private Sub ListProducts()
On Error GoTo SaveError
Me.LstMainProducts.ListItems.Clear
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


rs.Open "Select MainStoreProducts.* From MainStoreProducts  Order By MainStoreProducts.ProductName", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.LstMainProducts.ListItems.Add(, , Trim(rs!ProductName))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Val(Trim(rs!ReOrderLevel))
                List_Item.SubItems(2) = Trim(rs!Productid)
                'List_Item.SubItems(3) = Trim(rs!Discount)
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

