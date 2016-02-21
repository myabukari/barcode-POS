VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmFind 
   BackColor       =   &H00C29E21&
   Caption         =   "Search"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10065
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9735
      Begin VB.TextBox txtProduct 
         Appearance      =   0  'Flat
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
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProducts 
         Height          =   3375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5953
         _Version        =   393216
         BackColor       =   12754465
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"frmFind.frx":030A
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   3735
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   9495
      End
      Begin lvButton_H.lvButtons_H cmdFind 
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
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
         Image           =   "frmFind.frx":0399
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   5040
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   873
         Caption         =   "E&xit"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Image           =   "frmFind.frx":07EB
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "ProductName:"
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
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Accountid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer
Dim BalCD As Double, AccTransid As String, xflag As Boolean, AccountTransid As String
Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
If Me.txtProduct = "" Then
MsgBox "ENTER NAME OF PRODUCT TO SEARCH FOR", vbInformation, "PRODUCT NAME"
Me.txtProduct.SetFocus: Exit Sub
End If
On Error GoTo OkError
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      
       MsgBox strg, vbInformation:
      Exit Sub
   End If

   rs.Open "Select *   From Products  Where ProductName Like '" & Trim(Me.txtProduct) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   'flxgProducts.Height = 950 + (285 * (rs.RecordCount - 1))
   
   'If flxgProducts.Height >= 4455 Then
      'flxgProducts.Height = 4455
   'End If
    flxgProducts.Rows = rs.RecordCount + 1
   With flxgProducts
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("ProductName")
        .TextMatrix(X, 1) = rs.Fields("Department")
       .TextMatrix(X, 2) = rs.Fields("UnitPrice")
             
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 2
      .RowSel = 1
   End With
   flxgProducts.Visible = True
   If rs.State = 1 Then rs.Close
   
Else
     MsgBox "THE PRODUCT ENTERED IS NOT AVAILABLE", vbInformation, "SEARCH COMPLETE"
      If rs.State = 1 Then rs.Close
      
End If

Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "Items In Stock", , "Displaying"
     Exit Sub
End Sub

Private Sub Form_Load()
 Me.Height = 6840
  Me.Width = 10185
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub txtProduct_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()@#$%&*_-'"
If KeyAscii = vbKeyReturn Then
   Me.cmdFind.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub
