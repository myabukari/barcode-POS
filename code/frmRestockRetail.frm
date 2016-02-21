VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmRestockRetail 
   BackColor       =   &H00C29E21&
   Caption         =   "Retail Products Below Re-OrderLevel"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   Icon            =   "frmRestockRetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   7500
   Begin VB.Frame frame 
      BackColor       =   &H00C29E21&
      Height          =   7095
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6975
      Begin MSComctlLib.ListView lstProducts 
         Height          =   6735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   11880
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483643
         BackColor       =   0
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Stocklevel"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin lvButton_H.lvButtons_H cmdExit 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   8040
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      Caption         =   "&Exit"
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
      Image           =   "frmRestockRetail.frx":030A
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C29E21&
      Caption         =   "THE FOLLOWING PRODUCTS ARE RUNNING OUT OF STOCK IN THE RETAIL"
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
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmRestockRetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer

Private Sub ListProducts()
On Error GoTo SaveError
Me.lstProducts.ListItems.Clear
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select Products.*,ProductInventory.* From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID  Where ProductInventory.ReorderLevel>=ProductInventory.StockLevel    Order By Products.ProductName", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.lstProducts.ListItems.Add(, , Trim(rs!ProductName))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!StockLevel)
                'List_Item.SubItems(2) = Trim(rs!stocklevel)
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

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call ListProducts
CenterForm Me
Me.Width = 7620
Me.Height = 9105
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

