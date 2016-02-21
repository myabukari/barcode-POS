VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmConvert 
   BackColor       =   &H00C29E21&
   Caption         =   "Convert"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   2640
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProducts 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   12754465
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      FormatString    =   "<ProductName                                 |<UnitPrice                              |<ProductCode      "
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
      _Band(0).Cols   =   3
   End
   Begin lvButton_H.lvButtons_H cmdConvert 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1085
      Caption         =   "&Convert"
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
      Image           =   "frmConvert.frx":0000
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Debtorid As Integer
Dim Productid As String, UnitPrice As Variant, ctrl As Control, AccTransid As String

Private Sub cmdConvert_Click()
'On Error GoTo OkError
'Open Connecttion to Server
'   Me.OptConvert(1).SetFocus
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   

   rs.Open "Select * From Products ", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
'    Me.flxgProducts.Rows = rs.RecordCount + 1
   
      For X = 1 To rs.RecordCount
        
        UnitPrice = (rs.Fields("UnitPrice") / 10000)
        
        UnitPrice = Format$(UnitPrice, "#,###.00")
        
        Productid = rs.Fields("ProductID")
        cn.Execute "Update Products Set UnitPrice ='" & UnitPrice & "'Where ProductID ='" & Productid & "'", Y
      
        rs.MoveNext
        Next
        If rs.State = 1 Then rs.Close
 End If
 If rs.State = 1 Then rs.Close
 Call GetProducts
 
End Sub
Private Sub GetProducts()



'On Error GoTo OkError
'Open Connecttion to Server
'   Me.OptConvert(1).SetFocus
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   

   rs.Open "Select * From Products ", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
      rs.MoveFirst
        Me.flxgProducts.Rows = rs.RecordCount + 1
   
      For X = 1 To rs.RecordCount
     
        Me.flxgProducts.TextMatrix(X, 0) = rs.Fields("ProductName")
        Me.flxgProducts.TextMatrix(X, 1) = Format$(rs.Fields("UnitPrice"), "#,###.00")
        Me.flxgProducts.TextMatrix(X, 2) = rs.Fields("ProductID")
        rs.MoveNext
        Next
        If rs.State = 1 Then rs.Close
 End If

If rs.State = 1 Then rs.Close




End Sub

Private Sub Command1_Click()
'Me.Text2 = Format$(Me.Text1, "#,###.00")
UnitPrice = (Text1) / 10000
 UnitPrice = Format$(UnitPrice, "#,###.00")

Me.Text2 = UnitPrice
End Sub

Private Sub Form_Load()
GetProducts
End Sub
