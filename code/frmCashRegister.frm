VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmCashRegister 
   BackColor       =   &H00C29E21&
   Caption         =   "CashRegister"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   Icon            =   "frmCashRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9570
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   9255
      Begin VB.CommandButton cmdExit1 
         BackColor       =   &H00C0C0C0&
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
         Left            =   9240
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Ok"
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
         Left            =   9240
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Print"
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
         Left            =   9240
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear1 
         BackColor       =   &H00C0C0C0&
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
         Left            =   9240
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin lvButton_H.lvButtons_H cmdOk 
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Ok"
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
         Image           =   "frmCashRegister.frx":030A
         cBack           =   -2147483633
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   1800
         TabIndex        =   20
         Top             =   0
         Width           =   5895
         Begin lvButton_H.lvButtons_H cmdPrint 
            Height          =   375
            Left            =   1560
            TabIndex        =   23
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Print"
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
            Image           =   "frmCashRegister.frx":118C
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   3000
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
            Image           =   "frmCashRegister.frx":129E
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   4440
            TabIndex        =   25
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            Image           =   "frmCashRegister.frx":2B4F
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      Begin VB.CommandButton cmdRemove1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Remove"
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
         Left            =   9240
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Add"
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
         Left            =   9240
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3360
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgCashRegister 
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3836
         _Version        =   393216
         BackColor       =   16052405
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorBkg    =   16052405
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   "<Price                                      |<Quantity                       |<Amount                                      "
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
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
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
         Left            =   8520
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   7
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtAmountPaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtTotalAmont 
         BackColor       =   &H00C29E21&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   1920
         Width           =   6855
      End
      Begin VB.TextBox txtPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EBBF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   2535
         Left            =   0
         TabIndex        =   21
         Top             =   2400
         Width           =   9255
         Begin lvButton_H.lvButtons_H cmdAdd 
            Height          =   375
            Left            =   7200
            TabIndex        =   26
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "&Add"
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
            Image           =   "frmCashRegister.frx":2FA1
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdRemove 
            Height          =   375
            Left            =   7200
            TabIndex        =   27
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "&Remove"
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
            Image           =   "frmCashRegister.frx":4763
            cBack           =   -2147483633
         End
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Quantity:"
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
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Balance:"
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
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Amount Paid:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "TotalAmount:"
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
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Price:"
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
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
   End
   Begin lvButton_H.lvButtons_H lvButtons_H1 
      Height          =   6855
      Left            =   -480
      TabIndex        =   28
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12091
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   10123282
      cGradient       =   10123282
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmCashRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Frow2 As Integer, Fval2 As Integer
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ctrl As Control, StockQty As Integer, a As Double, yflag As Boolean, strgbal As Double

Private Sub cmdAdd_Click()
  Call AddPrices
End Sub

Private Sub cmdClear_Click()
On Error GoTo OkError
For X = 1 To Me.flxgCashRegister.Rows - 1
            flxgCashRegister.TextMatrix(X, 0) = ""
            flxgCashRegister.TextMatrix(X, 1) = ""
            flxgCashRegister.TextMatrix(X, 2) = ""
            
       Next
    Fval2 = 0
    Frow2 = 0
    Me.flxgCashRegister.Rows = 2
    
      Me.txtPrice = "": Me.txtQuantity = "": Me.txtAmount = ""
      Me.txtAmountPaid = "": Me.txtBalance = "": Me.txtTotalAmont = ""
      Me.txtPrice.SetFocus
      
      
         
   Exit Sub
OkError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
    MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
     Exit Sub
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Call Save
End Sub

Private Sub cmdPrint_Click()
Call Receipt
Call Save
End Sub

Private Sub cmdRemove_Click()
Dim a As Integer, strgbal As String

On Error GoTo OkError

 If Frow2 = 0 Then Exit Sub
 
  If Frow2 = 1 Then
   If flxgCashRegister.TextMatrix(Frow2, 0) = "" Then
   Frow2 = 0
   Me.txtPrice.Text = ""
   Me.txtQuantity = ""
   Me.txtAmount = ""
   Me.txtPrice.SetFocus
   Exit Sub
   End If
 End If
 
 
If Trim(Me.txtTotalAmont) <> "" Then
'a = Len(Trim(Me.txtNetCost)) - 1
'b = Len(Trim(flxgCashRegister.TextMatrix(Frow2, 3))) - 1
strgbal = CDbl(Me.txtTotalAmont) - CDbl(flxgCashRegister.TextMatrix(Frow2, 2))
Me.txtTotalAmont.Text = Format$(strgbal, "#,###.00")

End If
   If Frow2 = flxgCashRegister.Rows - 1 Then
     If flxgCashRegister.Rows <> 2 Then
         flxgCashRegister.Rows = flxgCashRegister.Rows - 1
     Else
        For xx = 0 To 2
           flxgCashRegister.TextMatrix(Frow2, xx) = ""
        Next
        Me.txtPrice.Text = ""
        Me.txtQuantity = ""
        Me.txtPrice.SetFocus
     End If
        Me.txtPrice.Text = ""
        Me.txtQuantity = ""
        Me.txtPrice.SetFocus
   Else
         For xx = Frow2 To flxgCashRegister.Rows - 2
            flxgCashRegister.TextMatrix(xx, 0) = flxgCashRegister.TextMatrix(xx + 1, 0)
            flxgCashRegister.TextMatrix(xx, 1) = flxgCashRegister.TextMatrix(xx + 1, 1)
            flxgCashRegister.TextMatrix(xx, 2) = flxgCashRegister.TextMatrix(xx + 1, 2)
           
        Next
        Me.txtPrice.Text = ""
        Me.txtQuantity = ""
        Me.txtPrice.SetFocus
        flxgCashRegister.Rows = flxgCashRegister.Rows - 1
   End If
  Fval2 = Fval2 - 1
  Frow2 = 0
  
  'oflag = False
   Exit Sub
OkError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,UNABLE TO REMOVE,TRY AGAIN", vbInformation, "COMPUTATION"
     Exit Sub
  
End Sub

Private Sub flxgCashRegister_Click()
If MsgBox("ARE YOU SURE  YOU WANT TO EDIT OR REPLACE THE PRODUCT CLICKED?", vbYesNo + vbQuestion, "CONFIRMATION") = vbYes Then
        xflag = True
        Me.txtPrice = flxgCashRegister.TextMatrix(flxgCashRegister.Row, 0)
        Me.txtQuantity = flxgCashRegister.TextMatrix(flxgCashRegister.Row, 1)
        Me.txtAmount = flxgCashRegister.TextMatrix(flxgCashRegister.Row, 2)
        
Frow2 = flxgCashRegister.Row
'eflag = True
'oflag = False
Me.txtQuantity.SetFocus
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Width = 9690
Me.Height = 7020
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub lvButtons_H2_Click()

End Sub

Private Sub txtAmountPaid_Change()
If Me.txtAmountPaid = "" Then
 Exit Sub
End If
On Error GoTo OkError
If CDbl(Val(Me.txtAmountPaid)) >= CDbl(Me.txtTotalAmont) Then
strgbal = CDbl(Val(Me.txtAmountPaid)) - CDbl(Me.txtTotalAmont)
Me.txtBalance.Text = Format$(strgbal, "#,###.00")
Else
Me.txtBalance.Text = ""
End If
Exit Sub
OkError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
    MsgBox "SORRY,TRY AGAIN", vbInformation, "TRY AGAIN"
     Exit Sub
End Sub

Private Sub txtAmountPaid_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789."
If KeyAscii = vbKeyReturn Then
   If Me.txtAmountPaid <> "" Then
    Me.cmdOk.SetFocus
   Else
    Me.txtPrice.SetFocus
   End If
End If
If KeyAscii = 112 Or KeyAscii = 80 Then
 Call Receipt
 Call Save
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789."
If KeyAscii = 43 Then
   Me.txtQuantity = ""
   Me.txtQuantity.SetFocus
End If
If KeyAscii = vbKeyReturn Then
   Me.txtAmountPaid.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub
Private Sub Cal()

For X = 1 To Me.flxgCashRegister.Rows - 1
    'If Trim(flxgCashRegister.TextMatrix(X, 3)) <> "" Then
    'a = Len(Trim(flxgCashRegister.TextMatrix(X, 3))) - 1
    'Else
    'Me.txtNetCost = "¢00"
    'Exit Sub
    'End If
    i = i + Val(flxgCashRegister.TextMatrix(X, 2))
    Next
    Me.txtTotalAmont.Text = Format$(i, "#,###.00")
    
    i = 0
End Sub
Private Sub Amount()
Me.txtAmount = Val(Me.txtPrice) * Val(Me.txtQuantity)
End Sub
Private Sub AddPrices()
 On Error GoTo OkError
    Call Amount
    If Trim(Me.txtPrice) = "" Then MsgBox "Please Enter Price Of Product", vbInformation: Me.txtPrice.SetFocus: Exit Sub
    If Trim(Me.txtQuantity) = "" Then MsgBox "Please Enter Quantity of Product", vbInformation: Me.txtQuantity.SetFocus: Exit Sub
    'If Val(Trim(Me.txtUnitPrice)) = 0 Then MsgBox "Please Enter Price of Product", vbInformation: Me.txtUnitPrice.SetFocus: Exit Sub
    'If Trim(Me.txtTotalCost) = "" Then MsgBox "Please Compute TotalCost", vbInformation: Me.txtTotalCost.SetFocus: Exit Sub
    If Frow2 <= 0 Then
      If Fval2 > 1 Then
         'For xx = 1 To flxgCashRegister.Rows - 1
           'If flxgCashRegister.TextMatrix(xx, 0) = Trim(Me.txtName) And flxgCashRegister.TextMatrix(xx, 7) = Productid Then
              'MsgBox flxgCashRegister.TextMatrix(xx, 0) & " Has Already Been Added!: You Can Only Edit What is Entered"
              ''Me.txtName = ""
              'Exit Sub
           'End If
         'Next
       End If
      Fval2 = Fval2 + 1
      flxgCashRegister.Rows = Fval2 + 1
      flxgCashRegister.TextMatrix(Fval2, 0) = Me.txtPrice
      flxgCashRegister.TextMatrix(Fval2, 1) = Me.txtQuantity
      flxgCashRegister.TextMatrix(Fval2, 2) = Me.txtAmount
      
       
    Else
        flxgCashRegister.TextMatrix(Frow2, 0) = Me.txtPrice
        flxgCashRegister.TextMatrix(Frow2, 1) = Me.txtQuantity
        flxgCashRegister.TextMatrix(Frow2, 2) = Me.txtAmount
      Frow2 = 0
    End If
    Me.txtPrice = ""
    Me.txtQuantity = ""
    Me.txtPrice.SetFocus
    Call Cal
       
   Exit Sub
OkError:
     If rs.State <> 0 Then
        rs.Close
     End If
    MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
     Exit Sub
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789."
If KeyAscii = 43 Then
   Call AddPrices
   Me.txtPrice.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub
Private Sub Receipt()
If Me.txtAmountPaid = "" Then
MsgBox "Nothing to Print", vbInformation, "Blank Print"
Me.txtPrice.SetFocus: Exit Sub
End If

Dim a As Integer, b As Integer, c As Integer, d As Integer
a = Len(Me.txtTotalAmont)
b = Len(Format$(Me.txtAmountPaid, "#,###.00"))
c = Len(Me.txtBalance)
On Error GoTo SaveError
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.FontSize = 7
Printer.Print Tab(15); "Abu MULTIMART"
Printer.Print Tab(17); "P.O.BOX 52"
Printer.Print Tab(18); "TAMALE"
Printer.Print ""
Printer.Print Tab(3); "V.A.T #:"
Printer.FontSize = 6
Printer.Print Tab(3); Format(Now, "hh:mm:ss ampm")
Printer.FontSize = 7
Printer.Print Tab(3); "------------------------------------------------------------"
Printer.Print Tab(3); "------------------------------------------------------------"
Printer.Print Tab(3); "Amount Due  "; Tab(25); ":"; Tab(35); Me.txtTotalAmont
Printer.Print Tab(3); "Discount    "; Tab(25); ":"; Tab(35 + (a - 4)); "0.00"
Printer.Print Tab(3); "Total Amount"; Tab(25); ":"; Tab(35); Me.txtTotalAmont
Printer.Print Tab(3); "Amount Paid "; Tab(25); ":"; Tab(35 + (a - b)); Format$(Me.txtAmountPaid, "#,###.00")
Printer.Print Tab(3); "Balance     "; Tab(25); ":"; Tab(35 + (a - c)); Me.txtBalance
Printer.Print Tab(3); "------------------------------------------------------------"
Printer.Print Tab(3); "------------------------------------------------------------"
Printer.FontSize = 6
Printer.Print Tab(3); "V.A.T Inclusive."
Printer.Print Tab(15); "Thanks for Shopping with us."
Printer.Print ""
Printer.Print Tab(3); "Developed by Abu Yakubu"
Printer.Print Tab(3); "MTX SYSTEMS"
Printer.Print Tab(3); "0242324486"
Printer.EndDoc
Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
    MsgBox "SORRY,TRY AGAIN", vbInformation, "TRY AGAIN"
     Exit Sub
End Sub
Private Sub Save()
Dim a As Double

If Trim(Me.txtTotalAmont) = "" Then
MsgBox "Perform Sales", vbInformation, "No Sales"
Me.txtPrice.SetFocus: Exit Sub
End If
On Error GoTo OkError
'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

  rs.Open "Select * from CashRegister Where SalesDate ='" & Date & "'", cn, adOpenForwardOnly, adLockReadOnly
  If rs.RecordCount > 0 Then
   a = rs.Fields("Amount") + Val(CDbl(Me.txtTotalAmont))
   cn.Execute "Update CashRegister Set Amount ='" & a & "' Where SalesDate ='" & Date & "'", Y
   If rs.State = 1 Then rs.Close
   yflag = True
  Else
   cn.Execute "Insert Into CashRegister  ([Amount],[SalesDate]) select '" & Trim(Val(CDbl(Me.txtTotalAmont.Text))) & "','" & Date & "'", Y
   If rs.State = 1 Then rs.Close
   yflag = True
  End If
  If Y > 0 Then
   'MsgBox "SAVED SUCCESSFULLY", vbInformation, "SAVED"
  Else
   MsgBox "SAVED FAILED", vbInformation, "TRY AGAIN"
  End If
  
  If yflag = True Then
   ' i = Me.flxgCashRegister.Rows - 1
      For X = 1 To Me.flxgCashRegister.Rows - 1
            flxgCashRegister.TextMatrix(X, 0) = ""
            flxgCashRegister.TextMatrix(X, 1) = ""
            flxgCashRegister.TextMatrix(X, 2) = ""
           
       Next
    Fval2 = 0
    Me.flxgCashRegister.Rows = 2
    Me.txtAmountPaid = ""
    Me.txtTotalAmont = ""
    Me.txtBalance = ""
    Me.txtAmount = ""
    Me.txtPrice.SetFocus
    yflag = False
   End If
   
 If cn.State = 1 Then cn.Close
 If rs.State = 1 Then rs.Close
 Exit Sub
OkError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
    MsgBox "SORRY,THERE IS AN ERROR,TRY AGAIN", vbInformation, "COMPUTATION"
     Exit Sub
  
End Sub
