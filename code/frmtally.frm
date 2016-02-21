VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmtally 
   BackColor       =   &H00C29E21&
   Caption         =   "TALLY/BIN CARD"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11415
      Begin VB.Frame Frame5 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   7800
         Width           =   11175
         Begin Crystal.CrystalReport CrystalTally 
            Left            =   4200
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            Image           =   "frmtally.frx":0000
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdPrintTallyReport 
            Height          =   375
            Left            =   6480
            TabIndex        =   26
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            Caption         =   "&Print Tally Report"
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
            Image           =   "frmtally.frx":0452
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   2400
            TabIndex        =   27
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            Image           =   "frmtally.frx":076C
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   4560
            TabIndex        =   28
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            Image           =   "frmtally.frx":201D
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdDetailedReport 
            Height          =   375
            Left            =   8520
            TabIndex        =   31
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "&Detailed Tally Report"
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
            Image           =   "frmtally.frx":246F
            cBack           =   -2147483633
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgProducts 
         Height          =   2175
         Left            =   1320
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   12754465
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   $"frmtally.frx":2789
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   6735
         Left            =   0
         TabIndex        =   5
         Top             =   960
         Width           =   11415
         Begin VB.ComboBox cboClient 
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
            Height          =   315
            ItemData        =   "frmtally.frx":281E
            Left            =   6120
            List            =   "frmtally.frx":2820
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1680
            Width           =   5055
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgTally 
            Height          =   3615
            Left            =   120
            TabIndex        =   20
            Top             =   2880
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6376
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColorBkg    =   12754465
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   2
            SelectionMode   =   1
            FormatString    =   $"frmtally.frx":2822
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
            _Band(0).Cols   =   9
         End
         Begin VB.ComboBox cboRemarks 
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
            Height          =   315
            ItemData        =   "frmtally.frx":290D
            Left            =   1320
            List            =   "frmtally.frx":291A
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   2280
            Width           =   3495
         End
         Begin VB.TextBox txtInitials 
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
            Left            =   1320
            TabIndex        =   16
            Top             =   1680
            Width           =   3495
         End
         Begin VB.ComboBox cboBalance 
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
            Height          =   315
            ItemData        =   "frmtally.frx":2939
            Left            =   9120
            List            =   "frmtally.frx":299A
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2055
         End
         Begin VB.ComboBox cboIssuedOut 
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
            Height          =   315
            ItemData        =   "frmtally.frx":2A14
            Left            =   6120
            List            =   "frmtally.frx":2A75
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1935
         End
         Begin VB.ComboBox cboReceivedIn 
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
            Height          =   315
            ItemData        =   "frmtally.frx":2AEF
            Left            =   1320
            List            =   "frmtally.frx":2B50
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox txtParticulars 
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
            Left            =   6120
            TabIndex        =   8
            Top             =   480
            Width           =   5055
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   480
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   15718770
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   16384003
            CurrentDate     =   39091
         End
         Begin lvButton_H.lvButtons_H cmdAdd 
            Height          =   375
            Left            =   9840
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
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
            Image           =   "frmtally.frx":2BCA
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdRemove 
            Height          =   375
            Left            =   9840
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1335
            _ExtentX        =   2355
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
            Image           =   "frmtally.frx":438C
            cBack           =   -2147483633
         End
         Begin MSComCtl2.DTPicker dtpTallyDate 
            Height          =   315
            Left            =   6120
            TabIndex        =   32
            Top             =   2280
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   15718770
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   16384003
            CurrentDate     =   39091
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C29E21&
            Caption         =   "Client:"
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
            Left            =   5400
            TabIndex        =   22
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C29E21&
            Caption         =   "Remarks:"
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
            TabIndex        =   19
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C29E21&
            Caption         =   "Initials:"
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
            TabIndex        =   17
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label7 
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
            Left            =   8280
            TabIndex        =   15
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C29E21&
            Caption         =   "Issue Out:"
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
            Left            =   5160
            TabIndex        =   13
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C29E21&
            Caption         =   "Received In:"
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
            TabIndex        =   12
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C29E21&
            Caption         =   "Particulars:"
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
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C29E21&
            Caption         =   "Date:"
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
            Left            =   720
            TabIndex        =   7
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.TextBox txtReorderlevel 
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
         Left            =   9120
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "Reorder Level:"
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
         Left            =   7680
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Description:"
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
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmtally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String, AdjNoOfPackages As Integer, AdjNoPerPackage As Integer, AdjTotal As Integer
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, xflag As Boolean
Dim Fval2 As Integer, Frow2 As Integer, xx As Integer, eflag As Boolean, Receiptid As String, X As Integer
Dim BalanceflagcboIssuedOut_Change As Boolean, yflag As Boolean, ShowProductflag As Boolean, Storeid As Integer, Balance As Integer
Dim tallyflag As Boolean
Private Sub cboIssuedOut_Change()
If Me.cboIssuedOut <> "" And Me.cboReceivedIn = "" Then
Call ComputeBalance(Productid, Balance, Trim(Val(Me.cboReceivedIn)), Trim(Val(Me.cboIssuedOut)))
Me.cboBalance = Balance
End If
End Sub

Private Sub cboIssuedOut_Click()
If Me.cboIssuedOut <> "" And Me.cboReceivedIn = "" Then
Call ComputeBalance(Productid, Balance, Trim(Val(Me.cboReceivedIn)), Trim(Val(Me.cboIssuedOut)))
Me.cboBalance = Balance
End If
End Sub

Private Sub cboReceivedIn_Change()
Balance = 0
'Me.cboBalance = ""
If Me.cboReceivedIn <> "" And Me.cboIssuedOut = "" Then
Call ComputeBalance(Productid, Balance, Trim(Val(Me.cboReceivedIn)), Trim(Val(Me.cboIssuedOut)))
Me.cboBalance = Balance
End If
End Sub

Private Sub cboReceivedIn_Click()
If Me.cboReceivedIn <> "" And Me.cboIssuedOut = "" Then

Call ComputeBalance(Productid, Balance, Trim(Val(Me.cboReceivedIn)), Trim(Val(Me.cboIssuedOut)))
Me.cboBalance = Balance
End If
End Sub

Private Sub cmdAdd_Click()
Dim X As Integer, a As Integer
On Error GoTo OkError
'If eflag = False Then
'For xx = 1 To flxgProducts.Rows - 1
'           If flxgProducts.TextMatrix(xx, 0) = Trim(Me.txtProductName) And flxgProducts.TextMatrix(xx, 6) = Productid And flxgProducts.TextMatrix(xx, 3) = Trim(Me.txtSource) And flxgProducts.TextMatrix(xx, 4) = Trim(Me.txtDestination) Then
'              MsgBox Me.txtProductName & " Has Already Been Added!: You Can Only Edit What is Entered", vbInformation, ""
'              ShowProductsflag = True
'              Me.txtProductName = ""
'              Me.cboQuantity = ""
'              ShowProductsflag = False
'              Me.txtProductName.SetFocus
'              Exit Sub
'           End If
'         Next
'End If
'Me.txtTotalQuantity = Val(Trim(Me.txtNoPackages)) * Val(Trim(Me.txtNoPerPackage))
''strg2 = Val(Me.txtUnitPrice) * Val(Me.cboQuantity) * (100 - Val(Me.txtDiscount)) / 100
'Me.txtTotalCost = Format$(strg2, "¢#,###.00")
If Trim(Me.txtDescription) = "" Then MsgBox "Please select a Product", vbInformation, "": Me.txtDescription.SetFocus: Exit Sub
If Trim(Me.cboReceivedIn) = "" And Trim(Me.cboIssuedOut) = "" Then MsgBox "Please select Received In or Issued Out", vbInformation, "":  Exit Sub
If Trim(Me.txtParticulars) = "" Then MsgBox "Please Enter Particulars", vbInformation, "Particulars": Me.txtParticulars.SetFocus: Exit Sub
'    If Trim(Me.cboQuantity) = "" Then MsgBox "Please Enter Quantity of Product", vbInformation, "Quantity": Me.cboQuantity.SetFocus: Exit Sub
    
    
    If MsgBox("ARE YOU SURE  ABOUT THE ISSUING DATE ENTERED?", vbYesNo + vbQuestion, "CONFIRM ISSUING DATE") = vbNo Then
    Me.dtpDate.SetFocus
    Exit Sub
    End If
    
    
    'If Trim(Me.txtTotalCost) = "" Then MsgBox "Please Compute TotalCost", vbInformation: Me.txtTotalCost.SetFocus: Exit Sub
    If Frow2 <= 0 Then
      If Fval2 > 0 Then
         For xx = 1 To flxgProducts.Rows - 1
           If flxgProducts.TextMatrix(xx, 7) = Productid Then
              MsgBox Me.txtDescription & " Has Already Been Added!: You Can Only Edit What is Entered", vbInformation, ""
              ShowProductsflag = True
'              Me.txtProductName = ""
'              Me.cboQuantity = ""
              ShowProductsflag = False
'              Me.txtProductName.SetFocus
              Exit Sub
           End If
         Next
       End If
      Fval2 = Fval2 + 1
      flxgProducts.Rows = Fval2 + 1
      flxgProducts.TextMatrix(Fval2, 0) = Me.dtpDate
      flxgProducts.TextMatrix(Fval2, 1) = Me.txtParticulars
      flxgProducts.TextMatrix(Fval2, 2) = Me.cboReceivedIn
      flxgProducts.TextMatrix(Fval2, 3) = Me.cboIssuedOut
      flxgProducts.TextMatrix(Fval2, 4) = Me.cboBalance
      flxgProducts.TextMatrix(Fval2, 5) = Me.txtInitials
      flxgProducts.TextMatrix(Fval2, 6) = Me.cboRemarks
      flxgProducts.TextMatrix(Fval2, 7) = Productid
       'flxgProducts.TextMatrix(Fval2, 7) = Trim(Productid)
       
    Else
       flxgProducts.TextMatrix(Frow2, 0) = Me.dtpDate
       flxgProducts.TextMatrix(Frow2, 1) = Me.txtParticulars
       flxgProducts.TextMatrix(Frow2, 2) = Me.cboReceivedIn
       flxgProducts.TextMatrix(Frow2, 3) = Me.cboIssuedOut
       flxgProducts.TextMatrix(Fval2, 4) = Me.cboBalance
        flxgProducts.TextMatrix(Frow2, 5) = txtInitials
        flxgProducts.TextMatrix(Frow2, 6) = cboRemarks
       flxgProducts.TextMatrix(Frow2, 7) = Productid
        'flxgProducts.TextMatrix(Frow2, 7) = Trim(Productid)
      Frow2 = 0
    
    End If
    Me.txtParticulars = ""
    Me.cboReceivedIn = ""
    Me.cboIssuedOut = ""
    Me.cboBalance = ""
    Me.txtInitials = ""
    Me.cboRemarks = ""
    
'    ShowProductsflag = True
'    Me.txtProductName = ""
'    ShowProductsflag = False
'    Me.txtProductName.SetFocus
'    Me.cmdSave.Enabled = True
    
    
    
    
    
'    If eflag = True Then
'    eflag = False
'   End If
   
   Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
    MsgBox "SORRY,TRY AGAIN", vbInformation, "ADD ITEMS"
     Exit Sub
End Sub

Private Sub cmdClear_Click()
Call ClearCtrls
End Sub

Private Sub cmdDetailedReport_Click()

If Me.txtDescription = "" Then
   MsgBox "YOU MUST ENTER PRODUCT NAME.", vbInformation, "PRODUCT NAME"
   Me.txtDescription.SetFocus: Exit Sub
End If


frmTallyRpt.Show
End Sub

Private Sub cmdPrintTallyReport_Click()
PrintReceipt
End Sub

Private Sub cmdSave_Click()
If Trim(Me.txtDescription) = "" Then
   MsgBox "YOU MUST ENTER PRODUCT NAME.", vbInformation, "PRODUCT NAME"
   Me.txtDescription.SetFocus: Exit Sub
End If

If Trim(Me.txtParticulars) = "" Then
   MsgBox "YOU MUST ENTER PARTICULARS.", vbInformation, "PARTICULARS"
   Me.txtParticulars.SetFocus: Exit Sub
End If


If Trim(Me.cboReceivedIn) = "" And Trim(Me.cboIssuedOut) = "" Then
   MsgBox "YOU MUST ENTER RECEIVEDIN OR ISSUEDOUT.", vbInformation, ""
   Exit Sub
End If

If Trim(Me.cboBalance) = "" Then
   MsgBox "YOU MUST ENTER BALANCE.", vbInformation, "BALANCE"
   Me.cboBalance.SetFocus: Exit Sub
End If

If Trim(Me.cboRemarks) = "" Then
   MsgBox "YOU MUST ENTER REMARKS.", vbInformation, "REMARKS"
   Me.cboRemarks.SetFocus: Exit Sub
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

'If Trim(Me.txtTotalQuantity) <> "" Then
'Me.txtStock = Trim(Me.txtTotalQuantity)
'End If

'If sflag = False Then
   'save part
'rs.Open "Select ProductName From Products Where ProductName ='" & Trim(Me.txtName) & "'", cn, adOpenForwardOnly, adLockReadOnly
'If rs.RecordCount > 0 Then
'   rs.Close: cn.Close
'   MsgBox "A ProductName Has Already Been Setup with the Name.", vbInformation
'   Me.MousePointer = vbDefault
'   Me.cmdSave.Enabled = True
'   Me.txtName.SetFocus: Exit Sub
'
'   Me.MousePointer = vbDefault
'   Me.cmdSave.Enabled = True
'End If
'rs.Close


'rs.Open "Select ProductCode From Products Where ProductCode ='" & Trim(Me.txtCode) & "'", cn, adOpenForwardOnly, adLockReadOnly
'If rs.RecordCount > 0 Then
'   rs.Close: cn.Close
'   MsgBox "A ProductName Has Already Been Setup with the Code.", vbInformation
'   Me.MousePointer = vbDefault
'   Me.cmdSave.Enabled = True
'   Me.txtName.SetFocus: Exit Sub
'
'   Me.MousePointer = vbDefault
'   Me.cmdSave.Enabled = True
'End If
'rs.Close

   Call ComputeBalance(Productid, Balance, Trim(Val(Me.cboReceivedIn)), Trim(Val(Me.cboIssuedOut)))
   
   If Balanceflag = True Then
   MsgBox "Please Enter Starting Balance", vbInformation, "Balance"
'   If rs.State = 1 Then rs.Close
   Balanceflag = False
   Me.cboBalance.SetFocus: Exit Sub
   End If
   
   
'   cn.BeginTrans
'" & Format$(dtpPaymentDate, "Medium Date") & " " & Format$(Now, "Medium Time") & "'
   cn.Execute "Insert Into Tally ([ProductID],[IssueDate],[Particulars],[ReceivedIn],[IssueOut],[Balance],[Initials],[Remarks],Client,IssueTime) select '" & Productid & "','" & Trim(Me.dtpDate) & "','" & Trim(Me.txtParticulars.Text) & "','" & Trim(Val(Me.cboReceivedIn.Text)) & "','" & Trim(Val(Me.cboIssuedOut)) & "','" & Balance & "','" & Trim(Me.txtInitials) & "','" & Trim(Me.cboRemarks.Text) & "','" & Trim(Me.cboClient.Text) & "','" & Format$(dtpDate, "Medium Date") & " " & Format$(Now, "Medium Time") & "'", Y
'   If Y > 0 Then
'   cn.Execute "Insert Into ProductInventory ([ProductID],[StockLevel],[ReorderLevel],[UserID],[Date]) select '" & Trim(ProductID) & "','" & Val(Me.txtStock.Text) & "','" & Val(Me.txtReorder.Text) & "','111','" & Date & "'", Y
'   End If
   If Y > 0 Then
'   cn.CommitTrans
     MsgBox "Product Tally processed Successfully!", vbInformation, "Tally/Bin System"
       Call ClearCtrls
       Call GetTally(Productid)
'     Me.txtName = "": Me.txtbasePrice = "": Me.txtbaseunit = "": Me.txtDiscount = ""
'     Me.txtManufacterer = "": Me.txtNoCartones = "": Me.txtReorder = "": Me.txtStock = ""
'     Me.txtTotalQuantity = "": Me.txtUnitPrice = ""
       Me.txtDescription.SetFocus
   Else
'   cn.RollbackTrans
      MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
      Me.txtDescription.SetFocus
   End If
 
   
'Else
'edit part

'   rs.Open "Select ProductName From Products Where ProductName ='" & Trim(Me.txtName) & "' and ProductID<>'" & ProductID & "'", cn, adOpenForwardOnly, adLockReadOnly
'   If rs.RecordCount > 0 Then
'      rs.Close: cn.Close
'      MsgBox "A Product Has Already Been Setup with the Name.", vbInformation
'      Me.txtName.SetFocus: Exit Sub
'   End If
'   rs.Close
'
'   rs.Open "Select ProductCode From Products Where ProductCode ='" & Trim(Me.txtCode) & "' and ProductID<>'" & ProductID & "'", cn, adOpenForwardOnly, adLockReadOnly
'   If rs.RecordCount > 0 Then
'      rs.Close: cn.Close
'      MsgBox "A Product Has Already Been Setup with the Code.", vbInformation
'      Me.txtName.SetFocus: Exit Sub
'   End If
'   rs.Close
'
'   rs.Open "Select ProductInventory.StockLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where  Products.ProductID='" & ProductID & "'", cn, adOpenForwardOnly, adLockReadOnly
'   If rs.RecordCount > 0 Then
'      If Trim(Me.txtTotalQuantity) <> "" Then
'      StockQty = Val(Trim(Me.txtTotalQuantity)) + rs.Fields("StockLevel")
'      Else
'      StockQty = Val(Trim(Me.txtStock))
'      End If
'      rs.Close
'   End If
  
   
'   cn.BeginTrans
'   cn.Execute "Update Products Set ProductName ='" & Trim(Me.txtName.Text) & "',BaseUnit='" & Val(Trim(Me.txtbaseunit.Text)) & "',UnitPrice='" & Trim(Me.txtUnitPrice.Text) & "',TotalQuantity='" & Val(Trim(Me.txtTotalQuantity.Text)) & "',PricePerCartone='" & Val(Trim(Me.txtbasePrice.Text)) & "',Discount='" & Val(Trim(Me.txtDiscount.Text)) & "',Manufacturer='" & Trim(Me.txtManufacterer.Text) & "',ExpiryDate='" & Trim(Me.dtpExp) & "',ManufucteryDate='" & Trim(Me.dtpmanu) & "',NoOfCartones='" & Val(Trim(Me.txtNoCartones.Text)) & "',Department='" & (Trim(Me.cboDepartment.Text)) & "',ProductCode='" & (Trim(Me.txtCode.Text)) & "' Where ProductID ='" & ProductID & "'", Y
'       If Y > 0 Then
'   cn.Execute "Update ProductInventory Set StockLevel ='" & StockQty & "',ReorderLevel='" & Val(Trim(Me.txtReorder.Text)) & "' Where ProductID ='" & ProductID & "'", Y
'       End If
'
'   If Y > 0 Then
'      cn.CommitTrans
'      MsgBox "Edit Successful!", vbInformation, "Edit Successful"
'     'Clear ctrls and setfocus to Clinic ctrl
'      Call ListProducts
'      Call ClearCtrls
'      Me.txtName = "": Me.txtbasePrice = "": Me.txtbaseunit = "": Me.txtDiscount = ""
'      Me.txtManufacterer = "": Me.txtNoCartones = "": Me.txtReorder = "": Me.txtStock = ""
'      Me.txtTotalQuantity = "": Me.txtUnitPrice = ""
'      Me.txtName.SetFocus
'   Else
'      MsgBox "Sorry, Unable to Edit Product Details:Please Try Again!", vbInformation, "Edit Failed"
'   End If
'
' End If
'
'sflag = False
'If cn.State = 1 Then cn.Close
'If rs.State = 1 Then rs.Close
'
'Me.MousePointer = vbDefault
'Me.cmdSave.Enabled = True


Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub



Private Sub dtpTallyDate_Click()
'tallyflag = True
'GetTally (Productid)
'tallyflag = False
End Sub

Private Sub dtpTallyDate_CloseUp()
tallyflag = True
GetTally (Productid)
tallyflag = False
End Sub

Private Sub flxgProducts_Click()

Me.txtDescription = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 0)
Me.txtReorderlevel = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 1)
Productid = Me.flxgProducts.TextMatrix(Me.flxgProducts.Row, 2)
Me.flxgProducts.Visible = False
Me.dtpDate.SetFocus
Balance = 0
GetTally (Productid)
End Sub

Private Sub Form_Load()
Me.dtpDate = Date
Me.dtpTallyDate = Date
End Sub

Private Sub txtDescription_Change()
If Me.txtDescription = "" Then
 Me.flxgProducts.Visible = False
 Exit Sub
End If
'On Error GoTo OkError
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      
       MsgBox strg, vbInformation:
      Exit Sub
   End If

  If ShowProductsflag = False Then

   rs.Open "Select * From Products inner Join ProductInventory on Products.ProductID=ProductInventory.ProductID  Where ProductName Like '" & Trim(Me.txtDescription) & "%" & "' Order By ProductName ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   Me.flxgProducts.Height = 950 + (285 * (rs.RecordCount - 1))
   
   If flxgProducts.Height >= 4455 Then
      flxgProducts.Height = 4455
   End If
    flxgProducts.Rows = rs.RecordCount + 1
   With flxgProducts
      For X = 1 To rs.RecordCount
       .TextMatrix(X, 0) = rs.Fields("ProductName")
       .TextMatrix(X, 1) = rs.Fields("ReorderLevel")
       .TextMatrix(X, 2) = rs.Fields("ProductID")
       
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 2
      .RowSel = 1
   End With
   flxgProducts.Visible = True
   If rs.State = 1 Then rs.Close
'   Me.cmdSave.Enabled = True
 Else
     flxgProducts.Visible = False
     If rs.State = 1 Then rs.Close
      
     
End If
If rs.State = 1 Then rs.Close
End If
Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
    MsgBox "Stores could not display", , "Displaying"
     Exit Sub
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.dtpDate.SetFocus
End If
End Sub

Private Sub txtParticulars_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   Me.cboReceivedIn.SetFocus
End If
End Sub
Private Sub ComputeBalance(Productid As String, Balance As Integer, Optional ReceivedIn As Integer, Optional IssuedOut As Integer)

'If Me.txtStoreName = "" Then
'Me.flxgStoreName.Visible = False
' Exit Sub
'End If
'On Error GoTo OkError
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
       MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   If rs.State = 1 Then rs.Close
   rs.Open "Select * From Tally where ProductID='" & Productid & "' order by IssueDate desc ,IssueTime desc ", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   rs.MoveFirst
   If Me.cboReceivedIn <> "" Then
   Balance = ReceivedIn + rs.Fields("Balance")
   End If
   
   If Me.cboIssuedOut <> "" Then
   Balance = rs.Fields("Balance") - IssuedOut
   End If
      
   If rs.State = 1 Then rs.Close
   
   Else
   
'   If Me.cboBalance = "" Then
'   Balanceflag = True
''   msg = "Please Enter Starting Balance"
'   If rs.State = 1 Then rs.Close
'    Exit Sub
'   End If
   
   If Me.cboReceivedIn <> "" Then
   Balance = ReceivedIn '+ Val(cboBalance)
   End If
   
'   If Me.cboIssuedOut <> "" Then
'   Balance = Val(cboBalance) - IssuedOut
'   End If
   
   If rs.State = 1 Then rs.Close
   
   End If
   If rs.State = 1 Then rs.Close

Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
    MsgBox "Stores could not display", , "Displaying"
     Exit Sub


End Sub

Private Sub GetTally(Productid As String)

'If Me.txtStoreName = "" Then
'Me.flxgStoreName.Visible = False
' Exit Sub
'End If
'On Error GoTo SaveError
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
       MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   
   If tallyflag = True Then
   rs.Open "Select * From Tally where ProductID='" & Productid & "' and IssueDate >= '" & Me.dtpTallyDate & "'order by IssueDate,IssueTime  ", cn, adOpenForwardOnly, adLockReadOnly
   Else
   rs.Open "Select * From Tally where ProductID='" & Productid & "' order by IssueDate,IssueTime  ", cn, adOpenForwardOnly, adLockReadOnly
   End If
   If rs.RecordCount > 0 Then
   rs.MoveFirst
   Me.flxgTally.Rows = rs.RecordCount + 1
  
   For X = 1 To rs.RecordCount
        Me.flxgTally.TextMatrix(X, 0) = rs.Fields("IssueDate")
        Me.flxgTally.TextMatrix(X, 1) = rs.Fields("Particulars")
        Me.flxgTally.TextMatrix(X, 2) = rs.Fields("ReceivedIn")
        Me.flxgTally.TextMatrix(X, 3) = rs.Fields("IssueOut")
        Me.flxgTally.TextMatrix(X, 4) = rs.Fields("Balance")
        Me.flxgTally.TextMatrix(X, 5) = rs.Fields("Initials")
        Me.flxgTally.TextMatrix(X, 6) = rs.Fields("Remarks")
        If rs.Fields("Client") <> "" Then
         Me.flxgTally.TextMatrix(X, 7) = rs.Fields("Client")
        End If
        Me.flxgTally.TextMatrix(X, 8) = rs.Fields("ProductID")
        
        rs.MoveNext
   Next
    If rs.State = 1 Then rs.Close
    
   Else
        For X = 0 To 8
         Me.flxgTally.TextMatrix(1, X) = ""
        Next
         Me.flxgTally.Rows = 2
   End If
    If rs.State = 1 Then rs.Close


Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Find Grades:Please Try Again!", vbInformation, "Please Try Again!"
     Exit Sub


End Sub
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next

For X = 0 To 8
Me.flxgTally.TextMatrix(1, X) = ""
Next
Me.flxgTally.Rows = 2
Productid = ""
Me.txtDescription.SetFocus

End Sub

Private Sub PrintReceipt()

On Error GoTo OkError

   
CrystalTally.ReportFileName = App.Path & "\rtpTallyCard.rpt"
CrystalTally.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"

'If Me.cboProducts <> "All" Then
CrystalTally.SelectionFormula = "{Products.ProductName} ='" & Me.txtDescription.Text & "'"
'Else
'CrystalProducts.SelectionFormula = ""
'End If
   CrystalTally.WindowState = crptMaximized
   CrystalTally.WindowShowRefreshBtn = True
   CrystalTally.WindowTitle = "TALLY " & Format$(Date, "yyyy")
   CrystalTally.Action = 0
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY RECEIPT,PLEASE TRY AGAIN", vbInformation, "RECEIPT"
       Exit Sub
End Sub

