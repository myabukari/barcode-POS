VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmCashRegisterRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "CashRegister Sales Report"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   Icon            =   "frmCashRegisterRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3240
   ScaleWidth      =   7185
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   6735
      Begin VB.CommandButton cmdOk1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
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
         Left            =   6720
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit1 
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
         Left            =   6720
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   855
         Left            =   1200
         TabIndex        =   8
         Top             =   0
         Width           =   4335
         Begin lvButton_H.lvButtons_H cmdOk 
            Height          =   375
            Left            =   480
            TabIndex        =   9
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Image           =   "frmCashRegisterRpt.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   2280
            TabIndex        =   10
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Image           =   "frmCashRegisterRpt.frx":118C
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Caption         =   "SPECIFY DATE RANGE"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin Crystal.CrystalReport CrystalCashRegister 
         Left            =   5280
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   4320
         TabIndex        =   1
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   15718770
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16515075
         CurrentDate     =   39087
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   15718770
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16515075
         CurrentDate     =   39087
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   6720
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "From:"
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
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "To:"
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
         Left            =   3840
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmCashRegisterRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError

'CrystalCashRegister.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalCashRegister.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalCashRegister.ReportFileName = App.Path & "\CashRegisterSales.rpt"
CrystalCashRegister.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"



If Me.dtpfrom <= Me.dtpto Then
CrystalCashRegister.SelectionFormula = "{CashRegister.SalesDate} >= #" & Me.dtpfrom & "# And {CashRegister.SalesDate} <= #" & Me.dtpto & "#"
End If

'If Me.dtpto >= Me.dtpfrom Then
'CrystalCashRegister.SelectionFormula = "{CashRegister.SalesDate} >=#" & Me.dtpfrom & "# and {CashRegister.SalesDate} <=#" & Me.dtpto & "#"
'End If


If Me.dtpfrom > Me.dtpto Then
MsgBox "WRONG DATE RANGE ENTERED", vbInformation, "DATE RANGE"
Exit Sub
End If

   CrystalCashRegister.WindowState = crptMaximized
   CrystalCashRegister.WindowShowRefreshBtn = True
   Me.CrystalCashRegister.WindowTitle = "CASH REGISTER SALES " & Format$(Date, "yyyy")
   
   CrystalCashRegister.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "SALES REVENUE"
       Exit Sub
End Sub

Private Sub Form_Load()
CenterForm Me
Me.dtpfrom = Date
Me.dtpto = Date
Me.Height = 3750
Me.Width = 7305
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

