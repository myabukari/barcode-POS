VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmDebtorsPaymentsRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "DebtorsPayments Report"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7785
   Icon            =   "frmDebtorsRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3480
   ScaleWidth      =   7785
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Caption         =   "SPECIFY DATE RANGE"
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
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   7215
      Begin Crystal.CrystalReport CrystalDebtorsReport 
         Left            =   5880
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   4320
         TabIndex        =   5
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         TabIndex        =   6
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin VB.Label Label2 
         BackColor       =   &H00C29E21&
         Caption         =   "To:"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "From:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   7200
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   855
         Left            =   1320
         TabIndex        =   1
         Top             =   0
         Width           =   4335
         Begin lvButton_H.lvButtons_H cmdOk 
            Height          =   375
            Left            =   480
            TabIndex        =   2
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
            Image           =   "frmDebtorsRpt.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   2280
            TabIndex        =   3
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
            Image           =   "frmDebtorsRpt.frx":118C
            cBack           =   -2147483633
         End
      End
   End
End
Attribute VB_Name = "frmDebtorsPaymentsRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Debtorid As Integer
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, AccTransid As String
Dim Balance As Double, BroughtforwardBalance As Double, BroughtforwardDate As String, msg As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo OkError

'CrystalDebtorsReport.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalDebtorsReport.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalDebtorsReport.ReportFileName = App.Path & "\rptDebtorsPayments.rpt"
CrystalDebtorsReport.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"
'Open Connecttion to Server



 
If Me.dtpfrom <= Me.dtpto Then

bFlag = OpenConnection(cn, strg)
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
 
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select  * From DebtorsPayments  where PaymentDate >='" & Me.dtpfrom & "' and PaymentDate <='" & Me.dtpto & "' order by PaymentDate,PaymentTime desc", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
           MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED DATE", vbInformation, "DATE RANGE"
        If rs.State = 1 Then rs.Close
        Exit Sub
        End If
If rs.State = 1 Then rs.Close




rs.Open "Select  Balance,PaymentDate From DebtorsPayments  where PaymentDate <'" & Me.dtpfrom & "' order by PaymentDate desc,PaymentTime desc", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
              BroughtforwardBalance = rs.Fields("Balance")
              BroughtforwardDate = rs.Fields("PaymentDate")
        If rs.State = 1 Then rs.Close
        Else
              BroughtforwardBalance = 0
              BroughtforwardDate = ""
        End If
If rs.State = 1 Then rs.Close

'inserts or updates current balance in BalanceBF field of BalanceBroughtforward table and sets BalanceBroughtforwardID to BF
'to aid showing Brought forward Balance in reports
rs.Open "Select  BalanceBF From BalanceBroughtforward  where BalanceBroughtforwardID ='BF'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
  cn.Execute "Update BalanceBroughtforward Set BalanceBF='" & BroughtforwardBalance & "',BalanceBFDate='" & BroughtforwardDate & "' where BalanceBroughtforwardID ='BF'", Y
  If rs.State = 1 Then rs.Close
Else
  cn.Execute "Insert Into BalanceBroughtforward (BalanceBF,BalanceBroughtforwardID,BalanceBFDate) Select '" & BroughtforwardBalance & "','BF','" & Me.dtpfrom & "'", Y
  If rs.State = 1 Then rs.Close
End If
 If rs.State = 1 Then rs.Close


CrystalDebtorsReport.SelectionFormula = "{DebtorsPayments.PaymentDate} >= #" & Me.dtpfrom & "# And {DebtorsPayments.PaymentDate} <= #" & Me.dtpto & "#"
End If

'If Me.dtpto >= Me.dtpfrom Then
'CrystalDebtorsReport.SelectionFormula = "{CashRegister.SalesDate} >=#" & Me.dtpfrom & "# and {CashRegister.SalesDate} <=#" & Me.dtpto & "#"
'End If


If Me.dtpfrom > Me.dtpto Then
MsgBox "WRONG DATE RANGE ENTERED", vbInformation, "DATE RANGE"
Exit Sub
End If

   CrystalDebtorsReport.WindowState = crptMaximized
   CrystalDebtorsReport.WindowShowRefreshBtn = True
   Me.CrystalDebtorsReport.WindowTitle = "DEBTORS PAYMENTS REPORTS" & Format$(Date, "yyyy")
   
   CrystalDebtorsReport.Action = 1
   
Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "DEBTORS PAYMENTS REPORTS"
       Exit Sub
End Sub

Private Sub Form_Load()
Me.dtpfrom = Date
Me.dtpto = Date
Me.Height = 3990
  Me.Width = 7905
  Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
  Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub
