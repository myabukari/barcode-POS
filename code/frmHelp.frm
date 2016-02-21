VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmHelp 
   BackColor       =   &H00E19E13&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleMode       =   0  'User
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   8640
      Width           =   1455
      Begin VB.Frame Frame3 
         BackColor       =   &H00C29E21&
         Height          =   855
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1455
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   735
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1296
            Caption         =   "E&xit"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   16777215
            Gradient        =   1
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            ImgAlign        =   4
            Image           =   "frmHelp.frx":0000
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   8655
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton cmdSales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales Revenue Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "frmHelp.frx":0452
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdProductsRpt 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "Products Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "frmHelp.frx":1C04
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdTips 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "frmHelp.frx":3520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "frmHelp.frx":4DC1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSetProducts 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SetUp Products"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "frmHelp.frx":5203
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdSellout 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daily Sales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "frmHelp.frx":550D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdPassword 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "frmHelp.frx":5DDF
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog commdialog 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Database Files (*.MDB) | *.MDB"
         InitDir         =   "C:"
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()

End Sub

Private Sub cmdBackUp_Click()

End Sub

Private Sub cmdCashRegister_Click()
frmCashRegister.Show
End Sub

Private Sub cmdCashRegisterSales_Click()
frmCashRegisterRpt.Show
End Sub

Private Sub cmdExit_Click()
Unload Me
Unload frmMDI
End Sub

Private Sub cmdPassword_Click()
frmChangePassword.Show
End Sub

Private Sub cmdProductsRpt_Click()
frmProductsRpt.Show
End Sub

Private Sub cmdSales_Click()
frmRevenueRpt.Show
End Sub

Private Sub cmdSearch_Click()
frmFind.Show
End Sub

Private Sub cmdSellout_Click()
frmSellProducts.Show
End Sub

Private Sub cmdSetProducts_Click()
frmProducts.Show
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
End Sub

Private Sub lvButtons_H2_Click()

End Sub
