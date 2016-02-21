VERSION 5.00
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmLogout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin lvButton_H.lvButtons_H cmdYES 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Y&ES"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmLogout.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.CommandButton cmdYES1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Y&ES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin lvButton_H.lvButtons_H cmdNo 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&No"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmLogout.frx":0E82
      cBack           =   -2147483633
   End
   Begin VB.CommandButton cmdNo1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&NO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C29E21&
      Caption         =   "EXIT FROM THE SYSTEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmLogout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub cmdNo_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub cmdYES_Click()
Unload Me
Unload frmMDI
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Height = 2055
Me.Width = 4950
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

