VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00A88311&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4155
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00A88311&
      Height          =   3915
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7665
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00A88311&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   7455
         Begin VB.Timer Timer1 
            Interval        =   100
            Left            =   6120
            Top             =   600
         End
         Begin lvButton_H.lvButtons_H lvButtons_H2 
            Height          =   135
            Left            =   0
            TabIndex        =   12
            Top             =   120
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   238
            CapAlign        =   2
            BackStyle       =   3
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
            cGradient       =   11043601
            Gradient        =   2
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H lvButtons_H3 
            Height          =   135
            Left            =   0
            TabIndex        =   13
            Top             =   1680
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   238
            CapAlign        =   2
            BackStyle       =   3
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
            cGradient       =   11043601
            Gradient        =   2
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H lvButtons_H4 
            Height          =   1335
            Left            =   720
            TabIndex        =   14
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   2355
            CapAlign        =   2
            BackStyle       =   3
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
            cGradient       =   11043601
            Gradient        =   2
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H lvButtons_H5 
            Height          =   1335
            Left            =   1200
            TabIndex        =   15
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   2355
            CapAlign        =   2
            BackStyle       =   3
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
            cGradient       =   11043601
            Gradient        =   2
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H lvButtons_H6 
            Height          =   1335
            Left            =   6720
            TabIndex        =   16
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   2355
            CapAlign        =   2
            BackStyle       =   3
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
            cGradient       =   11043601
            Gradient        =   2
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label2 
            BackColor       =   &H00616161&
            Height          =   135
            Left            =   0
            TabIndex        =   11
            Top             =   2160
            Width           =   7815
         End
         Begin VB.Label Label9 
            BackColor       =   &H00A88311&
            BackStyle       =   0  'Transparent
            Caption         =   "Software Providers of Hospitals,Banks,Supermarkets,Hotels,NGO's,etc"
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
            Height          =   495
            Left            =   1920
            TabIndex        =   9
            Top             =   1080
            Width           =   4215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00A88311&
            BackStyle       =   0  'Transparent
            Caption         =   "Developed by:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1920
            TabIndex        =   8
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label7 
            BackColor       =   &H00A88311&
            BackStyle       =   0  'Transparent
            Caption         =   "Abukari Yakubu Of"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1920
            TabIndex        =   7
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label1 
            BackColor       =   &H00A88311&
            BackStyle       =   0  'Transparent
            Caption         =   "IT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label4 
            BackColor       =   &H00A88311&
            BackStyle       =   0  'Transparent
            Caption         =   "Xtreme"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1920
            TabIndex        =   17
            Top             =   840
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00A88311&
         Caption         =   "Version 3.5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3015
         TabIndex        =   3
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00A88311&
         Caption         =   "POINT OF SALE MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   3930
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00A88311&
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00A88311&
         Caption         =   "This Software Package is Licensed to OFRAM ENTERPRISE Under The International Software License Agreement Rules."
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
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   3000
         Width           =   6315
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
PgBar.Value = PgBar.Value + 4
If PgBar.Value = 100 Then
Unload Me
frmMDI.Show
frmLogin.Show

End If
End Sub
