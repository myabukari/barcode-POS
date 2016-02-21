VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmItems 
   Caption         =   "ITEMS"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   9615
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   600
         TabIndex        =   17
         Top             =   120
         Width           =   8175
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
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
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
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
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdClear 
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
            Left            =   5160
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
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
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "D&elete"
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
            Left            =   1800
            Picture         =   "frmItems.frx":0000
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   9615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgItems 
         Height          =   2055
         Left            =   5520
         TabIndex        =   20
         Top             =   2400
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"frmItems.frx":0102
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.TextBox txtdiscount 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Top             =   2340
         Width           =   2775
      End
      Begin VB.TextBox txtTotalQuantity 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7080
         TabIndex        =   4
         Top             =   1740
         Width           =   2055
      End
      Begin VB.TextBox txtbasePrice 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7080
         TabIndex        =   2
         Top             =   1140
         Width           =   2055
      End
      Begin VB.TextBox txtUnitPrice 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   1740
         Width           =   2775
      End
      Begin VB.TextBox txtbaseunit 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   1140
         Width           =   2775
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   0
         Top             =   420
         Width           =   6735
      End
      Begin VB.Label Label5 
         Caption         =   "Discount Per Item:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lbl 
         Caption         =   "Total Quantity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Price per cartone:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Unit Price:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Base Unit(# per cartone):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Name of Item:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command5_Click()

End Sub

Private Sub txtbasePrice_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.txtUnitPrice.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtbaseunit_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.txtbasePrice.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtdiscount_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.cmdSave.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()"
If KeyAscii = vbKeyReturn Then
   Me.txtbaseunit.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtTotalQuantity_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.txtdiscount.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.txtTotalQuantity.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub
