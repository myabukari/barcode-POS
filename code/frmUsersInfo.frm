VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmUsersInfo 
   BackColor       =   &H00C29E21&
   Caption         =   "Users Information"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   Icon            =   "frmUsersInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   9630
   Begin MSComctlLib.ListView lstUsers 
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   -2147483625
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UserName"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Accesslevel"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status Of Account"
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton_H.lvButtons_H cmdExit 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   8895
      _ExtentX        =   15690
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
      Image           =   "frmUsersInfo.frx":030A
      cBack           =   -2147483633
   End
   Begin VB.CommandButton cmdExit1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   8895
   End
End
Attribute VB_Name = "frmUsersInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer

Private Sub ListUsers()
On Error GoTo SaveError
Me.lstUsers.ListItems.Clear
'Open Connecttion to Server

bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
   
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select * From Users  Order By UserName", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.lstUsers.ListItems.Add(, , Trim(rs!UserName))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Trim(rs!UserID)
                List_Item.SubItems(2) = Trim(rs!UserPassword)
                List_Item.SubItems(3) = Trim(rs!AccessLevel)
                List_Item.SubItems(4) = Trim(rs!StatusOfAccount)
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
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub Form_Load()
Call ListUsers
CenterForm Me
Me.Width = 9750
Me.Height = 5220
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub
