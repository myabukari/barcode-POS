VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmRecievalsRpt 
   BackColor       =   &H00C29E21&
   Caption         =   "WholeSale Recievals Reports"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   Icon            =   "frmRecievals.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   8010
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   7215
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
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
         Format          =   16384003
         CurrentDate     =   39087
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   720
         TabIndex        =   9
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
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
         Format          =   16384003
         CurrentDate     =   39087
      End
      Begin VB.Label Label4 
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
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
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
         TabIndex        =   10
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
   Begin Crystal.CrystalReport CrystalReceivals 
      Left            =   6840
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C29E21&
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Width           =   7215
      Begin lvButton_H.lvButtons_H cmdOk 
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
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
         Image           =   "frmRecievals.frx":030A
         cBack           =   -2147483633
      End
      Begin lvButton_H.lvButtons_H cmdExit 
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
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
         Image           =   "frmRecievals.frx":118C
         cBack           =   -2147483633
      End
      Begin VB.CommandButton cmdOk1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1200
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Caption         =   "SELECT STORE PRODUCT TO VIEW REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      Begin VB.ComboBox cboStore 
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
         Left            =   1800
         TabIndex        =   12
         Text            =   "All"
         Top             =   360
         Width           =   4095
      End
      Begin VB.ComboBox cboProducts 
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
         Left            =   1800
         TabIndex        =   2
         Text            =   "All"
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Store Name:"
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
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C29E21&
         Caption         =   "Product Name:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmRecievalsRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection, rs As New ADODB.Recordset, i As Integer
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer
Private Sub cboProducts_DropDown()
On Error GoTo OkError


'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

Me.cboProducts.Clear

 rs.Open "Select Distinct ProductName from Products  Order By ProductName Asc", cn, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
   rs.MoveFirst
   For i = 1 To rs.RecordCount
     Me.cboProducts.AddItem rs.Fields!ProductName
     
     rs.MoveNext
   Next
   If cboProducts.ListCount > 1 Then
      Me.cboProducts.AddItem "All"
   End If
End If
If rs.State = 1 Then
   rs.Close
End If
Exit Sub
OkError:
       If rs.State = 1 Then
          rs.Close
       End If
       MsgBox "TRY AGAIN", vbInformation, "Receivals REPORTS"
       Exit Sub



End Sub

Private Sub cboStore_DropDown()
On Error GoTo OkError


'Open Connecttion to Server
bFlag = OpenConnection(cn, strg)

If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   MsgBox strg, vbInformation:
   Exit Sub
End If

Me.cboStore.Clear

 rs.Open "Select Distinct StoreName from Stores  Order By StoreName Asc", cn, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
   rs.MoveFirst
   For i = 1 To rs.RecordCount
     Me.cboStore.AddItem rs.Fields!StoreName
     
     rs.MoveNext
   Next
   If cboStore.ListCount > 1 Then
      Me.cboStore.AddItem "All"
   End If
End If
If rs.State = 1 Then
   rs.Close
End If
Exit Sub
OkError:
       If rs.State = 1 Then
          rs.Close
       End If
       MsgBox "Sorry cannot display Store Names,Please try again ", vbInformation, "Store Names"
       Exit Sub


End Sub

Private Sub cmdExit_Click()
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdOk_Click()
'On Error GoTo OkError

If Me.dtpfrom > Me.dtpto Then
MsgBox "WRONG DATE RANGE ENTERED", vbInformation, "DATE RANGE"
Exit Sub
End If



If Me.dtpfrom <= Me.dtpto Then
bFlag = OpenConnection(cn, strg)
If bFlag = False Then
   If cn.State = 1 Then cn.Close
   If rs.State = 1 Then rs.Close
   Me.MousePointer = vbDefault
 
   MsgBox strg, vbInformation:
   Exit Sub
End If


rs.Open "Select  * From Receivals where Date >='" & Me.dtpfrom & "' and Date <='" & Me.dtpto & "'  order by Date desc", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
           
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED DATE", vbInformation, "DATE RANGE"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close

If Me.cboProducts <> "All" Then
rs.Open "Select  * From Receivals inner Join Products on Receivals.ProductID=Products.ProductID  where ProductName='" & Me.cboProducts & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED PRODUCT", vbInformation, "PRODUCT NOT IN REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close


If Me.cboStore <> "All" Then
rs.Open "Select  * From Stores inner Join Receivals on Stores.StoreID=Receivals.StoreID  where StoreName='" & Me.cboStore & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED STORE", vbInformation, "STORE NOT IN REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close


If Me.cboProducts <> "All" And Me.cboStore <> "All" Then
rs.Open "Select  * From (Products inner Join Receivals on Products.ProductID=Receivals.ProductID) inner Join Stores on Stores.StoreID=Receivals.StoreID  where ProductName='" & Me.cboProducts & "' and StoreName='" & Me.cboStore & "'", cn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
               MsgBox "THERE ARE NO REPORTS FOR THE SPECIFIED FIELDS", vbInformation, "WHOLASALE STOCKLEVEL REPORTS"
            If rs.State = 1 Then rs.Close
            Exit Sub
            End If
        
If rs.State = 1 Then rs.Close
End If
If rs.State = 1 Then rs.Close



'CrystalReceivals.Connect = "DSN=supermarket;UID=;PWD=;DSQ=Products"
'CrystalReceivals.DataFiles(0) = App.Path & "\database\Product.mdb"
CrystalReceivals.ReportFileName = App.Path & "\Receivals.rpt"
CrystalReceivals.Connect = "DSN=nxomen;UID=sa;PWD=Abu;DSQ=ZuksData"




If Me.cboStore = "All" And Me.cboProducts <> "All" Then
CrystalReceivals.SelectionFormula = "{Receivals.Date} >=#" & Me.dtpfrom & "# and {Receivals.Date} <=#" & Me.dtpto & "# and {Products.ProductName} ='" & Me.cboProducts.Text & "'"
End If

If Me.cboStore <> "All" And Me.cboProducts = "All" Then
CrystalReceivals.SelectionFormula = "{Receivals.Date} >=#" & Me.dtpfrom & "# and {Receivals.Date} <=#" & Me.dtpto & "# and {Stores.StoreName} ='" & Me.cboStore & "'"
End If

If Me.cboStore <> "All" And Me.cboProducts <> "All" Then
CrystalReceivals.SelectionFormula = "{Receivals.Date} >=#" & Me.dtpfrom & "# and {Receivals.Date} <=#" & Me.dtpto & "# and {Stores.StoreName} ='" & Me.cboStore & "' and {Products.ProductName} ='" & Me.cboProducts.Text & "'"
End If

If Me.cboProducts = "All" And Me.cboStore = "All" Then
CrystalReceivals.SelectionFormula = "{Receivals.Date} >=#" & Me.dtpfrom & "# and {Receivals.Date} <=#" & Me.dtpto & "#"
End If

'If Me.cboProducts = "All" And Me.cboStore = "All" And Me.OptDate = True Then
'CrystalReceivals.SelectionFormula = "{Receivals.Date} >=#" & Me.dtpfrom & "# and {Receivals.Date} <=#" & Me.dtpto & "#"
'End If

   CrystalReceivals.WindowState = crptMaximized
   CrystalReceivals.WindowShowRefreshBtn = True
   Me.CrystalReceivals.WindowTitle = "RECEIVALS REPORT " & Format$(Date, "yyyy")
    
     
   CrystalReceivals.Action = 1

Exit Sub
OkError:
       
       MsgBox "THERE WAS A PROBLEM TRYING TO DISPLAY REPORT,PLEASE TRY AGAIN", vbInformation, "RECEIVALS REPORT"
       Exit Sub

End Sub

Private Sub Form_Load()
Me.dtpfrom = Date
Me.dtpto = Date
CenterForm Me
Me.Height = 5865
Me.Width = 8130
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub OptDate_Click()

End Sub
