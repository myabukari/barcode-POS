VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00404040&
   Caption         =   "S M Y S"
   ClientHeight    =   4365
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10650
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMDI.frx":030A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CdgBackup 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select The Destination Folder Where The BMAS DATABASE Will Backed UP"
      Filter          =   "Database Files (*.MDB) | *.MDB"
      InitDir         =   "C:"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "INS"
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/14/2007"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:03 AM"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10650
      TabIndex        =   2
      Top             =   255
      Width           =   10650
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   10590
      TabIndex        =   0
      Top             =   0
      Width           =   10650
      Begin VB.TextBox txtU 
         Height          =   285
         Left            =   12600
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00F2EBBF&
         Caption         =   "O F R A M    E N T E R P R I S E"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0062490F&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   16095
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuRetail 
      Caption         =   "&Retail"
      Begin VB.Menu mnuSellOut 
         Caption         =   "Sell Out"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "R&etail-Reports"
      Begin VB.Menu mnuRetailStockLevel 
         Caption         =   "Retail StockLevel"
      End
      Begin VB.Menu mnusales 
         Caption         =   "Sales Revenue Report"
      End
      Begin VB.Menu mnuProductsReport 
         Caption         =   "Products Report"
      End
      Begin VB.Menu mnurecievals 
         Caption         =   "Stocking Retail Report"
      End
   End
   Begin VB.Menu mnuSetUps 
      Caption         =   "Set-Ups"
      Begin VB.Menu mnuSetUpRetailProducts 
         Caption         =   "SetUp Retail Products"
      End
      Begin VB.Menu mnuSetUpStores 
         Caption         =   "SetUp Stores"
      End
      Begin VB.Menu mnuSetUpWholeStocklevel 
         Caption         =   "SetUp Whole Stocklevel"
      End
      Begin VB.Menu mnuSeUptDebtors 
         Caption         =   "SetUp Debtors"
      End
      Begin VB.Menu mnuSetUpUsers 
         Caption         =   "Set Up Users"
      End
      Begin VB.Menu mnuSetUpSuppliers 
         Caption         =   "SetUp Suppliers"
      End
   End
   Begin VB.Menu mnuWholeSale 
      Caption         =   "&WholeSale"
      Begin VB.Menu mnuReceivals 
         Caption         =   "Recievals From Suppliers"
      End
      Begin VB.Menu mnuRestockRetail 
         Caption         =   "Supply to Retail"
      End
      Begin VB.Menu mnuGoodsTransfer 
         Caption         =   "Goods Transfer from WholeSales"
      End
   End
   Begin VB.Menu mnuWholeSaleReports 
      Caption         =   "W&holeSale-Reports"
      Begin VB.Menu mnuReceivalsReport 
         Caption         =   "Receivals Report"
      End
      Begin VB.Menu mnuWholeSaleStockLevel 
         Caption         =   "WholeSale StockLevels Report"
      End
      Begin VB.Menu mnuStockingRetail 
         Caption         =   "Stocking Retail Report"
      End
      Begin VB.Menu mnuGoodsTransferReport 
         Caption         =   "Goods Transfer from WholeSales Report"
      End
      Begin VB.Menu mnuWholeSaleWorth 
         Caption         =   "WholeSale Worth Report"
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "Accounts"
      Visible         =   0   'False
      Begin VB.Menu mnuSetDebtors 
         Caption         =   "Set Debtors"
      End
      Begin VB.Menu mnuDebtorsPayments 
         Caption         =   "Debtors Payments"
      End
      Begin VB.Menu mnuPaymentsReport 
         Caption         =   "Payments Report"
      End
      Begin VB.Menu mnuDebtorsBalances 
         Caption         =   "Debtors Balances Report"
      End
   End
   Begin VB.Menu mnuManagement 
      Caption         =   "&Management"
      Visible         =   0   'False
      Begin VB.Menu mnuSetSuppliers 
         Caption         =   "Set Suppliers/Creditors"
      End
      Begin VB.Menu mnuSetRetailProducts 
         Caption         =   "Set Retail Products"
      End
      Begin VB.Menu mnuUserAccounts 
         Caption         =   "User Accounts"
      End
      Begin VB.Menu mnuViewUsersDetails 
         Caption         =   "View Users Details"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuCreditorAccounts 
         Caption         =   "Creditor Accounts"
      End
      Begin VB.Menu mnuPricingProducts 
         Caption         =   "Pricing Products"
      End
   End
   Begin VB.Menu mnuManagementReports 
      Caption         =   "Ma&nagement-Reports"
      Visible         =   0   'False
      Begin VB.Menu mnuCreditorAccountsReports 
         Caption         =   "Creditor Accounts Reports"
      End
      Begin VB.Menu mnuDebtorAccountsReports 
         Caption         =   "Debtor Accounts Reports"
      End
      Begin VB.Menu mnuChequeDueDates 
         Caption         =   "ChequeDueDates for Today"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
mnufile.Enabled = False
Me.mnuRetail.Enabled = False
mnureports.Enabled = False
Me.mnuWholeSale.Enabled = False
Me.mnuWholeSaleReports.Enabled = False
Me.mnuManagement.Enabled = False
Me.mnuManagementReports.Enabled = False
Me.mnuSetUps.Enabled = False

Me.mnuAccounts.Enabled = False
frmLogin.Show
End Sub

Private Sub mnuBackup_Click()
'Dim strg1 As String, strg2 As String, fsDrive As Drive, fsFolder As Folder, fsFile As File, fsVal As New FileSystemObject, Fsize1, Fsize2 As Single
'CdgBackup.DialogTitle = "Select The Databse('Product') To Backup"
'CdgBackup.FileName = App.Path & "\database\Product.mdb"
'Me.CdgBackup.ShowOpen
'If Me.CdgBackup.FileTitle <> "" Then
   'If fsVal.FileExists(Me.CdgBackup.FileName) Then
      'strg1 = Me.CdgBackup.FileName
      'Set fsFile = fsVal.GetFile(strg1)
     ' Fsize1 = Round(fsFile.Size / 1000000, 2)
   'Else
      'If Me.CdgBackup.FileTitle = "Product.mdb" Then
        ' MsgBox "The Database " & "'Product.mdb'" & " Does Not Exist In " & Left(Me.CdgBackup.FileName, Len(Me.CdgBackup.FileName) - Len(Me.CdgBackup.FileTitle) - 1) & vbCrLf & "Please Check Very Well For The Correct Folder", vbInformation
     'Else
         'MsgBox Me.CdgBackup.FileTitle & " Is Not The Correct Name Of The Database(Product)" & vbCrLf & "The Correct Name of The Database File Is: 'Product'", vbInformation
      'End If
      'Exit Sub
   'End If
'Else
  'Exit Sub
'End If
'CdgBackup.DialogTitle = "Select The Folder Where The Database " & "' " & CdgBackup.FileTitle & " '" & " Would Be Backup To"
'Me.CdgBackup.ShowSave
'If Me.CdgBackup.FileTitle <> "" Then
   'If Me.CdgBackup.FileTitle = "DB2.mdb" Then
     'Set fsDrive = fsVal.GetDrive(fsVal.GetDriveName(Me.CdgBackup.FileName))
      'Fsize2 = Round(fsDrive.AvailableSpace / 1000000, 2)
      'If Fsize2 >= Fsize1 Then
        ' Fsize1 = Round(Fsize2 - Fsize1, 2)
        ' FrmBackUps.Label1 = FrmBackUps.Label1 & " " & fsDrive.DriveLetter & ":\"
         'FrmBackUps.Show
         'Do
           'DoEvents
           'Call fsVal.CopyFile(strg1, Me.CdgBackup.FileName, True)
         'Loop While (Round(fsDrive.AvailableSpace / 1000000, 2) > Fsize1)
         'Unload FrmBackUps
     ' Else
         ' If fsVal.FileExists(Me.CdgBackup.FileName) Then
          
            'Call fsVal.DeleteFile(Me.CdgBackup.FileName)
          'End If
      'End If
   'Else
     ' MsgBox "Select Or Enter The Database File Name" & "(" & "'Product'" & ")"
   'End If
'End If
End Sub

Private Sub mnuCashRegisterSalesReport_Click()
frmCashRegisterRpt.Show
End Sub

Private Sub mnuChange_Click()
frmChangePassword.Show
End Sub

Private Sub mnuChangePassword_Click()
frmChangePassword.Show
End Sub

Private Sub mnuChequeDueDates_Click()
frmDebtorsChequeDate.Show
End Sub

Private Sub mnuCreditorAccounts_Click()
frmCreditors.Show
End Sub

Private Sub mnuCreditorAccountsReports_Click()
frmCreditorsRpt.Show
End Sub

Private Sub mnuDebtorAccounts_Click()
frmDebtors.Show
End Sub

Private Sub mnuDebtorAccountsReports_Click()
frmAccountsRpt.Show
End Sub

Private Sub mnuDebtorsBalances_Click()
frmIndividualDebtorPaymentsRpt.Show
End Sub

Private Sub mnuDebtorsPayments_Click()
frmDebtorsPayments.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnupassword_Click()

End Sub

Private Sub mnuGoodsTransfer_Click()
frmGoodsTransfer.Show
End Sub

Private Sub mnuGoodsTransferReport_Click()
frmGoodsTransferRpt.Show
End Sub

Private Sub mnuPaymentsReport_Click()
frmDebtorsPaymentsRpt.Show
End Sub

Private Sub mnuPricingProducts_Click()
frmPricing.Show
End Sub

Private Sub mnuProductsReport_Click()
frmProductsRpt.Show
End Sub

Private Sub mnuReceivals_Click()
frmReceivals.Show
End Sub

Private Sub mnuReceivalsReport_Click()
frmRecievalsRpt.Show
End Sub

Private Sub mnurecievals_Click()
frmStockingRetailRpt.Show
End Sub

Private Sub mnuRestock_Click()
frmReOderProducts.Show
End Sub

Private Sub mnuRestockRetail_Click()
frmStockRetail.Show
End Sub

Private Sub mnuRetailStockLevel_Click()
frmStockLevelRpt.Show
End Sub

Private Sub mnusales_Click()
frmRevenueRpt.Show
End Sub

Private Sub mnuSellOut_Click()
'Dim cfom As Form
'Set cfom = New frmSellOut
frmSellProducts.Show
'frmSellOut.Show
'CenterForm (frmSellOut)
End Sub

Private Sub mnuSetDebtors_Click()
frmDebtor.Show
End Sub

Private Sub mnuSetProducts_Click()
frmProducts.Show
End Sub

Private Sub mnuSetRetailProducts_Click()
frmProducts.Show
End Sub

Private Sub mnuSetSuppliers_Click()
frmsuppliers.Show
End Sub

Private Sub mnuSetUpWholeSaleProducts_Click()
frmMainStoreProducts.Show
End Sub

Private Sub mnuSetUpRetailProducts_Click()
frmProducts.Show
End Sub

Private Sub mnuSetUpStores_Click()
frmStores.Show
End Sub

Private Sub mnuSetUpSuppliers_Click()
frmsuppliers.Show
End Sub

Private Sub mnuSetUpUsers_Click()
frmUsers.Show
End Sub

Private Sub mnuSetUpWholeStocklevel_Click()
frmAdjustWholeSaleStock.Show
End Sub

Private Sub mnuSeUptDebtors_Click()
frmDebtor.Show
End Sub

Private Sub mnuStockingRetail_Click()
frmStockingRetailRpt.Show
End Sub

Private Sub mnuStockLevel_Click()
frmStockLevelRpt.Show
End Sub

Private Sub mnusuppliers_Click()
frmsuppliers.Show
End Sub

Private Sub mnuUseCashRegister_Click()
frmCashRegister.Show
End Sub

Private Sub mnuUserAccounts_Click()
frmUsers.Show
End Sub

Private Sub mnuViewUsersDetails_Click()
frmUsersInfo.Show
End Sub

Private Sub mnuWholeSaleStockLevel_Click()
frmWholeSaleStockRpt.Show
End Sub

Private Sub mnuWholeSaleWorth_Click()
frmWholeSaleWorth.Show
End Sub
