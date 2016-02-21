VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93AA248C-8E0E-4697-879F-8A6D0D6838E8}#1.0#0"; "lvButton_H.ocx"
Begin VB.Form frmProducts 
   BackColor       =   &H00C29E21&
   Caption         =   "PRODUCT DETAILS"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   Icon            =   "frmProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   9930
   Begin MSComctlLib.ListView lstProducts 
      Height          =   3015
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Unit Price"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "StockLevel"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Discount Per Item"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Product Department"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ProductCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProductCode1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "NoPerCartone"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C29E21&
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   7320
      Width           =   9615
      Begin VB.Frame Frame4 
         BackColor       =   &H00C29E21&
         Height          =   735
         Left            =   360
         TabIndex        =   18
         Top             =   120
         Width           =   8895
         Begin lvButton_H.lvButtons_H cmdSave 
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            Caption         =   "&Save Particulars"
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
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmProducts.frx":030A
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdDelete 
            Height          =   375
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "D&elete"
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
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmProducts.frx":075C
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdFind 
            Height          =   375
            Left            =   4200
            TabIndex        =   26
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Find"
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
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmProducts.frx":0BAE
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdClear 
            Height          =   375
            Left            =   5760
            TabIndex        =   27
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmProducts.frx":1000
            cBack           =   -2147483633
         End
         Begin lvButton_H.lvButtons_H cmdExit 
            Height          =   375
            Left            =   7320
            TabIndex        =   28
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmProducts.frx":28B1
            cBack           =   -2147483633
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C29E21&
      Height          =   4215
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   9615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxgItems 
         Height          =   2055
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3625
         _Version        =   393216
         BackColor       =   15916706
         ForeColor       =   -2147483630
         Cols            =   16
         FixedCols       =   0
         BackColorFixed  =   8421504
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   12754465
         GridColor       =   -2147483628
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   $"frmProducts.frx":2D03
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
         _Band(0).Cols   =   16
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2DEA2&
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
         Left            =   7320
         TabIndex        =   4
         Top             =   1020
         Width           =   2055
      End
      Begin VB.TextBox txtUnitPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2DEA2&
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
         Left            =   2160
         TabIndex        =   1
         Top             =   1020
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C29E21&
         Height          =   1575
         Left            =   0
         TabIndex        =   21
         Top             =   2640
         Width           =   9615
         Begin VB.TextBox txtTotalQuantity 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2DEA2&
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
            Left            =   2160
            TabIndex        =   9
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox txtbaseunit 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2DEA2&
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
            Left            =   2160
            TabIndex        =   7
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txtNoCartones 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2DEA2&
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
            Left            =   4200
            TabIndex        =   8
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtStock 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2DEA2&
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
            Left            =   7320
            TabIndex        =   10
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtReorder 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2DEA2&
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
            Left            =   7320
            TabIndex        =   11
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C29E21&
            Caption         =   "Total Quantity:"
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
            Left            =   720
            TabIndex        =   34
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C29E21&
            Caption         =   "Number Per Cartone:"
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
            TabIndex        =   31
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C29E21&
            Caption         =   "Number Of Cartones:"
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
            Height          =   495
            Left            =   3120
            TabIndex        =   30
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C29E21&
            Caption         =   "ReorderLevel:"
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
            Left            =   6000
            TabIndex        =   23
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C29E21&
            Caption         =   "StockLevel:"
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
            Left            =   6240
            TabIndex        =   22
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2DEA2&
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
         Left            =   2160
         TabIndex        =   0
         Top             =   420
         Width           =   7215
      End
      Begin MSComCtl2.DTPicker dtpExp 
         Height          =   315
         Left            =   7440
         TabIndex        =   13
         Top             =   4680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   15454586
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22085635
         CurrentDate     =   39081
      End
      Begin VB.TextBox txtManufacterer 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2DEA2&
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
         Left            =   0
         TabIndex        =   14
         Top             =   3840
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpmanu 
         Height          =   315
         Left            =   7440
         TabIndex        =   12
         Top             =   4080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   15454586
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22085635
         CurrentDate     =   39081
      End
      Begin VB.ComboBox cboDepartment 
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
         ItemData        =   "frmProducts.frx":2E25
         Left            =   2160
         List            =   "frmProducts.frx":2E74
         TabIndex        =   3
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtProductCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2DEA2&
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
         Left            =   7320
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtbasePrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2DEA2&
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
         Left            =   2160
         TabIndex        =   2
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtdiscount 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2DEA2&
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
         Left            =   7320
         TabIndex        =   6
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C29E21&
         Caption         =   "Price per cartone:"
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
         Left            =   480
         TabIndex        =   38
         Top             =   1560
         Width           =   1575
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
         Left            =   720
         TabIndex        =   37
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C29E21&
         Caption         =   "Unit Price:"
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
         Left            =   1080
         TabIndex        =   36
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C29E21&
         Caption         =   "Product Department:"
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
         TabIndex        =   35
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C29E21&
         Caption         =   "Discount Per Item(%):"
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
         Left            =   5280
         TabIndex        =   33
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C29E21&
         Caption         =   "Product Code:"
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
         Left            =   6000
         TabIndex        =   32
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C29E21&
         Caption         =   "Bar Code:"
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
         Left            =   6360
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H005B4311&
         Caption         =   "Manufucterer:"
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
         Left            =   600
         TabIndex        =   20
         Top             =   3600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim bFlag As Boolean, strg As String, List_Item As ListItem, Productid As String
Dim sflag As Boolean, ListProductID As String, ctrl As Control, StockQty As Integer, ShowProductflag As Boolean
Private Sub Command5_Click()

End Sub

Private Sub cboDepartment_Click()
''If Me.cboDepartment = "AIR FRESHNERS" Then
''Me.txtProductCode = "AF"
''ElseIf cboDepartment = "BATTERIES" Then
''txtProductCode = "BA"
''ElseIf cboDepartment = "BISCUIT" Then
''txtProductCode = "B"
''ElseIf cboDepartment = "CANNED DELIGHTS" Then
''txtProductCode = "CD"
''ElseIf cboDepartment = "COSMETICS" Then
''txtProductCode = "C"
''ElseIf cboDepartment = "CUTLERY" Then
''txtProductCode = "CU"
''ElseIf cboDepartment = "DRINKS" Then
''txtProductCode = "D"
''ElseIf cboDepartment = "FLAVOURS" Then
''txtProductCode = "F"
''ElseIf cboDepartment = "FLOUR" Then
''txtProductCode = "FL"
''ElseIf cboDepartment = "INSECTICIDES" Then
''txtProductCode = "I"
''ElseIf cboDepartment = "LIGHTERS" Then
''txtProductCode = "L"
''ElseIf cboDepartment = "MAGARINES" Then
''txtProductCode = "MA"
''ElseIf cboDepartment = "MILK" Then
''txtProductCode = "M"
''ElseIf cboDepartment = "OILS" Then
''txtProductCode = "O"
''ElseIf cboDepartment = "PASTAS" Then
''txtProductCode = "P"
''ElseIf cboDepartment = "PERFUMES" Then
''txtProductCode = "PE"
''ElseIf cboDepartment = "POWDERS" Then
''txtProductCode = "PO"
''ElseIf cboDepartment = "RICE" Then
''txtProductCode = "R"
''ElseIf cboDepartment = "SALT" Then
''txtProductCode = "SA"
''ElseIf cboDepartment = "SOAPS" Then
''txtProductCode = "S"
''ElseIf cboDepartment = "SUGAR" Then
''txtProductCode = "SU"
''ElseIf cboDepartment = "SWEETS" Then
''txtProductCode = "SW"
''ElseIf cboDepartment = "TEA" Then
''txtProductCode = "T"
''ElseIf cboDepartment = "TOILETARIES" Then
''txtProductCode = "TO"
''ElseIf cboDepartment = "TEA" Then
''txtProductCode = "T"
''ElseIf cboDepartment = "TOMATO PASTE" Then
''txtProductCode = "TP"
''
''End If
End Sub

Private Sub cboDepartment_GotFocus()
Me.cboDepartment.BackColor = &HFFFF&
End Sub

Private Sub cboDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 40 Then
'Me.txtCode.SetFocus
'End If
If KeyCode = 38 Then
Me.txtbasePrice.SetFocus
End If
End Sub

Private Sub cboDepartment_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()@#$%&*_-"
If KeyAscii = vbKeyReturn Then
   Me.txtCode.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub cboDepartment_LostFocus()
Me.cboDepartment.BackColor = &HF2EBBF
End Sub

Private Sub cmdClear_Click()
 Call ClearCtrls
 Me.txtname.SetFocus
 sflag = False
 Me.cmdSave.Caption = "Save Particulars"
 Me.flxgItems.Visible = False
End Sub

Private Sub cmdDelete_Click()
If Me.txtname = "" Then
MsgBox "ENTER PRODUCT TO BE DELETED", vbInformation, "DELETE WHAT"
Me.txtname.SetFocus: Exit Sub
End If
If MsgBox("ARE YOU SURE  YOU WANT TO DELETE THE PRODUCTS DETAILS?", vbYesNo + vbQuestion, "CONFIRM DELETE") = vbYes Then
   
   Me.cmdDelete.Enabled = False
   
   On Error GoTo SaveError
       
   'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      Me.MousePointer = vbDefault
      Me.cmdDelete.Enabled = True
      MsgBox strg, vbInformation:
      Exit Sub
   End If
   
   cn.BeginTrans
    cn.Execute "Delete From Products Where ProductID ='" & Productid & "'", Y
    If Y > 0 Then
     cn.Execute "Delete From ProductInventory Where ProductID ='" & Productid & "'", Y
    End If
    If Y > 0 Then
      cn.CommitTrans
      MsgBox "Delete Successful!", vbInformation, "Delete Successful"
     'Clear ctrls and setfocus to supplier name ctrl
      ClearCtrls
      Me.txtname.SetFocus
   Else
      cn.RollbackTrans
      MsgBox "Sorry, Unable to Delete Products Details:Please Try Again!", vbInformation, "Delete Failed"
   End If
      Me.txtname.SetFocus
      sflag = False
      Me.cmdSave.Caption = "Save Particulars"
      Me.cmdDelete.Enabled = False
      If cn.State = 1 Then cn.Close
   
      Me.MousePointer = vbDefault
      Call ListProducts
      Exit Sub
SaveError:
        If cn.State = 1 Then cn.Close
        Me.MousePointer = vbDefault
        Me.cmdDelete.Enabled = True
   
        MsgBox "Sorry, Unable to Delete Products  Details:Please Try Again!", vbInformation, "Delete Failed"
        Exit Sub
End If

End Sub

Private Sub cmdExit_Click()
'If MsgBox("ARE YOU SURE  YOU WANT TO ClOSE?", vbYesNo + vbQuestion, "CONFIRM ClOSE") = vbYes Then
Unload Me
'End If
End Sub

Private Sub cmdFind_Click()
Dim strg1 As String
On Error GoTo SaveError

strg1 = InputBox("Enter The ProductName Or The First Few or All of the Characters of The ProductName.", "ProductName")
If Trim(strg1) <> "" Then

   
   Me.cmdFind.Enabled = False
          
   'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
      Me.cmdFind.Enabled = True
       MsgBox strg, vbInformation:
      Exit Sub
   End If
   strg1 = strg1 & "%"
     rs.Open "Select Products.*,ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID = ProductInventory.ProductID Where ProductName Like '" & strg1 & "' order by ProductName", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount <= 0 Then
      rs.Close: cn.Close
      MsgBox "There is No Products with the Name Entered.", vbInformation, "Search Failed"
      Me.MousePointer = vbDefault: Me.cmdFind.Enabled = True: Me.cmdFind.SetFocus: Exit Sub
   Else
      If rs.RecordCount = 1 Then
         sflag = True
         Me.cmdSave.Caption = "Edit Particulars"
         ShowProductflag = True
         Me.txtname = rs.Fields("ProductName")
         ShowProductflag = False
         Me.txtbaseunit = rs.Fields("BaseUnit")
         Me.txtUnitPrice = rs.Fields("UnitPrice")
         Me.txtTotalQuantity = "" 'rs.Fields("TotalQuantity")
         Me.txtbasePrice = rs.Fields("PricePerCartone")
         Me.txtDiscount = rs.Fields("Discount")
         Me.txtManufacterer = rs.Fields("Manufacturer")
         Me.dtpExp = rs.Fields("ExpiryDate")
         Me.dtpmanu = rs.Fields("ManufucteryDate")
         Me.txtNoCartones = rs.Fields("NoOfCartones")
         Productid = rs.Fields("ProductID")
         Me.txtStock = rs.Fields("StockLevel")
         Me.txtReorder = rs.Fields("ReorderLevel")
         Me.cboDepartment = rs.Fields("Department")
         If rs.Fields("ProductCode") <> "" Then
          Me.txtCode = rs.Fields("ProductCode")
         Else
          Me.txtCode = ""
         End If
         If rs.Fields("ProductCode1") <> "" Then
          Me.txtProductCode = rs.Fields("ProductCode1")
         Else
          Me.txtProductCode = ""
         End If
         rs.Close
        
         
         Me.txtname.SetFocus
         Me.cmdDelete.Enabled = True
        
      Else
         rs.MoveFirst
         Me.flxgItems.Rows = rs.RecordCount + 1
         For X = 1 To rs.RecordCount
           Me.flxgItems.TextMatrix(X, 0) = rs.Fields("ProductName")
           Me.flxgItems.TextMatrix(X, 1) = rs.Fields("BaseUnit")
           Me.flxgItems.TextMatrix(X, 2) = rs.Fields("UnitPrice")
           Me.flxgItems.TextMatrix(X, 3) = rs.Fields("PricePerCartone")
           Me.flxgItems.TextMatrix(X, 4) = rs.Fields("TotalQuantity")
           Me.flxgItems.TextMatrix(X, 5) = rs.Fields("Discount")
           Me.flxgItems.TextMatrix(X, 6) = rs.Fields("Manufacturer")
           Me.flxgItems.TextMatrix(X, 7) = rs.Fields("ExpiryDate")
           Me.flxgItems.TextMatrix(X, 8) = rs.Fields("ManufucteryDate")
           Me.flxgItems.TextMatrix(X, 9) = rs.Fields("NoOfCartones")
           Me.flxgItems.TextMatrix(X, 10) = rs.Fields("StockLevel")
           Me.flxgItems.TextMatrix(X, 11) = rs.Fields("ReorderLevel")
           Me.flxgItems.TextMatrix(X, 12) = rs.Fields("ProductID")
           If rs.Fields("Department") <> "" Then
           Me.flxgItems.TextMatrix(X, 13) = rs.Fields("Department")
           End If
           If rs.Fields("ProductCode") <> "" Then
           Me.flxgItems.TextMatrix(X, 14) = rs.Fields("ProductCode")
           Else
           Me.flxgItems.TextMatrix(X, 14) = ""
           End If
           
           If rs.Fields("ProductCode1") <> "" Then
           Me.flxgItems.TextMatrix(X, 15) = rs.Fields("ProductCode1")
           Else
           Me.flxgItems.TextMatrix(X, 15) = ""
           End If
           rs.MoveNext
           
         Next
         Me.flxgItems.Visible = True
         Me.flxgItems.SetFocus
         rs.Close
      End If
   End If
   

End If

If cn.State = 1 Then cn.Close
If rs.State = 1 Then rs.Close

Me.MousePointer = vbDefault
Me.cmdFind.Enabled = True
Me.cmdSave.Enabled = True
Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     Me.MousePointer = vbDefault
     Me.cmdFind.Enabled = True
     MsgBox "Sorry, Unable to Find Products Details:Please Try Again!", vbInformation, "Search Failed"
     Exit Sub
End Sub

Private Sub cmdFind1_Click()

End Sub

Private Sub cmdSave_Click()
 
If Trim(Me.txtname) = "" Then
   MsgBox "YOU MUST ENTER PRODUCT NAME.", vbInformation, "PRODUCT NAME"
   Me.txtname.SetFocus: Exit Sub
End If


If Trim(Me.txtCode) = "" And Trim(Me.txtProductCode) = "" Then
   MsgBox "YOU MUST ENTER PRODUCT CODE.", vbInformation, "PRODUCT CODE"
   Me.txtCode.SetFocus: Exit Sub
End If

'''If Trim(Me.txtCode) = "NULL" And Trim(Me.txtProductCode) = "" Then
'''   MsgBox "YOU MUST ENTER PRODUCT CODE.", vbInformation, "PRODUCT CODE"
'''   Me.txtProductCode.SetFocus: Exit Sub
'''End If
'''
'''If Trim(Me.txtCode) = "" And Trim(Me.txtProductCode) = "NULL" Then
'''   MsgBox "YOU MUST ENTER PRODUCT CODE.", vbInformation, "PRODUCT CODE"
'''   Me.txtCode.SetFocus: Exit Sub
'''End If
'''
'''If Trim(Me.txtCode) = "NULL" And Trim(Me.txtProductCode) = "NULL" Then
'''   MsgBox "YOU MUST ENTER PRODUCT CODE.", vbInformation, "PRODUCT CODE"
'''   Me.txtCode.SetFocus: Exit Sub
'''End If
'''
'''If Trim(Me.txtCode) = "" Then
'''   Me.txtCode = "NULL"
'''End If
'''
'''If Trim(Me.txtProductCode) = "" Then
'''   Me.txtProductCode = "NULL"
'''End If

If Trim(Me.txtUnitPrice) = "" Then
   MsgBox "YOU MUST ENTER UNIT PRICE.", vbInformation, "UNIT PRICE"
   Me.txtUnitPrice.SetFocus: Exit Sub
End If

'If Trim(Me.cboDepartment) = "" Then
'   MsgBox "YOU MUST ENTER PRODUCT DEPARTMENT.", vbInformation, "PRODUCT DEPARTMENT"
'   Me.cboDepartment.SetFocus: Exit Sub
'End If

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

If Trim(Me.txtTotalQuantity) <> "" Then
Me.txtStock = Trim(Val(Me.txtTotalQuantity))
End If

If sflag = False Then
   'save part
rs.Open "Select ProductName From Products Where ProductName ='" & Trim(Me.txtname) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "A Product Has Already Been Setup with the Name.", vbInformation, ""
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtname.SetFocus: Exit Sub
   
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
End If
   If rs.State = 1 Then rs.Close
   
If Me.txtCode <> "" Then
rs.Open "Select ProductCode From Products Where ProductCode ='" & Trim(Me.txtCode) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "A Product Has Already Been Setup with the Bar Code.", vbInformation, ""
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtCode.SetFocus: Exit Sub
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
End If
   If rs.State = 1 Then rs.Close
End If


If Me.txtProductCode <> "" Then
rs.Open "Select ProductCode1 From Products Where ProductCode1 ='" & Trim(Me.txtProductCode) & "'", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.Close: cn.Close
   MsgBox "A Product Has Already Been Setup with the ProductCode.", vbInformation, ""
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
   Me.txtProductCode.SetFocus: Exit Sub
   Me.MousePointer = vbDefault
   Me.cmdSave.Enabled = True
End If
   If rs.State = 1 Then rs.Close
End If


   Call Generate_ProductID(Productid)
   
   cn.BeginTrans
   cn.Execute "Insert Into Products ([ProductID],[ProductName],[BaseUnit],[UnitPrice],[TotalQuantity],[PricePerCartone],[Discount],[UserID],[Date],[Manufacturer],[ExpiryDate],[ManufucteryDate],[NoOfCartones],[Department],ProductCode,ProductCode1) select '" & Trim(Productid) & "','" & Trim(Me.txtname.Text) & "','" & Val(Me.txtbaseunit.Text) & "','" & Val(Trim(Me.txtUnitPrice.Text)) & "','" & Val(Me.txtTotalQuantity.Text) & "','" & Val(Me.txtbasePrice.Text) & "','" & Val(Me.txtDiscount.Text) & "','111','" & Date & "','" & Trim(Me.txtManufacterer.Text) & "','" & Trim(Me.dtpExp) & "','" & Trim(Me.dtpmanu) & "','" & Val(Me.txtNoCartones.Text) & "','" & Trim(Me.cboDepartment.Text) & "','" & Trim(Me.txtCode.Text) & "','" & Trim(Me.txtProductCode.Text) & "'", Y
   If Y > 0 Then
    cn.Execute "Insert Into ProductInventory ([ProductID],[StockLevel],[ReorderLevel],[UserID],[Date]) select '" & Trim(Productid) & "','" & Val(Me.txtStock.Text) & "','" & Val(Me.txtReorder.Text) & "','111','" & Date & "'", Y
   End If
   If Y > 0 Then
   cn.CommitTrans
     MsgBox "Product Particulars Saved Successfully!", vbInformation, "Save Successful"
       Call ClearCtrls
       Call ListProducts
     Me.txtname = "": Me.txtbasePrice = "": Me.txtbaseunit = "": Me.txtDiscount = ""
     Me.txtManufacterer = "": Me.txtNoCartones = "": Me.txtReorder = "": Me.txtStock = ""
     Me.txtTotalQuantity = "": Me.txtUnitPrice = ""
     Me.txtname.SetFocus
   Else
     cn.RollbackTrans
     MsgBox "Sorry, Unable to Save Product Details:Please Try Again!", vbInformation, "Save Failed"
     Me.txtname.SetFocus
   End If
 
   
Else
'edit part

   rs.Open "Select ProductName From Products Where ProductName ='" & Trim(Me.txtname) & "' and ProductID<>'" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.Close: cn.Close
      MsgBox "A Product Has Already Been Setup with the Name.", vbInformation, ""
      Me.txtname.SetFocus: Exit Sub
   End If
      If rs.State = 1 Then rs.Close
   
   
   If Me.txtCode <> "" Then
   rs.Open "Select ProductCode From Products Where ProductCode ='" & Trim(Me.txtCode) & "' and ProductID<>'" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.Close: cn.Close
      MsgBox "A Product Has Already Been Setup with the Code.", vbInformation, ""
      Me.txtCode.SetFocus: Exit Sub
   End If
      If rs.State = 1 Then rs.Close
   End If
   
   
   If Me.txtProductCode <> "" Then
   rs.Open "Select ProductCode1 From Products Where ProductCode1 ='" & Trim(Me.txtProductCode) & "' and ProductID<>'" & Productid & "' ", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
     rs.Close: cn.Close
     MsgBox "A Product Has Already Been Setup with the ProductCode.", vbInformation, ""
     Me.MousePointer = vbDefault
     Me.cmdSave.Enabled = True
     Me.txtProductCode.SetFocus: Exit Sub
     Me.MousePointer = vbDefault
     Me.cmdSave.Enabled = True
   End If
     If rs.State = 1 Then rs.Close
   End If
   
   
   rs.Open "Select ProductInventory.StockLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Where  Products.ProductID='" & Productid & "'", cn, adOpenForwardOnly, adLockReadOnly
   If rs.RecordCount > 0 Then
      If Trim(Me.txtTotalQuantity) <> "" Then
      StockQty = Val(Trim(Me.txtTotalQuantity)) + rs.Fields("StockLevel")
      Else
      StockQty = Val(Trim(Me.txtStock))
      End If
      rs.Close
   End If
  
   
   cn.BeginTrans
   cn.Execute "Update Products Set ProductName ='" & Trim(Me.txtname.Text) & "',BaseUnit='" & Val(Trim(Me.txtbaseunit.Text)) & "',UnitPrice='" & Val(Trim(Me.txtUnitPrice.Text)) & "',TotalQuantity='" & Val(Trim(Me.txtTotalQuantity.Text)) & "',PricePerCartone='" & Val(Trim(Me.txtbasePrice.Text)) & "',Discount='" & Val(Trim(Me.txtDiscount.Text)) & "',Manufacturer='" & Trim(Me.txtManufacterer.Text) & "',ExpiryDate='" & Trim(Me.dtpExp) & "',ManufucteryDate='" & Trim(Me.dtpmanu) & "',NoOfCartones='" & Val(Trim(Me.txtNoCartones.Text)) & "',Department='" & (Trim(Me.cboDepartment.Text)) & "',ProductCode='" & (Trim(Me.txtCode.Text)) & "',ProductCode1='" & (Trim(Me.txtProductCode.Text)) & "' Where ProductID ='" & Productid & "'", Y
       If Y > 0 Then
   cn.Execute "Update ProductInventory Set StockLevel ='" & StockQty & "',ReorderLevel='" & Val(Trim(Me.txtReorder.Text)) & "' Where ProductID ='" & Productid & "'", Y
       End If
         
   If Y > 0 Then
      cn.CommitTrans
      MsgBox "Product Details edited Successful!", vbInformation, "Edit Successful"
     'Clear ctrls and setfocus to Clinic ctrl
      Call ListProducts
      Call ClearCtrls
      Me.txtname = "": Me.txtbasePrice = "": Me.txtbaseunit = "": Me.txtDiscount = ""
      Me.txtManufacterer = "": Me.txtNoCartones = "": Me.txtReorder = "": Me.txtStock = ""
      Me.txtTotalQuantity = "": Me.txtUnitPrice = ""
      Me.txtname.SetFocus
   Else
      cn.RollbackTrans
      MsgBox "Sorry, Unable to Edit Product Details:Please Try Again!", vbInformation, "Edit Failed"
   End If

 End If

    sflag = False
    Me.cmdSave.Caption = "Save Particulars"
If cn.State = 1 Then cn.Close
If rs.State = 1 Then rs.Close

Me.MousePointer = vbDefault
Me.cmdSave.Enabled = True


Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "Sorry, Unable to Save Products Details:Please Try Again!", vbInformation, "Save Failed"
     Exit Sub

End Sub



Private Sub dtpExp_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   Me.txtStock.SetFocus
End If
End Sub

Private Sub dtpmanu_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   Me.txtManufacterer.SetFocus
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub flxgItems_Click()

Me.txtname = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 0)
Me.txtbaseunit = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 1)
Me.txtUnitPrice = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 2)
Me.txtbasePrice = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 3)
Me.txtTotalQuantity = "" 'Me.flxgItems.TextMatrix(Me.flxgItems.Row, 4)
Me.txtDiscount = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 5)
Me.txtManufacterer = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 6)
Me.dtpExp = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 7)
Me.dtpmanu = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 8)
'Me.txtNoCartones = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 9)
Me.txtStock = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 10)
Me.txtReorder = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 11)
Productid = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 12)
Me.cboDepartment = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 13)
If Me.flxgItems.TextMatrix(Me.flxgItems.Row, 14) <> "" Then
    Me.txtCode = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 14)
Else
    Me.txtCode = ""
End If

If Me.flxgItems.TextMatrix(Me.flxgItems.Row, 15) <> "" Then
    Me.txtProductCode = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 15)
Else
    Me.txtProductCode = ""
End If

Me.flxgItems.Visible = False
Me.cmdDelete.Enabled = True
Me.cmdSave.Enabled = True
sflag = True
Me.cmdSave.Caption = "Edit Particulars"
Me.txtname.SetFocus
End Sub



Private Sub flxgItems_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.txtname = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 0)
    Me.txtbaseunit = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 1)
    Me.txtUnitPrice = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 2)
    Me.txtbasePrice = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 3)
    Me.txtTotalQuantity = "" 'Me.flxgItems.TextMatrix(Me.flxgItems.Row, 4)
    Me.txtDiscount = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 5)
    Me.txtManufacterer = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 6)
    Me.dtpExp = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 7)
    Me.dtpmanu = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 8)
    'Me.txtNoCartones = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 9)
    Me.txtStock = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 10)
    Me.txtReorder = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 11)
    Productid = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 12)
    Me.cboDepartment = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 13)
    If Me.flxgItems.TextMatrix(Me.flxgItems.Row, 14) <> "" Then
        Me.txtCode = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 14)
    Else
        Me.txtCode = ""
    End If
    
    If Me.flxgItems.TextMatrix(Me.flxgItems.Row, 15) <> "" Then
        Me.txtProductCode = Me.flxgItems.TextMatrix(Me.flxgItems.Row, 15)
    Else
        Me.txtProductCode = ""
    End If
    
    Me.flxgItems.Visible = False
    Me.cmdDelete.Enabled = True
    Me.cmdSave.Enabled = True
    sflag = True
    Me.cmdSave.Caption = "Edit Particulars"
    Me.txtname.SetFocus
End If
End Sub

Private Sub Form_Load()
Call ListProducts
CenterForm Me
Me.Width = 10050
Me.Height = 9075
Me.Top = (frmMDI.ScaleHeight - Me.Height) / 2
Me.Left = (frmMDI.ScaleWidth - Me.Width) / 2
End Sub

Private Sub lstProducts_Click()
'On Error GoTo SaveError
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
rs.Open "Select Products.*,ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID", cn, adOpenForwardOnly, adLockReadOnly


    While Not rs.EOF
        If Me.lstProducts.SelectedItem.Text = rs("ProductName") Then
                ShowProductflag = True
                Me.txtname.Text = Trim(rs("ProductName"))
                ShowProductflag = False
                Productid = rs("ProductID")
                Me.txtbaseunit.Text = Trim(rs("BaseUnit"))
                Me.txtUnitPrice.Text = Trim(rs("UnitPrice"))
                Me.txtTotalQuantity.Text = "" 'Trim(rs("TotalQuantity"))
                Me.txtbasePrice.Text = Trim(rs("PricePerCartone"))
                Me.txtDiscount.Text = Trim(rs("Discount"))
                Me.txtManufacterer.Text = Trim(rs("Manufacturer"))
                Me.dtpExp = Trim(rs("ExpiryDate"))
                Me.dtpmanu = Trim(rs("ManufucteryDate"))
                Me.txtNoCartones.Text = "" ' Trim(rs("NoOfCartones"))
                Me.txtStock.Text = Trim(rs("StockLevel"))
                Me.txtReorder.Text = Trim(rs("ReorderLevel"))
                If Trim(rs("Department")) <> "" Then
                Me.cboDepartment = Trim(rs("Department"))
                End If
                If Trim(rs("ProductCode")) <> "" Then
                Me.txtCode = Trim(rs("ProductCode"))
                Else
                Me.txtCode = ""
                End If
                If Trim(rs("ProductCode1")) <> "" Then
                Me.txtProductCode = Trim(rs("ProductCode1"))
                Else
                Me.txtProductCode = ""
                End If
                
            rs.MoveNext
        Else
            rs.MoveNext
        End If
    Wend
    rs.Close
    Set rs = Nothing
    sflag = True
    Me.cmdSave.Caption = "Edit Particulars"
    Me.cmdDelete.Enabled = True
    Me.cmdSave.Enabled = True
    
    Exit Sub
SaveError:
     If cn.State = 1 Then cn.Close
     If rs.State = 1 Then rs.Close
     MsgBox "SORRY,TRY AGAIN!", vbInformation, "SAVE FAILED"
     Exit Sub

End Sub

Private Sub txtbasePrice_GotFocus()
Me.txtbasePrice.BackColor = &HFFFF&
End Sub

Private Sub txtbasePrice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
Me.cboDepartment.SetFocus
End If
If KeyCode = 38 Then
Me.txtUnitPrice.SetFocus
End If
End Sub

Private Sub txtbasePrice_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.cboDepartment.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtbasePrice_LostFocus()
Me.txtbasePrice.BackColor = &HF2EBBF
End Sub

Private Sub txtbaseunit_GotFocus()
Me.txtbaseunit.BackColor = &HFFFF&
End Sub

Private Sub txtbaseunit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
Me.txtTotalQuantity.SetFocus
End If
If KeyCode = 38 Then
Me.txtDiscount.SetFocus
End If
End Sub

Private Sub txtbaseunit_KeyPress(KeyAscii As Integer)
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

Private Sub txtbaseunit_LostFocus()
If Trim(Me.txtbaseunit.Text) <> "" And Trim(Me.txtNoCartones.Text) <> "" Then
Me.txtTotalQuantity = Trim(Val(Me.txtbaseunit)) * Trim(Val(Me.txtNoCartones))
End If
Me.txtbaseunit.BackColor = &HF2EBBF
End Sub

Private Sub txtCode_GotFocus()
Me.txtCode.BackColor = &HFFFF&
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
Me.txtProductCode.SetFocus
End If
If KeyCode = 38 Then
Me.cboDepartment.SetFocus
End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
'strk1 = "0123456789/|\;:.,()@#$%&*_-"
If KeyAscii = vbKeyReturn Then
   Me.txtProductCode.SetFocus
End If
'If KeyAscii > 26 Then
'   If KeyAscii <> 32 Then
'      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
'         KeyAscii = 0
'      End If
'   End If
'End If
End Sub

Private Sub txtCode_LostFocus()
Me.txtCode.BackColor = &HF2EBBF
'If Me.txtCode = "" Then
' Me.txtCode = "NULL"
'End If
End Sub

Private Sub txtdiscount_GotFocus()
Me.txtDiscount.BackColor = &HFFFF&
End Sub

Private Sub txtdiscount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
Me.txtbaseunit.SetFocus
End If
If KeyCode = 38 Then
Me.txtProductCode.SetFocus
End If
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.dtpmanu.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtdiscount_LostFocus()
Me.txtDiscount.BackColor = &HF2EBBF
End Sub

Private Sub txtManufacterer_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,(),&,%,@,#"
If KeyAscii = vbKeyReturn Then
   Me.dtpExp.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtName_Change()


If ShowProductflag = False Then

If Me.txtname = "" Then
flxgItems.Visible = False
Exit Sub
End If

On Error GoTo OkError
'Open Connecttion to Server
   
   bFlag = OpenConnection(cn, strg)
   If bFlag = False Then
      If cn.State = 1 Then cn.Close
      If rs.State = 1 Then rs.Close
      Me.MousePointer = vbDefault
       MsgBox strg, vbInformation:
      Exit Sub
   End If
   
If ShowProductflag = False Then  ' when true u can view grid to add products,but when false u can edit with no grid shown
   rs.Open "Select Products.*,ProductInventory.StockLevel,ProductInventory.ReorderLevel From Products Inner Join ProductInventory On Products.ProductID = ProductInventory.ProductID Where ProductName Like '" & Trim(Me.txtname) & "%" & "' order by ProductName", cn, adOpenForwardOnly, adLockReadOnly
   
   If rs.RecordCount > 0 Then
   flxgItems.Height = 950 + (285 * (rs.RecordCount - 1))
   
   If flxgItems.Height >= 3322 Then
      flxgItems.Height = 3322
   End If
    flxgItems.Rows = rs.RecordCount + 1
   With flxgItems
      For X = 1 To rs.RecordCount
           .TextMatrix(X, 0) = rs.Fields("ProductName")
           .TextMatrix(X, 1) = rs.Fields("BaseUnit")
           .TextMatrix(X, 2) = rs.Fields("UnitPrice")
           .TextMatrix(X, 3) = rs.Fields("PricePerCartone")
           .TextMatrix(X, 4) = rs.Fields("TotalQuantity")
           .TextMatrix(X, 5) = rs.Fields("Discount")
           .TextMatrix(X, 6) = rs.Fields("Manufacturer")
           .TextMatrix(X, 7) = rs.Fields("ExpiryDate")
           .TextMatrix(X, 8) = rs.Fields("ManufucteryDate")
           .TextMatrix(X, 9) = rs.Fields("NoOfCartones")
           .TextMatrix(X, 10) = rs.Fields("StockLevel")
           .TextMatrix(X, 11) = rs.Fields("ReorderLevel")
           .TextMatrix(X, 12) = rs.Fields("ProductID")
           
           If rs.Fields("Department") <> "" Then
           .TextMatrix(X, 13) = rs.Fields("Department")
           End If
           
           If rs.Fields("ProductCode") <> "" Then
           .TextMatrix(X, 14) = rs.Fields("ProductCode")
           Else
           .TextMatrix(X, 14) = ""
           End If
           
           If rs.Fields("ProductCode1") <> "" Then
           .TextMatrix(X, 15) = rs.Fields("ProductCode1")
           Else
           .TextMatrix(X, 15) = ""
           End If
             
        rs.MoveNext
      Next
      .Col = 0
      .Row = 1
      .ColSel = 1
      .RowSel = 1
   End With
   flxgItems.Visible = True
   If rs.State = 1 Then rs.Close
   Me.cmdSave.Enabled = True
Else
     flxgItems.Visible = False
     If rs.State = 1 Then rs.Close
     
End If
Else
      flxgItems.Visible = False
      'If cn.State = 1 Then cn.Close
      'If rs.State = 1 Then rs.Close
      'oflag = False
      xflag = False ': Exit Sub
End If
      If rs.State = 1 Then rs.Close
End If

Exit Sub
OkError:
     If rs.State = 1 Then rs.Close
     MsgBox "Items In Stock", , "Displaying"
     Exit Sub
End Sub

Private Sub txtname_GotFocus()
Me.txtname.BackColor = &HFFFF&
End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
  If Me.flxgItems.Visible = True Then
   Me.flxgItems.SetFocus
  End If
End If
If KeyCode = 40 And Me.flxgItems.Visible = False Then
 Me.txtUnitPrice.SetFocus
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
strk1 = "0123456789/|\;:.,()@#$%&*_-"
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

Private Sub txtname_LostFocus()
Me.txtname.BackColor = &HF2EBBF

End Sub

Private Sub txtNoCartones_KeyPress(KeyAscii As Integer)
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

Private Sub txtNoCartones_LostFocus()
If Trim(Me.txtbaseunit.Text) <> "" And Trim(Me.txtNoCartones.Text) <> "" Then
Me.txtTotalQuantity = Trim(Val(Me.txtbaseunit)) * Trim(Val(Me.txtNoCartones))
End If
'If (Trim(Me.txtbaseunit.Text) <> "" And Trim(Me.txtNoCartones.Text) = "") Then
'MsgBox "You must Enter Number of Cartones ", vbInformation, "Enter Number of Cartones"
'Me.txtNoCartones.SetFocus
'End If
'If (Trim(Me.txtbaseunit.Text) = "" And Trim(Me.txtNoCartones.Text) <> "") Then
'MsgBox "You must Enter BaseUnit ", vbInformation, "BaseUnit"
'Me.txtbaseunit.SetFocus
'End If
End Sub

Private Sub txtProductCode_GotFocus()
Me.txtProductCode.BackColor = &HFFFF&
End Sub

Private Sub txtProductCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
Me.txtDiscount.SetFocus
End If
If KeyCode = 38 Then
Me.txtCode.SetFocus
End If
End Sub

Private Sub txtProductCode_KeyPress(KeyAscii As Integer)
Dim strk1 As String
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then Exit Sub
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
'strk1 = "0123456789/|\;:.,()@#$%&*_-"
If KeyAscii = vbKeyReturn Then
   Me.txtDiscount.SetFocus
End If
'If KeyAscii > 26 Then
'   If KeyAscii <> 32 Then
'      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
'         KeyAscii = 0
'      End If
'   End If
'End If
End Sub

Private Sub txtProductCode_LostFocus()
Me.txtProductCode.BackColor = &HF2EBBF
'If Me.txtProductCode = "" Then
' Me.txtProductCode = "NULL"
'End If
End Sub

Private Sub txtReorder_GotFocus()
Me.txtReorder.BackColor = &HFFFF&
End Sub

Private Sub txtReorder_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
Me.cmdSave.SetFocus
End If
If KeyCode = 38 Then
Me.txtStock.SetFocus
End If
End Sub

Private Sub txtReorder_KeyPress(KeyAscii As Integer)
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

Private Sub txtReorder_LostFocus()
Me.txtReorder.BackColor = &HF2EBBF
End Sub

Private Sub txtStock_GotFocus()
If Trim(Me.txtTotalQuantity) <> "" Then
 Me.txtStock.Text = Trim(Me.txtTotalQuantity)
End If
 Me.txtStock.BackColor = &HFFFF&
End Sub

Private Sub txtStock_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 40 Then
 Me.txtReorder.SetFocus
End If
If KeyCode = 38 Then
 Me.txtTotalQuantity.SetFocus
End If

End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
Dim strk1 As String
strk1 = "0123456789.,"
If KeyAscii = vbKeyReturn Then
   Me.txtReorder.SetFocus
End If
If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If
End Sub

Private Sub txtStock_LostFocus()
 Me.txtStock.BackColor = &HF2EBBF
End Sub

Private Sub txtTotalQuantity_GotFocus()
If Trim(Me.txtbaseunit.Text) <> "" And Trim(Me.txtNoCartones.Text) <> "" Then
 Me.txtTotalQuantity = Trim(Val(Me.txtbaseunit)) * Trim(Val(Me.txtNoCartones))
End If
 Me.txtTotalQuantity.BackColor = &HFFFF&
End Sub

Private Sub txtTotalQuantity_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 40 Then
 Me.txtStock.SetFocus
End If
If KeyCode = 38 Then
 Me.txtbaseunit.SetFocus
End If

End Sub

Private Sub txtTotalQuantity_KeyPress(KeyAscii As Integer)
Dim strk1 As String

strk1 = "0123456789.-,"
If KeyAscii = vbKeyReturn Then
   Me.txtStock.SetFocus
End If

If KeyAscii > 26 Then
   If KeyAscii <> 32 Then
      If InStr(1, strk1, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End If

End Sub

Private Sub txtTotalQuantity_LostFocus()
Me.txtTotalQuantity.BackColor = &HF2EBBF
End Sub

Private Sub txtUnitPrice_GotFocus()
    Me.txtUnitPrice.BackColor = &HFFFF&
If Me.flxgItems.Visible = True Then
    Me.flxgItems.Visible = False
End If

End Sub

Private Sub txtUnitPrice_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 40 Then
 Me.txtbasePrice.SetFocus
End If
If KeyCode = 38 Then
 Me.txtname.SetFocus
End If

End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
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
Private Sub ListProducts()
On Error GoTo SaveError
Me.lstProducts.ListItems.Clear
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


rs.Open "Select Products.*,ProductInventory.StockLevel From Products Inner Join ProductInventory On Products.ProductID=ProductInventory.ProductID Order By Products.ProductName", cn, adOpenForwardOnly, adLockReadOnly
For i = 1 To rs.RecordCount
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                Set List_Item = Me.lstProducts.ListItems.Add(, , Trim(rs!ProductName))
                'List_Item.Icon = 1
                'List_Item.SmallIcon = 1
               ' List_Item.ForeColor = vbBlack
                
                List_Item.SubItems(1) = Format$(Trim(rs!UnitPrice), "#,###.00")
                List_Item.SubItems(2) = Trim(rs!StockLevel)
                List_Item.SubItems(3) = Trim(rs!Discount)
                If rs.Fields("ProductCode") <> "" Then
                List_Item.SubItems(5) = Trim(rs!ProductCode)
                End If
                If rs.Fields("ProductCode1") <> "" Then
                List_Item.SubItems(6) = Trim(rs!ProductCode1)
                List_Item.SubItems(7) = Trim(rs!BaseUnit)
                End If
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


Private Function Generate_ProductID(Product_ID As String) As Boolean

Dim strg As String, strg1 As String, strg2 As String, bFlag As Boolean
On Error GoTo SaveError
rs.Open "Select ProductID From Products  order by ProductID Desc", cn, adOpenForwardOnly, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
      strg1 = Trim(rs.Fields!Productid)

   strg1 = Val(strg1) + 1
'   strg1 = String$(5 - Len(strg1), "0") & strg1
Else
   strg1 = "1"
End If
Product_ID = strg1

If rs.State = 1 Then rs.Close

Generate_ProductID = True

Exit Function

'''If rs.RecordCount > 0 Then
'''   rs.MoveFirst
'''      strg1 = Trim(rs.Fields!Productid)
'''
'''   strg1 = Trim(Str(Val(strg1) + 1))
'''   strg1 = String$(5 - Len(strg1), "0") & strg1
'''Else
'''   strg1 = "00001"
'''End If
'''Product_ID = strg1
'''
'''If rs.State = 1 Then rs.Close
'''
'''Generate_ProductID = True
'''
'''Exit Function

SaveError:
     If rs.State = 1 Then rs.Close
    Generate_ProductID = False
     Exit Function
     
End Function
Private Sub ClearCtrls()
For Each ctrl In Me.Controls
   If (Trim(ctrl.Name) Like "txt*" Or Trim(ctrl.Name) Like "cbo*") Then
  ctrl = ""
   End If
Next
End Sub

Private Sub txtUnitPrice_LostFocus()
Me.txtUnitPrice.BackColor = &HF2EBBF
End Sub
