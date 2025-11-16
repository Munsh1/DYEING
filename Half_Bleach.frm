VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form HalfBleach 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Half Bleach"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   10965
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Half_Bleach.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Half_Bleach.frx":0268
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Half_Bleach.frx":06C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Half_Bleach.frx":0ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Half_Bleach.frx":0F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Half_Bleach.frx":1330
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Half_Bleach.frx":176C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Half_Bleach.frx":1BC0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   1695
      Index           =   1
      Left            =   120
      TabIndex        =   93
      Top             =   5400
      Width           =   7695
      Begin MSComctlLib.ListView lvwphase 
         Height          =   1320
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   2328
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7560
      Index           =   1
      Left            =   7920
      TabIndex        =   76
      Top             =   120
      Width           =   3000
      Begin VB.Frame Frame15 
         Height          =   800
         Left            =   120
         TabIndex        =   89
         Top             =   4680
         Width           =   2800
         Begin VB.CheckBox ImChk 
            Caption         =   "Item"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   0
            Width           =   615
         End
         Begin VB.ComboBox SrItem 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame14 
         Height          =   800
         Left            =   100
         TabIndex        =   87
         Top             =   3600
         Width           =   2800
         Begin VB.CheckBox ImTChk 
            Caption         =   "Item Type"
            Height          =   195
            Left            =   240
            TabIndex        =   98
            Top             =   0
            Width           =   1095
         End
         Begin VB.ComboBox SrItemType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame13 
         Height          =   800
         Left            =   100
         TabIndex        =   85
         Top             =   2640
         Width           =   2800
         Begin VB.CheckBox McChk 
            Caption         =   "Machine"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox srMachine 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Left            =   125
            TabIndex        =   86
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame12 
         Height          =   800
         Left            =   100
         TabIndex        =   83
         Top             =   1560
         Width           =   2800
         Begin VB.CheckBox PtChk 
            Caption         =   "Party"
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox srParty 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1155
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   2800
         Begin VB.CheckBox dtChk 
            Caption         =   "Date"
            Height          =   195
            Left            =   240
            TabIndex        =   80
            Top             =   0
            Width           =   735
         End
         Begin MSComCtl2.DTPicker SrDate2 
            Height          =   330
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   44630017
            CurrentDate     =   38298
         End
         Begin MSComCtl2.DTPicker SrDate 
            Height          =   330
            Left            =   125
            TabIndex        =   82
            Top             =   280
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   44630017
            CurrentDate     =   38235
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   77
         Top             =   5760
         Width           =   2775
         Begin VB.CheckBox hbChk 
            Caption         =   "H/B Code"
            Height          =   195
            Left            =   240
            TabIndex        =   100
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox srHalfBleachCode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   78
            Top             =   320
            Width           =   2535
         End
      End
      Begin LVbuttons.LaVolpeButton Cmdhide 
         Height          =   375
         Left            =   360
         TabIndex        =   91
         Top             =   6840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Hide"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   12632256
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Half_Bleach.frx":1E38
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "7"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Index           =   0
      Left            =   120
      TabIndex        =   53
      Top             =   2760
      Width           =   7695
      Begin VB.TextBox f_Soap_1_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2600
         TabIndex        =   18
         Top             =   260
         Width           =   1200
      End
      Begin VB.TextBox f_Soap_Temp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4600
         TabIndex        =   19
         Top             =   260
         Width           =   1200
      End
      Begin VB.TextBox f_Soap_Temp_Time 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6300
         TabIndex        =   20
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox f_Acid_Temp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4600
         TabIndex        =   37
         Top             =   2035
         Width           =   1200
      End
      Begin VB.ComboBox f_Soap_1_Code 
         Height          =   315
         Left            =   850
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   260
         Width           =   1200
      End
      Begin VB.ComboBox f_Castic_Code 
         Height          =   315
         Left            =   850
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   970
         Width           =   1200
      End
      Begin VB.TextBox f_Castic_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2600
         TabIndex        =   25
         Top             =   960
         Width           =   1200
      End
      Begin VB.ComboBox f_Chemical_3_Code 
         Height          =   315
         Left            =   850
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1680
         Width           =   1200
      End
      Begin VB.TextBox f_Soap_2_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6300
         TabIndex        =   27
         Top             =   970
         Width           =   1200
      End
      Begin VB.ComboBox f_Acid_Code 
         Height          =   315
         Left            =   850
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2035
         Width           =   1200
      End
      Begin VB.TextBox f_Cold_Wash_1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   850
         TabIndex        =   21
         Top             =   615
         Width           =   1200
      End
      Begin VB.ComboBox f_Chemical_2_Code 
         Height          =   315
         Left            =   4600
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1325
         Width           =   1200
      End
      Begin VB.TextBox f_Acid_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2600
         TabIndex        =   36
         Top             =   2035
         Width           =   1200
      End
      Begin VB.ComboBox f_Chemical_1_Code 
         Height          =   315
         Left            =   850
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1325
         Width           =   1200
      End
      Begin VB.TextBox f_Chemical_1_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2600
         TabIndex        =   29
         Top             =   1325
         Width           =   1200
      End
      Begin VB.ComboBox f_Soap_2_Code 
         Height          =   315
         Left            =   4600
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   970
         Width           =   1200
      End
      Begin VB.TextBox f_Chemical_2_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6300
         TabIndex        =   31
         Top             =   1325
         Width           =   1200
      End
      Begin VB.TextBox f_Cold_Wash_2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4600
         TabIndex        =   34
         Top             =   1680
         Width           =   1200
      End
      Begin VB.TextBox f_Soda_Qty 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4600
         TabIndex        =   23
         Top             =   610
         Width           =   1200
      End
      Begin VB.ComboBox f_Soda_Code 
         Height          =   315
         Left            =   2600
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox f_Acid_Temp_Time 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6300
         TabIndex        =   38
         Top             =   2035
         Width           =   1200
      End
      Begin VB.TextBox f_Chemical_3_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2600
         TabIndex        =   33
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label29 
         Caption         =   "Time"
         Height          =   255
         Left            =   5850
         TabIndex        =   75
         Top             =   2050
         Width           =   375
      End
      Begin VB.Label Label28 
         Caption         =   "Temp."
         Height          =   255
         Left            =   3900
         TabIndex        =   74
         Top             =   2050
         Width           =   495
      End
      Begin VB.Label Label27 
         Caption         =   "C / Wash"
         Height          =   255
         Left            =   3900
         TabIndex        =   73
         Top             =   1700
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Acid"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   2050
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "C / Wash"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Qty"
         Height          =   255
         Left            =   3900
         TabIndex        =   70
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Soda"
         Height          =   255
         Left            =   2100
         TabIndex        =   69
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   68
         Top             =   1700
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Qty"
         Height          =   255
         Left            =   5850
         TabIndex        =   67
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   66
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   65
         Top             =   2050
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1700
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   3900
         TabIndex        =   63
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Qty"
         Height          =   255
         Left            =   5850
         TabIndex        =   61
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "Soap"
         Height          =   255
         Left            =   3900
         TabIndex        =   60
         Top             =   1000
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   59
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Castic"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1000
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Time"
         Height          =   255
         Left            =   5850
         TabIndex        =   57
         Top             =   280
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Temp."
         Height          =   255
         Left            =   3900
         TabIndex        =   56
         Top             =   280
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   55
         Top             =   280
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Soap"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   280
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox f_Cone_KG_1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   95
         Top             =   2040
         Width           =   2000
      End
      Begin VB.TextBox f_Cone_KG_2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3200
         TabIndex        =   15
         Top             =   2020
         Width           =   2000
      End
      Begin VB.TextBox f_MachineCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   1
         Top             =   250
         Width           =   1215
      End
      Begin VB.ComboBox f_Item_3 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Width           =   2000
      End
      Begin VB.ComboBox f_Item_2 
         Height          =   315
         Left            =   3200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   2000
      End
      Begin VB.ComboBox f_Item_1 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   2000
      End
      Begin VB.ComboBox f_ItemType_3 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   2000
      End
      Begin VB.ComboBox f_ItemType_2 
         Height          =   315
         Left            =   3200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   2000
      End
      Begin VB.ComboBox f_ItemType_1 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   2000
      End
      Begin VB.TextBox f_Cone_KG_3 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5400
         TabIndex        =   16
         Top             =   2020
         Width           =   2000
      End
      Begin VB.TextBox f_Cone_3 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5400
         TabIndex        =   14
         Top             =   1680
         Width           =   2000
      End
      Begin VB.TextBox f_Cone_2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3200
         TabIndex        =   13
         Top             =   1680
         Width           =   2000
      End
      Begin VB.TextBox f_Cone_1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   1680
         Width           =   2000
      End
      Begin MSComCtl2.DTPicker f_HalfBleachDate 
         Height          =   300
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         _Version        =   393216
         Format          =   44630017
         CurrentDate     =   38365
      End
      Begin VB.TextBox f_HalfBleachCode 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   45
         Top             =   250
         Width           =   1515
      End
      Begin VB.ComboBox f_Party_3 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   2000
      End
      Begin VB.ComboBox f_Party_2 
         Height          =   315
         Left            =   3200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2000
      End
      Begin VB.ComboBox f_Party_1 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2000
      End
      Begin VB.Label Label30 
         Caption         =   "Machine"
         Height          =   255
         Left            =   2760
         TabIndex        =   92
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   255
         Left            =   4920
         TabIndex        =   52
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "KG"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Cone"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Item Type"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Party"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
   End
   Begin LVbuttons.LaVolpeButton CmdAllSearch 
      Height          =   405
      Left            =   5280
      TabIndex        =   43
      Top             =   7200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Search"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Half_Bleach.frx":1E54
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "5"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdClose 
      Height          =   405
      Left            =   6600
      TabIndex        =   44
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Half_Bleach.frx":1E70
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "6"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdNew 
      Height          =   405
      Left            =   4080
      TabIndex        =   40
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Half_Bleach.frx":1E8C
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "3"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   405
      Left            =   2880
      TabIndex        =   39
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Half_Bleach.frx":1EA8
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "2"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdDel 
      Height          =   405
      Left            =   1560
      TabIndex        =   41
      Top             =   7200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Half_Bleach.frx":1EC4
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "4"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdReport 
      Height          =   405
      Left            =   120
      TabIndex        =   42
      Top             =   7200
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Half_Bleach.frx":1EE0
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "8"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
End
Attribute VB_Name = "HalfBleach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ListID As String
Dim m_AddMode As Boolean
Dim CMDSearch As Boolean
Dim PreQty As Double
Dim QtyBit As Integer
Dim MsgBit As Integer
Dim ClickPane As Integer
Private Sub f_Acid_Code_LostFocus()
    If Me.f_Acid_Code.ListIndex = -1 Then
        Me.f_Acid_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Castic_Code_LostFocus()
    If Me.f_Castic_Code.ListIndex = -1 Then
        Me.f_Castic_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Chemical_1_Code_LostFocus()
    If Me.f_Chemical_1_Code.ListIndex = -1 Then
        Me.f_Chemical_1_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Chemical_2_Code_LostFocus()
    If Me.f_Chemical_2_Code.ListIndex = -1 Then
        Me.f_Chemical_2_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Chemical_3_Code_LostFocus()
    If Me.f_Chemical_3_Code.ListIndex = -1 Then
        Me.f_Chemical_3_Code.ListIndex = 0
    End If
End Sub
Private Sub f_HalfBleachDate_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Party_1.SetFocus
    End If
End Sub
Private Sub f_ItemType_1_Click()
    If Me.f_ItemType_1.ListIndex > 0 Then
        i = Me.f_ItemType_1.ItemData(Me.f_ItemType_1.ListIndex)
        FillCombo "Select ItemCode, ItemName from vwAvailableQty where Qty > 0 and ItemTypeCode = " & i, f_Item_1, "ItemName", "ItemCode"
    Else
        Me.f_Item_1.Clear
    End If
End Sub
Private Sub f_ItemType_2_Click()
    If Me.f_ItemType_2.ListIndex > 0 Then
        i = Me.f_ItemType_2.ItemData(Me.f_ItemType_2.ListIndex)
        FillCombo "Select ItemCode, ItemName from vwAvailableQty where Qty > 0 and ItemTypeCode = " & i, f_Item_2, "ItemName", "ItemCode"
    Else
        Me.f_Item_2.Clear
    End If
End Sub
Private Sub f_ItemType_3_Click()
    If Me.f_ItemType_3.ListIndex > 0 Then
        i = Me.f_ItemType_3.ItemData(Me.f_ItemType_3.ListIndex)
        FillCombo "Select ItemCode, ItemName from vwAvailableQty where Qty > 0 and ItemTypeCode = " & i, f_Item_3, "ItemName", "ItemCode"
    Else
        Me.f_Item_3.Clear
    End If
End Sub
Private Sub f_MachineCode_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_HalfBleachDate.SetFocus
    End If
End Sub
Private Sub f_Party_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Party_2.SetFocus
    End If
End Sub
Private Sub f_Party_1_LostFocus()
    If Me.f_Party_1.ListIndex = -1 Then
        Me.f_Party_1.ListIndex = 0
    End If
End Sub
Private Sub f_Party_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Party_3.SetFocus
    End If
End Sub
Private Sub f_Party_2_LostFocus()
    If Me.f_Party_2.ListIndex = -1 Then
        Me.f_Party_2.ListIndex = 0
    End If
End Sub
Private Sub f_Party_3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_ItemType_1.SetFocus
    End If
End Sub
Private Sub f_Party_3_LostFocus()
    If Me.f_Party_3.ListIndex = -1 Then
        Me.f_Party_3.ListIndex = 0
    End If
End Sub
Private Sub f_ItemType_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_ItemType_2.SetFocus
    End If
End Sub
Private Sub f_ItemType_1_LostFocus()
    If Me.f_ItemType_1.ListIndex = -1 Then
        Me.f_ItemType_1.ListIndex = 0
    End If
End Sub
Private Sub f_ItemType_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_ItemType_3.SetFocus
    End If
End Sub
Private Sub f_ItemType_2_LostFocus()
    If Me.f_ItemType_2.ListIndex = -1 Then
        Me.f_ItemType_2.ListIndex = 0
    End If
End Sub
Private Sub f_ItemType_3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Item_1.SetFocus
    End If
End Sub
Private Sub f_ItemType_3_LostFocus()
    If Me.f_ItemType_3.ListIndex = -1 Then
        Me.f_ItemType_3.ListIndex = 0
    End If
End Sub
Private Sub f_Item_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Item_2.SetFocus
    End If
End Sub
Private Sub f_Item_1_LostFocus()
    If Me.f_Item_1.ListIndex = -1 Then
        FillCombo "Select 0 as ItemCode, '-- Select --' as ItemName ", f_Item_1, "ItemName", "ItemCode"
        Me.f_Item_1.ListIndex = 0
    End If
End Sub
Private Sub f_Item_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Item_3.SetFocus
    End If
End Sub
Private Sub f_Item_2_LostFocus()
    If Me.f_Item_2.ListIndex = -1 Then
        FillCombo "Select 0 as ItemCode, '-- Select --' as ItemName ", f_Item_2, "ItemName", "ItemCode"
        Me.f_Item_2.ListIndex = 0
    End If
End Sub
Private Sub f_Item_3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cone_1.SetFocus
    End If
End Sub
Private Sub f_Item_3_LostFocus()
    If Me.f_Item_3.ListIndex = -1 Then
        FillCombo "Select 0 as ItemCode, '-- Select --' as ItemName ", f_Item_3, "ItemName", "ItemCode"
        Me.f_Item_3.ListIndex = 0
    End If
End Sub
Private Sub f_Cone_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cone_2.SetFocus
    End If
End Sub
Private Sub f_Cone_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cone_3.SetFocus
    End If
End Sub
Private Sub f_Cone_3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cone_KG_1.SetFocus
    End If
End Sub
Private Sub f_Cone_KG_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cone_KG_2.SetFocus
    End If
End Sub
Private Sub f_Cone_KG_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cone_KG_3.SetFocus
    End If
End Sub
Private Sub f_Cone_KG_3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap_1_Code.SetFocus
    End If
End Sub
Private Sub f_Soap_1_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap_1_Qty.SetFocus
    End If
End Sub
Private Sub f_Soap_1_Code_LostFocus()
    If Me.f_Soap_1_Code.ListIndex = -1 Then
        Me.f_Soap_1_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Soap_1_Qty_GotFocus()
    If Len(Trim(Me.f_Soap_1_Qty)) > 0 Then
        PreQty = Me.f_Soap_1_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Soap_1_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Soap_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Soap_Temp.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Soap_1_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Soap_1_Qty.Text)) > 0 Then
        vQty = Me.f_Soap_1_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Soap_1_Code.ItemData(Me.f_Soap_1_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Soap_Qty(Me.f_Soap_1_Code.ItemData(Me.f_Soap_1_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Soap_2_Code_LostFocus()
    If Me.f_Soap_2_Code.ListIndex = -1 Then
        Me.f_Soap_2_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Soap_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap_Temp_Time.SetFocus
    End If
End Sub
Private Sub f_Soap_Temp_Time_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cold_Wash_1.SetFocus
    End If
End Sub
Private Sub f_Cold_Wash_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soda_Code.SetFocus
    End If
End Sub
Private Sub f_Soda_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soda_Qty.SetFocus
    End If
End Sub
Private Sub f_Soda_Code_LostFocus()
    If Me.f_Soda_Code.ListIndex = -1 Then
        Me.f_Soda_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Soda_Qty_GotFocus()
    If Len(Trim(Me.f_Soda_Qty)) > 0 Then
        PreQty = Me.f_Soda_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Soda_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Soda_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Castic_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Soda_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Soda_Qty.Text)) > 0 Then
        vQty = Me.f_Soda_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Soda_Code.ItemData(Me.f_Soda_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Soda_Qty(Me.f_Soda_Code.ItemData(Me.f_Soda_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Castic_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Castic_Qty.SetFocus
    End If
End Sub
Private Sub f_Soap_2_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap_2_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_1_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_1_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_1_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_1_Qty)) > 0 Then
        PreQty = Me.f_Chemical_1_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_1_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Chemical_1_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Chemical_2_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Chemical_1_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Chemical_1_Qty.Text)) > 0 Then
        vQty = Me.f_Chemical_1_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Chemical_1_Code.ItemData(Me.f_Chemical_1_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Chemical_1_Qty(Me.f_Chemical_1_Code.ItemData(Me.f_Chemical_1_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Chemical_2_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_2_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_2_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_2_Qty)) > 0 Then
        PreQty = Me.f_Chemical_2_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_2_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Chemical_2_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Chemical_3_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Chemical_2_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Chemical_2_Qty.Text)) > 0 Then
        vQty = Me.f_Chemical_2_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Chemical_2_Code.ItemData(Me.f_Chemical_2_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Chemical_2_Qty(Me.f_Chemical_2_Code.ItemData(Me.f_Chemical_2_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Chemical_3_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_3_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_3_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_3_Qty)) > 0 Then
        PreQty = Me.f_Chemical_3_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_3_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Chemical_3_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Cold_Wash_2.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Chemical_3_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Chemical_3_Qty.Text)) > 0 Then
        vQty = Me.f_Chemical_3_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Chemical_3_Code.ItemData(Me.f_Chemical_3_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Chemical_3_Qty(Me.f_Chemical_3_Code.ItemData(Me.f_Chemical_3_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Cold_Wash_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Acid_Code.SetFocus
    End If
End Sub
Private Sub f_Acid_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Acid_Qty.SetFocus
    End If
End Sub
Private Sub f_Acid_Qty_GotFocus()
    If Len(Trim(Me.f_Acid_Qty)) > 0 Then
        PreQty = Me.f_Acid_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Acid_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Acid_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Acid_Temp.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Acid_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Acid_Qty.Text)) > 0 Then
        vQty = Me.f_Acid_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Acid_Code.ItemData(Me.f_Acid_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Acid_Qty(Me.f_Acid_Code.ItemData(Me.f_Acid_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Acid_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Acid_Temp_Time.SetFocus
    End If
End Sub
Private Sub f_Acid_Temp_Time_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.cmdSave.SetFocus
    End If
End Sub
Private Sub Form_Load()
  m_AddMode = True
  cmdSave.Enabled = False
  DBConn
  f_HalfBleachDate = Now
  SrDate = Now
  SrDate2 = Now
  
  FillColorCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", f_Party_1, "PartyName", "PartyCode"
  FillColorCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", f_Party_2, "PartyName", "PartyCode"
  FillColorCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", f_Party_3, "PartyName", "PartyCode"
  FillColorCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", srParty, "PartyName", "PartyCode"
  
  FillColorCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", f_ItemType_1, "ItemTypeName", "ItemTypeCode"
  FillColorCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", f_ItemType_2, "ItemTypeName", "ItemTypeCode"
  FillColorCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", f_ItemType_3, "ItemTypeName", "ItemTypeCode"
  FillColorCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", SrItemType, "ItemTypeName", "ItemTypeCode"
     
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_1_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_2_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_3_Code, "ItemName", "ItemCode"
  
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 10 order by 2", f_Soap_1_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 10 order by 2", f_Soap_2_Code, "ItemName", "ItemCode"
    
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 7 order by 2", f_Acid_Code, "ItemName", "ItemCode"
  
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 9 order by 2", f_Castic_Code, "ItemName", "ItemCode"
  
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 8 order by 2", f_Soda_Code, "ItemName", "ItemCode"
   
  lvwphase.ColumnHeaders.Add Text:="Code", Width:=600
  lvwphase.ColumnHeaders.Add Text:="Date", Width:=1200
  lvwphase.ColumnHeaders.Add Text:="Party Name", Width:=1700
  lvwphase.ColumnHeaders.Add Text:="Machine #", Width:=1000
  lvwphase.ColumnHeaders.Add Text:="Item Type", Width:=1500
  lvwphase.ColumnHeaders.Add Text:="Item", Width:=1490
  
  Call fillList

End Sub
Public Sub setVal()
    Dim rstSave As New ADODB.Recordset
    If (Len(Trim(m_ListID)) = 0) Then
        Set rstSave = FillRecordSet("select * from HalfBleach where 1 = 2")
        rstSave.AddNew
        m_ListID = ValAutoNumber("HalfBleach", "HalfBleachCode")
        rstSave("HalfBleachCode") = m_ListID
    Else
       Set rstSave = FillRecordSet("select * from HalfBleach where HalfBleachCode =" & m_ListID)
    End If
    rstSave("MachineCode") = Me.f_MachineCode.Text
    rstSave("HalfBleachDate") = Me.f_HalfBleachDate.value
    
    If Me.f_Party_1.ItemData(Me.f_Party_1.ListIndex) > 0 Then
        rstSave("Party_1_Code") = Me.f_Party_1.ItemData(Me.f_Party_1.ListIndex)
    Else
        rstSave("Party_1_Code") = 0
    End If

    If Me.f_Party_2.ItemData(Me.f_Party_2.ListIndex) > 0 Then
        rstSave("Party_2_Code") = Me.f_Party_2.ItemData(Me.f_Party_2.ListIndex)
    Else
        rstSave("Party_2_Code") = 0
    End If

    If Me.f_Party_3.ItemData(Me.f_Party_3.ListIndex) > 0 Then
        rstSave("Party_3_Code") = Me.f_Party_3.ItemData(Me.f_Party_3.ListIndex)
    Else
        rstSave("Party_3_Code") = 0
    End If

    If Me.f_ItemType_1.ItemData(Me.f_ItemType_1.ListIndex) > 0 Then
        rstSave("ItemType_1_Code") = Me.f_ItemType_1.ItemData(Me.f_ItemType_1.ListIndex)
    Else
        rstSave("ItemType_1_Code") = 0
    End If
    
    If Me.f_ItemType_2.ItemData(Me.f_ItemType_2.ListIndex) > 0 Then
        rstSave("ItemType_2_Code") = Me.f_ItemType_2.ItemData(Me.f_ItemType_2.ListIndex)
    Else
        rstSave("ItemType_2_Code") = 0
    End If
    
    If Me.f_ItemType_3.ItemData(Me.f_ItemType_3.ListIndex) > 0 Then
        rstSave("ItemType_3_Code") = Me.f_ItemType_3.ItemData(Me.f_ItemType_3.ListIndex)
    Else
        rstSave("ItemType_3_Code") = 0
    End If
    
    If Me.f_Item_1.ItemData(Me.f_Item_1.ListIndex) > 0 Then
        rstSave("Item_1_Code") = Me.f_Item_1.ItemData(Me.f_Item_1.ListIndex)
    Else
        rstSave("Item_1_Code") = 0
    End If
    
    If Me.f_Item_2.ListIndex = -1 Then
        FillCombo "Select 0 as ItemCode, '-- Select --' as ItemName ", f_Item_2, "ItemName", "ItemCode"
        Me.f_Item_2.ListIndex = 0
    End If
    
    If Me.f_Item_2.ItemData(Me.f_Item_2.ListIndex) > 0 Then
        rstSave("Item_2_Code") = Me.f_Item_2.ItemData(Me.f_Item_2.ListIndex)
    Else
        rstSave("Item_2_Code") = 0
    End If
    
    If Me.f_Item_3.ListIndex = -1 Then
        FillCombo "Select 0 as ItemCode, '-- Select --' as ItemName ", f_Item_3, "ItemName", "ItemCode"
        Me.f_Item_3.ListIndex = 0
    End If
    
    If Me.f_Item_3.ItemData(Me.f_Item_3.ListIndex) > 0 Then
        rstSave("Item_3_Code") = Me.f_Item_3.ItemData(Me.f_Item_3.ListIndex)
    Else
        rstSave("Item_3_Code") = 0
    End If
    
    If Len(Trim(Me.f_Cone_1.Text)) > 0 Then
        rstSave("Cone_1") = Me.f_Cone_1.Text
    Else
        rstSave("Cone_1") = 0
    End If
    
    If Len(Trim(Me.f_Cone_2.Text)) > 0 Then
        rstSave("Cone_2") = Me.f_Cone_2.Text
    Else
        rstSave("Cone_2") = 0
    End If
    
    If Len(Trim(Me.f_Cone_3.Text)) > 0 Then
        rstSave("Cone_3") = Me.f_Cone_3.Text
    Else
        rstSave("Cone_3") = 0
    End If
    
    If Len(Trim(Me.f_Cone_KG_1.Text)) > 0 Then
        rstSave("Cone_KG_1") = Me.f_Cone_KG_1.Text
    Else
        rstSave("Cone_KG_1") = 0
    End If
    
    If Len(Trim(Me.f_Cone_KG_2.Text)) > 0 Then
        rstSave("Cone_KG_2") = Me.f_Cone_KG_2.Text
    Else
        rstSave("Cone_KG_2") = 0
    End If
    
    If Len(Trim(Me.f_Cone_KG_3.Text)) > 0 Then
        rstSave("Cone_KG_3") = Me.f_Cone_KG_3.Text
    Else
        rstSave("Cone_KG_3") = 0
    End If
    
    If Me.f_Soap_1_Code.ItemData(Me.f_Soap_1_Code.ListIndex) > 0 Then
        rstSave("Soap_1_Code") = Me.f_Soap_1_Code.ItemData(Me.f_Soap_1_Code.ListIndex)
        If Len(Trim(Me.f_Soap_1_Qty.Text)) > 0 Then
            rstSave("Soap_1_Qty") = Me.f_Soap_1_Qty.Text
        Else
            rstSave("Soap_1_Qty") = 0
        End If
    Else
        rstSave("Soap_1_Code") = 0
        rstSave("Soap_1_Qty") = 0
    End If
    
    If Len(Trim(Me.f_Soap_Temp.Text)) > 0 Then
        rstSave("Soap_Temp") = Me.f_Soap_Temp.Text
        rstSave("Soap_Temp_Time") = Me.f_Soap_Temp_Time.Text
    Else
        rstSave("Soap_Temp") = 0
        rstSave("Soap_Temp_Time") = 0
    End If
    
    If Len(Trim(Me.f_Cold_Wash_1.Text)) > 0 Then
        rstSave("Cold_Wash_1") = Me.f_Cold_Wash_1.Text
    Else
        rstSave("Cold_Wash_1") = 0
    End If
    
    If Me.f_Soda_Code.ItemData(Me.f_Soda_Code.ListIndex) > 0 Then
        rstSave("Soda_Code") = Me.f_Soda_Code.ItemData(Me.f_Soda_Code.ListIndex)
        If Len(Trim(Me.f_Soda_Qty.Text)) > 0 Then
            rstSave("Soda_Qty") = Me.f_Soda_Qty.Text
        Else
            rstSave("Soda_Qty") = 0
        End If
    Else
        rstSave("Soda_Code") = 0
        rstSave("Soda_Qty") = 0
    End If
    
    If Me.f_Castic_Code.ItemData(Me.f_Castic_Code.ListIndex) > 0 Then
        rstSave("Castic_Code") = Me.f_Castic_Code.ItemData(Me.f_Castic_Code.ListIndex)
        If Len(Trim(Me.f_Castic_Qty.Text)) > 0 Then
            rstSave("Castic_Qty") = Me.f_Castic_Qty.Text
        Else
            rstSave("Castic_Qty") = 0
        End If
    Else
        rstSave("Castic_Code") = 0
        rstSave("Castic_Qty") = 0
    End If
    
    If Me.f_Soap_2_Code.ItemData(Me.f_Soap_2_Code.ListIndex) > 0 Then
        rstSave("Soap_2_Code") = Me.f_Soap_2_Code.ItemData(Me.f_Soap_2_Code.ListIndex)
        If Len(Trim(Me.f_Soap_2_Qty.Text)) > 0 Then
            rstSave("Soap_2_Qty") = Me.f_Soap_2_Qty.Text
        Else
            rstSave("Soap_2_Qty") = 0
        End If
    Else
        rstSave("Soap_2_Code") = 0
        rstSave("Soap_2_Qty") = 0
    End If
    
    If Me.f_Chemical_1_Code.ItemData(Me.f_Chemical_1_Code.ListIndex) > 0 Then
        rstSave("Chemical_1_Code") = Me.f_Chemical_1_Code.ItemData(Me.f_Chemical_1_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_1_Qty.Text)) > 0 Then
            rstSave("Chemical_1_Qty") = Me.f_Chemical_1_Qty.Text
        Else
            rstSave("Chemical_1_Qty") = 0
        End If
    Else
        rstSave("Chemical_1_Code") = 0
        rstSave("Chemical_1_Qty") = 0
    End If
    
    If Me.f_Chemical_2_Code.ItemData(Me.f_Chemical_2_Code.ListIndex) > 0 Then
        rstSave("Chemical_2_Code") = Me.f_Chemical_2_Code.ItemData(Me.f_Chemical_2_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_2_Qty.Text)) > 0 Then
            rstSave("Chemical_2_Qty") = Me.f_Chemical_2_Qty.Text
        Else
            rstSave("Chemical_2_Qty") = 0
        End If
    Else
        rstSave("Chemical_2_Code") = 0
        rstSave("Chemical_2_Qty") = 0
    End If

    If Me.f_Chemical_3_Code.ItemData(Me.f_Chemical_3_Code.ListIndex) > 0 Then
        rstSave("Chemical_3_Code") = Me.f_Chemical_3_Code.ItemData(Me.f_Chemical_3_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_3_Qty.Text)) > 0 Then
            rstSave("Chemical_3_Qty") = Me.f_Chemical_3_Qty.Text
        Else
            rstSave("Chemical_3_Qty") = 0
        End If
    Else
        rstSave("Chemical_3_Code") = 0
        rstSave("Chemical_3_Qty") = 0
    End If

    If Len(Trim(Me.f_Cold_Wash_2.Text)) > 0 Then
        rstSave("Cold_Wash_2") = Me.f_Cold_Wash_2.Text
    Else
        rstSave("Cold_Wash_2") = 0
    End If

    If Me.f_Acid_Code.ItemData(Me.f_Acid_Code.ListIndex) > 0 Then
        rstSave("Acid_Code") = Me.f_Acid_Code.ItemData(Me.f_Acid_Code.ListIndex)
        If Len(Trim(Me.f_Acid_Qty.Text)) > 0 Then
            rstSave("Acid_Qty") = Me.f_Acid_Qty.Text
        Else
            rstSave("Acid_Qty") = 0
        End If
    Else
        rstSave("Acid_Code") = 0
        rstSave("Acid_Qty") = 0
    End If
    
    If Len(Trim(Me.f_Acid_Temp.Text)) > 0 Then
        rstSave("Acid_Temp") = Me.f_Acid_Temp.Text
        rstSave("Acid_Temp_Time") = Me.f_Acid_Temp_Time.Text
    Else
        rstSave("Acid_Temp") = 0
        rstSave("Acid_Temp_Time") = 0
    End If
    
rstSave.Update
rstSave.Close
Set rstSave = Nothing
m_AddMode = False
Call fillList
End Sub
Private Sub EnableSave()
    If Len(Trim(Me.f_MachineCode)) > 0 And Len(Trim(f_Party_1)) > 0 And Len(Trim(Me.f_ItemType_1)) > 0 And Len(Trim(f_Item_1)) > 0 And Len(Trim(f_Cone_1)) > 0 And Len(Trim(f_Cone_KG_1)) > 0 Then
        Me.cmdSave.Enabled = True
    Else
        Me.cmdSave.Enabled = False
    End If
End Sub
Private Sub f_Soap_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Soap_1_Qty.Text)) > 0 Then
        vQty = Me.f_Soap_1_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Soap_1_Code.ItemData(Me.f_Soap_1_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Soap_Qty(Me.f_Soap_1_Code.ItemData(Me.f_Soap_1_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub chkQty_Soap_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CStr(rstGetQty("Quantity"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(vQty)) > 0) Then
            If (IsNull(vQty)) Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty)), CLng(AvbQty)) < IIf(m_AddMode = False, CLng(vQty), CLng(vQty))) Then
                strAns = MsgBox("Quantity not Available !" & Chr(13) & "Would your like to Continue ", vbYesNo + vbInformation)
                If strAns = vbNo Then
                    QtyBit = 0
                    MsgBit = 0
                    Exit Sub
                    Call EnableSave
                Else
                    vTemp = 1
                    Me.f_Soap_Temp.SetFocus
                End If
            ElseIf vQty <= 0 Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            Else
                QtyBit = 1
                Call EnableSave
            End If
        End If
End Sub
Private Sub chkQty_Soda_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CStr(rstGetQty("Quantity"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(vQty)) > 0) Then
            If (IsNull(vQty)) Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty)), CLng(AvbQty)) < IIf(m_AddMode = False, CLng(vQty), CLng(vQty))) Then
                strAns = MsgBox("Quantity not Available !" & Chr(13) & "Would your like to Continue ", vbYesNo + vbInformation)
                If strAns = vbNo Then
                    QtyBit = 0
                    MsgBit = 0
                    Exit Sub
                    Call EnableSave
                Else
                    vTemp = 1
                    Me.f_Castic_Code.SetFocus
                End If
            ElseIf vQty <= 0 Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            Else
                QtyBit = 1
                Call EnableSave
            End If
        End If
End Sub
Private Sub f_Castic_Qty_GotFocus()
    If Len(Trim(Me.f_Castic_Qty)) > 0 Then
        PreQty = Me.f_Castic_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Castic_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Castic_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Soap_2_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Castic_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Castic_Qty.Text)) > 0 Then
        vQty = Me.f_Castic_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Castic_Code.ItemData(Me.f_Castic_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Castic_Qty(Me.f_Castic_Code.ItemData(Me.f_Castic_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub chkQty_Castic_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CStr(rstGetQty("Quantity"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(vQty)) > 0) Then
            If (IsNull(vQty)) Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty)), CLng(AvbQty)) < IIf(m_AddMode = False, CLng(vQty), CLng(vQty))) Then
                strAns = MsgBox("Quantity not Available !" & Chr(13) & "Would your like to Continue ", vbYesNo + vbInformation)
                If strAns = vbNo Then
                    QtyBit = 0
                    MsgBit = 0
                    Exit Sub
                    Call EnableSave
                Else
                    vTemp = 1
                    Me.f_Soap_2_Code.SetFocus
                End If
            ElseIf vQty <= 0 Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            Else
                QtyBit = 1
                Call EnableSave
            End If
        End If
End Sub
Private Sub f_Soap_2_Qty_GotFocus()
    If Len(Trim(Me.f_Soap_2_Qty)) > 0 Then
        PreQty = Me.f_Soap_2_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Soap_2_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Soap_2_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Chemical_1_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Soap_2_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Soap_2_Qty.Text)) > 0 Then
        vQty = Me.f_Soap_2_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Soap_2_Code.ItemData(Me.f_Soap_2_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Soap_2_Qty(Me.f_Soap_2_Code.ItemData(Me.f_Soap_2_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub chkQty_Soap_2_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CStr(rstGetQty("Quantity"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(vQty)) > 0) Then
            If (IsNull(vQty)) Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty)), CLng(AvbQty)) < IIf(m_AddMode = False, CLng(vQty), CLng(vQty))) Then
                strAns = MsgBox("Quantity not Available !" & Chr(13) & "Would your like to Continue ", vbYesNo + vbInformation)
                If strAns = vbNo Then
                    QtyBit = 0
                    MsgBit = 0
                    Exit Sub
                    Call EnableSave
                Else
                    vTemp = 1
                    Me.f_Chemical_1_Code.SetFocus
                End If
            ElseIf vQty <= 0 Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            Else
                QtyBit = 1
                Call EnableSave
            End If
        End If
End Sub
Private Sub chkQty_Chemical_1_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CStr(rstGetQty("Quantity"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(vQty)) > 0) Then
            If (IsNull(vQty)) Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty)), CLng(AvbQty)) < IIf(m_AddMode = False, CLng(vQty), CLng(vQty))) Then
                strAns = MsgBox("Quantity not Available !" & Chr(13) & "Would your like to Continue ", vbYesNo + vbInformation)
                If strAns = vbNo Then
                    QtyBit = 0
                    MsgBit = 0
                    Exit Sub
                    Call EnableSave
                Else
                    vTemp = 1
                    Me.f_Chemical_2_Code.SetFocus
                End If
            ElseIf vQty <= 0 Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            Else
                QtyBit = 1
                Call EnableSave
            End If
        End If
End Sub
Private Sub chkQty_Chemical_2_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CStr(rstGetQty("Quantity"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(vQty)) > 0) Then
            If (IsNull(vQty)) Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty)), CLng(AvbQty)) < IIf(m_AddMode = False, CLng(vQty), CLng(vQty))) Then
                strAns = MsgBox("Quantity not Available !" & Chr(13) & "Would your like to Continue ", vbYesNo + vbInformation)
                If strAns = vbNo Then
                    QtyBit = 0
                    MsgBit = 0
                    Exit Sub
                    Call EnableSave
                Else
                    vTemp = 1
                    Me.f_Chemical_3_Code.SetFocus
                End If
            ElseIf vQty <= 0 Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            Else
                QtyBit = 1
                Call EnableSave
            End If
        End If
End Sub
Private Sub chkQty_Chemical_3_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CStr(rstGetQty("Quantity"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(vQty)) > 0) Then
            If (IsNull(vQty)) Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty)), CLng(AvbQty)) < IIf(m_AddMode = False, CLng(vQty), CLng(vQty))) Then
                strAns = MsgBox("Quantity not Available !" & Chr(13) & "Would your like to Continue ", vbYesNo + vbInformation)
                If strAns = vbNo Then
                    QtyBit = 0
                    MsgBit = 0
                    Exit Sub
                    Call EnableSave
                Else
                    vTemp = 1
                    Me.f_Cold_Wash_2.SetFocus
                End If
            ElseIf vQty <= 0 Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            Else
                QtyBit = 1
                Call EnableSave
            End If
        End If
End Sub
Private Sub chkQty_Acid_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CStr(rstGetQty("Quantity"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(vQty)) > 0) Then
            If (IsNull(vQty)) Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty)), CLng(AvbQty)) < IIf(m_AddMode = False, CLng(vQty), CLng(vQty))) Then
                strAns = MsgBox("Quantity not Available !" & Chr(13) & "Would your like to Continue ", vbYesNo + vbInformation)
                If strAns = vbNo Then
                    QtyBit = 0
                    MsgBit = 0
                    Exit Sub
                    Call EnableSave
                Else
                    vTemp = 1
                    Me.f_Acid_Temp.SetFocus
                End If
            ElseIf vQty <= 0 Then
                MsgBox "Quantity must be greater then zero"
                QtyBit = 0
                Exit Sub
                Call EnableSave
            Else
                QtyBit = 1
                Call EnableSave
            End If
        End If
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CmdDel_Click()
    If (Len(Trim(m_ListID)) > 0) Then
        Dim strAns As String
        Dim AvbQty As Integer
        Dim rstGetQty As New ADODB.Recordset
        
        strAns = MsgBox("Do you want to delete this record...?", vbYesNo + vbInformation)
        If strAns = vbYes Then
            cnDatabase.Execute "update HalfBleach set Is_Active = 0 where HalfBleachCode =" & m_ListID
            Call fillList
            MsgBox ("Record deleted succesfully..."), vbInformation
            Me.cmdSave.Enabled = False
            Call AddNewRecord
        End If
        m_ListID = ""
        m_AddMode = True
        Me.CmdNew.SetFocus
        End If
End Sub
Private Sub CmdNew_Click()
   Call AddNewRecord
    Me.f_MachineCode.SetFocus
End Sub
Private Sub Cmdhide_Click()
        HalfBleach.Width = 8000
        HalfBleach.Left = 1700
        Me.srHalfBleachCode = ""
        Me.srMachine = ""
        Me.SrItem.ListIndex = -1
        Me.SrItemType.ListIndex = -1
        Me.srParty.ListIndex = -1
        Call fillList
End Sub
Private Sub cmdSave_Click()
If Len(Trim(Me.f_MachineCode)) > 0 And Len(Trim(f_Party_1)) > 0 And Len(Trim(Me.f_ItemType_1)) > 0 And Len(Trim(f_Item_1)) > 0 And Len(Trim(f_Cone_1)) > 0 And Len(Trim(f_Cone_KG_1)) > 0 Then
            Call setVal
            MsgBox ("Record saved successfully"), vbInformation
            Me.f_MachineCode.SetFocus
            Call AddNewRecord
            Call fillList
Else
    MsgBox "Provide data in all Fields"
End If
End Sub
Private Sub CmdAllSearch_Click()
        HalfBleach.Left = 0
        HalfBleach.Width = 11100
        Call SrfillList
End Sub
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Set rstList = FillRecordSet("SELECT HalfBleachCode, HalfBleachDate, PartyName, MachineCode, ItemTypeName, (Select ItemName from Item where ItemCode = Item_1_Code) as ItemName " & _
                                "FROM Party INNER JOIN (ItemType INNER JOIN HalfBleach ON ItemType.ItemTypeCode = HalfBleach.ItemType_1_Code) ON Party.PartyCode = HalfBleach.Party_1_Code where Is_Active = 1 order by HalfBleachCode desc")
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!HalfBleachCode, _
                   Key:=CStr("Id=" & rstList!HalfBleachCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!HalfBleachDate
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!MachineCode
                 .Add Text:=rstList!ItemTypeName
                 .Add Text:=rstList!ItemName
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Public Sub AddNewRecord()
    m_ListID = ""
    Me.f_HalfBleachDate.value = Now
    Me.f_MachineCode.Text = ""
    Me.f_Party_1.ListIndex = -1
    Me.f_Party_2.ListIndex = -1
    Me.f_Party_3.ListIndex = -1
    Me.f_ItemType_1.ListIndex = -1
    Me.f_ItemType_2.ListIndex = -1
    Me.f_ItemType_3.ListIndex = -1
    Me.f_Item_1.ListIndex = -1
    Me.f_Item_2.ListIndex = -1
    Me.f_Item_3.ListIndex = -1
    Me.f_Cone_1.Text = ""
    Me.f_Cone_2.Text = ""
    Me.f_Cone_3.Text = ""
    Me.f_Cone_KG_1.Text = ""
    Me.f_Cone_KG_2.Text = ""
    Me.f_Cone_KG_3.Text = ""
    Me.f_Soap_1_Code.ListIndex = -1
    Me.f_Soap_1_Qty.Text = ""
    Me.f_Soap_Temp.Text = ""
    Me.f_Soap_Temp_Time.Text = ""
    Me.f_Cold_Wash_1.Text = ""
    Me.f_Soda_Code.ListIndex = -1
    Me.f_Soda_Qty.Text = ""
    Me.f_Castic_Code.ListIndex = -1
    Me.f_Castic_Qty.Text = ""
    Me.f_Soap_2_Code.ListIndex = -1
    Me.f_Soap_2_Qty.Text = ""
    Me.f_Chemical_1_Code.ListIndex = -1
    Me.f_Chemical_1_Qty.Text = ""
    Me.f_Chemical_2_Code.ListIndex = -1
    Me.f_Chemical_2_Qty.Text = ""
    Me.f_Chemical_3_Code.ListIndex = -1
    Me.f_Chemical_3_Qty.Text = ""
    Me.f_Cold_Wash_2.Text = ""
    Me.f_Acid_Code.ListIndex = -1
    Me.f_Acid_Qty.Text = ""
    Me.f_Acid_Temp.Text = ""
    Me.f_Acid_Temp_Time.Text = ""
End Sub
Private Sub hbChk_Click()
    If hbChk.value = Checked Then
        Me.srHalfBleachCode.Enabled = True
    Else
        Me.srHalfBleachCode.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub ImChk_Click()
    If ImChk.value = Checked Then
        Me.SrItem.Enabled = True
    Else
        Me.SrItem.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub ImTChk_Click()
    If ImTChk.value = Checked Then
        Me.SrItemType.Enabled = True
    Else
        Me.SrItemType.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub lvwphase_Click()
    cmdSave.Enabled = True
    m_AddMode = False
    If Me.lvwphase.ListItems.Count > 0 Then
        m_ListID = Me.lvwphase.SelectedItem.Text
        ClickPane = 1
        Call getVal
    End If
End Sub
Private Sub lvwphase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSave.Enabled = True
        m_AddMode = False
        If Me.lvwphase.ListItems.Count > 0 Then
            m_ListID = Me.lvwphase.SelectedItem.Text
            ClickPane = 1
            Call getVal
        End If
    End If
End Sub
Private Sub getVal()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("Select * From HalfBleach Where HalfBleachCode = " & m_ListID)
    If Not (rstGetVal.EOF) Then
        Me.f_HalfBleachCode.Text = rstGetVal("HalfBleachCode")
        Me.f_MachineCode.Text = IIf(IsNull(rstGetVal("MachineCode")), 0, rstGetVal("MachineCode"))
        Me.f_HalfBleachDate.value = IIf(IsNull(rstGetVal("HalfBleachDate")), Now, rstGetVal("HalfBleachDate"))
        Call selectValueInCombo(Me.f_Party_1, IIf(IsNull(rstGetVal("Party_1_Code")), -1, rstGetVal("Party_1_Code")))
        Call selectValueInCombo(Me.f_Party_2, IIf(IsNull(rstGetVal("Party_2_Code")), -1, rstGetVal("Party_2_Code")))
        Call selectValueInCombo(Me.f_Party_3, IIf(IsNull(rstGetVal("Party_3_Code")), -1, rstGetVal("Party_3_Code")))
        Call selectValueInCombo(Me.f_ItemType_1, IIf(IsNull(rstGetVal("ItemType_1_Code")), -1, rstGetVal("ItemType_1_Code")))
        Call selectValueInCombo(Me.f_ItemType_2, IIf(IsNull(rstGetVal("ItemType_2_Code")), -1, rstGetVal("ItemType_2_Code")))
        Call selectValueInCombo(Me.f_ItemType_3, IIf(IsNull(rstGetVal("ItemType_3_Code")), -1, rstGetVal("ItemType_3_Code")))
        Call selectValueInCombo(Me.f_Item_1, IIf(IsNull(rstGetVal("Item_1_Code")), -1, rstGetVal("Item_1_Code")))
        Call selectValueInCombo(Me.f_Item_2, IIf(IsNull(rstGetVal("Item_2_Code")), -1, rstGetVal("Item_2_Code")))
        Call selectValueInCombo(Me.f_Item_3, IIf(IsNull(rstGetVal("Item_3_Code")), -1, rstGetVal("Item_3_Code")))
        Me.f_Cone_1.Text = IIf(IsNull(rstGetVal("Cone_1")), 0, rstGetVal("Cone_1"))
        Me.f_Cone_2.Text = IIf(IsNull(rstGetVal("Cone_2")), 0, rstGetVal("Cone_2"))
        Me.f_Cone_3.Text = IIf(IsNull(rstGetVal("Cone_3")), 0, rstGetVal("Cone_3"))
        Me.f_Cone_KG_1.Text = IIf(IsNull(rstGetVal("Cone_KG_1")), 0, rstGetVal("Cone_KG_1"))
        Me.f_Cone_KG_2.Text = IIf(IsNull(rstGetVal("Cone_KG_2")), 0, rstGetVal("Cone_KG_2"))
        Me.f_Cone_KG_3.Text = IIf(IsNull(rstGetVal("Cone_KG_3")), 0, rstGetVal("Cone_KG_3"))
        Call selectValueInCombo(Me.f_Soap_1_Code, IIf(IsNull(rstGetVal("Soap_1_Code")), -1, rstGetVal("Soap_1_Code")))
        Me.f_Soap_1_Qty.Text = IIf(IsNull(rstGetVal("Soap_1_Qty")), 0, rstGetVal("Soap_1_Qty"))
        Me.f_Soap_Temp.Text = IIf(IsNull(rstGetVal("Soap_Temp")), 0, rstGetVal("Soap_Temp"))
        Me.f_Soap_Temp_Time.Text = IIf(IsNull(rstGetVal("Soap_Temp_Time")), 0, rstGetVal("Soap_Temp_Time"))
        Me.f_Cold_Wash_1.Text = IIf(IsNull(rstGetVal("Cold_Wash_1")), 0, rstGetVal("Cold_Wash_1"))
        Call selectValueInCombo(Me.f_Soda_Code, IIf(IsNull(rstGetVal("Soda_Code")), -1, rstGetVal("Soda_Code")))
        Me.f_Soda_Qty.Text = IIf(IsNull(rstGetVal("Soda_Qty")), 0, rstGetVal("Soda_Qty"))
        Call selectValueInCombo(Me.f_Castic_Code, IIf(IsNull(rstGetVal("Castic_Code")), -1, rstGetVal("Castic_Code")))
        Me.f_Castic_Qty.Text = IIf(IsNull(rstGetVal("Castic_Qty")), 0, rstGetVal("Castic_Qty"))
        Call selectValueInCombo(Me.f_Soap_2_Code, IIf(IsNull(rstGetVal("Soap_2_Code")), -1, rstGetVal("Soap_2_Code")))
        Me.f_Soap_2_Qty.Text = IIf(IsNull(rstGetVal("Soap_2_Qty")), 0, rstGetVal("Soap_2_Qty"))
        Call selectValueInCombo(Me.f_Chemical_1_Code, IIf(IsNull(rstGetVal("Chemical_1_Code")), -1, rstGetVal("Chemical_1_Code")))
        Me.f_Chemical_1_Qty.Text = IIf(IsNull(rstGetVal("Chemical_1_Qty")), 0, rstGetVal("Chemical_1_Qty"))
        Call selectValueInCombo(Me.f_Chemical_2_Code, IIf(IsNull(rstGetVal("Chemical_2_Code")), -1, rstGetVal("Chemical_2_Code")))
        Me.f_Chemical_2_Qty.Text = IIf(IsNull(rstGetVal("Chemical_2_Qty")), 0, rstGetVal("Chemical_2_Qty"))
        Call selectValueInCombo(Me.f_Chemical_3_Code, IIf(IsNull(rstGetVal("Chemical_3_Code")), -1, rstGetVal("Chemical_3_Code")))
        Me.f_Chemical_3_Qty.Text = IIf(IsNull(rstGetVal("Chemical_3_Qty")), 0, rstGetVal("Chemical_3_Qty"))
        Me.f_Cold_Wash_2.Text = IIf(IsNull(rstGetVal("Cold_Wash_2")), 0, rstGetVal("Cold_Wash_2"))
        Call selectValueInCombo(Me.f_Acid_Code, IIf(IsNull(rstGetVal("Acid_Code")), -1, rstGetVal("Acid_Code")))
        Me.f_Acid_Qty.Text = IIf(IsNull(rstGetVal("Acid_Qty")), 0, rstGetVal("Acid_Qty"))
        Me.f_Acid_Temp.Text = IIf(IsNull(rstGetVal("Acid_Temp")), 0, rstGetVal("Acid_Temp"))
        Me.f_Acid_Temp_Time.Text = IIf(IsNull(rstGetVal("Acid_Temp_Time")), 0, rstGetVal("Acid_Temp_Time"))
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
End Sub
Private Sub SrfillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Dim sql As String
    Dim cbo1 As String
    Dim cbo2 As String
    Dim cbo3 As String
    Dim cbo4 As String
    Dim cbo5 As String
    If dtChk.value = Checked Then
        srdt = " And (HalfBleachDate between #" & Me.SrDate.value - 1 & " # and #" & Me.SrDate2.value + 1 & " #)"
    Else
        srdt = ""
    End If
    
    If PtChk.value = Checked And Me.srParty.ListIndex > -1 Then
        cbo1 = " And HalfBleach.Party_1_Code = " & Me.srParty.ItemData(Me.srParty.ListIndex)
    Else
        cbo1 = ""
    End If
    
    If ImTChk.value = Checked And Me.SrItemType.ListIndex > -1 Then
        cbo2 = " And HalfBleach.ItemType_1_Code = " & Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
    Else
        cbo2 = ""
    End If
    
    If ImChk.value = Checked And Me.SrItem.ListIndex > -1 Then
        cbo3 = " And HalfBleach.Item_1_Code = " & Me.SrItem.ItemData(Me.SrItem.ListIndex)
    Else
        cbo3 = ""
    End If
   
    If McChk.value = Checked And Len(Trim(Me.srMachine)) > 0 Then
        cbo4 = " And HalfBleach.MachineCode like '%" & Me.srMachine & "%'"
    Else
        cbo4 = ""
     End If
    
    If hbChk.value = Checked And Len(Trim(Me.srHalfBleachCode)) > 0 Then
        cbo5 = " And HalfBleach.HalfBleachCode = " & Me.srHalfBleachCode
    Else
        cbo5 = ""
    End If
    
    sql = "SELECT HalfBleachCode, HalfBleachDate, PartyName, MachineCode, ItemTypeName, (Select ItemName from Item where ItemCode = Item_1_Code) as ItemName " & _
          "FROM Party INNER JOIN (ItemType INNER JOIN HalfBleach ON ItemType.ItemTypeCode = HalfBleach.ItemType_1_Code) ON Party.PartyCode = HalfBleach.Party_1_Code where Is_Active = 1 " & _
          srdt & _
          cbo1 & _
          cbo2 & _
          cbo3 & _
          cbo4 & _
          cbo5 & _
          " order by HalfBleachCode desc"
                                
    Debug.Print sql
    Set rstList = FillRecordSet(sql)
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!HalfBleachCode, _
                   Key:=CStr("Id=" & rstList!HalfBleachCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!HalfBleachDate
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!MachineCode
                 .Add Text:=rstList!ItemTypeName
                 .Add Text:=rstList!ItemName
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub McChk_Click()
    If McChk.value = Checked Then
        Me.srMachine.Enabled = True
    Else
        Me.srMachine.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub PtChk_Click()
    If PtChk.value = Checked Then
        Me.srParty.Enabled = True
    Else
        Me.srParty.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub srHalfBleachCode_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrItem_Click()
    Call SrfillList
End Sub
Private Sub SrItemType_Change()
    Call SrfillList
End Sub
Private Sub SrItemType_Click()
    If Me.SrItemType.ListIndex > 0 Then
        i = Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
        FillCombo "Select ItemCode, ItemName from Item where ItemTypeCode = " & i, SrItem, "ItemName", "ItemCode"
    Else
        Me.SrItem.Clear
    End If
    Call SrfillList
End Sub
Private Sub dtChk_Click()
    If dtChk.value = Checked Then
        Me.SrDate.Enabled = True
        Me.SrDate2.Enabled = True
    Else
        Me.SrDate.Enabled = False
        Me.SrDate2.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub SrDate_Change()
        Call SrfillList
End Sub
Private Sub SrDate2_Change()
    Call SrfillList
End Sub
Private Sub srMachine_KeyUp(KeyCode As Integer, Shift As Integer)
Call SrfillList
End Sub
Private Sub srParty_Change()
    Call SrfillList
End Sub
Private Sub SrParty_Click()
    Call SrfillList
End Sub
