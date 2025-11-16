VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CottonDyeing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotton Dyeing"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   10935
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   120
      TabIndex        =   119
      Top             =   1320
      Width           =   7695
      Begin VB.TextBox f_Chemical_4_Qty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6300
         TabIndex        =   17
         Top             =   550
         Width           =   1215
      End
      Begin VB.ComboBox f_Chemical_4_Code 
         Height          =   315
         Left            =   4605
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   550
         Width           =   1215
      End
      Begin VB.TextBox f_Chemical_2_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6300
         TabIndex        =   13
         Top             =   200
         Width           =   1200
      End
      Begin VB.TextBox f_Chemical_1_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2595
         TabIndex        =   11
         Top             =   200
         Width           =   1200
      End
      Begin VB.ComboBox f_Chemical_1_Code 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   200
         Width           =   1200
      End
      Begin VB.ComboBox f_Chemical_2_Code 
         Height          =   315
         Left            =   4605
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   200
         Width           =   1200
      End
      Begin VB.TextBox f_Chemical_3_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2595
         TabIndex        =   15
         Top             =   550
         Width           =   1200
      End
      Begin VB.ComboBox f_Chemical_3_Code 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   550
         Width           =   1200
      End
      Begin VB.Label Label30 
         Caption         =   "Qty"
         Height          =   255
         Left            =   5895
         TabIndex        =   127
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   3900
         TabIndex        =   126
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   125
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   3900
         TabIndex        =   123
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   121
         Top             =   270
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Qty"
         Height          =   255
         Left            =   5895
         TabIndex        =   120
         Top             =   270
         Width           =   255
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   7440
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
            Picture         =   "Cotton_Dyeing.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cotton_Dyeing.frx":0268
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cotton_Dyeing.frx":06C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cotton_Dyeing.frx":0ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cotton_Dyeing.frx":0F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cotton_Dyeing.frx":1330
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cotton_Dyeing.frx":176C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cotton_Dyeing.frx":1BC0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   118
      Top             =   6120
      Width           =   7695
      Begin MSComctlLib.ListView lvwphase 
         Height          =   1080
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   1905
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
      Height          =   7920
      Index           =   1
      Left            =   7920
      TabIndex        =   108
      Top             =   120
      Width           =   3000
      Begin VB.Frame Frame5 
         Height          =   800
         Index           =   0
         Left            =   120
         TabIndex        =   117
         Top             =   6480
         Width           =   2775
         Begin VB.CheckBox hbChk 
            Caption         =   "H/B Code"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox srHalfBleachCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   86
            Top             =   320
            Width           =   2535
         End
      End
      Begin VB.Frame Frame4 
         Height          =   800
         Left            =   120
         TabIndex        =   116
         Top             =   5640
         Width           =   2775
         Begin VB.CheckBox Dychk 
            Caption         =   "Dyeing Code"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   0
            Width           =   1455
         End
         Begin VB.TextBox srCottonDyeingCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   84
            Top             =   320
            Width           =   2535
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1155
         Left            =   100
         TabIndex        =   114
         Top             =   240
         Width           =   2800
         Begin VB.CheckBox dtChk 
            Caption         =   "Date"
            Height          =   195
            Left            =   240
            TabIndex        =   70
            Top             =   0
            Width           =   735
         End
         Begin MSComCtl2.DTPicker SrDate2 
            Height          =   315
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   44761089
            CurrentDate     =   38298
         End
         Begin MSComCtl2.DTPicker SrDate 
            Height          =   315
            Left            =   125
            TabIndex        =   71
            Top             =   280
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
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
            Format          =   44761089
            CurrentDate     =   38235
         End
      End
      Begin VB.Frame Frame12 
         Height          =   800
         Left            =   100
         TabIndex        =   113
         Top             =   1440
         Width           =   2800
         Begin VB.CheckBox PtChk 
            Caption         =   "Party"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox srParty 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame13 
         Height          =   800
         Left            =   100
         TabIndex        =   112
         Top             =   2280
         Width           =   2800
         Begin VB.CheckBox McChk 
            Caption         =   "Machine"
            Height          =   195
            Left            =   240
            TabIndex        =   75
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
            Height          =   330
            Left            =   125
            TabIndex        =   76
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame14 
         Height          =   800
         Left            =   100
         TabIndex        =   111
         Top             =   3120
         Width           =   2800
         Begin VB.CheckBox ImTChk 
            Caption         =   "Item Type"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   0
            Width           =   1095
         End
         Begin VB.ComboBox SrItemType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame15 
         Height          =   800
         Left            =   100
         TabIndex        =   110
         Top             =   3960
         Width           =   2800
         Begin VB.CheckBox ImChk 
            Caption         =   "Item"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   0
            Width           =   615
         End
         Begin VB.ComboBox SrItem 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame18 
         Height          =   800
         Left            =   120
         TabIndex        =   109
         Top             =   4800
         Width           =   2775
         Begin VB.CheckBox ClChk 
            Caption         =   "Color"
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox SrColor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   82
            Top             =   320
            Width           =   2535
         End
      End
      Begin LVbuttons.LaVolpeButton Cmdhide 
         Height          =   375
         Left            =   360
         TabIndex        =   88
         Top             =   7400
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
         MICON           =   "Cotton_Dyeing.frx":1E38
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
   Begin VB.Frame Frame3 
      Caption         =   "Color"
      Height          =   950
      Left            =   120
      TabIndex        =   105
      Top             =   2280
      Width           =   7695
      Begin VB.ComboBox f_Color_1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   200
         Width           =   1200
      End
      Begin VB.ComboBox f_Color_2 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   200
         Width           =   1200
      End
      Begin VB.ComboBox f_Color_3 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   200
         Width           =   1200
      End
      Begin VB.ComboBox f_Color_4 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   550
         Width           =   1200
      End
      Begin VB.ComboBox f_Color_5 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   550
         Width           =   1200
      End
      Begin VB.ComboBox f_Color_6 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   550
         Width           =   1200
      End
      Begin VB.TextBox f_Color_1_Qty 
         Height          =   315
         Left            =   1320
         TabIndex        =   19
         Top             =   200
         Width           =   1200
      End
      Begin VB.TextBox f_Color_4_Qty 
         Height          =   315
         Left            =   1320
         TabIndex        =   25
         Top             =   550
         Width           =   1200
      End
      Begin VB.TextBox f_Color_2_Qty 
         Height          =   315
         Left            =   3840
         TabIndex        =   21
         Top             =   200
         Width           =   1200
      End
      Begin VB.TextBox f_Color_5_Qty 
         Height          =   315
         Left            =   3840
         TabIndex        =   27
         Top             =   550
         Width           =   1200
      End
      Begin VB.TextBox f_Color_3_Qty 
         Height          =   315
         Left            =   6360
         TabIndex        =   23
         Top             =   200
         Width           =   1200
      End
      Begin VB.TextBox f_Color_6_Qty 
         Height          =   315
         Left            =   6360
         TabIndex        =   29
         Top             =   550
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2850
      Index           =   0
      Left            =   120
      TabIndex        =   96
      Top             =   3240
      Width           =   7695
      Begin VB.TextBox f_Acid_Temp_Time 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6800
         TabIndex        =   42
         Top             =   850
         Width           =   700
      End
      Begin VB.TextBox f_Cold_Wash_5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6800
         TabIndex        =   62
         Top             =   2430
         Width           =   700
      End
      Begin VB.TextBox f_Chemical_6_Temp_Time 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5300
         TabIndex        =   61
         Top             =   2430
         Width           =   700
      End
      Begin VB.TextBox f_Chemical_6_Temp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4100
         TabIndex        =   60
         Top             =   2430
         Width           =   800
      End
      Begin VB.TextBox f_Chemical_6_Qty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2600
         TabIndex        =   59
         Top             =   2430
         Width           =   800
      End
      Begin VB.ComboBox f_Chemical_6_Code 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2430
         Width           =   1200
      End
      Begin VB.TextBox f_Chemical_5_Temp_Time 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6300
         TabIndex        =   57
         Top             =   2120
         Width           =   1200
      End
      Begin VB.TextBox f_Chemical_5_Temp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4100
         TabIndex        =   56
         Top             =   2120
         Width           =   800
      End
      Begin VB.TextBox f_Chemical_5_Qty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2600
         TabIndex        =   55
         Top             =   2120
         Width           =   800
      End
      Begin VB.ComboBox f_Chemical_5_Code 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   2120
         Width           =   1200
      End
      Begin VB.TextBox f_Cold_Wash_4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6300
         TabIndex        =   53
         Top             =   1800
         Width           =   1200
      End
      Begin VB.TextBox f_Hot_Wash_2_Temp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4100
         TabIndex        =   52
         Top             =   1800
         Width           =   800
      End
      Begin VB.TextBox f_Hot_Wash_2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2600
         TabIndex        =   51
         Top             =   1800
         Width           =   800
      End
      Begin VB.TextBox f_Cold_Wash_3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   50
         Top             =   1810
         Width           =   800
      End
      Begin VB.TextBox f_Soap_Temp_Time 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6300
         TabIndex        =   49
         Top             =   1480
         Width           =   1200
      End
      Begin VB.TextBox f_Soap_Temp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4100
         TabIndex        =   48
         Top             =   1480
         Width           =   800
      End
      Begin VB.TextBox f_Soap_Qty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2600
         TabIndex        =   47
         Top             =   1480
         Width           =   800
      End
      Begin VB.ComboBox f_Soap_Code 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1480
         Width           =   1200
      End
      Begin VB.TextBox f_Hot_Wash_1_Temp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6300
         TabIndex        =   45
         Top             =   1170
         Width           =   1200
      End
      Begin VB.TextBox f_Hot_Wash_1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4100
         TabIndex        =   44
         Top             =   1170
         Width           =   800
      End
      Begin VB.TextBox f_Cold_Wash_2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   43
         Top             =   1170
         Width           =   800
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   66500
         TabIndex        =   128
         Text            =   "Text5"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox f_Acid_Temp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5300
         TabIndex        =   41
         Top             =   850
         Width           =   700
      End
      Begin VB.TextBox f_Acid_Qty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   40
         Top             =   850
         Width           =   800
      End
      Begin VB.ComboBox f_Acid_Code 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   850
         Width           =   1200
      End
      Begin VB.TextBox f_Cold_Wash_1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   38
         Top             =   850
         Width           =   800
      End
      Begin VB.TextBox f_Salt_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2500
         TabIndex        =   31
         Top             =   150
         Width           =   1200
      End
      Begin VB.TextBox f_Salt_Temp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4300
         TabIndex        =   32
         Top             =   150
         Width           =   1200
      End
      Begin VB.TextBox f_Salt_Temp_Time 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6300
         TabIndex        =   33
         Top             =   150
         Width           =   1200
      End
      Begin VB.ComboBox f_Salt_Code 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   150
         Width           =   1200
      End
      Begin VB.TextBox f_Soda_Temp_Time 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6300
         TabIndex        =   37
         Top             =   500
         Width           =   1200
      End
      Begin VB.TextBox f_Soda_Qty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2500
         TabIndex        =   35
         Top             =   500
         Width           =   1200
      End
      Begin VB.ComboBox f_Soda_Code 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   500
         Width           =   1200
      End
      Begin VB.TextBox f_Soda_Temp 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4300
         TabIndex        =   36
         Top             =   500
         Width           =   1200
      End
      Begin VB.Label Label51 
         Caption         =   "C/Wash"
         Height          =   255
         Left            =   6120
         TabIndex        =   153
         Top             =   2450
         Width           =   615
      End
      Begin VB.Label Label50 
         Caption         =   "Time"
         Height          =   255
         Left            =   4920
         TabIndex        =   152
         Top             =   2450
         Width           =   375
      End
      Begin VB.Label Label49 
         Caption         =   "Temp."
         Height          =   255
         Left            =   3500
         TabIndex        =   151
         Top             =   2450
         Width           =   495
      End
      Begin VB.Label Label48 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   150
         Top             =   2450
         Width           =   255
      End
      Begin VB.Label Label47 
         Caption         =   "Time"
         Height          =   255
         Left            =   5600
         TabIndex        =   149
         Top             =   2150
         Width           =   375
      End
      Begin VB.Label Label46 
         Caption         =   "Temp."
         Height          =   255
         Left            =   3500
         TabIndex        =   148
         Top             =   2150
         Width           =   495
      End
      Begin VB.Label Label45 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   147
         Top             =   2150
         Width           =   255
      End
      Begin VB.Label Label44 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   2450
         Width           =   735
      End
      Begin VB.Label Label43 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   2150
         Width           =   735
      End
      Begin VB.Label Label42 
         Caption         =   "C/Wash"
         Height          =   255
         Left            =   5600
         TabIndex        =   144
         Top             =   1820
         Width           =   615
      End
      Begin VB.Label Label41 
         Caption         =   "Temp."
         Height          =   255
         Left            =   3500
         TabIndex        =   143
         Top             =   1820
         Width           =   495
      End
      Begin VB.Label Label40 
         Caption         =   "H/Wash"
         Height          =   255
         Left            =   1700
         TabIndex        =   142
         Top             =   1820
         Width           =   615
      End
      Begin VB.Label Label39 
         Caption         =   "C/Wash"
         Height          =   255
         Left            =   120
         TabIndex        =   141
         Top             =   1820
         Width           =   615
      End
      Begin VB.Label Label38 
         Caption         =   "Time"
         Height          =   255
         Left            =   5600
         TabIndex        =   140
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label Label37 
         Caption         =   "Temp."
         Height          =   255
         Left            =   3500
         TabIndex        =   139
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label36 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   138
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label Label35 
         Caption         =   "Soap"
         Height          =   255
         Left            =   120
         TabIndex        =   137
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "Temp."
         Height          =   255
         Left            =   5600
         TabIndex        =   136
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label33 
         Caption         =   "H/Wash"
         Height          =   255
         Left            =   3260
         TabIndex        =   135
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label32 
         Caption         =   "C/Wash"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label31 
         Caption         =   "Time"
         Height          =   255
         Left            =   6300
         TabIndex        =   133
         Top             =   880
         Width           =   375
      End
      Begin VB.Label Label29 
         Caption         =   "Temp."
         Height          =   255
         Left            =   4800
         TabIndex        =   132
         Top             =   880
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "Qty"
         Height          =   255
         Left            =   3480
         TabIndex        =   131
         Top             =   885
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "Acid"
         Height          =   255
         Left            =   1800
         TabIndex        =   130
         Top             =   885
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "C/Wash"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   880
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "Time"
         Height          =   255
         Left            =   5600
         TabIndex        =   104
         Top             =   530
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Temp."
         Height          =   255
         Left            =   3750
         TabIndex        =   103
         Top             =   530
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   102
         Top             =   530
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Salt"
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Soda"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   530
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Time"
         Height          =   255
         Left            =   5600
         TabIndex        =   99
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Temp."
         Height          =   255
         Left            =   3750
         TabIndex        =   98
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Qty"
         Height          =   255
         Left            =   2100
         TabIndex        =   97
         Top             =   180
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   120
      TabIndex        =   87
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox f_Party_1 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   855
         Width           =   1875
      End
      Begin VB.TextBox f_Cone_1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5880
         TabIndex        =   8
         Top             =   840
         Width           =   675
      End
      Begin VB.ComboBox f_ItemType_1 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   855
         Width           =   1755
      End
      Begin VB.ComboBox f_Item_1 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   855
         Width           =   1755
      End
      Begin VB.TextBox f_Cone_KG_1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6600
         TabIndex        =   9
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox f_MachineCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   500
      End
      Begin VB.TextBox f_Color 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3750
         TabIndex        =   2
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox f_HalfBleachCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Width           =   700
      End
      Begin MSComCtl2.DTPicker f_CottonDyeingDate 
         Height          =   300
         Left            =   6240
         TabIndex        =   4
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         Format          =   44761089
         CurrentDate     =   38365
      End
      Begin VB.TextBox f_CottonDyeingCode 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   700
      End
      Begin VB.Label Label26 
         Caption         =   "M/C #"
         Height          =   255
         Left            =   4650
         TabIndex        =   115
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "H/B Code"
         Height          =   255
         Left            =   1720
         TabIndex        =   107
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Color"
         Height          =   255
         Left            =   3330
         TabIndex        =   106
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   255
         Left            =   5760
         TabIndex        =   95
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "KG"
         Height          =   255
         Left            =   6720
         TabIndex        =   94
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Cone"
         Height          =   255
         Left            =   6000
         TabIndex        =   93
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Item"
         Height          =   255
         Left            =   4440
         TabIndex        =   92
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Item Type"
         Height          =   255
         Left            =   2760
         TabIndex        =   91
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Party"
         Height          =   255
         Left            =   840
         TabIndex        =   90
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Dye Code"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   270
         Width           =   735
      End
   End
   Begin LVbuttons.LaVolpeButton CmdAllSearch 
      Height          =   405
      Left            =   5280
      TabIndex        =   65
      Top             =   7680
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
      MICON           =   "Cotton_Dyeing.frx":1E54
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
      TabIndex        =   66
      Top             =   7680
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
      MICON           =   "Cotton_Dyeing.frx":1E70
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
      TabIndex        =   64
      Top             =   7680
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
      MICON           =   "Cotton_Dyeing.frx":1E8C
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
      TabIndex        =   63
      Top             =   7680
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
      MICON           =   "Cotton_Dyeing.frx":1EA8
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
      TabIndex        =   67
      Top             =   7680
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
      MICON           =   "Cotton_Dyeing.frx":1EC4
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
      TabIndex        =   68
      Top             =   7680
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
      MICON           =   "Cotton_Dyeing.frx":1EE0
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
Attribute VB_Name = "CottonDyeing"
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
Private Sub f_Cold_Wash_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Acid_Code.SetFocus
    End If
End Sub
Private Sub f_Cold_Wash_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Hot_Wash_1.SetFocus
    End If
End Sub
Private Sub f_Cold_Wash_3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Hot_Wash_2.SetFocus
    End If
End Sub
Private Sub f_Cold_Wash_4_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_5_Code.SetFocus
    End If
End Sub
Private Sub f_Cold_Wash_5_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.cmdSave.SetFocus
    End If
End Sub
Private Sub f_Hot_Wash_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Hot_Wash_1_Temp.SetFocus
    End If
End Sub
Private Sub f_Hot_Wash_1_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap_Code.SetFocus
    End If
End Sub
Private Sub f_Hot_Wash_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Hot_Wash_2_Temp.SetFocus
    End If
End Sub
Private Sub f_Hot_Wash_2_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cold_Wash_4.SetFocus
    End If
End Sub
Private Sub f_Soap_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap_Qty.SetFocus
    End If
End Sub
Private Sub f_Soap_Code_LostFocus()
    If Me.f_Soap_Code.ListIndex = -1 Then
        Me.f_Soap_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Soap_Qty_GotFocus()
    If Len(Trim(Me.f_Soap_Qty)) > 0 Then
        PreQty = Me.f_Soap_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Soap_Qty_KeyPress(KeyAscii As Integer)
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
Private Sub f_Soap_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Soap_Qty.Text)) > 0 Then
        vQty = Me.f_Soap_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Soap_Code.ItemData(Me.f_Soap_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Soap_Qty(Me.f_Soap_Code.ItemData(Me.f_Soap_Code.ListIndex), vQty)
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
Private Sub f_Soap_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap_Temp_Time.SetFocus
    End If
End Sub
Private Sub f_Soap_Temp_Time_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cold_Wash_3.SetFocus
    End If
End Sub
Private Sub Form_Load()
  m_AddMode = True
  cmdSave.Enabled = False
  DBConn
  f_CottonDyeingDate = Now
  SrDate = Now
  SrDate2 = Now
  Dim sql As String
  FillColorCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", f_Party_1, "PartyName", "PartyCode"
  FillColorCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", srParty, "PartyName", "PartyCode"
  
  FillColorCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", f_ItemType_1, "ItemTypeName", "ItemTypeCode"
  FillColorCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", SrItemType, "ItemTypeName", "ItemTypeCode"
     
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_1_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_2_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_3_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_4_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_5_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_6_Code, "ItemName", "ItemCode"
  
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 15 order by 2", f_Salt_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 16 order by 2", f_Soda_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 10 order by 2", f_Soap_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 7 order by 2", f_Acid_Code, "ItemName", "ItemCode"

  sql = "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2"
  FillColorCombo sql, f_Color_1, "ItemName", "ItemCode"
  FillColorCombo sql, f_Color_2, "ItemName", "ItemCode"
  FillColorCombo sql, f_Color_3, "ItemName", "ItemCode"
  FillColorCombo sql, f_Color_4, "ItemName", "ItemCode"
  FillColorCombo sql, f_Color_5, "ItemName", "ItemCode"
  FillColorCombo sql, f_Color_6, "ItemName", "ItemCode"

  lvwphase.ColumnHeaders.Add Text:="Code", Width:=600
  lvwphase.ColumnHeaders.Add Text:="Date", Width:=1200
  lvwphase.ColumnHeaders.Add Text:="Party Name", Width:=1700
  lvwphase.ColumnHeaders.Add Text:="Machine #", Width:=1000
  lvwphase.ColumnHeaders.Add Text:="Item Type", Width:=1500
  lvwphase.ColumnHeaders.Add Text:="Item", Width:=1490
  
  Call fillList

End Sub
Private Sub f_Party_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_ItemType_1.SetFocus
    End If
End Sub
Private Sub f_Party_1_LostFocus()
    If Me.f_Party_1.ListIndex = -1 Then
        Me.f_Party_1.ListIndex = 0
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
Private Sub f_ItemType_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Item_1.SetFocus
    End If
End Sub
Private Sub f_ItemType_1_LostFocus()
    If Me.f_ItemType_1.ListIndex = -1 Then
        Me.f_ItemType_1.ListIndex = 0
    End If
End Sub
Private Sub f_Item_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cone_1.SetFocus
    End If
End Sub
Private Sub f_Item_1_LostFocus()
    If Me.f_Item_1.ListIndex = -1 Then
        FillCombo "Select 0 as ItemCode, '-- Select --' as ItemName ", f_Item_1, "ItemName", "ItemCode"
        Me.f_Item_1.ListIndex = 0
    End If
End Sub
Private Sub f_Cone_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cone_KG_1.SetFocus
    End If
End Sub
Private Sub f_Cone_KG_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_1_Code.SetFocus
    End If
End Sub
Private Sub f_Chemical_1_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_1_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_1_Code_LostFocus()
    If Me.f_Chemical_1_Code.ListIndex = -1 Then
        Me.f_Chemical_1_Code.ListIndex = 0
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
         '       MsgBox QtyBit
                Call EnableSave
            End If
        End If
End Sub
Private Sub f_Chemical_2_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_2_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_2_Code_LostFocus()
    If Me.f_Chemical_2_Code.ListIndex = -1 Then
        Me.f_Chemical_2_Code.ListIndex = 0
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
Private Sub f_Chemical_3_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_3_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_3_Code_LostFocus()
    If Me.f_Chemical_3_Code.ListIndex = -1 Then
        Me.f_Chemical_3_Code.ListIndex = 0
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
            Me.f_Chemical_4_Code.SetFocus
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
                    Me.f_Chemical_4_Code.SetFocus
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
Private Sub f_Chemical_4_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_4_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_4_Code_LostFocus()
    If Me.f_Chemical_4_Code.ListIndex = -1 Then
        Me.f_Chemical_4_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Chemical_4_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_4_Qty)) > 0 Then
        PreQty = Me.f_Chemical_4_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_4_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Chemical_4_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Color_1.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Chemical_4_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Chemical_4_Qty.Text)) > 0 Then
        vQty = Me.f_Chemical_4_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Chemical_4_Code.ItemData(Me.f_Chemical_4_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Chemical_4_Qty(Me.f_Chemical_4_Code.ItemData(Me.f_Chemical_4_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub chkQty_Chemical_4_Qty(vItemCode As Integer, vQty As Double)
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
                    Me.f_Color_1.SetFocus
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
Private Sub f_Chemical_5_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_5_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_5_Code_LostFocus()
    If Me.f_Chemical_5_Code.ListIndex = -1 Then
        Me.f_Chemical_5_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Chemical_5_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_5_Qty)) > 0 Then
        PreQty = Me.f_Chemical_5_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_5_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Chemical_5_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Chemical_5_Temp.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Chemical_5_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Chemical_5_Qty.Text)) > 0 Then
        vQty = Me.f_Chemical_5_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Chemical_5_Code.ItemData(Me.f_Chemical_5_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Chemical_5_Qty(Me.f_Chemical_5_Code.ItemData(Me.f_Chemical_5_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub chkQty_Chemical_5_Qty(vItemCode As Integer, vQty As Double)
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
                    Me.f_Chemical_5_Temp.SetFocus
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
Private Sub f_Chemical_5_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_5_Temp_Time.SetFocus
    End If
End Sub
Private Sub f_Chemical_5_Temp_Time_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_6_Code.SetFocus
    End If
End Sub
Private Sub f_Chemical_6_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_6_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_6_Code_LostFocus()
    If Me.f_Chemical_6_Code.ListIndex = -1 Then
        Me.f_Chemical_6_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Chemical_6_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_6_Qty)) > 0 Then
        PreQty = Me.f_Chemical_6_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_6_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Chemical_6_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Chemical_6_Temp.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Chemical_6_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Chemical_6_Qty.Text)) > 0 Then
        vQty = Me.f_Chemical_6_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Chemical_6_Code.ItemData(Me.f_Chemical_6_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Chemical_6_Qty(Me.f_Chemical_6_Code.ItemData(Me.f_Chemical_6_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub chkQty_Chemical_6_Qty(vItemCode As Integer, vQty As Double)
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
                    Me.f_Chemical_6_Temp.SetFocus
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
Private Sub f_Chemical_6_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_6_Temp_Time.SetFocus
    End If
End Sub
Private Sub f_Chemical_6_Temp_Time_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cold_Wash_5.SetFocus
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
            Me.f_Soda_Temp.SetFocus
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
                    Me.f_Soda_Temp.SetFocus
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
Private Sub f_Soda_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soda_Temp_Time.SetFocus
    End If
End Sub
Private Sub f_Soda_Temp_Time_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cold_Wash_1.SetFocus
    End If
End Sub
Private Sub f_Acid_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Acid_Qty.SetFocus
    End If
End Sub
Private Sub f_Acid_Code_LostFocus()
    If Me.f_Acid_Code.ListIndex = -1 Then
        Me.f_Acid_Code.ListIndex = 0
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
Private Sub f_Acid_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Acid_Temp_Time.SetFocus
    End If
End Sub
Private Sub f_Acid_Temp_Time_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cold_Wash_2.SetFocus
    End If
End Sub
Private Sub f_Salt_Code_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Salt_Qty.SetFocus
    End If
End Sub
Private Sub f_Salt_Code_LostFocus()
    If Me.f_Salt_Code.ListIndex = -1 Then
        Me.f_Salt_Code.ListIndex = 0
    End If
End Sub
Private Sub f_Salt_Qty_GotFocus()
    If Len(Trim(Me.f_Salt_Qty)) > 0 Then
        PreQty = Me.f_Salt_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Salt_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Salt_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Salt_Temp.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Salt_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Salt_Qty.Text)) > 0 Then
        vQty = Me.f_Salt_Qty.Text
    Else
        vQty = 0
    End If
    If Me.f_Salt_Code.ItemData(Me.f_Salt_Code.ListIndex) > 0 And MsgBit = 0 Then
        Call chkQty_Salt_Qty(Me.f_Salt_Code.ItemData(Me.f_Salt_Code.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub chkQty_Salt_Qty(vItemCode As Integer, vQty As Double)
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
                    Me.f_Salt_Temp.SetFocus
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
Private Sub f_Salt_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Salt_Temp_Time.SetFocus
    End If
End Sub
Private Sub f_Salt_Temp_Time_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soda_Code.SetFocus
    End If
End Sub
Private Sub EnableSave()
    If Len(Trim(Me.f_HalfBleachCode)) > 0 And Len(Trim(Me.f_MachineCode)) > 0 And Len(Trim(f_Party_1)) > 0 And Len(Trim(Me.f_ItemType_1)) > 0 And Len(Trim(f_Item_1)) > 0 And Len(Trim(f_Cone_1)) > 0 And Len(Trim(f_Cone_KG_1)) > 0 Then
        Me.cmdSave.Enabled = True
    Else
        Me.cmdSave.Enabled = False
    End If
End Sub
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Set rstList = FillRecordSet("SELECT ProcessCode as CottonDyeingCode, ProcessDate as CottonDyeingDate, PartyName, MachineNo as MachineCode, ItemTypeName, (Select ItemName from Item where ItemCode = Den) as ItemName " & _
                                "FROM Party INNER JOIN (ItemType INNER JOIN Process ON ItemType.ItemTypeCode = Process.ItemTypeCode) ON Party.PartyCode = Process.PartyCode where Is_Active = 1 and Is_Cotton_Dyeing = 1 order by ProcessCode desc")
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!CottonDyeingCode, _
                   Key:=CStr("Id=" & rstList!CottonDyeingCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!CottonDyeingDate
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
Private Sub SrfillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Dim sql As String
    Dim cbo1 As String
    Dim cbo2 As String
    Dim cbo3 As String
    Dim cbo4 As String
    Dim cbo5 As String
    Dim cbo6 As String
    Dim cbo7 As String
    Dim cbo8 As String
    If dtChk.value = Checked Then
        srdt = " And (ProcessDate between #" & Me.SrDate.value - 1 & " # and #" & Me.SrDate2.value + 1 & " #)"
    Else
        srdt = ""
    End If
    
    If PtChk.value = Checked And Me.srParty.ListIndex > -1 Then
        cbo1 = " And Process.PartyCode = " & Me.srParty.ItemData(Me.srParty.ListIndex)
    Else
        cbo1 = ""
    End If
    
    If ImTChk.value = Checked And Me.SrItemType.ListIndex > -1 Then
        cbo2 = " And Process.ItemTypeCode = " & Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
    Else
        cbo2 = ""
    End If
    
    If ImChk.value = Checked And Me.SrItem.ListIndex > -1 Then
        cbo3 = " And Process.Den = " & Me.SrItem.ItemData(Me.SrItem.ListIndex)
    Else
        cbo3 = ""
    End If
   
    If McChk.value = Checked And Len(Trim(Me.srMachine)) > 0 Then
        cbo4 = " And Process.MachineNo like '%" & Me.srMachine & "%'"
    Else
        cbo4 = ""
     End If
    
    If ClChk.value = Checked And Len(Trim(Me.SrColor)) > 0 Then
        cbo5 = " And Process.NewColor like '%" & Me.SrColor & "%'"
    Else
        cbo5 = ""
    End If
    
    If Dychk.value = Checked And Len(Trim(Me.srCottonDyeingCode)) > 0 Then
        cbo6 = " And Process.ProcessCode = " & Me.srCottonDyeingCode
    Else
        cbo6 = ""
    End If
    
    If hbChk.value = Checked And Len(Trim(Me.srHalfBleachCode)) > 0 Then
        cbo7 = " And Process.HalfBleachCode = " & Me.srHalfBleachCode
    Else
        cbo7 = ""
    End If
    
    sql = "SELECT ProcessCode as CottonDyeingCode, ProcessDate as CottonDyeingDate, PartyName, MachineNo as MachineCode, ItemTypeName, (Select ItemName from Item where ItemCode = Den) as ItemName " & _
                                "FROM Party INNER JOIN (ItemType INNER JOIN Process ON ItemType.ItemTypeCode = Process.ItemTypeCode) ON Party.PartyCode = Process.PartyCode where Is_Active = 1 and Is_Cotton_Dyeing = 1" & _
          srdt & _
          cbo1 & _
          cbo2 & _
          cbo3 & _
          cbo4 & _
          cbo5 & _
          cbo6 & _
          cbo7 & _
          " order by ProcessCode desc"
                                
    Debug.Print sql
    Set rstList = FillRecordSet(sql)
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!CottonDyeingCode, _
                   Key:=CStr("Id=" & rstList!CottonDyeingCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!CottonDyeingDate
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
Private Sub getVal()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("Select * From Process Where Is_Cotton_Dyeing = 1 and ProcessCode = " & m_ListID)
    If Not (rstGetVal.EOF) Then
        Me.f_CottonDyeingCode.Text = rstGetVal("ProcessCode")
        Me.f_HalfBleachCode.Text = IIf(IsNull(rstGetVal("HalfBleachCode")), 0, rstGetVal("HalfBleachCode"))
        Me.f_MachineCode.Text = IIf(IsNull(rstGetVal("MachineNo")), 0, rstGetVal("MachineNo"))
        Me.f_Color.Text = IIf(IsNull(rstGetVal("NewColor")), "", rstGetVal("NewColor"))
        Me.f_CottonDyeingDate.value = IIf(IsNull(rstGetVal("ProcessDate")), Now, rstGetVal("ProcessDate"))
        Call selectValueInCombo(Me.f_Party_1, IIf(IsNull(rstGetVal("PartyCode")), -1, rstGetVal("PartyCode")))
        Call selectValueInCombo(Me.f_ItemType_1, IIf(IsNull(rstGetVal("ItemTypeCode")), -1, rstGetVal("ItemTypeCode")))
        Call selectValueInCombo(Me.f_Item_1, IIf(IsNull(rstGetVal("Den")), -1, rstGetVal("Den")))
        Me.f_Cone_1.Text = IIf(IsNull(rstGetVal("Cone")), 0, rstGetVal("Cone"))
        Me.f_Cone_KG_1.Text = IIf(IsNull(rstGetVal("ConeKG")), 0, rstGetVal("ConeKG"))
        Call selectValueInCombo(Me.f_Color_1, IIf(IsNull(rstGetVal("Color_1")), -1, rstGetVal("Color_1")))
        Me.f_Color_1_Qty.Text = IIf(IsNull(rstGetVal("Color_1_Qty")), 0, rstGetVal("Color_1_Qty"))
        Call selectValueInCombo(Me.f_Color_2, IIf(IsNull(rstGetVal("Color_2")), 0, rstGetVal("Color_2")))
        Me.f_Color_2_Qty.Text = IIf(IsNull(rstGetVal("Color_2_Qty")), 0, rstGetVal("Color_2_Qty"))
        Call selectValueInCombo(Me.f_Color_3, IIf(IsNull(rstGetVal("Color_3")), 0, rstGetVal("Color_3")))
        Me.f_Color_3_Qty.Text = IIf(IsNull(rstGetVal("Color_3_Qty")), 0, rstGetVal("Color_3_Qty"))
        Call selectValueInCombo(Me.f_Color_4, IIf(IsNull(rstGetVal("Color_4")), 0, rstGetVal("Color_4")))
        Me.f_Color_4_Qty.Text = IIf(IsNull(rstGetVal("Color_4_Qty")), 0, rstGetVal("Color_4_Qty"))
        Call selectValueInCombo(Me.f_Color_5, IIf(IsNull(rstGetVal("Color_5")), 0, rstGetVal("Color_5")))
        Me.f_Color_5_Qty.Text = IIf(IsNull(rstGetVal("Color_5_Qty")), 0, rstGetVal("Color_5_Qty"))
        Call selectValueInCombo(Me.f_Color_6, IIf(IsNull(rstGetVal("Color_6")), 0, rstGetVal("Color_6")))
        Me.f_Color_6_Qty.Text = IIf(IsNull(rstGetVal("Color_6_Qty")), 0, rstGetVal("Color_6_Qty"))
        Call selectValueInCombo(Me.f_Chemical_1_Code, IIf(IsNull(rstGetVal("Chemical")), -1, rstGetVal("Chemical")))
        Me.f_Chemical_1_Qty.Text = IIf(IsNull(rstGetVal("Chemical_Qty")), 0, rstGetVal("Chemical_Qty"))
        Call selectValueInCombo(Me.f_Chemical_2_Code, IIf(IsNull(rstGetVal("Chemical2")), -1, rstGetVal("Chemical2")))
        Me.f_Chemical_2_Qty.Text = IIf(IsNull(rstGetVal("Chemical2_Qty")), 0, rstGetVal("Chemical2_Qty"))
        Call selectValueInCombo(Me.f_Chemical_3_Code, IIf(IsNull(rstGetVal("Chemical_3_Code")), -1, rstGetVal("Chemical_3_Code")))
        Me.f_Chemical_3_Qty.Text = IIf(IsNull(rstGetVal("Chemical_3_Qty")), 0, rstGetVal("Chemical_3_Qty"))
        Call selectValueInCombo(Me.f_Chemical_4_Code, IIf(IsNull(rstGetVal("Chemical_4_Code")), -1, rstGetVal("Chemical_4_Code")))
        Me.f_Chemical_4_Qty.Text = IIf(IsNull(rstGetVal("Chemical_4_Qty")), 0, rstGetVal("Chemical_4_Qty"))
        Call selectValueInCombo(Me.f_Chemical_5_Code, IIf(IsNull(rstGetVal("Chemical_5_Code")), -1, rstGetVal("Chemical_5_Code")))
        Me.f_Chemical_5_Qty.Text = IIf(IsNull(rstGetVal("Chemical_5_Qty")), 0, rstGetVal("Chemical_5_Qty"))
        Me.f_Chemical_5_Temp.Text = IIf(IsNull(rstGetVal("Chemical_5_Temp")), 0, rstGetVal("Chemical_5_Temp"))
        Me.f_Chemical_5_Temp_Time.Text = IIf(IsNull(rstGetVal("Chemical_5_Temp_Time")), 0, rstGetVal("Chemical_5_Temp_Time"))
        Call selectValueInCombo(Me.f_Chemical_6_Code, IIf(IsNull(rstGetVal("Chemical_6_Code")), -1, rstGetVal("Chemical_6_Code")))
        Me.f_Chemical_6_Qty.Text = IIf(IsNull(rstGetVal("Chemical_6_Qty")), 0, rstGetVal("Chemical_6_Qty"))
        Me.f_Chemical_6_Temp.Text = IIf(IsNull(rstGetVal("Chemical_6_Temp")), 0, rstGetVal("Chemical_6_Temp"))
        Me.f_Chemical_6_Temp_Time.Text = IIf(IsNull(rstGetVal("Chemical_6_Temp_Time")), 0, rstGetVal("Chemical_6_Temp_Time"))
        Call selectValueInCombo(Me.f_Salt_Code, IIf(IsNull(rstGetVal("salt_Code")), -1, rstGetVal("salt_Code")))
        Me.f_Salt_Qty.Text = IIf(IsNull(rstGetVal("salt_Qty")), 0, rstGetVal("salt_Qty"))
        Me.f_Salt_Temp.Text = IIf(IsNull(rstGetVal("salt_Temp")), 0, rstGetVal("salt_Temp"))
        Me.f_Salt_Temp_Time.Text = IIf(IsNull(rstGetVal("salt_Temp_Time")), 0, rstGetVal("salt_Temp_Time"))
        Call selectValueInCombo(Me.f_Soda_Code, IIf(IsNull(rstGetVal("Soda_Code")), -1, rstGetVal("Soda_Code")))
        Me.f_Soda_Qty.Text = IIf(IsNull(rstGetVal("Soda_Qty")), 0, rstGetVal("Soda_Qty"))
        Me.f_Soda_Temp.Text = IIf(IsNull(rstGetVal("Soda_Temp")), 0, rstGetVal("Soda_Temp"))
        Me.f_Soda_Temp_Time.Text = IIf(IsNull(rstGetVal("Soda_Temp_Time")), 0, rstGetVal("Soda_Temp_Time"))
        Call selectValueInCombo(Me.f_Acid_Code, IIf(IsNull(rstGetVal("Acid")), -1, rstGetVal("Acid")))
        Me.f_Acid_Qty.Text = IIf(IsNull(rstGetVal("Acid_Qty")), 0, rstGetVal("Acid_Qty"))
        Me.f_Acid_Temp.Text = IIf(IsNull(rstGetVal("Acid_Temp")), 0, rstGetVal("Acid_Temp"))
        Me.f_Acid_Temp_Time.Text = IIf(IsNull(rstGetVal("Acid_Temp_Time")), 0, rstGetVal("Acid_Temp_Time"))
        Call selectValueInCombo(Me.f_Soap_Code, IIf(IsNull(rstGetVal("Soap")), -1, rstGetVal("Soap")))
        Me.f_Soap_Qty.Text = IIf(IsNull(rstGetVal("Soap_Qty")), 0, rstGetVal("Soap_Qty"))
        Me.f_Soap_Temp.Text = IIf(IsNull(rstGetVal("Soap_Temp")), 0, rstGetVal("Soap_Temp"))
        Me.f_Soap_Temp_Time.Text = IIf(IsNull(rstGetVal("Soap_Temp_Time")), 0, rstGetVal("Soap_Temp_Time"))
        Me.f_Cold_Wash_1.Text = IIf(IsNull(rstGetVal("Cold_Wash_1")), 0, rstGetVal("Cold_Wash_1"))
        Me.f_Cold_Wash_2.Text = IIf(IsNull(rstGetVal("Cold_Wash_2")), 0, rstGetVal("Cold_Wash_2"))
        Me.f_Cold_Wash_3.Text = IIf(IsNull(rstGetVal("Cold_Wash_3")), 0, rstGetVal("Cold_Wash_3"))
        Me.f_Cold_Wash_4.Text = IIf(IsNull(rstGetVal("Cold_Wash_4")), 0, rstGetVal("Cold_Wash_4"))
        Me.f_Cold_Wash_5.Text = IIf(IsNull(rstGetVal("Cold_Wash_5")), 0, rstGetVal("Cold_Wash_5"))
        Me.f_Hot_Wash_1.Text = IIf(IsNull(rstGetVal("Hot_Wash_1")), 0, rstGetVal("Hot_Wash_1"))
        Me.f_Hot_Wash_1_Temp.Text = IIf(IsNull(rstGetVal("Hot_Wash_1_Temp")), 0, rstGetVal("Hot_Wash_1_Temp"))
        Me.f_Hot_Wash_2.Text = IIf(IsNull(rstGetVal("Hot_Wash_2")), 0, rstGetVal("Hot_Wash_2"))
        Me.f_Hot_Wash_2_Temp.Text = IIf(IsNull(rstGetVal("Hot_Wash_2_Temp")), 0, rstGetVal("Hot_Wash_2_Temp"))
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
End Sub
Public Sub setVal()
    Dim rstSave As New ADODB.Recordset
    If (Len(Trim(m_ListID)) = 0) Then
        Set rstSave = FillRecordSet("select * from Process where 1 = 2")
        rstSave.AddNew
        m_ListID = ValAutoNumber("Process", "ProcessCode")
        rstSave("ProcessCode") = m_ListID
    Else
       Set rstSave = FillRecordSet("select * from Process where ProcessCode =" & m_ListID)
    End If
    rstSave("HalfBleachCode") = Me.f_HalfBleachCode.Text
    rstSave("NewColor") = Me.f_Color.Text
    rstSave("MachineNo") = Me.f_MachineCode.Text
    rstSave("ProcessDate") = Me.f_CottonDyeingDate.value
    
    If Me.f_Party_1.ItemData(Me.f_Party_1.ListIndex) > 0 Then
        rstSave("PartyCode") = Me.f_Party_1.ItemData(Me.f_Party_1.ListIndex)
    Else
        rstSave("PartyCode") = 0
    End If

    If Me.f_ItemType_1.ItemData(Me.f_ItemType_1.ListIndex) > 0 Then
        rstSave("ItemTypeCode") = Me.f_ItemType_1.ItemData(Me.f_ItemType_1.ListIndex)
    Else
        rstSave("ItemTypeCode") = 0
    End If
    
    If Me.f_Item_1.ItemData(Me.f_Item_1.ListIndex) > 0 Then
        rstSave("Den") = Me.f_Item_1.ItemData(Me.f_Item_1.ListIndex)
    Else
        rstSave("Den") = 0
    End If
    
    If Len(Trim(Me.f_Cone_1.Text)) > 0 Then
        rstSave("Cone") = Me.f_Cone_1.Text
    Else
        rstSave("Cone") = 0
    End If
    
    If Len(Trim(Me.f_Cone_KG_1.Text)) > 0 Then
        rstSave("ConeKG") = Me.f_Cone_KG_1.Text
    Else
        rstSave("ConeKG") = 0
    End If
    
    If Me.f_Color_1.ItemData(Me.f_Color_1.ListIndex) > 0 Then
        rstSave("Color_1") = Me.f_Color_1.ItemData(Me.f_Color_1.ListIndex)
        If Len(Trim(Me.f_Color_1_Qty.Text)) > 0 Then
            rstSave("Color_1_Qty") = Me.f_Color_1_Qty.Text
        Else
            rstSave("Color_1_Qty") = 0
        End If
    Else
        rstSave("Color_1") = 0
        rstSave("Color_1_Qty") = 0
    End If
    
    If Me.f_Color_2.ItemData(Me.f_Color_2.ListIndex) > 0 Then
        rstSave("Color_2") = Me.f_Color_2.ItemData(Me.f_Color_2.ListIndex)
        If Len(Trim(Me.f_Color_2_Qty.Text)) > 0 Then
            rstSave("Color_2_Qty") = Me.f_Color_2_Qty.Text
        Else
            rstSave("Color_2_Qty") = 0
        End If
    Else
        rstSave("Color_2") = 0
        rstSave("Color_2_Qty") = 0
    End If
    
    If Me.f_Color_3.ItemData(Me.f_Color_3.ListIndex) > 0 Then
        rstSave("Color_3") = Me.f_Color_3.ItemData(Me.f_Color_3.ListIndex)
        If Len(Trim(Me.f_Color_3_Qty.Text)) > 0 Then
            rstSave("Color_3_Qty") = Me.f_Color_3_Qty.Text
        Else
            rstSave("Color_3_Qty") = 0
        End If
    Else
        rstSave("Color_3") = 0
        rstSave("Color_3_Qty") = 0
    End If
    
    If Me.f_Color_4.ItemData(Me.f_Color_4.ListIndex) > 0 Then
        rstSave("Color_4") = Me.f_Color_4.ItemData(Me.f_Color_4.ListIndex)
        If Len(Trim(Me.f_Color_4_Qty.Text)) > 0 Then
            rstSave("Color_4_Qty") = Me.f_Color_4_Qty.Text
        Else
            rstSave("Color_4_Qty") = 0
        End If
    Else
        rstSave("Color_4") = 0
        rstSave("Color_4_Qty") = 0
    End If
    
    If Me.f_Color_5.ItemData(Me.f_Color_5.ListIndex) > 0 Then
        rstSave("Color_5") = Me.f_Color_5.ItemData(Me.f_Color_5.ListIndex)
        If Len(Trim(Me.f_Color_5_Qty.Text)) > 0 Then
            rstSave("Color_5_Qty") = Me.f_Color_5_Qty.Text
        Else
            rstSave("Color_5_Qty") = 0
        End If
    Else
        rstSave("Color_5") = 0
        rstSave("Color_5_Qty") = 0
    End If
    
    If Me.f_Color_6.ItemData(Me.f_Color_6.ListIndex) > 0 Then
        rstSave("Color_6") = Me.f_Color_6.ItemData(Me.f_Color_6.ListIndex)
        If Len(Trim(Me.f_Color_6_Qty.Text)) > 0 Then
            rstSave("Color_6_Qty") = Me.f_Color_6_Qty.Text
        Else
            rstSave("Color_6_Qty") = 0
        End If
    Else
        rstSave("Color_6") = 0
        rstSave("Color_6_Qty") = 0
    End If
    
    If Me.f_Chemical_1_Code.ItemData(Me.f_Chemical_1_Code.ListIndex) > 0 Then
        rstSave("Chemical") = Me.f_Chemical_1_Code.ItemData(Me.f_Chemical_1_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_1_Qty.Text)) > 0 Then
            rstSave("Chemical_Qty") = Me.f_Chemical_1_Qty.Text
        Else
            rstSave("Chemical_Qty") = 0
        End If
    Else
        rstSave("Chemical") = 0
        rstSave("Chemical_Qty") = 0
    End If
    
    If Me.f_Chemical_2_Code.ItemData(Me.f_Chemical_2_Code.ListIndex) > 0 Then
        rstSave("Chemical2") = Me.f_Chemical_2_Code.ItemData(Me.f_Chemical_2_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_2_Qty.Text)) > 0 Then
            rstSave("Chemical2_Qty") = Me.f_Chemical_2_Qty.Text
        Else
            rstSave("Chemical2_Qty") = 0
        End If
    Else
        rstSave("Chemical2") = 0
        rstSave("Chemical2_Qty") = 0
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

    If Me.f_Chemical_4_Code.ItemData(Me.f_Chemical_4_Code.ListIndex) > 0 Then
        rstSave("Chemical_4_Code") = Me.f_Chemical_4_Code.ItemData(Me.f_Chemical_4_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_4_Qty.Text)) > 0 Then
            rstSave("Chemical_4_Qty") = Me.f_Chemical_4_Qty.Text
        Else
            rstSave("Chemical_4_Qty") = 0
        End If
    Else
        rstSave("Chemical_4_Code") = 0
        rstSave("Chemical_4_Qty") = 0
    End If
    
    If Me.f_Chemical_5_Code.ItemData(Me.f_Chemical_5_Code.ListIndex) > 0 Then
        rstSave("Chemical_5_Code") = Me.f_Chemical_5_Code.ItemData(Me.f_Chemical_5_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_5_Qty.Text)) > 0 Then
            rstSave("Chemical_5_Qty") = Me.f_Chemical_5_Qty.Text
        Else
            rstSave("Chemical_5_Qty") = 0
        End If
        If Len(Trim(Me.f_Chemical_5_Temp.Text)) > 0 Then
            rstSave("Chemical_5_Temp") = Me.f_Chemical_5_Temp.Text
        Else
            rstSave("Chemical_5_Temp") = 0
        End If
        If Len(Trim(Me.f_Chemical_5_Temp_Time.Text)) > 0 Then
            rstSave("Chemical_5_Temp_Time") = Me.f_Chemical_5_Temp_Time.Text
        Else
            rstSave("Chemical_5_Temp_Time") = 0
        End If
    Else
        rstSave("Chemical_5_Code") = 0
        rstSave("Chemical_5_Qty") = 0
        rstSave("Chemical_5_Temp") = 0
        rstSave("Chemical_5_Temp_Time") = 0
    End If
    
    If Me.f_Chemical_6_Code.ItemData(Me.f_Chemical_6_Code.ListIndex) > 0 Then
        rstSave("Chemical_6_Code") = Me.f_Chemical_6_Code.ItemData(Me.f_Chemical_6_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_6_Qty.Text)) > 0 Then
            rstSave("Chemical_6_Qty") = Me.f_Chemical_6_Qty.Text
        Else
            rstSave("Chemical_6_Qty") = 0
        End If
        If Len(Trim(Me.f_Chemical_6_Temp.Text)) > 0 Then
            rstSave("Chemical_6_Temp") = Me.f_Chemical_6_Temp.Text
        Else
            rstSave("Chemical_6_Temp") = 0
        End If
        If Len(Trim(Me.f_Chemical_6_Temp_Time.Text)) > 0 Then
            rstSave("Chemical_6_Temp_Time") = Me.f_Chemical_6_Temp_Time.Text
        Else
            rstSave("Chemical_6_Temp_Time") = 0
        End If
    Else
        rstSave("Chemical_6_Code") = 0
        rstSave("Chemical_6_Qty") = 0
        rstSave("Chemical_6_Temp") = 0
        rstSave("Chemical_6_Temp_Time") = 0
    End If
    
    If Me.f_Salt_Code.ItemData(Me.f_Salt_Code.ListIndex) > 0 Then
        rstSave("Salt_Code") = Me.f_Salt_Code.ItemData(Me.f_Salt_Code.ListIndex)
        If Len(Trim(Me.f_Salt_Qty.Text)) > 0 Then
            rstSave("Salt_Qty") = Me.f_Salt_Qty.Text
        Else
            rstSave("Salt_Qty") = 0
        End If
        If Len(Trim(Me.f_Salt_Temp.Text)) > 0 Then
            rstSave("Salt_Temp") = Me.f_Salt_Temp.Text
        Else
            rstSave("Salt_Temp") = 0
        End If
        If Len(Trim(Me.f_Salt_Temp_Time.Text)) > 0 Then
            rstSave("Salt_Temp_Time") = Me.f_Salt_Temp_Time.Text
        Else
            rstSave("Salt_Temp_Time") = 0
        End If
    Else
        rstSave("Salt_Code") = 0
        rstSave("Salt_Qty") = 0
        rstSave("Salt_Temp") = 0
        rstSave("Salt_Temp_Time") = 0
    End If
    
    If Me.f_Soda_Code.ItemData(Me.f_Soda_Code.ListIndex) > 0 Then
        rstSave("Soda_Code") = Me.f_Soda_Code.ItemData(Me.f_Soda_Code.ListIndex)
        If Len(Trim(Me.f_Soda_Qty.Text)) > 0 Then
            rstSave("Soda_Qty") = Me.f_Soda_Qty.Text
        Else
            rstSave("Soda_Qty") = 0
        End If
        If Len(Trim(Me.f_Soda_Temp.Text)) > 0 Then
            rstSave("Soda_Temp") = Me.f_Soda_Temp.Text
        Else
            rstSave("Soda_Temp") = 0
        End If
        If Len(Trim(Me.f_Soda_Temp_Time.Text)) > 0 Then
            rstSave("Soda_Temp_Time") = Me.f_Soda_Temp_Time.Text
        Else
            rstSave("Soda_Temp_Time") = 0
        End If
    Else
        rstSave("Soda_Code") = 0
        rstSave("Soda_Qty") = 0
        rstSave("Soda_Temp") = 0
        rstSave("Soda_Temp_Time") = 0
    End If
    
    If Me.f_Acid_Code.ItemData(Me.f_Acid_Code.ListIndex) > 0 Then
        rstSave("Acid") = Me.f_Acid_Code.ItemData(Me.f_Acid_Code.ListIndex)
        If Len(Trim(Me.f_Acid_Qty.Text)) > 0 Then
            rstSave("Acid_Qty") = Me.f_Acid_Qty.Text
        Else
            rstSave("Acid_Qty") = 0
        End If
        If Len(Trim(Me.f_Acid_Temp.Text)) > 0 Then
            rstSave("Acid_Temp") = Me.f_Acid_Temp.Text
        Else
            rstSave("Acid_Temp") = 0
        End If
        If Len(Trim(Me.f_Acid_Temp_Time.Text)) > 0 Then
            rstSave("Acid_Temp_Time") = Me.f_Acid_Temp_Time.Text
        Else
            rstSave("Acid_Temp_Time") = 0
        End If
    Else
        rstSave("Acid") = 0
        rstSave("Acid_Qty") = 0
        rstSave("Acid_Temp") = 0
        rstSave("Acid_Temp_Time") = 0
    End If
    
    If Me.f_Soap_Code.ItemData(Me.f_Soap_Code.ListIndex) > 0 Then
        rstSave("Soap") = Me.f_Soap_Code.ItemData(Me.f_Soap_Code.ListIndex)
        If Len(Trim(Me.f_Soap_Qty.Text)) > 0 Then
            rstSave("Soap_Qty") = Me.f_Soap_Qty.Text
        Else
            rstSave("Soap_Qty") = 0
        End If
        If Len(Trim(Me.f_Soap_Temp.Text)) > 0 Then
            rstSave("Soap_Temp") = Me.f_Soap_Temp.Text
        Else
            rstSave("Soap_Temp") = 0
        End If
        If Len(Trim(Me.f_Soap_Temp_Time.Text)) > 0 Then
            rstSave("Soap_Temp_Time") = Me.f_Soap_Temp_Time.Text
        Else
            rstSave("Soap_Temp_Time") = 0
        End If
    Else
        rstSave("Soap") = 0
        rstSave("Soap_Qty") = 0
        rstSave("Soap_Temp") = 0
        rstSave("Soap_Temp_Time") = 0
    End If
    
    If Len(Trim(Me.f_Cold_Wash_1.Text)) > 0 Then
        rstSave("Cold_Wash_1") = Me.f_Cold_Wash_1.Text
    Else
        rstSave("Cold_Wash_1") = 0
    End If

    If Len(Trim(Me.f_Cold_Wash_2.Text)) > 0 Then
        rstSave("Cold_Wash_2") = Me.f_Cold_Wash_2.Text
    Else
        rstSave("Cold_Wash_2") = 0
    End If

    If Len(Trim(Me.f_Cold_Wash_3.Text)) > 0 Then
        rstSave("Cold_Wash_3") = Me.f_Cold_Wash_3.Text
    Else
        rstSave("Cold_Wash_3") = 0
    End If

    If Len(Trim(Me.f_Cold_Wash_4.Text)) > 0 Then
        rstSave("Cold_Wash_4") = Me.f_Cold_Wash_4.Text
    Else
        rstSave("Cold_Wash_4") = 0
    End If

    If Len(Trim(Me.f_Cold_Wash_5.Text)) > 0 Then
        rstSave("Cold_Wash_5") = Me.f_Cold_Wash_5.Text
    Else
        rstSave("Cold_Wash_5") = 0
    End If
    
    If Len(Trim(Me.f_Hot_Wash_1.Text)) > 0 Then
        rstSave("Hot_Wash_1") = Me.f_Hot_Wash_1.Text
        If Len(Trim(Me.f_Hot_Wash_1_Temp.Text)) > 0 Then
            rstSave("Hot_Wash_1_Temp") = Me.f_Hot_Wash_1_Temp.Text
        Else
            rstSave("Hot_Wash_1_Temp") = 0
        End If
    Else
        rstSave("Hot_Wash_1") = 0
        rstSave("Hot_Wash_1_Temp") = 0
    End If

    If Len(Trim(Me.f_Hot_Wash_2.Text)) > 0 Then
        rstSave("Hot_Wash_2") = Me.f_Hot_Wash_2.Text
        If Len(Trim(Me.f_Hot_Wash_2_Temp.Text)) > 0 Then
            rstSave("Hot_Wash_2_Temp") = Me.f_Hot_Wash_2_Temp.Text
        Else
            rstSave("Hot_Wash_2_Temp") = 0
        End If
    Else
        rstSave("Hot_Wash_2") = 0
        rstSave("Hot_Wash_2_Temp") = 0
    End If
    rstSave("Is_Cotton_Dyeing") = 1
    
rstSave.Update
rstSave.Close
Set rstSave = Nothing
m_AddMode = False
Call fillList
End Sub
Private Sub cmdSave_Click()
If Len(Trim(Me.f_HalfBleachCode)) > 0 And Len(Trim(Me.f_MachineCode)) > 0 And Len(Trim(f_Party_1)) > 0 And Len(Trim(Me.f_ItemType_1)) > 0 And Len(Trim(f_Item_1)) > 0 And Len(Trim(f_Cone_1)) > 0 And Len(Trim(f_Cone_KG_1)) > 0 Then
            Call setVal
            MsgBox ("Record saved successfully"), vbInformation
            Me.f_HalfBleachCode.SetFocus
            Call AddNewRecord
            Call fillList
Else
    MsgBox "Provide data in all Fields"
End If
End Sub
Public Sub AddNewRecord()
    m_ListID = ""
    Me.f_HalfBleachCode.Text = ""
    Me.f_Color.Text = ""
    Me.f_MachineCode.Text = ""
    Me.f_CottonDyeingDate.value = Now
    Me.f_Party_1.ListIndex = -1
    Me.f_ItemType_1.ListIndex = -1
    Me.f_Item_1.ListIndex = -1
    Me.f_Cone_1.Text = ""
    Me.f_Cone_KG_1.Text = ""
    Me.f_Color_1.ListIndex = -1
    Me.f_Color_1_Qty.Text = ""
    Me.f_Color_2.ListIndex = -1
    Me.f_Color_2_Qty.Text = ""
    Me.f_Color_3.ListIndex = -1
    Me.f_Color_3_Qty.Text = ""
    Me.f_Color_4.ListIndex = -1
    Me.f_Color_4_Qty.Text = ""
    Me.f_Color_5.ListIndex = -1
    Me.f_Color_5_Qty.Text = ""
    Me.f_Color_6.ListIndex = -1
    Me.f_Color_6_Qty.Text = ""
    Me.f_Chemical_1_Code.ListIndex = -1
    Me.f_Chemical_1_Qty.Text = ""
    Me.f_Chemical_2_Code.ListIndex = -1
    Me.f_Chemical_2_Qty.Text = ""
    Me.f_Chemical_3_Code.ListIndex = -1
    Me.f_Chemical_3_Qty.Text = ""
    Me.f_Chemical_4_Code.ListIndex = -1
    Me.f_Chemical_4_Qty.Text = ""
    Me.f_Chemical_5_Code.ListIndex = -1
    Me.f_Chemical_5_Qty.Text = ""
    Me.f_Chemical_5_Temp.Text = ""
    Me.f_Chemical_5_Temp_Time.Text = ""
    Me.f_Chemical_6_Code.ListIndex = -1
    Me.f_Chemical_6_Qty.Text = ""
    Me.f_Chemical_6_Temp.Text = ""
    Me.f_Chemical_6_Temp_Time.Text = ""
    Me.f_Salt_Code.ListIndex = -1
    Me.f_Salt_Qty.Text = ""
    Me.f_Salt_Temp.Text = ""
    Me.f_Salt_Temp_Time.Text = ""
    Me.f_Soda_Code.ListIndex = -1
    Me.f_Soda_Qty.Text = ""
    Me.f_Soda_Temp.Text = ""
    Me.f_Soda_Temp_Time.Text = ""
    Me.f_Acid_Code.ListIndex = -1
    Me.f_Acid_Qty.Text = ""
    Me.f_Acid_Temp.Text = ""
    Me.f_Acid_Temp_Time.Text = ""
    Me.f_Soap_Code.ListIndex = -1
    Me.f_Soap_Qty.Text = ""
    Me.f_Soap_Temp.Text = ""
    Me.f_Soap_Temp_Time.Text = ""
    Me.f_Cold_Wash_1.Text = ""
    Me.f_Cold_Wash_2.Text = ""
    Me.f_Cold_Wash_3.Text = ""
    Me.f_Cold_Wash_4.Text = ""
    Me.f_Cold_Wash_5.Text = ""
    Me.f_Hot_Wash_1.Text = ""
    Me.f_Hot_Wash_1_Temp.Text = ""
    Me.f_Hot_Wash_2.Text = ""
    Me.f_Hot_Wash_2_Temp.Text = ""
    Me.f_CottonDyeingCode.Text = ""
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
            cnDatabase.Execute "update CottonDyeing set Is_Active = 0 where CottonDyeingCode =" & m_ListID
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
    Me.f_HalfBleachCode.SetFocus
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
Private Sub srHalfBleachCode_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrItem_Click()
    Call SrfillList
End Sub
Private Sub SrItemType_Change()
    Call SrfillList
End Sub
Private Sub SrColor_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub srCottonDyeingCode_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub ClChk_Click()
    If ClChk.value = Checked Then
        Me.SrColor.Enabled = True
    Else
        Me.SrColor.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub CmdAllSearch_Click()
        CottonDyeing.Left = 0
        CottonDyeing.Width = 11100
        Call SrfillList
End Sub
Private Sub Cmdhide_Click()
        CottonDyeing.Width = 8000
        CottonDyeing.Left = 1700
        Me.srHalfBleachCode = ""
        Me.srCottonDyeingCode = ""
        Me.SrColor = ""
        Me.srMachine = ""
        Me.SrItem.ListIndex = -1
        Me.SrItemType.ListIndex = -1
        Me.srParty.ListIndex = -1
        Call fillList
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
Private Sub Dychk_Click()
    If Dychk.value = Checked Then
        Me.srCottonDyeingCode.Enabled = True
    Else
        Me.srCottonDyeingCode.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub f_Color_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_1_Qty.SetFocus
    End If
End Sub
Private Sub f_Color_1_LostFocus()
    If Me.f_Color_1.ListIndex = -1 Then
        Me.f_Color_1.ListIndex = 0
    End If
End Sub
Private Sub f_Color_1_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_2.SetFocus
    End If
End Sub
Private Sub f_Color_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_2_Qty.SetFocus
    End If
End Sub
Private Sub f_Color_2_LostFocus()
    If Me.f_Color_2.ListIndex = -1 Then
        Me.f_Color_2.ListIndex = 0
    End If
End Sub
Private Sub f_Color_2_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_3.SetFocus
    End If
End Sub
Private Sub f_Color_3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_3_Qty.SetFocus
    End If
End Sub
Private Sub f_Color_3_LostFocus()
    If Me.f_Color_3.ListIndex = -1 Then
        Me.f_Color_3.ListIndex = 0
    End If
End Sub
Private Sub f_Color_3_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_4.SetFocus
    End If
End Sub
Private Sub f_Color_4_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_4_Qty.SetFocus
    End If
End Sub
Private Sub f_Color_4_LostFocus()
    If Me.f_Color_4.ListIndex = -1 Then
        Me.f_Color_4.ListIndex = 0
    End If
End Sub
Private Sub f_Color_4_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_5.SetFocus
    End If
End Sub
Private Sub f_Color_5_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_5_Qty.SetFocus
    End If
End Sub
Private Sub f_Color_5_LostFocus()
    If Me.f_Color_5.ListIndex = -1 Then
        Me.f_Color_5.ListIndex = 0
    End If
End Sub
Private Sub f_Color_5_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_6.SetFocus
    End If
End Sub
Private Sub f_Color_6_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_6_Qty.SetFocus
    End If
End Sub
Private Sub f_Color_6_LostFocus()
    If Me.f_Color_6.ListIndex = -1 Then
        Me.f_Color_6.ListIndex = 0
    End If
End Sub
Private Sub f_Color_6_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Salt_Code.SetFocus
    End If
End Sub
Private Sub f_CottonDyeingCode_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_HalfBleachCode.SetFocus
    End If
End Sub
Private Sub f_HalfBleachCode_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color.SetFocus
    End If
End Sub
Private Sub f_Color_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_MachineCode.SetFocus
    End If
End Sub
Private Sub f_MachineCode_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_CottonDyeingDate.SetFocus
    End If
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
Private Sub SrItemType_Click()
    If Me.SrItemType.ListIndex > 0 Then
        i = Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
        FillCombo "Select ItemCode, ItemName from Item where ItemTypeCode = " & i, SrItem, "ItemName", "ItemCode"
    Else
        Me.SrItem.Clear
    End If
    Call SrfillList
End Sub
