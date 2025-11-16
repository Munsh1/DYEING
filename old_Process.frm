VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form old_Process 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                     ----- Process -----"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11880
   Begin Crystal.CrystalReport crptDaily 
      Left            =   0
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   5640
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
            Picture         =   "old_Process.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "old_Process.frx":0268
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "old_Process.frx":06C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "old_Process.frx":0ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "old_Process.frx":0F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "old_Process.frx":1330
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "old_Process.frx":176C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "old_Process.frx":1BC0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   1695
      Left            =   120
      TabIndex        =   83
      Top             =   5880
      Width           =   8535
      Begin MSComctlLib.ListView lvwphase 
         Height          =   1320
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   8205
         _ExtentX        =   14473
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
   Begin VB.Frame Frame4 
      Caption         =   "Re. Color"
      Height          =   1100
      Left            =   120
      TabIndex        =   82
      Top             =   4750
      Width           =   8535
      Begin VB.TextBox f_R_Color_4_Qty 
         Height          =   315
         Left            =   1750
         TabIndex        =   78
         Top             =   650
         Width           =   1000
      End
      Begin VB.TextBox f_R_Color_1_Qty 
         Height          =   315
         Left            =   1750
         TabIndex        =   75
         Top             =   250
         Width           =   1000
      End
      Begin VB.TextBox f_R_Color_5_Qty 
         Height          =   315
         Left            =   4450
         TabIndex        =   79
         Top             =   650
         Width           =   1000
      End
      Begin VB.TextBox f_R_Color_2_Qty 
         Height          =   315
         Left            =   4450
         TabIndex        =   76
         Top             =   250
         Width           =   1000
      End
      Begin VB.TextBox f_R_Color_6_Qty 
         Height          =   315
         Left            =   7150
         TabIndex        =   80
         Top             =   650
         Width           =   1000
      End
      Begin VB.TextBox f_R_Color_3_Qty 
         Height          =   315
         Left            =   7150
         TabIndex        =   77
         Top             =   250
         Width           =   1000
      End
      Begin VB.ComboBox f_R_Color_6 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   650
         Width           =   1500
      End
      Begin VB.ComboBox f_R_Color_5 
         Height          =   315
         Left            =   2925
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   650
         Width           =   1500
      End
      Begin VB.ComboBox f_R_Color_4 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   650
         Width           =   1500
      End
      Begin VB.ComboBox f_R_Color_3 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   250
         Width           =   1500
      End
      Begin VB.ComboBox f_R_Color_2 
         Height          =   315
         Left            =   2925
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   250
         Width           =   1500
      End
      Begin VB.ComboBox f_R_Color_1 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   250
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Color"
      Height          =   1100
      Left            =   240
      TabIndex        =   81
      Top             =   2640
      Width           =   8295
      Begin VB.TextBox f_Color_3_Qty 
         Height          =   315
         Left            =   6930
         TabIndex        =   72
         Top             =   250
         Width           =   1000
      End
      Begin VB.TextBox f_Color_5_Qty 
         Height          =   315
         Left            =   4270
         TabIndex        =   74
         Top             =   650
         Width           =   1000
      End
      Begin VB.TextBox f_Color_2_Qty 
         Height          =   315
         Left            =   4270
         TabIndex        =   71
         Top             =   250
         Width           =   1000
      End
      Begin VB.TextBox f_Color_4_Qty 
         Height          =   315
         Left            =   1635
         TabIndex        =   73
         Top             =   650
         Width           =   1000
      End
      Begin VB.TextBox f_Color_1_Qty 
         Height          =   315
         Left            =   1635
         TabIndex        =   70
         Top             =   250
         Width           =   1000
      End
      Begin VB.ComboBox f_Color_5 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   650
         Width           =   1500
      End
      Begin VB.ComboBox f_Color_4 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   650
         Width           =   1500
      End
      Begin VB.ComboBox f_Color_3 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   250
         Width           =   1500
      End
      Begin VB.ComboBox f_Color_2 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   250
         Width           =   1500
      End
      Begin VB.ComboBox f_Color_1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   250
         Width           =   1500
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
      Height          =   7935
      Left            =   8800
      TabIndex        =   0
      Top             =   120
      Width           =   3000
      Begin VB.Frame Frame19 
         Height          =   855
         Left            =   120
         TabIndex        =   124
         Top             =   6360
         Width           =   2775
         Begin VB.CheckBox PCChk 
            Caption         =   "PC Code"
            Height          =   255
            Left            =   240
            TabIndex        =   127
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox srPC2 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1440
            TabIndex        =   126
            Top             =   360
            Width           =   1000
         End
         Begin VB.TextBox srPC1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   240
            TabIndex        =   125
            Top             =   360
            Width           =   1000
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   120
         TabIndex        =   119
         Top             =   5520
         Width           =   2775
         Begin VB.TextBox SrColor 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   67
            Top             =   320
            Width           =   2535
         End
         Begin VB.CheckBox ClChk 
            Caption         =   "Color"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.CheckBox ImChk 
         Caption         =   "Item"
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   4560
         Width           =   735
      End
      Begin VB.CheckBox ImTChk 
         Caption         =   "Item Type"
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CheckBox McChk 
         Caption         =   "Machine"
         Height          =   255
         Left            =   360
         TabIndex        =   60
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox PtChk 
         Caption         =   "Party"
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   1680
         Width           =   735
      End
      Begin VB.Frame Frame15 
         Height          =   800
         Left            =   100
         TabIndex        =   109
         Top             =   4560
         Width           =   2800
         Begin VB.ComboBox SrItem 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            TabIndex        =   65
            Text            =   "SrItem"
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame14 
         Height          =   800
         Left            =   100
         TabIndex        =   108
         Top             =   3600
         Width           =   2800
         Begin VB.ComboBox SrItemType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            TabIndex        =   63
            Text            =   "SrItemType"
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame13 
         Height          =   800
         Left            =   100
         TabIndex        =   107
         Top             =   2640
         Width           =   2800
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
            TabIndex        =   61
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame12 
         Height          =   800
         Left            =   100
         TabIndex        =   106
         Top             =   1680
         Width           =   2800
         Begin VB.ComboBox srParty 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            TabIndex        =   59
            Text            =   "srParty"
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1155
         Left            =   100
         TabIndex        =   105
         Top             =   360
         Width           =   2800
         Begin MSComCtl2.DTPicker SrDate2 
            Height          =   330
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   44695553
            CurrentDate     =   38298
         End
         Begin VB.CheckBox dtChk 
            Caption         =   "Date"
            Height          =   195
            Left            =   240
            TabIndex        =   55
            Top             =   0
            Width           =   735
         End
         Begin MSComCtl2.DTPicker SrDate 
            Height          =   330
            Left            =   125
            TabIndex        =   56
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
            Format          =   44695553
            CurrentDate     =   38235
         End
      End
      Begin LVbuttons.LaVolpeButton Cmdhide 
         Height          =   375
         Left            =   480
         TabIndex        =   68
         Top             =   7440
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
         MICON           =   "old_Process.frx":1E38
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
   Begin LVbuttons.LaVolpeButton CmdAllSearch 
      Height          =   405
      Left            =   6300
      TabIndex        =   50
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
      MICON           =   "old_Process.frx":1E54
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
      Left            =   7560
      TabIndex        =   51
      Top             =   7680
      Width           =   1100
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
      MICON           =   "old_Process.frx":1E70
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
   Begin LVbuttons.LaVolpeButton CmdRecipe 
      Height          =   405
      Left            =   2760
      TabIndex        =   52
      Top             =   7680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Recipe"
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
      MICON           =   "old_Process.frx":1E8C
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "7"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdNew 
      Height          =   405
      Left            =   5160
      TabIndex        =   49
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
      MICON           =   "old_Process.frx":1EA8
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
      Left            =   4000
      TabIndex        =   48
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
      MICON           =   "old_Process.frx":1EC4
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
      Left            =   1520
      TabIndex        =   53
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
      MICON           =   "old_Process.frx":1EE0
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
      TabIndex        =   54
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
      MICON           =   "old_Process.frx":1EFC
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
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   84
      Top             =   120
      Width           =   8535
      Begin VB.TextBox f_Chemical_4_Qty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4480
         TabIndex        =   40
         Top             =   4120
         Width           =   1000
      End
      Begin VB.TextBox f_Chemical_3_Qty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3400
         TabIndex        =   38
         Top             =   4120
         Width           =   1000
      End
      Begin VB.ComboBox f_Chemical_4_Code 
         Height          =   315
         Left            =   4630
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3780
         Width           =   870
      End
      Begin VB.ComboBox f_Chemical_3_Code 
         Height          =   315
         Left            =   3550
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   3780
         Width           =   870
      End
      Begin VB.TextBox f_NewColor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2500
         TabIndex        =   36
         Top             =   4120
         Width           =   840
      End
      Begin VB.Frame Frame17 
         Height          =   825
         Left            =   120
         TabIndex        =   115
         Top             =   3600
         Width           =   2325
         Begin VB.TextBox f_Acid2_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1230
            TabIndex        =   35
            Top             =   480
            Width           =   1000
         End
         Begin VB.TextBox f_Soap2_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   100
            TabIndex        =   33
            Top             =   480
            Width           =   1000
         End
         Begin VB.ComboBox f_Acid2 
            Height          =   315
            Left            =   1370
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   140
            Width           =   870
         End
         Begin VB.ComboBox f_Soap2 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   140
            Width           =   870
         End
         Begin VB.Label Label22 
            Caption         =   "Ac"
            Height          =   255
            Left            =   1160
            TabIndex        =   117
            Top             =   180
            Width           =   255
         End
         Begin VB.Label Label17 
            Caption         =   "Ø"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   180
            Width           =   135
         End
      End
      Begin VB.Frame Frame16 
         Height          =   825
         Left            =   6315
         TabIndex        =   110
         Top             =   1680
         Width           =   2065
         Begin VB.CheckBox f_Re_RecipeCode 
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox f_RecipeCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   26
            Top             =   480
            Width           =   800
         End
         Begin VB.Label Label20 
            Caption         =   "Re. Recipe"
            Height          =   225
            Left            =   120
            TabIndex        =   112
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Recipe"
            Height          =   225
            Left            =   1200
            TabIndex        =   111
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.TextBox f_Remarks 
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   5550
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   3900
         Width           =   2895
      End
      Begin VB.ComboBox f_PartyCode 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "f_PartyCode"
         Top             =   450
         Width           =   3255
      End
      Begin VB.TextBox f_MachineNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6240
         TabIndex        =   4
         Top             =   450
         Width           =   1000
      End
      Begin VB.TextBox f_Den 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7320
         TabIndex        =   5
         Top             =   450
         Width           =   1000
      End
      Begin VB.Frame Frame6 
         Height          =   825
         Left            =   120
         TabIndex        =   89
         Top             =   1680
         Width           =   2325
         Begin VB.ComboBox f_Soap 
            Height          =   315
            Left            =   260
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   140
            Width           =   870
         End
         Begin VB.TextBox f_Soap_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   130
            TabIndex        =   18
            Top             =   480
            Width           =   1000
         End
         Begin VB.TextBox f_SoapTime 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   19
            Top             =   460
            Width           =   1000
         End
         Begin VB.Label Label14 
            Caption         =   "Time"
            Height          =   225
            Left            =   1440
            TabIndex        =   101
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Ø"
            Height          =   225
            Left            =   120
            TabIndex        =   100
            Top             =   180
            Width           =   135
         End
      End
      Begin VB.Frame Frame7 
         Height          =   825
         Left            =   2640
         TabIndex        =   88
         Top             =   1680
         Width           =   3435
         Begin VB.TextBox f_Hydro_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1215
            TabIndex        =   23
            Top             =   480
            Width           =   1005
         End
         Begin VB.ComboBox f_Hydro 
            Height          =   315
            Left            =   1335
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   140
            Width           =   870
         End
         Begin VB.ComboBox f_Castic 
            Height          =   315
            Left            =   260
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   140
            Width           =   870
         End
         Begin VB.TextBox f_Castic_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   130
            TabIndex        =   21
            Top             =   460
            Width           =   1000
         End
         Begin VB.TextBox f_CasticTime 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2300
            TabIndex        =   24
            Top             =   460
            Width           =   1000
         End
         Begin VB.Label Label13 
            Caption         =   "H"
            Height          =   225
            Left            =   1200
            TabIndex        =   113
            Top             =   165
            Width           =   135
         End
         Begin VB.Label Label16 
            Caption         =   "Time"
            Height          =   225
            Left            =   2600
            TabIndex        =   103
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "C"
            Height          =   225
            Left            =   120
            TabIndex        =   102
            Top             =   180
            Width           =   135
         End
      End
      Begin VB.Frame Frame8 
         Height          =   825
         Left            =   120
         TabIndex        =   87
         Top             =   840
         Width           =   2415
         Begin VB.ComboBox f_Cone 
            Height          =   315
            Left            =   80
            TabIndex        =   7
            Text            =   "f_Cone"
            Top             =   475
            Width           =   1600
         End
         Begin VB.ComboBox f_ItemTypeCode 
            Height          =   315
            Left            =   80
            TabIndex        =   6
            Text            =   "f_ItemTypeCode"
            Top             =   150
            Width           =   1815
         End
         Begin VB.TextBox f_ConeKG 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1700
            TabIndex        =   8
            Top             =   475
            Width           =   645
         End
         Begin VB.Label Label7 
            Caption         =   "KG."
            Height          =   225
            Left            =   1920
            TabIndex        =   95
            Top             =   180
            Width           =   375
         End
      End
      Begin VB.Frame Frame9 
         Height          =   825
         Left            =   2565
         TabIndex        =   86
         Top             =   840
         Width           =   1980
         Begin VB.TextBox f_Temp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   80
            TabIndex        =   9
            Top             =   475
            Width           =   900
         End
         Begin VB.TextBox f_TempTime 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1000
            TabIndex        =   10
            Top             =   475
            Width           =   900
         End
         Begin VB.Label Label9 
            Caption         =   "Time"
            Height          =   225
            Left            =   1155
            TabIndex        =   97
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "Temp."
            Height          =   225
            Left            =   360
            TabIndex        =   96
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.Frame Frame10 
         Height          =   825
         Left            =   4590
         TabIndex        =   85
         Top             =   840
         Width           =   3800
         Begin VB.ComboBox f_Chemical2 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   150
            Width           =   975
         End
         Begin VB.TextBox f_Chemical2_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1298
            TabIndex        =   14
            Top             =   475
            Width           =   1200
         End
         Begin VB.ComboBox f_Acid 
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   150
            Width           =   975
         End
         Begin VB.ComboBox f_Chemical 
            Height          =   315
            Left            =   305
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   150
            Width           =   975
         End
         Begin VB.TextBox f_Acid_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   16
            Top             =   475
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   80
            TabIndex        =   12
            Top             =   475
            Width           =   1200
         End
         Begin VB.Label Label21 
            Caption         =   "Ch"
            Height          =   225
            Left            =   1285
            TabIndex        =   114
            Top             =   180
            Width           =   255
         End
         Begin VB.Label Label11 
            Caption         =   "Ac"
            Height          =   225
            Left            =   2540
            TabIndex        =   99
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "Ch"
            Height          =   225
            Index           =   0
            Left            =   80
            TabIndex        =   98
            Top             =   180
            Width           =   255
         End
      End
      Begin MSComCtl2.DTPicker f_ProcessDate 
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   44695553
         CurrentDate     =   38235
      End
      Begin MSComCtl2.DTPicker f_ProcessTime 
         Height          =   315
         Left            =   4800
         TabIndex        =   3
         Top             =   450
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   44695554
         CurrentDate     =   38235
      End
      Begin VB.Label Label23 
         Caption         =   "Ch"
         Height          =   225
         Left            =   4440
         TabIndex        =   123
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Ch"
         Height          =   225
         Index           =   1
         Left            =   3360
         TabIndex        =   120
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "New Color"
         Height          =   255
         Left            =   2565
         TabIndex        =   118
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Remarks"
         Height          =   225
         Left            =   6480
         TabIndex        =   104
         Top             =   3645
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Cones"
         Height          =   225
         Left            =   7560
         TabIndex        =   94
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Machine No."
         Height          =   225
         Left            =   6240
         TabIndex        =   93
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Time"
         Height          =   225
         Left            =   5040
         TabIndex        =   92
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   225
         Left            =   3600
         TabIndex        =   91
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Party"
         Height          =   225
         Left            =   600
         TabIndex        =   90
         Top             =   180
         Width           =   495
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Ch"
      Height          =   225
      Index           =   3
      Left            =   600
      TabIndex        =   122
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "Ch"
      Height          =   225
      Index           =   2
      Left            =   4560
      TabIndex        =   121
      Top             =   3960
      Width           =   255
   End
End
Attribute VB_Name = "old_Process"
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
Private Sub ClChk_Click()
    If ClChk.value = Checked Then
        Me.SrColor.Enabled = True
    Else
        Me.SrColor.Enabled = False
    End If
    Call SrfillList
End Sub
Private Sub CmdAllSearch_Click()
        old_Process.Left = 0
        old_Process.Width = 12000
        Call SrfillList
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
            cnDatabase.Execute "update Process set Is_Active = 0 where ProcessCode =" & m_ListID
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
    Me.f_PartyCode.SetFocus
End Sub
Private Sub CmdRecipe_Click()
        Recipe.Show
        Recipe.Width = 8250
        Recipe.Height = 6400
        Recipe.Left = 2000
        Recipe.Top = 500
End Sub
Private Sub Cmdhide_Click()
        old_Process.Width = 8900
        old_Process.Left = 1700
        Me.SrItem.ListIndex = -1
        Me.SrItemType.ListIndex = -1
        Me.srParty.ListIndex = -1
        Call fillList
End Sub
Private Sub cmdReport_Click()
If Len(Trim(m_ListID)) > 0 Then
     crptDaily.ReportFileName = App.Path & "\Reports\Rpt_Process.rpt"
     crptDaily.Connect = conStr
     selcformula = "{vwRpt_Process_1.ProcessCode}=" & m_ListID
     
     crptDaily.Formulas(0) = "ProCode ='" & m_ListID & "'"
     crptDaily.SelectionFormula = selcformula
     crptDaily.WindowState = crptMaximized
     crptDaily.Action = 1
 End If
End Sub
Private Sub cmdSave_Click()
If Len(Trim(Me.f_PartyCode)) > 0 And Len(Trim(f_MachineNo)) > 0 And Len(Trim(Me.f_ItemTypeCode)) > 0 And Len(Trim(f_Cone)) > 0 And Len(Trim(f_ConeKG)) > 0 And Len(Trim(f_Temp)) > 0 And Len(Trim(f_TempTime)) > 0 And Len(Trim(f_RecipeCode)) > 0 Then
            Call setVal
            MsgBox ("Record saved successfully"), vbInformation
            Me.f_PartyCode.SetFocus
            Call AddNewRecord
            Call fillList
Else
    MsgBox "Provide data in all Fields"
End If
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
Private Sub f_Acid_KeyPress(KeyAscii As Integer)
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
            Me.f_Soap.SetFocus
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
    If Len(Trim(Me.f_Acid)) > 0 And MsgBit = 0 Then
        Call chkQty_Acid_Qty(Me.f_Acid.ItemData(Me.f_Acid.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Acid2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Acid2_Qty.SetFocus
    End If
End Sub
Private Sub f_Acid2_Qty_GotFocus()
    If Len(Trim(Me.f_Acid2_Qty)) > 0 Then
        PreQty = Me.f_Acid2_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Acid2_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Acid2_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_NewColor.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Acid2_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Acid2_Qty.Text)) > 0 Then
        vQty = Me.f_Acid2_Qty.Text
    Else
        vQty = 0
    End If
    If Len(Trim(Me.f_Acid2)) > 0 And MsgBit = 0 Then
        Call chkQty_Acid2_Qty(Me.f_Acid2.ItemData(Me.f_Acid2.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Castic_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Castic_Qty.SetFocus
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
            Me.f_Hydro.SetFocus
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
    If Len(Trim(Me.f_Castic)) > 0 And MsgBit = 0 Then
        Call chkQty_Castic_Qty(Me.f_Castic.ItemData(Me.f_Castic.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_CasticTime_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Re_RecipeCode.SetFocus
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
                AvbQty = CDbl(rstGetQty("Quantity"))
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
            Me.f_Remarks.SetFocus
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
                AvbQty = CDbl(rstGetQty("Quantity"))
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
                    Me.f_Remarks.SetFocus
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
Private Sub f_Chemical_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Chemical_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Chemical2.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Chemical_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Chemical_Qty.Text)) > 0 Then
        vQty = Me.f_Chemical_Qty.Text
    Else
        vQty = 0
    End If
    If Len(Trim(Me.f_Chemical)) > 0 And MsgBit = 0 Then
        Call chkQty_Chemical_Qty(Me.f_Chemical.ItemData(Me.f_Chemical.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Chemical_Qty_GotFocus()
If Len(Trim(Me.f_Chemical_Qty)) > 0 Then
    PreQty = Me.f_Chemical_Qty.Text
Else
    PreQty = 0
End If
End Sub
Private Sub f_Chemical2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical2_Qty.SetFocus
    End If
End Sub
Private Sub f_Chemical2_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical2_Qty)) > 0 Then
        PreQty = Me.f_Chemical2_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical2_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Chemical2_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Acid.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Chemical2_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Chemical2_Qty.Text)) > 0 Then
        vQty = Me.f_Chemical2_Qty.Text
    Else
        vQty = 0
    End If
    If Len(Trim(Me.f_Chemical2)) > 0 And MsgBit = 0 Then
        Call chkQty_Chemical2_Qty(Me.f_Chemical2.ItemData(Me.f_Chemical2.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Color_1_Click()
    If Me.f_Color_1.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_Color_1.ItemData(Me.f_Color_1.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
            If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
            End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_Color_1_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_Color_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_2.SetFocus
    End If
End Sub
Private Sub f_Color_2_Click()
    If Me.f_Color_2.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_Color_2.ItemData(Me.f_Color_2.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
        If Not (rstGetQty.EOF) Then
             Qty = rstGetQty("Quantity")
         End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_Color_2_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_Color_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_3.SetFocus
    End If
End Sub
Private Sub f_Color_3_Click()
    If Me.f_Color_3.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_Color_3.ItemData(Me.f_Color_3.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
        If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_Color_3_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_Color_3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_4.SetFocus
    End If
End Sub
Private Sub f_Color_4_Click()
    If Me.f_Color_4.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_Color_4.ItemData(Me.f_Color_4.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
        If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_Color_4_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_Color_4_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_5.SetFocus
    End If
End Sub
Private Sub f_Color_5_Click()
    If Me.f_Color_5.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_Color_5.ItemData(Me.f_Color_5.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
        If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_Color_5_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_Color_5_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap2.SetFocus
    End If
End Sub
Private Sub f_Cone_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_ConeKG.SetFocus
    End If
End Sub
Private Sub f_ConeKG_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Temp.SetFocus
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Den_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_ItemTypeCode.SetFocus
    End If
    
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Hydro_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Hydro_Qty.SetFocus
    End If
End Sub
Private Sub f_Hydro_Qty_GotFocus()
    If Len(Trim(Me.f_Hydro_Qty)) > 0 Then
        PreQty = Me.f_Hydro_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Hydro_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Hydro_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_CasticTime.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Hydro_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Hydro_Qty.Text)) > 0 Then
        vQty = Me.f_Hydro_Qty.Text
    Else
        vQty = 0
    End If
    If Len(Trim(Me.f_Hydro)) > 0 And MsgBit = 0 Then
        Call chkQty_hydro_Qty(Me.f_Hydro.ItemData(Me.f_Hydro.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_ItemTypeCode_Click()
    If ClickPane = 0 And Me.f_ItemTypeCode.ListIndex > -1 Then
        i = Me.f_ItemTypeCode.ItemData(Me.f_ItemTypeCode.ListIndex)
        If i = 1 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type1 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 2 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type2 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 3 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type3 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 4 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type4 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 5 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type5 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 6 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type6 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 7 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type7 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 8 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type8 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 9 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type9 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 10 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type10 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 11 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type11 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 12 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type12 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 13 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type13 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 14 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type14 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 15 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type15 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ElseIf i = 16 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type16 where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        Else
            FillCombo "Select ItemCode, ItemName from vwAvailableQty where Qty > 0 and ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        End If
    ElseIf ClickPane = 1 And Me.f_ItemTypeCode.ListIndex > -1 Then
        i = Me.f_ItemTypeCode.ItemData(Me.f_ItemTypeCode.ListIndex)
        FillCombo "Select ItemCode, ItemName from Item where ItemTypeCode = " & i, f_Cone, "ItemName", "ItemCode"
        ClickPane = 0
    Else
        Me.f_Cone.Clear
    End If
End Sub
Private Sub f_ItemTypeCode_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Cone.SetFocus
    End If
End Sub
Private Sub f_MachineNo_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Den.SetFocus
    End If
End Sub
Private Sub f_NewColor_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical_3_Code.SetFocus
    End If
End Sub
Private Sub f_R_Color_1_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_R_Color_2.SetFocus
    End If
End Sub
Private Sub f_R_Color_2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_R_Color_3.SetFocus
    End If
End Sub
Private Sub f_R_Color_3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_R_Color_4.SetFocus
    End If
End Sub
Private Sub f_R_Color_4_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_R_Color_5.SetFocus
    End If
End Sub
Private Sub f_R_Color_5_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_R_Color_6.SetFocus
    End If
End Sub
Private Sub f_R_Color_6_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 And cmdSave.Enabled = True Then
        Me.cmdSave.SetFocus
    End If
End Sub
Private Sub f_Re_RecipeCode_Click()
If Len(Trim(Me.f_RecipeCode)) > 0 Then
    Call f_RecipeCode_LostFocus
End If
If Me.f_Re_RecipeCode.value = Checked Then
    Me.f_Re_RecipeCode.value = 1
    Me.f_Color_1_Qty.Text = 0
    Me.f_Color_2_Qty.Text = 0
    Me.f_Color_3_Qty.Text = 0
    Me.f_Color_4_Qty.Text = 0
    Me.f_Color_5_Qty.Text = 0
    'Me.f_Color_6_Qty.Text = 0
Else
    Me.f_Re_RecipeCode = 0
    Me.f_R_Color_1_Qty.Text = 0
    Me.f_R_Color_2_Qty.Text = 0
    Me.f_R_Color_3_Qty.Text = 0
    Me.f_R_Color_4_Qty.Text = 0
    Me.f_R_Color_5_Qty.Text = 0
    Me.f_R_Color_6_Qty.Text = 0
End If
End Sub
Private Sub f_Re_RecipeCode_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_RecipeCode.SetFocus
    End If
'    If (KeyAscii >= 48 And KeyAscii <= 57) Then
'        KeyAscii = KeyAscii
'    Else
'        KeyAscii = 0
'    End If
End Sub
Private Sub f_RecipeCode_Change()
If (Len(Me.f_RecipeCode.Text) > 0) Then
        Me.Frame3.Enabled = True
    Else
        Me.Frame3.Enabled = False
    End If
End Sub
Private Sub f_RecipeCode_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_1.SetFocus
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_RecipeCode_LostFocus()
 Dim sql As String
If Me.f_Re_RecipeCode.value = Checked And Len(Trim(Me.f_RecipeCode)) > 0 Then
    sql = "Select ItemCode, ItemName from Item where ItemCode in (Select ItemCode from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & ")"
    FillColorCombo sql, f_R_Color_1, "ItemName", "ItemCode"
    FillColorCombo sql, f_R_Color_2, "ItemName", "ItemCode"
    FillColorCombo sql, f_R_Color_3, "ItemName", "ItemCode"
    FillColorCombo sql, f_R_Color_4, "ItemName", "ItemCode"
    FillColorCombo sql, f_R_Color_5, "ItemName", "ItemCode"
    FillColorCombo sql, f_R_Color_6, "ItemName", "ItemCode"
    cmdSave.Enabled = True
 ElseIf Me.f_Re_RecipeCode.value = Unchecked And Len(Trim(Me.f_RecipeCode)) > 0 Then
    sql = "Select ItemCode, ItemName from Item where ItemCode in (Select ItemCode from RecipeDetail where RecipeMasterCode = " & Me.f_RecipeCode.Text & ")"
    FillColorCombo sql, f_Color_1, "ItemName", "ItemCode"
    FillColorCombo sql, f_Color_2, "ItemName", "ItemCode"
    FillColorCombo sql, f_Color_3, "ItemName", "ItemCode"
    FillColorCombo sql, f_Color_4, "ItemName", "ItemCode"
    FillColorCombo sql, f_Color_5, "ItemName", "ItemCode"
    'FillColorCombo sql, f_Color_6, "ItemName", "ItemCode"
    cmdSave.Enabled = True
End If
End Sub
Private Sub f_Remarks_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_R_Color_1.SetFocus
    End If
End Sub
Private Sub f_Soap_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap_Qty.SetFocus
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
            Me.f_SoapTime.SetFocus
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
    If Len(Trim(Me.f_Soap)) > 0 And MsgBit = 0 Then
        Call chkQty_Soap_Qty(Me.f_Soap.ItemData(Me.f_Soap.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_Soap2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Soap2_Qty.SetFocus
    End If
End Sub
Private Sub f_Soap2_Qty_GotFocus()
    If Len(Trim(Me.f_Soap2_Qty)) > 0 Then
        PreQty = Me.f_Soap2_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Soap2_Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        MsgBit = 0
        Call f_Soap2_Qty_LostFocus
        MsgBit = 1
        If QtyBit = 1 Then
            Me.f_Acid2.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub f_Soap2_Qty_LostFocus()
    Dim vQty As Double
    If Len(Trim(Me.f_Soap2_Qty.Text)) > 0 Then
        vQty = Me.f_Soap2_Qty.Text
    Else
        vQty = 0
    End If
    If Len(Trim(Me.f_Soap2)) > 0 And MsgBit = 0 Then
        Call chkQty_Soap2_Qty(Me.f_Soap2.ItemData(Me.f_Soap2.ListIndex), vQty)
        MsgBit = 0
    End If
End Sub
Private Sub f_SoapTime_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Castic.SetFocus
    End If
End Sub
Private Sub f_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_TempTime.SetFocus
    End If
End Sub
Private Sub f_TempTime_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Chemical.SetFocus
    End If
End Sub
Private Sub Form_Load()
  m_AddMode = True
  cmdSave.Enabled = False
  DBConn
  f_ProcessDate = Now
  f_ProcessTime = Now
  SrDate = Now
  SrDate2 = Now
  
  FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", f_ItemTypeCode, "ItemTypeName", "ItemTypeCode"
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical2, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_3_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_4_Code, "ItemName", "ItemCode"
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 7 order by 2", f_Acid, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 7 order by 2", f_Acid2, "ItemName", "ItemCode"
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 10 order by 2", f_Soap, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 10 order by 2", f_Soap2, "ItemName", "ItemCode"
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 8 order by 2", f_Hydro, "ItemName", "ItemCode"
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 9 order by 2", f_Castic, "ItemName", "ItemCode"
  FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", f_PartyCode, "PartyName", "PartyCode"
    
  FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", SrItemType, "ItemTypeName", "ItemTypeCode"
  FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", srParty, "PartyName", "PartyCode"
  
  lvwphase.ColumnHeaders.Add Text:="Code", Width:=600
  lvwphase.ColumnHeaders.Add Text:="Date", Width:=1200
  lvwphase.ColumnHeaders.Add Text:="Party Name", Width:=1700
  lvwphase.ColumnHeaders.Add Text:="Machine #", Width:=1000
  lvwphase.ColumnHeaders.Add Text:="Item Type", Width:=1200
  lvwphase.ColumnHeaders.Add Text:="Item", Width:=700
  lvwphase.ColumnHeaders.Add Text:="Recipe", Width:=800
  lvwphase.ColumnHeaders.Add Text:="Re Recipe", Width:=1000
  
  Call fillList
End Sub
Public Sub enabledDisabled(flag As Boolean)
    Me.f_Color_1_Qty.Enabled = flag
    Me.f_Color_2_Qty.Enabled = flag
    Me.f_Color_3_Qty.Enabled = flag
    Me.f_Color_4_Qty.Enabled = flag
    Me.f_Color_5_Qty.Enabled = flag
    'Me.f_Color_6_Qty.Enabled = flag
    
    Me.f_R_Color_1_Qty.Enabled = flag
    Me.f_R_Color_2_Qty.Enabled = flag
    Me.f_R_Color_3_Qty.Enabled = flag
    Me.f_R_Color_4_Qty.Enabled = flag
    Me.f_R_Color_5_Qty.Enabled = flag
    Me.f_R_Color_6_Qty.Enabled = flag
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
    rstSave("PartyCode") = Me.f_PartyCode.ItemData(Me.f_PartyCode.ListIndex)
    rstSave("ProcessDate") = Me.f_ProcessDate.value
    rstSave("ProcessTime") = Me.f_ProcessTime.value
    rstSave("MachineNo") = Me.f_MachineNo.Text
    If Len(Trim(Me.f_Den.Text)) > 0 Then
        rstSave("Den") = Me.f_Den.Text
    Else
        rstSave("Den") = 0
    End If
    If Len(Trim(Me.f_ItemTypeCode)) > 0 Then
        rstSave("ItemTypeCode") = Me.f_ItemTypeCode.ItemData(Me.f_ItemTypeCode.ListIndex)
        rstSave("Cone") = Me.f_Cone.ItemData(Me.f_Cone.ListIndex)
        rstSave("ConeKG") = Me.f_ConeKG.Text
    End If
    rstSave("Temp") = Me.f_Temp.Text
    rstSave("TempTime") = Me.f_TempTime.Text
    If Len(Trim(Me.f_Chemical)) > 0 Then
        rstSave("Chemical") = Me.f_Chemical.ItemData(Me.f_Chemical.ListIndex)
        rstSave("Chemical_Qty") = Me.f_Chemical_Qty.Text
'        rstSave("Chemical_Cost") = (getAvgCost(Me.f_Chemical.ItemData(Me.f_Chemical.ListIndex)) * (Me.f_Chemical_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Chemical2)) > 0 And Me.f_Chemical2 <> "-- Select --" Then
        rstSave("Chemical2") = Me.f_Chemical2.ItemData(Me.f_Chemical2.ListIndex)
        rstSave("Chemical2_Qty") = Me.f_Chemical2_Qty.Text
'       rstSave("Chemical2_Cost") = (getAvgCost(Me.f_Chemical2.ItemData(Me.f_Chemical2.ListIndex)) * (Me.f_Chemical2_Qty.Text / 1000))
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
    
    If Len(Trim(Me.f_Acid)) > 0 Then
        rstSave("Acid") = Me.f_Acid.ItemData(Me.f_Acid.ListIndex)
        rstSave("Acid_Qty") = Me.f_Acid_Qty.Text
'        rstSave("Acid_Cost") = (getAvgCost(Me.f_Acid.ItemData(Me.f_Acid.ListIndex)) * (Me.f_Acid_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Acid2)) > 0 And Me.f_Acid2 <> "-- Select --" Then
        rstSave("Acid2") = Me.f_Acid2.ItemData(Me.f_Acid2.ListIndex)
        rstSave("Acid2_Qty") = Me.f_Acid2_Qty.Text
'        rstSave("Acid2_Cost") = (getAvgCost(Me.f_Acid2.ItemData(Me.f_Acid2.ListIndex)) * (Me.f_Acid2_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Soap)) > 0 Then
        rstSave("Soap") = Me.f_Soap.ItemData(Me.f_Soap.ListIndex)
        rstSave("Soap_Qty") = Me.f_Soap_Qty.Text
        rstSave("SoapTime") = Me.f_SoapTime.Text
'        rstSave("Soap_Cost") = (getAvgCost(Me.f_Soap.ItemData(Me.f_Soap.ListIndex)) * (Me.f_Soap_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Soap2)) > 0 And Me.f_Soap2 <> "-- Select --" Then
        rstSave("Soap2") = Me.f_Soap2.ItemData(Me.f_Soap2.ListIndex)
        rstSave("Soap2_Qty") = Me.f_Soap2_Qty.Text
'        rstSave("Soap2_Cost") = (getAvgCost(Me.f_Soap2.ItemData(Me.f_Soap2.ListIndex)) * (Me.f_Soap2_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Hydro)) > 0 Then
        rstSave("Hydro") = Me.f_Hydro.ItemData(Me.f_Hydro.ListIndex)
        rstSave("Hydro_Qty") = Me.f_Hydro_Qty.Text
'        rstSave("Hydro_Cost") = (getAvgCost(Me.f_Hydro.ItemData(Me.f_Hydro.ListIndex)) * (Me.f_Hydro_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Castic)) > 0 Then
        rstSave("Castic") = Me.f_Castic.ItemData(Me.f_Castic.ListIndex)
        rstSave("Castic_Qty") = Me.f_Castic_Qty.Text
        rstSave("CasticTime") = Me.f_CasticTime.Text
'        rstSave("Castic_Cost") = (getAvgCost(Me.f_Castic.ItemData(Me.f_Castic.ListIndex)) * (Me.f_Castic_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_RecipeCode)) > 0 Then
        rstSave("RecipeCode") = Me.f_RecipeCode.Text
    End If
    If Me.f_Re_RecipeCode.value = Checked Then
    'Len(Trim(Me.f_Re_RecipeCode)) > 0 Then
        rstSave("Re_RecipeCode") = 1
    'Me.f_Re_RecipeCode.Text
    Else
        rstSave("Re_RecipeCode") = 0
    End If
    If Len(Trim(Me.f_Remarks)) > 0 Then
        rstSave("Remarks") = Me.f_Remarks.Text
    End If
    
'    enabledDisabled (True)
    
    If Len(Trim(Me.f_Color_1)) > 0 And Me.f_Color_1 <> "-- Select --" Then
        rstSave("Color_1") = Me.f_Color_1.ItemData(Me.f_Color_1.ListIndex)
        rstSave("Color_1_Qty") = Me.f_Color_1_Qty.Text
'        rstSave("Color_1_Cost") = (getAvgCost(Me.f_Color_1.ItemData(Me.f_Color_1.ListIndex)) * (Me.f_Color_1_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Color_2)) > 0 And Me.f_Color_2 <> "-- Select --" Then
        rstSave("Color_2") = Me.f_Color_2.ItemData(Me.f_Color_2.ListIndex)
        rstSave("Color_2_Qty") = Me.f_Color_2_Qty.Text
'        rstSave("Color_2_Cost") = (getAvgCost(Me.f_Color_2.ItemData(Me.f_Color_2.ListIndex)) * (Me.f_Color_2_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Color_3)) > 0 And Me.f_Color_3 <> "-- Select --" Then
        rstSave("Color_3") = Me.f_Color_3.ItemData(Me.f_Color_3.ListIndex)
        rstSave("Color_3_Qty") = Me.f_Color_3_Qty.Text
'        rstSave("Color_3_Cost") = (getAvgCost(Me.f_Color_3.ItemData(Me.f_Color_3.ListIndex)) * (Me.f_Color_3_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Color_4)) > 0 And Me.f_Color_4 <> "-- Select --" Then
        rstSave("Color_4") = Me.f_Color_4.ItemData(Me.f_Color_4.ListIndex)
        rstSave("Color_4_Qty") = Me.f_Color_4_Qty.Text
'        rstSave("Color_4_Cost") = (getAvgCost(Me.f_Color_4.ItemData(Me.f_Color_4.ListIndex)) * (Me.f_Color_4_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_Color_5)) > 0 And Me.f_Color_5 <> "-- Select --" Then
        rstSave("Color_5") = Me.f_Color_5.ItemData(Me.f_Color_5.ListIndex)
        rstSave("Color_5_Qty") = Me.f_Color_5_Qty.Text
'        rstSave("Color_5_Cost") = (getAvgCost(Me.f_Color_5.ItemData(Me.f_Color_5.ListIndex)) * (Me.f_Color_5_Qty.Text / 1000))
    End If
'    If Len(Trim(Me.f_Color_6)) > 0 And Me.f_Color_6 <> "-- Select --" Then
'        rstSave("Color_6") = Me.f_Color_6.ItemData(Me.f_Color_6.ListIndex)
'        rstSave("Color_6_Qty") = Me.f_Color_6_Qty.Text
'        rstSave("Color_6_Cost") = (getAvgCost(Me.f_Color_6.ItemData(Me.f_Color_6.ListIndex)) * (Me.f_Color_6_Qty.Text / 1000))
'    End If
    If Len(Trim(Me.f_R_Color_1)) > 0 And Me.f_R_Color_1 <> "-- Select --" Then
        rstSave("R_Color_1") = Me.f_R_Color_1.ItemData(Me.f_R_Color_1.ListIndex)
        rstSave("R_Color_1_Qty") = Me.f_R_Color_1_Qty.Text
'        rstSave("R_Color_1_Cost") = (getAvgCost(Me.f_R_Color_1.ItemData(Me.f_R_Color_1.ListIndex)) * (Me.f_R_Color_1_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_R_Color_2)) > 0 And Me.f_R_Color_2 <> "-- Select --" Then
        rstSave("R_Color_2") = Me.f_R_Color_2.ItemData(Me.f_R_Color_2.ListIndex)
        rstSave("R_Color_2_Qty") = Me.f_R_Color_2_Qty.Text
'        rstSave("R_Color_2_Cost") = (getAvgCost(Me.f_R_Color_2.ItemData(Me.f_R_Color_2.ListIndex)) * (Me.f_R_Color_2_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_R_Color_3)) > 0 And Me.f_R_Color_3 <> "-- Select --" Then
        rstSave("R_Color_3") = Me.f_R_Color_3.ItemData(Me.f_R_Color_3.ListIndex)
        rstSave("R_Color_3_Qty") = Me.f_R_Color_3_Qty.Text
'        rstSave("R_Color_3_Cost") = (getAvgCost(Me.f_R_Color_3.ItemData(Me.f_R_Color_3.ListIndex)) * (Me.f_R_Color_3_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_R_Color_4)) > 0 And Me.f_R_Color_4 <> "-- Select --" Then
        rstSave("R_Color_4") = Me.f_R_Color_4.ItemData(Me.f_R_Color_4.ListIndex)
        rstSave("R_Color_4_Qty") = Me.f_R_Color_4_Qty.Text
'        rstSave("R_Color_4_Cost") = (getAvgCost(Me.f_R_Color_4.ItemData(Me.f_R_Color_4.ListIndex)) * (Me.f_R_Color_4_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_R_Color_5)) > 0 And Me.f_R_Color_5 <> "-- Select --" Then
        rstSave("R_Color_5") = Me.f_R_Color_5.ItemData(Me.f_R_Color_5.ListIndex)
        rstSave("R_Color_5_Qty") = Me.f_R_Color_5_Qty.Text
'        rstSave("R_Color_5_Cost") = (getAvgCost(Me.f_R_Color_5.ItemData(Me.f_R_Color_5.ListIndex)) * (Me.f_R_Color_5_Qty.Text / 1000))
    End If
    If Len(Trim(Me.f_R_Color_6)) > 0 And Me.f_R_Color_6 <> "-- Select --" Then
        rstSave("R_Color_6") = Me.f_R_Color_6.ItemData(Me.f_R_Color_6.ListIndex)
        rstSave("R_Color_6_Qty") = Me.f_R_Color_6_Qty.Text
'        rstSave("R_Color_6_Cost") = (getAvgCost(Me.f_R_Color_6.ItemData(Me.f_R_Color_6.ListIndex)) * (Me.f_R_Color_6_Qty.Text / 1000))
    End If
    rstSave("NewColor") = Me.f_NewColor.Text
    rstSave("Is_Cotton_Dyeing") = 0
rstSave.Update
rstSave.Close
Set rstSave = Nothing
'enabledDisabled (False)
m_AddMode = False
Call fillList
End Sub
Private Sub EnableSave()
    If Len(Trim(Me.f_PartyCode)) > 0 And Len(Trim(f_MachineNo)) > 0 And Len(Trim(Me.f_ItemTypeCode)) > 0 And Len(Trim(f_Cone)) > 0 And Len(Trim(f_ConeKG)) > 0 And Len(Trim(f_Temp)) > 0 And Len(Trim(f_TempTime)) > 0 And Len(Trim(f_RecipeCode)) > 0 Then
        Me.cmdSave.Enabled = True
    Else
        Me.cmdSave.Enabled = False
    End If
End Sub
Public Sub AddNewRecord()
    m_ListID = ""
    Me.f_ProcessDate.value = Now
    Me.f_ProcessTime.value = Now
    Me.f_PartyCode.ListIndex = -1
    Me.f_MachineNo.Text = ""
    Me.f_Den.Text = ""
    Me.f_ItemTypeCode.ListIndex = -1
    Me.f_Cone.ListIndex = -1
    Me.f_ConeKG.Text = ""
    Me.f_Temp.Text = "134"
    Me.f_TempTime.Text = ""
    Me.f_Chemical.ListIndex = -1
    Me.f_Chemical_Qty.Text = ""
    Me.f_Chemical2.ListIndex = -1
    Me.f_Chemical2_Qty.Text = ""
    Me.f_Chemical_3_Code.ListIndex = -1
    Me.f_Chemical_3_Qty.Text = ""
    Me.f_Chemical_4_Code.ListIndex = -1
    Me.f_Chemical_4_Qty.Text = ""
    Me.f_Acid.ListIndex = -1
    Me.f_Acid_Qty.Text = ""
    Me.f_Soap.ListIndex = -1
    Me.f_Soap_Qty.Text = ""
    Me.f_SoapTime.Text = ""
    Me.f_Castic.ListIndex = -1
    Me.f_Castic_Qty.Text = ""
    Me.f_Hydro.ListIndex = -1
    Me.f_Hydro_Qty.Text = ""
    Me.f_CasticTime.Text = ""
    Me.f_RecipeCode.Text = ""
    Me.f_Re_RecipeCode.value = 0
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
'    Me.f_Color_6.ListIndex = -1
'    Me.f_Color_6_Qty.Text = ""
    Me.f_Soap2.ListIndex = -1
    Me.f_Soap2_Qty.Text = ""
    Me.f_Acid2.ListIndex = -1
    Me.f_Acid2_Qty.Text = ""
    Me.f_NewColor.Text = ""
    Me.f_Remarks.Text = ""
    Me.f_R_Color_1.ListIndex = -1
    Me.f_R_Color_1_Qty.Text = ""
    Me.f_R_Color_2.ListIndex = -1
    Me.f_R_Color_2_Qty.Text = ""
    Me.f_R_Color_3.ListIndex = -1
    Me.f_R_Color_3_Qty.Text = ""
    Me.f_R_Color_4.ListIndex = -1
    Me.f_R_Color_4_Qty.Text = ""
    Me.f_R_Color_5.ListIndex = -1
    Me.f_R_Color_5_Qty.Text = ""
    Me.f_R_Color_6.ListIndex = -1
    Me.f_R_Color_6_Qty.Text = ""
End Sub
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Set rstList = FillRecordSet("SELECT top 60 ProcessCode, ProcessDate, PartyName, MachineNo, ItemTypeName, (Select ItemName from Item where ItemCode = Cone) as Cone, isNull(RecipeCode, 0) as RecipeCode, isNull(Re_RecipeCode, 0) as Re_RecipeCode " & _
                                "FROM Party INNER JOIN (ItemType INNER JOIN Process ON ItemType.ItemTypeCode = Process.ItemTypeCode) ON Party.PartyCode = Process.PartyCode where Is_Active = 1 and Is_Cotton_Dyeing = 0 order by ProcessCode desc")
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!ProcessCode, _
                   Key:=CStr("Id=" & rstList!ProcessCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!ProcessDate
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!MachineNo
                 .Add Text:=rstList!ItemTypeName
                 .Add Text:=rstList!Cone
                 .Add Text:=rstList!RecipeCode
                 .Add Text:=rstList!Re_RecipeCode
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub getVal()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("Select * From Process Where Is_Cotton_Dyeing = 0 and ProcessCode = " & m_ListID)
    If Not (rstGetVal.EOF) Then
        Call selectValueInCombo(Me.f_PartyCode, rstGetVal("PartyCode"))
        Me.f_ProcessDate.value = IIf(IsNull(rstGetVal("ProcessDate")), Now, rstGetVal("ProcessDate"))
        Me.f_ProcessTime.value = IIf(IsNull(rstGetVal("ProcessTime")), Now, rstGetVal("ProcessTime"))
        Me.f_MachineNo.Text = rstGetVal("MachineNo")
        Me.f_Den.Text = rstGetVal("Den")
        Call selectValueInCombo(Me.f_ItemTypeCode, rstGetVal("ItemTypeCode"))
        Call selectValueInCombo(Me.f_Cone, rstGetVal("Cone"))
        Me.f_ConeKG.Text = rstGetVal("ConeKG")
        Me.f_Temp.Text = rstGetVal("Temp")
        Me.f_TempTime.Text = rstGetVal("TempTime")
        Call selectValueInCombo(Me.f_Chemical, rstGetVal("Chemical"))
        Me.f_Chemical_Qty.Text = rstGetVal("Chemical_Qty")
        Call selectValueInCombo(Me.f_Chemical2, IIf(IsNull(rstGetVal("Chemical2")), -1, rstGetVal("Chemical2")))
        Me.f_Chemical2_Qty.Text = IIf(IsNull(rstGetVal("Chemical2_Qty")), 0, rstGetVal("Chemical2_Qty"))
        Call selectValueInCombo(Me.f_Chemical_3_Code, IIf(IsNull(rstGetVal("Chemical_3_Code")), -1, rstGetVal("Chemical_3_Code")))
        Me.f_Chemical_3_Qty.Text = IIf(IsNull(rstGetVal("Chemical_3_Qty")), 0, rstGetVal("Chemical_3_Qty"))
        Call selectValueInCombo(Me.f_Chemical_4_Code, IIf(IsNull(rstGetVal("Chemical_4_Code")), -1, rstGetVal("Chemical_4_Code")))
        Me.f_Chemical_4_Qty.Text = IIf(IsNull(rstGetVal("Chemical_4_Qty")), 0, rstGetVal("Chemical_4_Qty"))
        Call selectValueInCombo(Me.f_Acid, rstGetVal("Acid"))
        Me.f_Acid_Qty.Text = rstGetVal("Acid_Qty")
        Call selectValueInCombo(Me.f_Soap, IIf(IsNull(rstGetVal("Soap")), 0, rstGetVal("Soap")))
        Me.f_Soap_Qty.Text = IIf(IsNull(rstGetVal("Soap_Qty")), 0, rstGetVal("Soap_Qty"))
        Me.f_SoapTime.Text = IIf(IsNull(rstGetVal("SoapTime")), "", rstGetVal("SoapTime"))
        Call selectValueInCombo(Me.f_Castic, IIf(IsNull(rstGetVal("Castic")), 0, rstGetVal("Castic")))
        Me.f_Castic_Qty.Text = IIf(IsNull(rstGetVal("Castic_Qty")), 0, rstGetVal("Castic_Qty"))
        Call selectValueInCombo(Me.f_Hydro, IIf(IsNull(rstGetVal("Hydro")), 0, rstGetVal("Hydro")))
        Me.f_Hydro_Qty.Text = IIf(IsNull(rstGetVal("Hydro_Qty")), 0, rstGetVal("Hydro_Qty"))
        Me.f_CasticTime.Text = IIf(IsNull(rstGetVal("CasticTime")), "", rstGetVal("CasticTime"))
        Me.f_RecipeCode.Text = IIf(IsNull(rstGetVal("RecipeCode")), 0, rstGetVal("RecipeCode"))
        If rstGetVal("Re_RecipeCode") = 1 Then
            Me.f_Re_RecipeCode.value = Checked
        Else
            Me.f_Re_RecipeCode.value = Unchecked
            'IIf(IsNull(rstGetVal("Re_RecipeCode")), 0, rstGetVal("Re_RecipeCode"))
        End If
        Dim sql As String
        If Me.f_Re_RecipeCode.value = Checked And Len(Trim(Me.f_RecipeCode)) > 0 Then
            sql = "Select ItemCode, ItemName from Item where ItemCode in (Select ItemCode from RecipeDetail where RecipeMasterCode = " & Me.f_RecipeCode.Text & ")"
            FillColorCombo sql, f_R_Color_1, "ItemName", "ItemCode"
            FillColorCombo sql, f_R_Color_2, "ItemName", "ItemCode"
            FillColorCombo sql, f_R_Color_3, "ItemName", "ItemCode"
            FillColorCombo sql, f_R_Color_4, "ItemName", "ItemCode"
            FillColorCombo sql, f_R_Color_5, "ItemName", "ItemCode"
            FillColorCombo sql, f_R_Color_6, "ItemName", "ItemCode"
        ElseIf Me.f_Re_RecipeCode.value = Unchecked And Len(Trim(Me.f_RecipeCode)) > 0 Then
            sql = "Select ItemCode, ItemName from Item where ItemCode in (Select ItemCode from RecipeDetail where RecipeMasterCode = " & Me.f_RecipeCode.Text & ")"
            FillColorCombo sql, f_Color_1, "ItemName", "ItemCode"
            FillColorCombo sql, f_Color_2, "ItemName", "ItemCode"
            FillColorCombo sql, f_Color_3, "ItemName", "ItemCode"
            FillColorCombo sql, f_Color_4, "ItemName", "ItemCode"
            FillColorCombo sql, f_Color_5, "ItemName", "ItemCode"
'            FillColorCombo sql, f_Color_6, "ItemName", "ItemCode"
        End If
       
        Call selectValueInCombo(Me.f_Color_1, IIf(IsNull(rstGetVal("Color_1")), -1, rstGetVal("Color_1")))
        Me.f_Color_1_Qty.Text = IIf(IsNull(rstGetVal("Color_1_Qty")), 0, rstGetVal("Color_1_Qty"))
        Call selectValueInCombo(Me.f_Color_2, IIf(IsNull(rstGetVal("Color_2")), -1, rstGetVal("Color_2")))
        Me.f_Color_2_Qty.Text = IIf(IsNull(rstGetVal("Color_2_Qty")), 0, rstGetVal("Color_2_Qty"))
        Call selectValueInCombo(Me.f_Color_3, IIf(IsNull(rstGetVal("Color_3")), -1, rstGetVal("Color_3")))
        Me.f_Color_3_Qty.Text = IIf(IsNull(rstGetVal("Color_3_Qty")), 0, rstGetVal("Color_3_Qty"))
        Call selectValueInCombo(Me.f_Color_4, IIf(IsNull(rstGetVal("Color_4")), -1, rstGetVal("Color_4")))
        Me.f_Color_4_Qty.Text = IIf(IsNull(rstGetVal("Color_4_Qty")), 0, rstGetVal("Color_4_Qty"))
        Call selectValueInCombo(Me.f_Color_5, IIf(IsNull(rstGetVal("Color_5")), -1, rstGetVal("Color_5")))
        Me.f_Color_5_Qty.Text = IIf(IsNull(rstGetVal("Color_5_Qty")), 0, rstGetVal("Color_5_Qty"))
'        Call selectValueInCombo(Me.f_Color_6, IIf(IsNull(rstGetVal("Color_6")), -1, rstGetVal("Color_6")))
'        Me.f_Color_6_Qty.Text = IIf(IsNull(rstGetVal("Color_6_Qty")), 0, rstGetVal("Color_6_Qty"))
        Call selectValueInCombo(Me.f_Soap2, IIf(IsNull(rstGetVal("Soap2")), -1, rstGetVal("Soap2")))
        Me.f_Soap2_Qty.Text = IIf(IsNull(rstGetVal("Soap2_Qty")), 0, rstGetVal("Soap2_Qty"))
        Call selectValueInCombo(Me.f_Acid2, IIf(IsNull(rstGetVal("Acid2")), -1, rstGetVal("Acid2")))
        Me.f_Acid2_Qty.Text = IIf(IsNull(rstGetVal("Acid2_Qty")), 0, rstGetVal("Acid2_Qty"))
        Me.f_Remarks.Text = IIf(IsNull(rstGetVal("Remarks")), " ", rstGetVal("Remarks"))
        Call selectValueInCombo(Me.f_R_Color_1, IIf(IsNull(rstGetVal("R_Color_1")), -1, rstGetVal("R_Color_1")))
        Me.f_R_Color_1_Qty.Text = IIf(IsNull(rstGetVal("R_Color_1_Qty")), 0, rstGetVal("R_Color_1_Qty"))
        Call selectValueInCombo(Me.f_R_Color_2, IIf(IsNull(rstGetVal("R_Color_2")), -1, rstGetVal("R_Color_2")))
        Me.f_R_Color_2_Qty.Text = IIf(IsNull(rstGetVal("R_Color_2_Qty")), 0, rstGetVal("R_Color_2_Qty"))
        Call selectValueInCombo(Me.f_R_Color_3, IIf(IsNull(rstGetVal("R_Color_3")), -1, rstGetVal("R_Color_3")))
        Me.f_R_Color_3_Qty.Text = IIf(IsNull(rstGetVal("R_Color_3_Qty")), 0, rstGetVal("R_Color_3_Qty"))
        Call selectValueInCombo(Me.f_R_Color_4, IIf(IsNull(rstGetVal("R_Color_4")), -1, rstGetVal("R_Color_4")))
        Me.f_R_Color_4_Qty.Text = IIf(IsNull(rstGetVal("R_Color_4_Qty")), 0, rstGetVal("R_Color_4_Qty"))
        Call selectValueInCombo(Me.f_R_Color_5, IIf(IsNull(rstGetVal("R_Color_5")), -1, rstGetVal("R_Color_5")))
        Me.f_R_Color_5_Qty.Text = IIf(IsNull(rstGetVal("R_Color_5_Qty")), 0, rstGetVal("R_Color_5_Qty"))
        Call selectValueInCombo(Me.f_R_Color_6, IIf(IsNull(rstGetVal("R_Color_6")), -1, rstGetVal("R_Color_6")))
        Me.f_R_Color_6_Qty.Text = IIf(IsNull(rstGetVal("R_Color_6_Qty")), 0, rstGetVal("R_Color_6_Qty"))
        Me.f_NewColor.Text = IIf(IsNull(rstGetVal("NewColor")), " ", rstGetVal("NewColor"))
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
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
Private Sub chkQty_Chemical_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CDbl(rstGetQty("Quantity"))
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
                    Me.f_Chemical2.SetFocus
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
Private Sub chkQty_Chemical2_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CDbl(rstGetQty("Quantity"))
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
                    Me.f_Acid.SetFocus
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
                AvbQty = CDbl(rstGetQty("Quantity"))
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
                    Me.f_Soap.SetFocus
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
Private Sub chkQty_Soap_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CDbl(rstGetQty("Quantity"))
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
                    Me.f_SoapTime.SetFocus
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
Private Sub chkQty_Castic_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CDbl(rstGetQty("Quantity"))
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
                    Me.f_Hydro.SetFocus
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
Private Sub chkQty_hydro_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CDbl(rstGetQty("Quantity"))
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
                    Me.f_CasticTime.SetFocus
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
Private Sub chkQty_Soap2_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CDbl(rstGetQty("Quantity"))
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
                    Me.f_Acid2.SetFocus
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
Private Sub chkQty_Acid2_Qty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
    Dim vTemp As Integer
    Dim rstGetQty As New ADODB.Recordset
    Set rstGetQty = FillRecordSet("Select Qty * 1000 as Quantity from vwAvailableQty where ItemCode = " & vItemCode)
    AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Quantity"))) Then
                AvbQty = CDbl(rstGetQty("Quantity"))
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
                    Me.f_Remarks.SetFocus
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

Private Sub PCChk_Click()
    If PCChk.value = Checked Then
        Me.srPC1.Enabled = True
        Me.srPC2.Enabled = True
    Else
        Me.srPC1.Enabled = False
        Me.srPC2.Enabled = False
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
Private Sub SrColor_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrPC1_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrPC2_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrItemType_Click()
    If Me.SrItemType.ListIndex > -1 Then
        i = Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = " & i & " order by 2", SrItem, "ItemName", "ItemCode"
    Else
        Me.SrItem.Clear
    End If
    Call SrfillList
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
        'srdt = " And (ProcessDate between #" & Me.SrDate.value - 1 & " # and #" & Me.SrDate2.value + 1 & " #)"
        srdt = " And (ProcessDate between Convert(datetime, '" & Me.SrDate.value - 1 & "')  and Convert(datetime, '" & Me.SrDate2.value + 1 & "'))"
    Else
        srdt = ""
    End If
    
    If PtChk.value = Checked And Me.srParty.ListIndex > -1 Then
        cbo1 = " And Process.partycode = " & Me.srParty.ItemData(Me.srParty.ListIndex)
    Else
        cbo1 = ""
    End If
    
    If ImTChk.value = Checked And Me.SrItemType.ListIndex > -1 Then
        cbo2 = " And Process.ItemTypeCode = " & Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
    Else
        cbo2 = ""
    End If
    
    If ImChk.value = Checked And Me.SrItem.ListIndex > -1 Then
        cbo3 = " And Process.Cone = " & Me.SrItem.ItemData(Me.SrItem.ListIndex)
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
    
    If PCChk.value = Checked And Len(Trim(Me.srPC1)) > 0 And Len(Trim(Me.srPC2)) > 0 Then
        cbo6 = " And (ProcessCode between " & Me.srPC1 & " and " & Me.srPC2 & " )"
    Else
        cbo6 = ""
    End If
    sql = " SELECT top 100 ProcessCode, ProcessDate, PartyName, MachineNo, ItemTypeName, Cone, isNull(RecipeCode, 0) as RecipeCode, isNull(Re_RecipeCode, 0) as Re_RecipeCode " & _
          " FROM Party INNER JOIN (ItemType INNER JOIN Process ON ItemType.ItemTypeCode = Process.ItemTypeCode) ON Party.PartyCode = Process.PartyCode " & _
          " Where Process.Is_Active = 1 and Is_Cotton_Dyeing = 0 " & _
          srdt & _
          cbo1 & _
          cbo2 & _
          cbo3 & _
          cbo4 & _
          cbo5 & _
          cbo6 & _
          " order by ProcessDate desc"
                                
    Debug.Print sql
    Set rstList = FillRecordSet(sql)
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!ProcessCode, _
                   Key:=CStr("Id=" & rstList!ProcessCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!ProcessDate
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!MachineNo
                 .Add Text:=rstList!ItemTypeName
                 .Add Text:=rstList!Cone
                 .Add Text:=rstList!RecipeCode
                 .Add Text:=rstList!Re_RecipeCode
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub srMachine_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrParty_Click()
    Call SrfillList
End Sub
Private Sub SrItem_Click()
    Call SrfillList
End Sub
Private Sub SrDate_Change()
'    If Me.SrDate.value >= Me.SrDate2.value Then
        Call SrfillList
 '   End If
End Sub
Private Sub SrDate2_Change()
  '  If Me.SrDate.value >= Me.SrDate2.value Then
        Call SrfillList
  '  End If
End Sub
Private Sub f_R_Color_1_Click()
    If Me.f_R_Color_1.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_R_Color_1.ItemData(Me.f_R_Color_1.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
            If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
            End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_R_Color_1_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_R_Color_2_Click()
    If Me.f_R_Color_2.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_R_Color_2.ItemData(Me.f_R_Color_2.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
            If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
            End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_R_Color_2_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_R_Color_3_Click()
    If Me.f_R_Color_3.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_R_Color_3.ItemData(Me.f_R_Color_3.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
            If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
            End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_R_Color_3_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_R_Color_4_Click()
    If Me.f_R_Color_4.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_R_Color_4.ItemData(Me.f_R_Color_4.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
            If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
            End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_R_Color_4_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_R_Color_5_Click()
    If Me.f_R_Color_5.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_R_Color_5.ItemData(Me.f_R_Color_5.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
            If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
            End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_R_Color_5_Qty.Text = (Qty * kg)
    End If
End Sub
Private Sub f_R_Color_6_Click()
    If Me.f_R_Color_6.ListIndex > 0 Then
        Dim rstGetQty As New ADODB.Recordset
        i = Me.f_R_Color_6.ItemData(Me.f_R_Color_6.ListIndex)
        Set rstGetQty = FillRecordSet("Select ItemCode, Quantity from RecipeDetail where RecipeMasterCode = " & f_RecipeCode.Text & " and ItemCode = " & i)
            If Not (rstGetQty.EOF) Then
                Qty = rstGetQty("Quantity")
            End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        kg = Round(Me.f_ConeKG.Text)
        Me.f_R_Color_6_Qty.Text = (Qty * kg)
    End If
End Sub

