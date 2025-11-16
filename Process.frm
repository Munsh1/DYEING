VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Process 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process"
   ClientHeight    =   11600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15500
   StartUpPosition =   2  'Center screen
   LinkTopic       =   "Form1"
   MaxButton       =   1   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleMode       =   0  'User
   ScaleWidth      =   14997.82
   Begin Crystal.CrystalReport crptDaily 
      Left            =   0
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Process.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Process.frx":0268
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Process.frx":06C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Process.frx":0ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Process.frx":0F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Process.frx":1330
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Process.frx":176C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Process.frx":1BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Process.frx":1E38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   1695
      Left            =   0
      TabIndex        =   171
      Top             =   8040
      Width           =   10766
      Begin MSComctlLib.ListView lvwphase 
         Height          =   1320
         Left            =   75
         TabIndex        =   184
         Top             =   240
         Width           =   10605
         _ExtentX        =   18706
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
      Height          =   10245
      Left            =   10813
      TabIndex        =   0
      Top             =   0
      Width           =   3000
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   120
         TabIndex        =   221
         Top             =   7200
         Width           =   2775
         Begin VB.TextBox srSER2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   222
            Top             =   300
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox srSER1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   169
            Top             =   300
            Width           =   1000
         End
         Begin VB.CheckBox SERChk 
            Caption         =   "Serial #"
            Height          =   195
            Left            =   240
            TabIndex        =   168
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame19 
         Height          =   735
         Left            =   120
         TabIndex        =   182
         Top             =   6360
         Width           =   2775
         Begin VB.CheckBox PCChk 
            Caption         =   "PC Code"
            Height          =   255
            Left            =   240
            TabIndex        =   165
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox srPC2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   167
            Top             =   300
            Width           =   1000
         End
         Begin VB.TextBox srPC1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   166
            Top             =   300
            Width           =   1000
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   120
         TabIndex        =   181
         Top             =   5520
         Width           =   2775
         Begin VB.TextBox SrColor 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   164
            Top             =   320
            Width           =   2535
         End
         Begin VB.CheckBox ClChk 
            Caption         =   "Color"
            Height          =   255
            Left            =   240
            TabIndex        =   163
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.CheckBox ImChk 
         Caption         =   "Item"
         Height          =   255
         Left            =   360
         TabIndex        =   161
         Top             =   4560
         Width           =   735
      End
      Begin VB.CheckBox ImTChk 
         Caption         =   "Item Type"
         Height          =   255
         Left            =   360
         TabIndex        =   159
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CheckBox McChk 
         Caption         =   "Machine"
         Height          =   255
         Left            =   360
         TabIndex        =   157
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox PtChk 
         Caption         =   "Party"
         Height          =   255
         Left            =   360
         TabIndex        =   155
         Top             =   1680
         Width           =   735
      End
      Begin VB.Frame Frame15 
         Height          =   800
         Left            =   120
         TabIndex        =   180
         Top             =   4560
         Width           =   2800
         Begin VB.ComboBox SrItem 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            TabIndex        =   162
            Text            =   "SrItem"
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame14 
         Height          =   800
         Left            =   100
         TabIndex        =   179
         Top             =   3600
         Width           =   2800
         Begin VB.ComboBox SrItemType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            TabIndex        =   160
            Text            =   "SrItemType"
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame13 
         Height          =   800
         Left            =   100
         TabIndex        =   178
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
            TabIndex        =   158
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame12 
         Height          =   800
         Left            =   100
         TabIndex        =   177
         Top             =   1680
         Width           =   2800
         Begin VB.ComboBox srParty 
            Enabled         =   0   'False
            Height          =   315
            Left            =   125
            TabIndex        =   156
            Text            =   "srParty"
            Top             =   280
            Width           =   2600
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1155
         Left            =   100
         TabIndex        =   176
         Top             =   360
         Width           =   2800
         Begin MSComCtl2.DTPicker SrDate2 
            Height          =   330
            Left            =   120
            TabIndex        =   154
            Top             =   720
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   47841281
            CurrentDate     =   38298
         End
         Begin VB.CheckBox dtChk 
            Caption         =   "Date"
            Height          =   195
            Left            =   240
            TabIndex        =   152
            Top             =   0
            Width           =   735
         End
         Begin MSComCtl2.DTPicker SrDate 
            Height          =   330
            Left            =   125
            TabIndex        =   153
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
            Format          =   47841281
            CurrentDate     =   38235
         End
      End
      Begin LVbuttons.LaVolpeButton Cmdhide 
         Height          =   375
         Left            =   480
         TabIndex        =   170
         Top             =   9480
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
         MICON           =   "Process.frx":228A
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
      Left            =   8340
      TabIndex        =   147
      Top             =   9840
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
      MICON           =   "Process.frx":22A6
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
      Left            =   9600
      TabIndex        =   148
      Top             =   9840
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
      MICON           =   "Process.frx":22C2
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
      Left            =   4800
      TabIndex        =   149
      Top             =   9840
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
      MICON           =   "Process.frx":22DE
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
      Left            =   7200
      TabIndex        =   145
      Top             =   9840
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
      MICON           =   "Process.frx":22FA
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
      Left            =   6042
      TabIndex        =   144
      Top             =   9840
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
      MICON           =   "Process.frx":2316
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
      Left            =   3555
      TabIndex        =   150
      Top             =   9840
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
      MICON           =   "Process.frx":2332
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
      Left            =   2160
      TabIndex        =   151
      Top             =   9840
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
      MICON           =   "Process.frx":234E
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
      Height          =   8055
      Left            =   0
      TabIndex        =   172
      Top             =   0
      Width           =   10766
      Begin VB.TextBox f_NewColor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9600
         TabIndex        =   5
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox f_Remarks 
         Height          =   1455
         Left            =   6240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   146
         Top             =   6480
         Width           =   4400
      End
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   50
         TabIndex        =   206
         Top             =   6325
         Width           =   6120
         Begin VB.Frame Frame16 
            Height          =   825
            Left            =   6120
            TabIndex        =   216
            Top             =   960
            Visible         =   0   'False
            Width           =   2055
            Begin VB.TextBox f_RecipeCode 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   218
               Top             =   480
               Width           =   800
            End
            Begin VB.CheckBox f_Re_RecipeCode 
               Height          =   255
               Left            =   360
               TabIndex        =   217
               Top             =   480
               Width           =   255
            End
            Begin VB.Label Label19 
               Caption         =   "Recipe"
               Height          =   225
               Left            =   1200
               TabIndex        =   220
               Top             =   180
               Width           =   735
            End
            Begin VB.Label Label20 
               Caption         =   "Re. Recipe"
               Height          =   225
               Left            =   120
               TabIndex        =   219
               Top             =   180
               Width           =   855
            End
         End
         Begin VB.TextBox f_Acid2_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            TabIndex        =   143
            Top             =   1250
            Width           =   650
         End
         Begin VB.TextBox f_Acid_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   141
            Top             =   1250
            Width           =   650
         End
         Begin VB.ComboBox f_Acid2 
            Height          =   315
            Left            =   3315
            TabIndex        =   142
            Top             =   1250
            Width           =   1200
         End
         Begin VB.ComboBox f_Acid 
            Height          =   315
            Left            =   720
            TabIndex        =   140
            Top             =   1250
            Width           =   1200
         End
         Begin VB.ComboBox f_Soap2 
            Height          =   315
            Left            =   720
            TabIndex        =   125
            Top             =   200
            Width           =   1200
         End
         Begin VB.TextBox f_Soap2_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   126
            Top             =   200
            Width           =   650
         End
         Begin VB.ComboBox f_Soap3 
            Height          =   315
            Left            =   3315
            TabIndex        =   127
            Top             =   200
            Width           =   1200
         End
         Begin VB.TextBox f_Soap3_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4545
            TabIndex        =   128
            Top             =   200
            Width           =   650
         End
         Begin VB.TextBox f_SoapTime2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5250
            TabIndex        =   129
            Top             =   200
            Width           =   650
         End
         Begin VB.ComboBox f_Hydro2 
            Height          =   315
            Left            =   720
            TabIndex        =   130
            Top             =   550
            Width           =   1200
         End
         Begin VB.TextBox f_Hydro_Qty2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   131
            Top             =   550
            Width           =   650
         End
         Begin VB.ComboBox f_Castic2 
            Height          =   315
            Left            =   3315
            TabIndex        =   132
            Top             =   550
            Width           =   1200
         End
         Begin VB.TextBox f_Castic_Qty2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4545
            TabIndex        =   133
            Top             =   550
            Width           =   650
         End
         Begin VB.TextBox f_CasticTime2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5250
            TabIndex        =   134
            Top             =   550
            Width           =   650
         End
         Begin VB.ComboBox f_Hydro3 
            Height          =   315
            Left            =   720
            TabIndex        =   135
            Top             =   900
            Width           =   1200
         End
         Begin VB.ComboBox f_Castic3 
            Height          =   315
            Left            =   3315
            TabIndex        =   137
            Top             =   900
            Width           =   1200
         End
         Begin VB.TextBox f_Hydro_Qty3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   136
            Top             =   900
            Width           =   650
         End
         Begin VB.TextBox f_Castic_Qty3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            TabIndex        =   138
            Top             =   900
            Width           =   650
         End
         Begin VB.TextBox f_CasticTime3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5250
            TabIndex        =   139
            Top             =   900
            Width           =   650
         End
         Begin VB.Label Label18 
            Caption         =   "Acid"
            Height          =   255
            Left            =   2880
            TabIndex        =   215
            Top             =   1250
            Width           =   375
         End
         Begin VB.Label Label17 
            Caption         =   "Acid"
            Height          =   255
            Left            =   120
            TabIndex        =   214
            Top             =   1250
            Width           =   495
         End
         Begin VB.Label Label15 
            Caption         =   "Castic"
            Height          =   255
            Left            =   2800
            TabIndex        =   212
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Castic"
            Height          =   255
            Left            =   2800
            TabIndex        =   211
            Top             =   550
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Soap"
            Height          =   255
            Left            =   2800
            TabIndex        =   210
            Top             =   200
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Soap"
            Height          =   255
            Left            =   100
            TabIndex        =   209
            Top             =   200
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Hydro"
            Height          =   255
            Left            =   100
            TabIndex        =   208
            Top             =   550
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Hydro"
            Height          =   255
            Left            =   100
            TabIndex        =   207
            Top             =   900
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   50
         TabIndex        =   202
         Top             =   5010
         Width           =   10680
         Begin VB.TextBox f_TempTime3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10080
            TabIndex        =   114
            Top             =   555
            Width           =   500
         End
         Begin VB.TextBox f_R_Color_10_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9420
            TabIndex        =   124
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_R_Color_9_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7550
            TabIndex        =   122
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_R_Color_8_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   120
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_R_Color_7_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3800
            TabIndex        =   118
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_R_Color_6_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   116
            Top             =   920
            Width           =   650
         End
         Begin VB.ComboBox f_R_Color_10 
            Height          =   315
            Left            =   8220
            TabIndex        =   123
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_R_Color_9 
            Height          =   315
            Left            =   6350
            TabIndex        =   121
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_R_Color_8 
            Height          =   315
            Left            =   4470
            TabIndex        =   119
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_R_Color_7 
            Height          =   315
            Left            =   2600
            TabIndex        =   117
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_R_Color_6 
            Height          =   315
            Left            =   720
            TabIndex        =   115
            Top             =   920
            Width           =   1200
         End
         Begin VB.TextBox f_Temp3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10080
            TabIndex        =   103
            Top             =   195
            Width           =   500
         End
         Begin VB.TextBox f_Chemical_15_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9420
            TabIndex        =   102
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_15_Code 
            Height          =   315
            Left            =   8220
            TabIndex        =   101
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_14_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7550
            TabIndex        =   100
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_14_Code 
            Height          =   315
            Left            =   6350
            TabIndex        =   99
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_13_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   98
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_13_Code 
            Height          =   315
            Left            =   4470
            TabIndex        =   97
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_12_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3800
            TabIndex        =   96
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_12_Code 
            Height          =   315
            Left            =   2600
            TabIndex        =   95
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_11_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   94
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_11_Code 
            Height          =   315
            Left            =   720
            TabIndex        =   93
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_15 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9420
            TabIndex        =   113
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_15 
            Height          =   315
            Left            =   8220
            TabIndex        =   112
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_14 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7550
            TabIndex        =   111
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_14 
            Height          =   315
            Left            =   6350
            TabIndex        =   110
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_13 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   109
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_13 
            Height          =   315
            Left            =   4470
            TabIndex        =   108
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_12 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3800
            TabIndex        =   107
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_12 
            Height          =   315
            Left            =   2600
            TabIndex        =   106
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_11 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   105
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_11 
            Height          =   315
            Left            =   720
            TabIndex        =   104
            Top             =   555
            Width           =   1200
         End
         Begin VB.Label Label9 
            Caption         =   "Color"
            Height          =   255
            Left            =   50
            TabIndex        =   205
            Top             =   920
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Chemical"
            Height          =   255
            Left            =   45
            TabIndex        =   204
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Recipe-3"
            Height          =   255
            Left            =   45
            TabIndex        =   203
            Top             =   555
            Width           =   735
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1335
         Left            =   50
         TabIndex        =   198
         Top             =   3680
         Width           =   10680
         Begin VB.ComboBox f_RecipeColor_6 
            Height          =   315
            Left            =   720
            TabIndex        =   72
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   73
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_7 
            Height          =   315
            Left            =   2600
            TabIndex        =   74
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3800
            TabIndex        =   75
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_8 
            Height          =   315
            Left            =   4470
            TabIndex        =   76
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_8 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   77
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_9 
            Height          =   315
            Left            =   6350
            TabIndex        =   78
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_9 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7550
            TabIndex        =   79
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_10 
            Height          =   315
            Left            =   8220
            TabIndex        =   80
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_10 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9420
            TabIndex        =   81
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_6_Code 
            Height          =   315
            Left            =   720
            TabIndex        =   61
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_6_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   62
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_7_Code 
            Height          =   315
            Left            =   2600
            TabIndex        =   63
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_7_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3800
            TabIndex        =   64
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_8_Code 
            Height          =   315
            Left            =   4470
            TabIndex        =   65
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_8_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   66
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_9_Code 
            Height          =   315
            Left            =   6350
            TabIndex        =   67
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_9_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7550
            TabIndex        =   68
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_10_Code 
            Height          =   315
            Left            =   8220
            TabIndex        =   69
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_10_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9420
            TabIndex        =   70
            Top             =   195
            Width           =   650
         End
         Begin VB.TextBox f_Temp2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10080
            TabIndex        =   71
            Top             =   195
            Width           =   500
         End
         Begin VB.ComboBox f_R_Color_1 
            Height          =   315
            Left            =   720
            TabIndex        =   83
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_R_Color_2 
            Height          =   315
            Left            =   2600
            TabIndex        =   85
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_R_Color_3 
            Height          =   315
            Left            =   4470
            TabIndex        =   87
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_R_Color_4 
            Height          =   315
            Left            =   6350
            TabIndex        =   89
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_R_Color_5 
            Height          =   315
            Left            =   8220
            TabIndex        =   91
            Top             =   920
            Width           =   1200
         End
         Begin VB.TextBox f_R_Color_1_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   84
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_R_Color_2_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3800
            TabIndex        =   86
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_R_Color_3_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   88
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_R_Color_4_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7550
            TabIndex        =   90
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_R_Color_5_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9420
            TabIndex        =   92
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_TempTime2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10080
            TabIndex        =   82
            Top             =   555
            Width           =   500
         End
         Begin VB.Label Label34 
            Caption         =   "Recipe-2"
            Height          =   255
            Left            =   45
            TabIndex        =   201
            Top             =   555
            Width           =   735
         End
         Begin VB.Label Label33 
            Caption         =   "Chemical"
            Height          =   255
            Left            =   45
            TabIndex        =   200
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "Color"
            Height          =   255
            Left            =   50
            TabIndex        =   199
            Top             =   920
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1695
         Left            =   50
         TabIndex        =   191
         Top             =   1980
         Width           =   10680
         Begin VB.TextBox f_CasticTime 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   58
            Top             =   1280
            Width           =   500
         End
         Begin VB.TextBox f_Hydro_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8335
            TabIndex        =   60
            Top             =   1280
            Width           =   650
         End
         Begin VB.ComboBox f_Hydro 
            Height          =   315
            Left            =   7140
            TabIndex        =   59
            Top             =   1280
            Width           =   1200
         End
         Begin VB.TextBox f_Castic_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4955
            TabIndex        =   57
            Top             =   1280
            Width           =   650
         End
         Begin VB.ComboBox f_Castic 
            Height          =   315
            Left            =   3750
            TabIndex        =   56
            Top             =   1280
            Width           =   1200
         End
         Begin VB.TextBox f_SoapTime 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2600
            TabIndex        =   55
            Top             =   1280
            Width           =   650
         End
         Begin VB.TextBox f_Soap_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   54
            Top             =   1280
            Width           =   650
         End
         Begin VB.ComboBox f_Soap 
            Height          =   315
            Left            =   720
            TabIndex        =   53
            Top             =   1280
            Width           =   1200
         End
         Begin VB.TextBox f_TempTime 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10080
            TabIndex        =   42
            Top             =   555
            Width           =   500
         End
         Begin VB.TextBox f_Color_5_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9400
            TabIndex        =   52
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_Color_4_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7550
            TabIndex        =   50
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_Color_3_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   48
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_Color_2_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3800
            TabIndex        =   46
            Top             =   920
            Width           =   650
         End
         Begin VB.TextBox f_Color_1_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   44
            Top             =   920
            Width           =   650
         End
         Begin VB.ComboBox f_Color_5 
            Height          =   315
            Left            =   8220
            TabIndex        =   51
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_Color_4 
            Height          =   315
            Left            =   6350
            TabIndex        =   49
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_Color_3 
            Height          =   315
            Left            =   4470
            TabIndex        =   47
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_Color_2 
            Height          =   315
            Left            =   2600
            TabIndex        =   45
            Top             =   920
            Width           =   1200
         End
         Begin VB.ComboBox f_Color_1 
            Height          =   315
            Left            =   720
            TabIndex        =   43
            Top             =   920
            Width           =   1200
         End
         Begin VB.TextBox f_Temp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10080
            TabIndex        =   31
            Top             =   195
            Width           =   500
         End
         Begin VB.TextBox f_Chemical_5_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9400
            TabIndex        =   30
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_5_Code 
            Height          =   315
            Left            =   8220
            TabIndex        =   29
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_4_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7550
            TabIndex        =   28
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_4_Code 
            Height          =   315
            Left            =   6350
            TabIndex        =   27
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_3_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   26
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical_3_Code 
            Height          =   315
            Left            =   4470
            TabIndex        =   25
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical2_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3800
            TabIndex        =   24
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical2 
            Height          =   315
            Left            =   2600
            TabIndex        =   23
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_Chemical_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   22
            Top             =   195
            Width           =   650
         End
         Begin VB.ComboBox f_Chemical 
            Height          =   315
            Left            =   720
            TabIndex        =   21
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9400
            TabIndex        =   41
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_5 
            Height          =   315
            Left            =   8220
            TabIndex        =   40
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7550
            TabIndex        =   39
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_4 
            Height          =   315
            Left            =   6350
            TabIndex        =   38
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5680
            TabIndex        =   37
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_3 
            Height          =   315
            Left            =   4470
            TabIndex        =   36
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty_2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3800
            TabIndex        =   35
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor_2 
            Height          =   315
            Left            =   2600
            TabIndex        =   34
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox f_RecipeColor_Qty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   33
            Top             =   555
            Width           =   650
         End
         Begin VB.ComboBox f_RecipeColor 
            Height          =   315
            Left            =   720
            TabIndex        =   32
            Top             =   555
            Width           =   1200
         End
         Begin VB.Label Label32 
            Caption         =   "Hydro"
            Height          =   255
            Left            =   6720
            TabIndex        =   197
            Top             =   1275
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "Castic"
            Height          =   255
            Left            =   3240
            TabIndex        =   196
            Top             =   1275
            Width           =   495
         End
         Begin VB.Label Label30 
            Caption         =   "Soap"
            Height          =   255
            Left            =   50
            TabIndex        =   195
            Top             =   1280
            Width           =   495
         End
         Begin VB.Label Label29 
            Caption         =   "Color"
            Height          =   255
            Left            =   50
            TabIndex        =   194
            Top             =   920
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "Chemical"
            Height          =   255
            Left            =   45
            TabIndex        =   193
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "Recipe"
            Height          =   255
            Left            =   45
            TabIndex        =   192
            Top             =   555
            Width           =   615
         End
      End
      Begin VB.TextBox f_SerialNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   4
         Top             =   200
         Width           =   855
      End
      Begin VB.Frame Frame20 
         Height          =   1500
         Left            =   50
         TabIndex        =   183
         Top             =   480
         Width           =   10680
         Begin LVbuttons.LaVolpeButton Clr3 
            Height          =   300
            Left            =   9900
            TabIndex        =   225
            Top             =   1080
            Width           =   700
            _ExtentX        =   1244
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Clr"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "Process.frx":236A
            ALIGN           =   1
            IMGLST          =   "ImageList1"
            IMGICON         =   "9"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   0
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
         Begin LVbuttons.LaVolpeButton Clr2 
            Height          =   300
            Left            =   9900
            TabIndex        =   224
            Top             =   730
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Clr"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "Process.frx":2386
            ALIGN           =   1
            IMGLST          =   "ImageList1"
            IMGICON         =   "9"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   0
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
         Begin LVbuttons.LaVolpeButton Clr1 
            Height          =   300
            Left            =   9900
            TabIndex        =   223
            Top             =   390
            Width           =   700
            _ExtentX        =   1244
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Clr"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            MICON           =   "Process.frx":23A2
            ALIGN           =   1
            IMGLST          =   "ImageList1"
            IMGICON         =   "9"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   0
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
         Begin VB.TextBox f_Den_3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9135
            TabIndex        =   20
            Top             =   1080
            Width           =   700
         End
         Begin VB.TextBox f_Den_2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9135
            TabIndex        =   15
            Top             =   730
            Width           =   700
         End
         Begin VB.TextBox f_ConeKG_3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8410
            TabIndex        =   19
            Top             =   1080
            Width           =   700
         End
         Begin VB.TextBox f_ConeKG_2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8410
            TabIndex        =   14
            Top             =   720
            Width           =   700
         End
         Begin VB.ComboBox f_Cone_3 
            Height          =   315
            Left            =   5625
            TabIndex        =   18
            Top             =   1080
            Width           =   2800
         End
         Begin VB.ComboBox f_Cone_2 
            Height          =   315
            Left            =   5625
            TabIndex        =   13
            Top             =   730
            Width           =   2800
         End
         Begin VB.ComboBox f_ItemTypeCode_3 
            Height          =   315
            Left            =   3100
            TabIndex        =   17
            Top             =   1080
            Width           =   2540
         End
         Begin VB.ComboBox f_ItemTypeCode_2 
            Height          =   315
            Left            =   3100
            TabIndex        =   12
            Top             =   720
            Width           =   2540
         End
         Begin VB.ComboBox f_PartyCode_3 
            Height          =   315
            Left            =   100
            TabIndex        =   16
            Top             =   1080
            Width           =   3000
         End
         Begin VB.ComboBox f_PartyCode_2 
            Height          =   315
            ItemData        =   "Process.frx":23BE
            Left            =   100
            List            =   "Process.frx":23C0
            TabIndex        =   11
            Top             =   730
            Width           =   3000
         End
         Begin VB.TextBox f_Den 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9135
            TabIndex        =   10
            Top             =   390
            Width           =   700
         End
         Begin VB.TextBox f_ConeKG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8410
            TabIndex        =   9
            Top             =   390
            Width           =   700
         End
         Begin VB.ComboBox f_Cone 
            Height          =   315
            Left            =   5625
            TabIndex        =   8
            Top             =   390
            Width           =   2800
         End
         Begin VB.ComboBox f_ItemTypeCode 
            Height          =   315
            Left            =   3100
            TabIndex        =   7
            Top             =   390
            Width           =   2540
         End
         Begin VB.ComboBox f_PartyCode 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   390
            Width           =   3000
         End
         Begin VB.Label Label26 
            Caption         =   "Cone"
            Height          =   255
            Left            =   9135
            TabIndex        =   189
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label25 
            Caption         =   "KG"
            Height          =   255
            Left            =   8410
            TabIndex        =   188
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label24 
            Caption         =   "Item"
            Height          =   255
            Left            =   5625
            TabIndex        =   187
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Item Type"
            Height          =   255
            Left            =   3100
            TabIndex        =   186
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Party"
            Height          =   255
            Left            =   120
            TabIndex        =   185
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.TextBox f_MachineNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Top             =   200
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker f_ProcessDate 
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Top             =   200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47841281
         CurrentDate     =   38235
      End
      Begin MSComCtl2.DTPicker f_ProcessTime 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47841282
         CurrentDate     =   38235
      End
      Begin VB.Label Label16 
         Caption         =   "New Color"
         Height          =   255
         Left            =   8760
         TabIndex        =   213
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Serial No."
         Height          =   255
         Left            =   6720
         TabIndex        =   190
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Machine No."
         Height          =   225
         Left            =   4440
         TabIndex        =   175
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Time"
         Height          =   225
         Left            =   2280
         TabIndex        =   174
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   225
         Left            =   120
         TabIndex        =   173
         Top             =   200
         Width           =   375
      End
   End
End
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ListID As String
Dim m_AddMode As Boolean
Dim CMDSearch As Boolean
Dim PreQty As Double
Dim vQty As Double
Dim QtyBit As Integer
Dim MsgBit As Integer
Dim ClickPane As Integer
Private Sub Clr2_Click()
    If Len(m_ListID) > 0 Then
        Dim strAns As String
        Dim rstGetQty As New ADODB.Recordset
        
        strAns = MsgBox("Do you want to clear this record...?", vbYesNo + vbInformation)
        If strAns = vbYes Then
            cnDatabase.Execute "update Process set PartyCode2 = null, ItemTypeCode2 = null, Den2 = null, Cone2 = null, ConeKG2 = null Where ProcessCode = " & m_ListID
   
            Me.f_PartyCode_2.ListIndex = -1
            Me.f_ItemTypeCode_2.ListIndex = -1
            Me.f_Cone_2.ListIndex = -1
            Me.f_ConeKG_2.Text = ""
            Me.f_Den_2.Text = ""
            
            Call fillList
            MsgBox ("Record deleted succesfully..."), vbInformation
        End If
    End If
End Sub
Private Sub Clr3_Click()
    If Len(m_ListID) > 0 Then
        Dim strAns As String
        Dim rstGetQty As New ADODB.Recordset
        
        strAns = MsgBox("Do you want to clear this record...?", vbYesNo + vbInformation)
        If strAns = vbYes Then
            cnDatabase.Execute "update Process set PartyCode3 = null, ItemTypeCode3 = null, Den3 = null, Cone3 = null, ConeKG3 = null Where ProcessCode = " & m_ListID
   
            Me.f_PartyCode_3.ListIndex = -1
            Me.f_ItemTypeCode_3.ListIndex = -1
            Me.f_Cone_3.ListIndex = -1
            Me.f_ConeKG_3.Text = ""
            Me.f_Den_3.Text = ""
            
            Call fillList
            MsgBox ("Record deleted succesfully..."), vbInformation
        End If
    End If
End Sub
Private Sub f_MachineNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_SerialNo.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_MachineNo_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_SerialNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_NewColor.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_SerialNo_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_NewColor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_PartyCode.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_NewColor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_PartyCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_ItemTypeCode.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_PartyCode_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_ItemTypeCode_Click()
    If ClickPane = 0 And Me.f_ItemTypeCode.ListIndex > -1 Then
        i = Me.f_ItemTypeCode.ItemData(Me.f_ItemTypeCode.ListIndex)
        If i = 1 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type1 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 2 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type2 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 3 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type3 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 4 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type4 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 5 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type5 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 6 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type6 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 7 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type7 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 8 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type8 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 9 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type9 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 10 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type10 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 11 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type11 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 12 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type12 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 13 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type13 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 14 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type14 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 15 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type15 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 16 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type16 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 17 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type17 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 18 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type18 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 19 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type19 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 20 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type20 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 21 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type21 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 22 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type22 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 23 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type23 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 24 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type24 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
        ElseIf i = 25 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type25 where Qty > 0 ", f_Cone, "ItemName", "ItemCode"
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
    If KeyAscii = 13 Then
        Me.f_Cone.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_ItemTypeCode_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Cone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_ConeKG.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Cone_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_ConeKG_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Den.SetFocus
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_ConeKG_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Den_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_PartyCode_2.SetFocus
    End If
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Den_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_PartyCode_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_ItemTypeCode_2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_PartyCode_2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_ItemTypeCode_2_Click()
    If ClickPane = 0 And Me.f_ItemTypeCode_2.ListIndex > -1 Then
        i = Me.f_ItemTypeCode_2.ItemData(Me.f_ItemTypeCode_2.ListIndex)
        If i = 1 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type1 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 2 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type2 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 3 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type3 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 4 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type4 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 5 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type5 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 6 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type6 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 7 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type7 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 8 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type8 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 9 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type9 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 10 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type10 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 11 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type11 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 12 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type12 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 13 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type13 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 14 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type14 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 15 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type15 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 16 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type16 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 17 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type17 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 18 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type18 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 19 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type19 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 20 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type20 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 21 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type21 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 22 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type22 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 23 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type23 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 24 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type24 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        ElseIf i = 25 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type25 where Qty > 0 ", f_Cone_2, "ItemName", "ItemCode"
        End If
    ElseIf ClickPane = 1 And Me.f_ItemTypeCode_2.ListIndex > -1 Then
        i = Me.f_ItemTypeCode_2.ItemData(Me.f_ItemTypeCode_2.ListIndex)
        FillCombo "Select ItemCode, ItemName from Item where ItemTypeCode = " & i, f_Cone_2, "ItemName", "ItemCode"
        ClickPane = 0
    Else
        Me.f_Cone_2.Clear
    End If
End Sub
Private Sub f_ItemTypeCode_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Cone_2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_ItemTypeCode_2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Cone_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_ConeKG_2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Cone_2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_ConeKG_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Den_2.SetFocus
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_ConeKG_2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Den_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_PartyCode_3.SetFocus
    End If
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Den_2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_PartyCode_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_ItemTypeCode_3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_PartyCode_3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_ItemTypeCode_3_Click()
    If ClickPane = 0 And Me.f_ItemTypeCode_3.ListIndex > -1 Then
        i = Me.f_ItemTypeCode_3.ItemData(Me.f_ItemTypeCode_3.ListIndex)
        If i = 1 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type1 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 2 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type2 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 3 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type3 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 4 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type4 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 5 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type5 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 6 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type6 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 7 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type7 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 8 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type8 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 9 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type9 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 10 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type10 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 11 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type11 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 12 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type12 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 13 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type13 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 14 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type14 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 15 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type15 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 16 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type16 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 17 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type17 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 18 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type18 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 19 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type19 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 20 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type20 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 21 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type21 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 22 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type22 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 23 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type23 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 24 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type24 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        ElseIf i = 25 Then
            FillCombo "Select ItemCode, ItemName from vwAvailableQty_Type25 where Qty > 0 ", f_Cone_3, "ItemName", "ItemCode"
        End If
    ElseIf ClickPane = 1 And Me.f_ItemTypeCode_3.ListIndex > -1 Then
        i = Me.f_ItemTypeCode_3.ItemData(Me.f_ItemTypeCode_3.ListIndex)
        FillCombo "Select ItemCode, ItemName from Item where ItemTypeCode = " & i, f_Cone_3, "ItemName", "ItemCode"
        ClickPane = 0
    Else
        Me.f_Cone_3.Clear
    End If
End Sub
Private Sub f_ItemTypeCode_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Cone_3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_ItemTypeCode_3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Cone_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_ConeKG_3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Cone_3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_ConeKG_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Den_3.SetFocus
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_ConeKG_3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Den_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical.SetFocus
    End If
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Den_3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical.ItemData(Me.f_Chemical.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical.ListIndex = -1 Then
            Me.f_Chemical2.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical2_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical2_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical2_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical2_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical2_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical2.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical2_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical2_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical2.ItemData(Me.f_Chemical2.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical2.ListIndex = -1 Then
            Me.f_Chemical_3_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical2_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_3_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_3_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_3_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_3_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_3_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_3_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_3_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_3_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_3_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_3_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_3_Code.ItemData(Me.f_Chemical_3_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_3_Code.ListIndex = -1 Then
            Me.f_Chemical_4_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_3_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_4_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_4_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_4_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_4_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_4_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_4_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_4_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_4_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_4_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_4_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_4_Code.ItemData(Me.f_Chemical_4_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_4_Code.ListIndex = -1 Then
            Me.f_Chemical_5_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_4_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_5_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_5_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_5_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_5_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_5_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_5_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_5_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_5_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_5_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_5_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_5_Code.ItemData(Me.f_Chemical_5_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_5_Code.ListIndex = -1 Then
            Me.f_Temp.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_5_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Temp_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_RecipeColor.SetFocus
    End If
End Sub
Private Sub f_RecipeColor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_4.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_4.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_4_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_5.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_4_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_5.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_5_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_TempTime.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_5_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_TempTime_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Color_1.SetFocus
    End If
End Sub
Private Sub f_Color_1_Click()
    If Me.f_RecipeColor.ListIndex > -1 And Me.f_Color_1.ListIndex > -1 Then
        If Me.f_Color_1.ItemData(Me.f_Color_1.ListIndex) = Me.f_RecipeColor.ItemData(Me.f_RecipeColor.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_Color_1_Qty.Text = (CDbl(Me.f_RecipeColor_Qty.Text) * kg)
        Else
            Me.f_Color_1_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_Color_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Color_2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Color_1_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Color_2_Click()
    If Me.f_RecipeColor_2.ListIndex > -1 And Me.f_Color_2.ListIndex > -1 Then
        If Me.f_Color_2.ItemData(Me.f_Color_2.ListIndex) = Me.f_RecipeColor_2.ItemData(Me.f_RecipeColor_2.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_Color_2_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_2.Text) * kg)
        Else
            Me.f_Color_2_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_Color_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Color_3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Color_2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Color_3_Click()
    If Me.f_RecipeColor_3.ListIndex > -1 And Me.f_Color_3.ListIndex > -1 Then
        If Me.f_Color_3.ItemData(Me.f_Color_3.ListIndex) = Me.f_RecipeColor_3.ItemData(Me.f_RecipeColor_3.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_Color_3_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_3.Text) * kg)
        Else
            Me.f_Color_3_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_Color_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Color_4.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Color_3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Color_4_Click()
    If Me.f_RecipeColor_4.ListIndex > -1 And Me.f_Color_4.ListIndex > -1 Then
        If Me.f_Color_4.ItemData(Me.f_Color_4.ListIndex) = Me.f_RecipeColor_4.ItemData(Me.f_RecipeColor_4.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_Color_4_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_4.Text) * kg)
        Else
            Me.f_Color_4_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_Color_4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Color_5.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Color_4_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Color_5_Click()
    If Me.f_RecipeColor_5.ListIndex > -1 And Me.f_Color_5.ListIndex > -1 Then
        If Me.f_Color_5.ItemData(Me.f_Color_5.ListIndex) = Me.f_RecipeColor_5.ItemData(Me.f_RecipeColor_5.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_Color_5_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_5.Text) * kg)
        Else
            Me.f_Color_5_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_Color_5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Soap.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Color_5_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Soap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Soap_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Soap_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Soap_Qty_GotFocus()
    If Len(Trim(Me.f_Soap_Qty)) > 0 Then
        PreQty = Me.f_Soap_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Soap_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Soap.ListIndex > -1 Then
            If Len(Trim(Me.f_Soap_Qty.Text)) > 0 Then
                vQty = Me.f_Soap_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Soap.ItemData(Me.f_Soap.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Soap.ListIndex = -1 Then
            Me.f_SoapTime.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Soap_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_SoapTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Castic.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Castic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Castic_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Castic_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Castic_Qty_GotFocus()
    If Len(Trim(Me.f_Castic_Qty)) > 0 Then
        PreQty = Me.f_Castic_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Castic_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Castic.ListIndex > -1 Then
            If Len(Trim(Me.f_Castic_Qty.Text)) > 0 Then
                vQty = Me.f_Castic_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Castic.ItemData(Me.f_Castic.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Castic.ListIndex = -1 Then
            Me.f_CasticTime.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Castic_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_CasticTime_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_Hydro.SetFocus
    End If
End Sub
Private Sub f_Hydro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Hydro_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Hydro_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Hydro_Qty_GotFocus()
    If Len(Trim(Me.f_Hydro_Qty)) > 0 Then
        PreQty = Me.f_Hydro_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Hydro_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Hydro.ListIndex > -1 Then
            If Len(Trim(Me.f_Hydro_Qty.Text)) > 0 Then
                vQty = Me.f_Hydro_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Hydro.ItemData(Me.f_Hydro.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Hydro.ListIndex = -1 Then
            Me.f_Chemical_6_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Hydro_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_6_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_6_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_6_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_6_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_6_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_6_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_6_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_6_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_6_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_6_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_6_Code.ItemData(Me.f_Chemical_6_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_6_Code.ListIndex = -1 Then
            Me.f_Chemical_7_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_6_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_7_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_7_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_7_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_7_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_7_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_7_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_7_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_7_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_7_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_7_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_7_Code.ItemData(Me.f_Chemical_7_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_7_Code.ListIndex = -1 Then
            Me.f_Chemical_8_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_7_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_8_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_8_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_8_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_8_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_8_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_8_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_8_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_8_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_8_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_8_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_8_Code.ItemData(Me.f_Chemical_8_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_8_Code.ListIndex = -1 Then
            Me.f_Chemical_9_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_8_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_9_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_9_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_9_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_9_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_9_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_9_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_9_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_9_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_9_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_9_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_9_Code.ItemData(Me.f_Chemical_9_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_9_Code.ListIndex = -1 Then
            Me.f_Chemical_10_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_9_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_10_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_10_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_10_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_10_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_10_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_10_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_10_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_10_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_10_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_10_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_10_Code.ItemData(Me.f_Chemical_10_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_10_Code.ListIndex = -1 Then
            Me.f_Temp2.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_10_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Temp2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_RecipeColor_6.SetFocus
    End If
End Sub
Private Sub f_RecipeColor_6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_6.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_6_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_7.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_6_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_7.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_7_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_8.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_7_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_8.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_8_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_9.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_8_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_9.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_9_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_10.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_9_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_10.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_10_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_TempTime2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_10_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_TempTime2_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_R_Color_1.SetFocus
    End If
End Sub
Private Sub f_R_Color_1_Click()
    If Me.f_RecipeColor_6.ListIndex > -1 And Me.f_R_Color_1.ListIndex > -1 Then
        If Me.f_R_Color_1.ItemData(Me.f_R_Color_1.ListIndex) = Me.f_RecipeColor_6.ItemData(Me.f_RecipeColor_6.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_1_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_6.Text) * kg)
        Else
            Me.f_R_Color_1_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_R_Color_2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_1_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_R_Color_2_Click()
    If Me.f_RecipeColor_7.ListIndex > -1 And Me.f_R_Color_2.ListIndex > -1 Then
        If Me.f_R_Color_2.ItemData(Me.f_R_Color_2.ListIndex) = Me.f_RecipeColor_7.ItemData(Me.f_RecipeColor_7.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_2_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_7.Text) * kg)
        Else
            Me.f_R_Color_2_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_R_Color_3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_R_Color_3_Click()
    If Me.f_RecipeColor_8.ListIndex > -1 And Me.f_R_Color_3.ListIndex > -1 Then
        If Me.f_R_Color_3.ItemData(Me.f_R_Color_3.ListIndex) = Me.f_RecipeColor_8.ItemData(Me.f_RecipeColor_8.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_3_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_8.Text) * kg)
        Else
            Me.f_R_Color_3_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_R_Color_4.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_R_Color_4_Click()
    If Me.f_RecipeColor_9.ListIndex > -1 And Me.f_R_Color_4.ListIndex > -1 Then
        If Me.f_R_Color_4.ItemData(Me.f_R_Color_4.ListIndex) = Me.f_RecipeColor_9.ItemData(Me.f_RecipeColor_9.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_4_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_9.Text) * kg)
        Else
            Me.f_R_Color_4_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_R_Color_5.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_4_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_R_Color_5_Click()
    If Me.f_RecipeColor_10.ListIndex > -1 And Me.f_R_Color_5.ListIndex > -1 Then
        If Me.f_R_Color_5.ItemData(Me.f_R_Color_5.ListIndex) = Me.f_RecipeColor_10.ItemData(Me.f_RecipeColor_10.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_5_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_10.Text) * kg)
        Else
            Me.f_R_Color_5_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_11_Code.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_5_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_11_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_11_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_11_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_11_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_11_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_11_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_11_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_11_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_11_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_11_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_11_Code.ItemData(Me.f_Chemical_11_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_11_Code.ListIndex = -1 Then
            Me.f_Chemical_12_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_11_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_12_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_12_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_12_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_12_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_12_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_12_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_12_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_12_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_12_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_12_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_12_Code.ItemData(Me.f_Chemical_12_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_12_Code.ListIndex = -1 Then
            Me.f_Chemical_13_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_12_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_13_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_13_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_13_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_13_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_13_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_13_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_13_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_13_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_13_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_13_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_13_Code.ItemData(Me.f_Chemical_13_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_13_Code.ListIndex = -1 Then
            Me.f_Chemical_14_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_13_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_14_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_14_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_14_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_14_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_14_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_14_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_14_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_14_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_14_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_14_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_14_Code.ItemData(Me.f_Chemical_14_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_14_Code.ListIndex = -1 Then
            Me.f_Chemical_15_Code.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_14_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_15_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Chemical_15_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_15_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Chemical_15_Qty_GotFocus()
    If Len(Trim(Me.f_Chemical_15_Qty.Text)) > 0 Then
        PreQty = Me.f_Chemical_15_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Chemical_15_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Chemical_15_Code.ListIndex > -1 Then
            If Len(Trim(Me.f_Chemical_15_Qty.Text)) > 0 Then
                vQty = Me.f_Chemical_15_Qty
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Chemical_15_Code.ItemData(Me.f_Chemical_15_Code.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Chemical_15_Code.ListIndex = -1 Then
            Me.f_Temp3.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Chemical_15_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Temp3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_RecipeColor_11.SetFocus
    End If
End Sub
Private Sub f_RecipeColor_11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_11.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_11_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_12.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_11_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_12.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_12_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_13.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_12_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_13.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_13_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_14.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_13_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_14.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_14_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_15.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_14_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_RecipeColor_Qty_15.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_15_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_TempTime3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_RecipeColor_Qty_15_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_TempTime3_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
        Me.f_R_Color_6.SetFocus
    End If
End Sub
Private Sub f_R_Color_6_Click()
    If Me.f_RecipeColor_11.ListIndex > -1 And Me.f_R_Color_6.ListIndex > -1 Then
        If Me.f_R_Color_6.ItemData(Me.f_R_Color_6.ListIndex) = Me.f_RecipeColor_11.ItemData(Me.f_RecipeColor_11.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_6_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_11.Text) * kg)
        Else
            Me.f_R_Color_6_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_R_Color_7.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_6_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_R_Color_7_Click()
    If Me.f_RecipeColor_12.ListIndex > -1 And Me.f_R_Color_7.ListIndex > -1 Then
        If Me.f_R_Color_7.ItemData(Me.f_R_Color_7.ListIndex) = Me.f_RecipeColor_12.ItemData(Me.f_RecipeColor_12.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_7_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_12.Text) * kg)
        Else
            Me.f_R_Color_7_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_R_Color_8.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_7_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_R_Color_8_Click()
    If Me.f_RecipeColor_13.ListIndex > -1 And Me.f_R_Color_8.ListIndex > -1 Then
        If Me.f_R_Color_8.ItemData(Me.f_R_Color_8.ListIndex) = Me.f_RecipeColor_13.ItemData(Me.f_RecipeColor_13.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_8_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_13.Text) * kg)
        Else
            Me.f_R_Color_8_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_R_Color_9.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_8_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_R_Color_9_Click()
    If Me.f_RecipeColor_14.ListIndex > -1 And Me.f_R_Color_9.ListIndex > -1 Then
        If Me.f_R_Color_9.ItemData(Me.f_R_Color_9.ListIndex) = Me.f_RecipeColor_14.ItemData(Me.f_RecipeColor_14.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_9_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_14.Text) * kg)
        Else
            Me.f_R_Color_9_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_R_Color_10.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_9_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_R_Color_10_Click()
    If Me.f_RecipeColor_15.ListIndex > -1 And Me.f_R_Color_10.ListIndex > -1 Then
        If Me.f_R_Color_10.ItemData(Me.f_R_Color_10.ListIndex) = Me.f_RecipeColor_15.ItemData(Me.f_RecipeColor_15.ListIndex) Then
            If Len(Trim(Me.f_ConeKG.Text)) = 0 Then
                Me.f_ConeKG.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_2.Text)) = 0 Then
                Me.f_ConeKG_2.Text = 0
            End If
            If Len(Trim(Me.f_ConeKG_3.Text)) = 0 Then
                Me.f_ConeKG_3.Text = 0
            End If
            kg = Round(CDbl(Me.f_ConeKG.Text) + CDbl(Me.f_ConeKG_2.Text) + CDbl(Me.f_ConeKG_3.Text))
            Me.f_R_Color_10_Qty.Text = (CDbl(Me.f_RecipeColor_Qty_15.Text) * kg)
        Else
            Me.f_R_Color_10_Qty.Text = ""
        End If
    End If
End Sub
Private Sub f_R_Color_10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Soap2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_R_Color_10_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Soap2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Soap2_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Soap2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Soap2_Qty_GotFocus()
    If Len(Trim(Me.f_Soap2_Qty)) > 0 Then
        PreQty = Me.f_Soap2_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Soap2_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Soap2.ListIndex > -1 Then
            If Len(Trim(Me.f_Soap2_Qty.Text)) > 0 Then
                vQty = Me.f_Soap2_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Soap2.ItemData(Me.f_Soap2.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Soap2.ListIndex = -1 Then
            Me.f_Soap3.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Soap2_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Soap3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Soap3_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Soap3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Soap3_Qty_GotFocus()
    If Len(Trim(Me.f_Soap3_Qty)) > 0 Then
        PreQty = Me.f_Soap3_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Soap3_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Soap3.ListIndex > -1 Then
            If Len(Trim(Me.f_Soap3_Qty.Text)) > 0 Then
                vQty = Me.f_Soap3_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Soap3.ItemData(Me.f_Soap3.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Soap3.ListIndex = -1 Then
            Me.f_SoapTime2.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Soap3_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_SoapTime2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Hydro2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Hydro2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Hydro_Qty2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Hydro2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Hydro_Qty2_GotFocus()
    If Len(Trim(Me.f_Hydro_Qty2)) > 0 Then
        PreQty = Me.f_Hydro_Qty2.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Hydro_Qty2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Hydro2.ListIndex > -1 Then
            If Len(Trim(Me.f_Hydro_Qty2.Text)) > 0 Then
                vQty = Me.f_Hydro_Qty2.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Hydro2.ItemData(Me.f_Hydro2.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Hydro2.ListIndex = -1 Then
            Me.f_Castic2.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Hydro_Qty2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Castic2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Castic_Qty2.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Castic2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Castic_Qty2_GotFocus()
    If Len(Trim(Me.f_Castic_Qty2)) > 0 Then
        PreQty = Me.f_Castic_Qty2.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Castic_Qty2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Castic2.ListIndex > -1 Then
            If Len(Trim(Me.f_Castic_Qty2.Text)) > 0 Then
                vQty = Me.f_Castic_Qty2.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Castic2.ItemData(Me.f_Castic2.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Castic2.ListIndex = -1 Then
            Me.f_CasticTime2.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Castic_Qty2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_CasticTime2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Hydro3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Hydro3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Hydro_Qty3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Hydro3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Hydro_Qty3_GotFocus()
    If Len(Trim(Me.f_Hydro_Qty3)) > 0 Then
        PreQty = Me.f_Hydro_Qty3.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Hydro_Qty3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Hydro3.ListIndex > -1 Then
            If Len(Trim(Me.f_Hydro_Qty3.Text)) > 0 Then
                vQty = Me.f_Hydro_Qty3.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Hydro3.ItemData(Me.f_Hydro3.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Hydro3.ListIndex = -1 Then
            Me.f_Castic3.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Hydro_Qty3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Castic3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Castic_Qty3.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Castic3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Castic_Qty3_GotFocus()
    If Len(Trim(Me.f_Castic_Qty3)) > 0 Then
        PreQty = Me.f_Castic_Qty3.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Castic_Qty3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Castic3.ListIndex > -1 Then
            If Len(Trim(Me.f_Castic_Qty3.Text)) > 0 Then
                vQty = Me.f_Castic_Qty3.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Castic3.ItemData(Me.f_Castic3.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Castic3.ListIndex = -1 Then
            Me.f_CasticTime3.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Castic_Qty3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_CasticTime3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Acid.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Acid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Acid_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Acid_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Acid_Qty_GotFocus()
    If Len(Trim(Me.f_Acid_Qty)) > 0 Then
        PreQty = Me.f_Acid_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Acid_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Acid.ListIndex > -1 Then
            If Len(Trim(Me.f_Acid_Qty.Text)) > 0 Then
                vQty = Me.f_Acid_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Acid.ItemData(Me.f_Acid.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Acid.ListIndex = -1 Then
            Me.f_Acid2.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Acid_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Acid2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.f_Acid2_Qty.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub f_Acid2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Acid2_Qty_GotFocus()
    If Len(Trim(Me.f_Acid2_Qty)) > 0 Then
        PreQty = Me.f_Acid2_Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub f_Acid2_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBit = 0
        If Me.f_Acid2.ListIndex > -1 Then
            If Len(Trim(Me.f_Acid2.Text)) > 0 Then
                vQty = Me.f_Acid2_Qty.Text
            Else
                vQty = 0
            End If
            Call chkQty(Me.f_Acid2.ItemData(Me.f_Acid2.ListIndex), vQty)
        End If
        If MsgBit = 1 Or Me.f_Acid2.ListIndex = -1 Then
            Me.f_Remarks.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    Call EnableSave
End Sub
Private Sub f_Acid2_Qty_KeyUp(KeyCode As Integer, Shift As Integer)
    Call EnableSave
End Sub
Private Sub f_Remarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(Me.f_PartyCode)) > 0 And Len(Trim(f_MachineNo)) > 0 And Len(Trim(Me.f_ItemTypeCode)) > 0 And Len(Trim(f_Cone)) > 0 And Len(Trim(f_ConeKG)) > 0 Then
            Me.cmdSave.SetFocus
        End If
    End If
    Call EnableSave
End Sub
Public Sub AddNewRecord()
    m_ListID = ""
    m_AddMode = True
    Me.f_ProcessDate.value = Now
    Me.f_ProcessTime.value = Now
    Me.f_PartyCode.ListIndex = -1
    Me.f_PartyCode_2.ListIndex = -1
    Me.f_PartyCode_3.ListIndex = -1
    Me.f_MachineNo.Text = ""
    Me.f_Den_2.Text = ""
    Me.f_Den_3.Text = ""
    Me.f_SerialNo.Text = ""
    Me.f_Den.Text = ""
    Me.f_ItemTypeCode.ListIndex = -1
    Me.f_ItemTypeCode_2.ListIndex = -1
    Me.f_ItemTypeCode_3.ListIndex = -1
    Me.f_Cone.ListIndex = -1
    Me.f_Cone_2.ListIndex = -1
    Me.f_Cone_3.ListIndex = -1
    Me.f_ConeKG.Text = ""
    Me.f_ConeKG_2.Text = ""
    Me.f_ConeKG_3.Text = ""
    Me.f_Temp.Text = "134"
    Me.f_TempTime.Text = ""
    Me.f_Temp2.Text = ""
    Me.f_TempTime2.Text = ""
    Me.f_TempTime3.Text = ""
    Me.f_Chemical.ListIndex = -1
    Me.f_Chemical_Qty.Text = ""
    Me.f_Chemical2.ListIndex = -1
    Me.f_Chemical2_Qty.Text = ""
    Me.f_Chemical_3_Code.ListIndex = -1
    Me.f_Chemical_3_Qty.Text = ""
    Me.f_Chemical_4_Code.ListIndex = -1
    Me.f_Chemical_4_Qty.Text = ""
    Me.f_Chemical_5_Code.ListIndex = -1
    Me.f_Chemical_5_Qty.Text = ""
    Me.f_Chemical_6_Code.ListIndex = -1
    Me.f_Chemical_6_Qty.Text = ""
    Me.f_Chemical_7_Code.ListIndex = -1
    Me.f_Chemical_7_Qty.Text = ""
    Me.f_Chemical_8_Code.ListIndex = -1
    Me.f_Chemical_8_Qty.Text = ""
    Me.f_Chemical_9_Code.ListIndex = -1
    Me.f_Chemical_9_Qty.Text = ""
    Me.f_Chemical_10_Code.ListIndex = -1
    Me.f_Chemical_10_Qty.Text = ""
    Me.f_Chemical_11_Code.ListIndex = -1
    Me.f_Chemical_11_Qty.Text = ""
    Me.f_Chemical_12_Code.ListIndex = -1
    Me.f_Chemical_12_Qty.Text = ""
    Me.f_Chemical_13_Code.ListIndex = -1
    Me.f_Chemical_13_Qty.Text = ""
    Me.f_Chemical_14_Code.ListIndex = -1
    Me.f_Chemical_14_Qty.Text = ""
    Me.f_Chemical_15_Code.ListIndex = -1
    Me.f_Chemical_15_Qty.Text = ""
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
    Me.f_Soap2.ListIndex = -1
    Me.f_Soap2_Qty.Text = ""
    Me.f_Acid2.ListIndex = -1
    Me.f_Acid2_Qty.Text = ""
    Me.f_Soap3.ListIndex = -1
    Me.f_Soap3_Qty.Text = ""
    Me.f_SoapTime2.Text = ""
    Me.f_Castic2.ListIndex = -1
    Me.f_Castic_Qty2.Text = ""
    Me.f_Castic3.ListIndex = -1
    Me.f_Castic_Qty3.Text = ""
    Me.f_CasticTime2.Text = ""
    Me.f_CasticTime3.Text = ""
    Me.f_Hydro2.ListIndex = -1
    Me.f_Hydro_Qty2.Text = ""
    Me.f_Hydro3.ListIndex = -1
    Me.f_Hydro_Qty3.Text = ""
    Me.f_RecipeColor.ListIndex = -1
    Me.f_RecipeColor_Qty.Text = ""
    Me.f_RecipeColor_2.ListIndex = -1
    Me.f_RecipeColor_Qty_2.Text = ""
    Me.f_RecipeColor_3.ListIndex = -1
    Me.f_RecipeColor_Qty_3.Text = ""
    Me.f_RecipeColor_4.ListIndex = -1
    Me.f_RecipeColor_Qty_4.Text = ""
    Me.f_RecipeColor_5.ListIndex = -1
    Me.f_RecipeColor_Qty_5.Text = ""
    Me.f_RecipeColor_6.ListIndex = -1
    Me.f_RecipeColor_Qty_6.Text = ""
    Me.f_RecipeColor_7.ListIndex = -1
    Me.f_RecipeColor_Qty_7.Text = ""
    Me.f_RecipeColor_8.ListIndex = -1
    Me.f_RecipeColor_Qty_8.Text = ""
    Me.f_RecipeColor_9.ListIndex = -1
    Me.f_RecipeColor_Qty_9.Text = ""
    Me.f_RecipeColor_10.ListIndex = -1
    Me.f_RecipeColor_Qty_10.Text = ""
    Me.f_RecipeColor_11.ListIndex = -1
    Me.f_RecipeColor_Qty_11.Text = ""
    Me.f_RecipeColor_12.ListIndex = -1
    Me.f_RecipeColor_Qty_12.Text = ""
    Me.f_RecipeColor_13.ListIndex = -1
    Me.f_RecipeColor_Qty_13.Text = ""
    Me.f_RecipeColor_14.ListIndex = -1
    Me.f_RecipeColor_Qty_14.Text = ""
    Me.f_RecipeColor_15.ListIndex = -1
    Me.f_RecipeColor_Qty_15.Text = ""
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
    Me.f_R_Color_7.ListIndex = -1
    Me.f_R_Color_7_Qty.Text = ""
    Me.f_R_Color_8.ListIndex = -1
    Me.f_R_Color_8_Qty.Text = ""
    Me.f_R_Color_9.ListIndex = -1
    Me.f_R_Color_9_Qty.Text = ""
    Me.f_R_Color_10.ListIndex = -1
    Me.f_R_Color_10_Qty.Text = ""
    Me.f_Castic2.ListIndex = -1
    Me.f_Castic_Qty2.Text = ""
    Me.f_Castic3.ListIndex = -1
    Me.f_Castic_Qty3.Text = ""
End Sub
Private Sub EnableSave()
    If Len(Trim(Me.f_PartyCode)) > 0 And Len(Trim(f_MachineNo)) > 0 And Len(Trim(Me.f_ItemTypeCode)) > 0 And Len(Trim(f_Cone)) > 0 And Len(Trim(f_ConeKG)) > 0 Then
        Me.cmdSave.Enabled = True
    Else
        Me.cmdSave.Enabled = False
    End If
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
        Process.Left = 1000
        Process.Width = 14000
        Call AddNewRecord
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
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_2, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_3, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_4, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_5, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_6, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_7, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_8, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_9, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_10, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_11, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_12, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_13, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_14, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_15, "ItemName", "ItemCode"

  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_1, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_2, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_3, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_4, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_5, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_1, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_2, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_3, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_4, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_5, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_6, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_7, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_8, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_9, "ItemName", "ItemCode"
  'FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_10, "ItemName", "ItemCode"
  Me.f_ProcessDate.SetFocus
End Sub
Private Sub Cmdhide_Click()
        Process.Width = 10900
        Process.Left = 1700
        Me.SrItem.ListIndex = -1
        Me.SrItemType.ListIndex = -1
        Me.srParty.ListIndex = -1
        Me.SrItem.ListIndex = -1
        Call AddNewRecord
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
If Len(Trim(Me.f_PartyCode)) > 0 And Len(Trim(f_MachineNo)) > 0 And Len(Trim(Me.f_ItemTypeCode)) > 0 And Len(Trim(f_Cone)) > 0 And Len(Trim(f_ConeKG)) > 0 Then
            Call setVal
            MsgBox ("Record saved successfully"), vbInformation
            Me.f_ProcessDate.SetFocus
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
Private Sub Form_Load()
  m_AddMode = True
  cmdSave.Enabled = False
  DBConn
  f_ProcessDate = Now
  f_ProcessTime = Now
  SrDate = Now
  SrDate2 = Now
  ClickPane = 0
  
  FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", f_PartyCode, "PartyName", "PartyCode"
  FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", f_PartyCode_2, "PartyName", "PartyCode"
  FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", f_PartyCode_3, "PartyName", "PartyCode"
  
  FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", f_ItemTypeCode, "ItemTypeName", "ItemTypeCode"
  FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", f_ItemTypeCode_2, "ItemTypeName", "ItemTypeCode"
  FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", f_ItemTypeCode_3, "ItemTypeName", "ItemTypeCode"
  
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical2, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_3_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_4_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_5_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_6_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_7_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_8_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_9_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_10_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_11_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_12_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_13_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_14_Code, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 3 order by 2", f_Chemical_15_Code, "ItemName", "ItemCode"

  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_2, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_3, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_4, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_5, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_6, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_7, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_8, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_9, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_10, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_11, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_12, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_13, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_14, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_RecipeColor_15, "ItemName", "ItemCode"

  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_1, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_2, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_3, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_4, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_Color_5, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_1, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_2, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_3, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_4, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_5, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_6, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_7, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_8, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_9, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 4 order by 2", f_R_Color_10, "ItemName", "ItemCode"


  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 7 order by 2", f_Acid, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 7 order by 2", f_Acid2, "ItemName", "ItemCode"
  
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 10 order by 2", f_Soap, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 10 order by 2", f_Soap2, "ItemName", "ItemCode"
  FillColorCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 10 order by 2", f_Soap3, "ItemName", "ItemCode"
  
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 8 order by 2", f_Hydro, "ItemName", "ItemCode"
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 8 order by 2", f_Hydro2, "ItemName", "ItemCode"
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 8 order by 2", f_Hydro3, "ItemName", "ItemCode"
  
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 9 order by 2", f_Castic, "ItemName", "ItemCode"
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 9 order by 2", f_Castic2, "ItemName", "ItemCode"
  FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = 9 order by 2", f_Castic3, "ItemName", "ItemCode"
  
  
  FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", SrItemType, "ItemTypeName", "ItemTypeCode"
  FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", srParty, "PartyName", "PartyCode"
  
  lvwphase.ColumnHeaders.Add Text:="Code", Width:=600
  lvwphase.ColumnHeaders.Add Text:="Date", Width:=1100
  lvwphase.ColumnHeaders.Add Text:="Serial", Width:=700
  lvwphase.ColumnHeaders.Add Text:="Party Name", Width:=1700
  lvwphase.ColumnHeaders.Add Text:="M/C #", Width:=650
  lvwphase.ColumnHeaders.Add Text:="Item Type", Width:=2200
  lvwphase.ColumnHeaders.Add Text:="Item", Width:=2000
  lvwphase.ColumnHeaders.Add Text:="Recipe", Width:=800
  lvwphase.ColumnHeaders.Add Text:="Re Recipe", Width:=1000
  
  Call fillList
End Sub
Public Sub setVal()
Dim rstSave As New ADODB.Recordset
    If m_AddMode = True Then
    'If (Len(Trim(m_ListID)) = 0) Then
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
    If Len(Trim(Me.f_SerialNo.Text)) > 0 Then
        rstSave("SerialNo") = Me.f_SerialNo.Text
    Else
        rstSave("SerialNo") = 0
    End If
    If Len(Trim(Me.f_Den.Text)) > 0 Then
        rstSave("Den") = Me.f_Den.Text
    Else
        rstSave("Den") = 0
    End If
    If Me.f_ItemTypeCode.ListIndex > -1 Then
        rstSave("ItemTypeCode") = Me.f_ItemTypeCode.ItemData(Me.f_ItemTypeCode.ListIndex)
        rstSave("Cone") = Me.f_Cone.ItemData(Me.f_Cone.ListIndex)
        rstSave("ConeKG") = Me.f_ConeKG.Text
    End If
    rstSave("Temp") = Me.f_Temp.Text
    rstSave("TempTime") = Me.f_TempTime.Text
    If Me.f_Chemical.ListIndex > -1 Then
        rstSave("Chemical") = Me.f_Chemical.ItemData(Me.f_Chemical.ListIndex)
        rstSave("Chemical_Qty") = IIf(IsNull(Me.f_Chemical_Qty.Text), 0, Me.f_Chemical_Qty.Text)
        
    End If
    If Me.f_Chemical2.ListIndex > -1 And Me.f_Chemical2 <> "-- Select --" Then
        rstSave("Chemical2") = Me.f_Chemical2.ItemData(Me.f_Chemical2.ListIndex)
        rstSave("Chemical2_Qty") = IIf(IsNull(Me.f_Chemical2_Qty.Text), 0, Me.f_Chemical2_Qty.Text)
    End If
    
    If Me.f_Chemical_3_Code.ListIndex > -1 Then
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

    If Me.f_Chemical_4_Code.ListIndex > -1 Then
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
    
    If Me.f_Acid.ListIndex > -1 Then
        rstSave("Acid") = Me.f_Acid.ItemData(Me.f_Acid.ListIndex)
        rstSave("Acid_Qty") = IIf(IsNull(Me.f_Acid_Qty.Text), 0, Me.f_Acid_Qty.Text)
    End If
    If Me.f_Acid2.ListIndex > -1 And Me.f_Acid2 <> "-- Select --" Then
        rstSave("Acid2") = Me.f_Acid2.ItemData(Me.f_Acid2.ListIndex)
        rstSave("Acid2_Qty") = IIf(IsNull(Me.f_Acid2_Qty.Text), 0, Me.f_Acid2_Qty.Text)
    End If
    If Me.f_Soap.ListIndex > -1 Then
        rstSave("Soap") = Me.f_Soap.ItemData(Me.f_Soap.ListIndex)
        rstSave("Soap_Qty") = IIf(IsNull(Me.f_Soap_Qty.Text), 0, Me.f_Soap_Qty.Text)
        rstSave("SoapTime") = Me.f_SoapTime.Text
    End If
    If Me.f_Soap2.ListIndex > -1 And Me.f_Soap2 <> "-- Select --" Then
        rstSave("Soap2") = Me.f_Soap2.ItemData(Me.f_Soap2.ListIndex)
        rstSave("Soap2_Qty") = IIf(IsNull(Me.f_Soap2_Qty.Text), 0, Me.f_Soap2_Qty.Text)
    End If
    If Me.f_Hydro.ListIndex > -1 Then
        rstSave("Hydro") = Me.f_Hydro.ItemData(Me.f_Hydro.ListIndex)
        rstSave("Hydro_Qty") = IIf(IsNull(Me.f_Hydro_Qty.Text), 0, Me.f_Hydro_Qty.Text)
    End If
    If Me.f_Castic.ListIndex > -1 Then
        rstSave("Castic") = Me.f_Castic.ItemData(Me.f_Castic.ListIndex)
        rstSave("Castic_Qty") = IIf(IsNull(Me.f_Castic_Qty.Text), 0, Me.f_Castic_Qty.Text)
        rstSave("CasticTime") = Me.f_CasticTime.Text
    End If
    If Len(Trim(Me.f_RecipeCode)) > 0 Then
        rstSave("RecipeCode") = Me.f_RecipeCode.Text
    End If
    If Me.f_Re_RecipeCode.value = Checked Then
        rstSave("Re_RecipeCode") = 1
    Else
        rstSave("Re_RecipeCode") = 0
    End If
    If Len(Trim(Me.f_Remarks)) > 0 Then
        rstSave("Remarks") = Me.f_Remarks.Text
    End If
    
    If Me.f_Color_1.ListIndex > -1 And Me.f_Color_1 <> "-- Select --" Then
        rstSave("Color_1") = Me.f_Color_1.ItemData(Me.f_Color_1.ListIndex)
        rstSave("Color_1_Qty") = IIf(IsNull(Me.f_Color_1_Qty.Text), 0, Me.f_Color_1_Qty.Text)
    End If
    If Me.f_Color_2.ListIndex > -1 And Me.f_Color_2 <> "-- Select --" Then
        rstSave("Color_2") = Me.f_Color_2.ItemData(Me.f_Color_2.ListIndex)
        rstSave("Color_2_Qty") = IIf(IsNull(Me.f_Color_2_Qty.Text), 0, Me.f_Color_2_Qty.Text)
    End If
    If Me.f_Color_3.ListIndex > -1 And Me.f_Color_3 <> "-- Select --" Then
        rstSave("Color_3") = Me.f_Color_3.ItemData(Me.f_Color_3.ListIndex)
        rstSave("Color_3_Qty") = IIf(IsNull(Me.f_Color_3_Qty.Text), 0, Me.f_Color_3_Qty.Text)
    End If
    If Me.f_Color_4.ListIndex > -1 And Me.f_Color_4 <> "-- Select --" Then
        rstSave("Color_4") = Me.f_Color_4.ItemData(Me.f_Color_4.ListIndex)
        rstSave("Color_4_Qty") = IIf(IsNull(Me.f_Color_4_Qty.Text), 0, Me.f_Color_4_Qty.Text)
    End If
    If Me.f_Color_5.ListIndex > -1 And Me.f_Color_5 <> "-- Select --" Then
        rstSave("Color_5") = Me.f_Color_5.ItemData(Me.f_Color_5.ListIndex)
        rstSave("Color_5_Qty") = IIf(IsNull(Me.f_Color_5_Qty.Text), 0, Me.f_Color_5_Qty.Text)
    End If
    If Me.f_R_Color_1.ListIndex > -1 And Me.f_R_Color_1 <> "-- Select --" Then
        rstSave("R_Color_1") = Me.f_R_Color_1.ItemData(Me.f_R_Color_1.ListIndex)
        rstSave("R_Color_1_Qty") = IIf(IsNull(Me.f_R_Color_1_Qty.Text), 0, Me.f_R_Color_1_Qty.Text)
    End If
    If Me.f_R_Color_2.ListIndex > -1 And Me.f_R_Color_2 <> "-- Select --" Then
        rstSave("R_Color_2") = Me.f_R_Color_2.ItemData(Me.f_R_Color_2.ListIndex)
        rstSave("R_Color_2_Qty") = IIf(IsNull(Me.f_R_Color_2_Qty.Text), 0, Me.f_R_Color_2_Qty.Text)
    End If
    If Me.f_R_Color_3.ListIndex > -1 And Me.f_R_Color_3 <> "-- Select --" Then
        rstSave("R_Color_3") = Me.f_R_Color_3.ItemData(Me.f_R_Color_3.ListIndex)
        rstSave("R_Color_3_Qty") = IIf(IsNull(Me.f_R_Color_3_Qty.Text), 0, Me.f_R_Color_3_Qty.Text)
    End If
    If Me.f_R_Color_4.ListIndex > -1 And Me.f_R_Color_4 <> "-- Select --" Then
        rstSave("R_Color_4") = Me.f_R_Color_4.ItemData(Me.f_R_Color_4.ListIndex)
        rstSave("R_Color_4_Qty") = IIf(IsNull(Me.f_R_Color_4_Qty.Text), 0, Me.f_R_Color_4_Qty.Text)
    End If
    If Me.f_R_Color_5.ListIndex > -1 And Me.f_R_Color_5 <> "-- Select --" Then
        rstSave("R_Color_5") = Me.f_R_Color_5.ItemData(Me.f_R_Color_5.ListIndex)
        rstSave("R_Color_5_Qty") = IIf(IsNull(Me.f_R_Color_5_Qty.Text), 0, Me.f_R_Color_5_Qty.Text)
    End If
    If Me.f_R_Color_6.ListIndex > -1 And Me.f_R_Color_6 <> "-- Select --" Then
        rstSave("R_Color_6") = Me.f_R_Color_6.ItemData(Me.f_R_Color_6.ListIndex)
        rstSave("R_Color_6_Qty") = IIf(IsNull(Me.f_R_Color_6_Qty.Text), 0, Me.f_R_Color_6_Qty.Text)
    End If
    rstSave("NewColor") = Me.f_NewColor.Text
    rstSave("Is_Cotton_Dyeing") = 0
    
    If Me.f_PartyCode_2.ListIndex > -1 Then
        rstSave("PartyCode2") = Me.f_PartyCode_2.ItemData(Me.f_PartyCode_2.ListIndex)
    End If
    If Me.f_PartyCode_3.ListIndex > -1 Then
        rstSave("PartyCode3") = Me.f_PartyCode_3.ItemData(Me.f_PartyCode_3.ListIndex)
    End If
    If Me.f_ItemTypeCode_2.ListIndex > -1 Then
        rstSave("ItemTypeCode2") = Me.f_ItemTypeCode_2.ItemData(Me.f_ItemTypeCode_2.ListIndex)
        rstSave("Cone2") = Me.f_Cone_2.ItemData(Me.f_Cone_2.ListIndex)
        rstSave("ConeKG2") = Me.f_ConeKG_2.Text
    End If
    If Me.f_ItemTypeCode_3.ListIndex > -1 Then
        rstSave("ItemTypeCode3") = Me.f_ItemTypeCode_3.ItemData(Me.f_ItemTypeCode_3.ListIndex)
        rstSave("Cone3") = Me.f_Cone_3.ItemData(Me.f_Cone_3.ListIndex)
        rstSave("ConeKG3") = Me.f_ConeKG_3.Text
    End If
    If Len(Trim(Me.f_Den_2.Text)) > 0 Then
        rstSave("Den2") = Me.f_Den_2.Text
    Else
        rstSave("Den2") = 0
    End If
    If Len(Trim(Me.f_Den_3.Text)) > 0 Then
        rstSave("Den3") = Me.f_Den_3.Text
    Else
        rstSave("Den3") = 0
    End If
    If Me.f_Chemical_5_Code.ListIndex > -1 Then
        rstSave("Chemical_5_Code") = Me.f_Chemical_5_Code.ItemData(Me.f_Chemical_5_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_5_Qty.Text)) > 0 Then
            rstSave("Chemical_5_Qty") = Me.f_Chemical_5_Qty.Text
        Else
            rstSave("Chemical_5_Qty") = 0
        End If
    Else
        rstSave("Chemical_5_Code") = 0
        rstSave("Chemical_5_Qty") = 0
    End If
    If Me.f_Chemical_6_Code.ListIndex > -1 Then
        rstSave("Chemical_6_Code") = Me.f_Chemical_6_Code.ItemData(Me.f_Chemical_6_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_6_Qty.Text)) > 0 Then
            rstSave("Chemical_6_Qty") = Me.f_Chemical_6_Qty.Text
        Else
            rstSave("Chemical_6_Qty") = 0
        End If
    Else
        rstSave("Chemical_6_Code") = 0
        rstSave("Chemical_6_Qty") = 0
    End If
    If Me.f_Chemical_7_Code.ListIndex > -1 Then
        rstSave("Chemical_7_Code") = Me.f_Chemical_7_Code.ItemData(Me.f_Chemical_7_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_7_Qty.Text)) > 0 Then
            rstSave("Chemical_7_Qty") = Me.f_Chemical_7_Qty.Text
        Else
            rstSave("Chemical_7_Qty") = 0
        End If
    Else
        rstSave("Chemical_7_Code") = 0
        rstSave("Chemical_7_Qty") = 0
    End If
    If Me.f_Chemical_8_Code.ListIndex > -1 Then
        rstSave("Chemical_8_Code") = Me.f_Chemical_8_Code.ItemData(Me.f_Chemical_8_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_8_Qty.Text)) > 0 Then
            rstSave("Chemical_8_Qty") = Me.f_Chemical_8_Qty.Text
        Else
            rstSave("Chemical_8_Qty") = 0
        End If
    Else
        rstSave("Chemical_8_Code") = 0
        rstSave("Chemical_8_Qty") = 0
    End If
    If Me.f_Chemical_9_Code.ListIndex > -1 Then
        rstSave("Chemical_9_Code") = Me.f_Chemical_9_Code.ItemData(Me.f_Chemical_9_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_9_Qty.Text)) > 0 Then
            rstSave("Chemical_9_Qty") = Me.f_Chemical_9_Qty.Text
        Else
            rstSave("Chemical_9_Qty") = 0
        End If
    Else
        rstSave("Chemical_9_Code") = 0
        rstSave("Chemical_9_Qty") = 0
    End If
    If Me.f_Chemical_10_Code.ListIndex > -1 Then
        rstSave("Chemical_10_Code") = Me.f_Chemical_10_Code.ItemData(Me.f_Chemical_10_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_10_Qty.Text)) > 0 Then
            rstSave("Chemical_10_Qty") = Me.f_Chemical_10_Qty.Text
        Else
            rstSave("Chemical_10_Qty") = 0
        End If
    Else
        rstSave("Chemical_10_Code") = 0
        rstSave("Chemical_10_Qty") = 0
    End If
    If Me.f_Chemical_11_Code.ListIndex > -1 Then
        rstSave("Chemical_11_Code") = Me.f_Chemical_11_Code.ItemData(Me.f_Chemical_11_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_11_Qty.Text)) > 0 Then
            rstSave("Chemical_11_Qty") = Me.f_Chemical_11_Qty.Text
        Else
            rstSave("Chemical_11_Qty") = 0
        End If
    Else
        rstSave("Chemical_11_Code") = 0
        rstSave("Chemical_11_Qty") = 0
    End If
    If Me.f_Chemical_12_Code.ListIndex > -1 Then
        rstSave("Chemical_12_Code") = Me.f_Chemical_12_Code.ItemData(Me.f_Chemical_12_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_12_Qty.Text)) > 0 Then
            rstSave("Chemical_12_Qty") = Me.f_Chemical_12_Qty.Text
        Else
            rstSave("Chemical_12_Qty") = 0
        End If
    Else
        rstSave("Chemical_12_Code") = 0
        rstSave("Chemical_12_Qty") = 0
    End If
    If Me.f_Chemical_13_Code.ListIndex > -1 Then
        rstSave("Chemical_13_Code") = Me.f_Chemical_13_Code.ItemData(Me.f_Chemical_13_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_13_Qty.Text)) > 0 Then
            rstSave("Chemical_13_Qty") = Me.f_Chemical_13_Qty.Text
        Else
            rstSave("Chemical_13_Qty") = 0
        End If
    Else
        rstSave("Chemical_13_Code") = 0
        rstSave("Chemical_13_Qty") = 0
    End If
    If Me.f_Chemical_14_Code.ListIndex > -1 Then
        rstSave("Chemical_14_Code") = Me.f_Chemical_14_Code.ItemData(Me.f_Chemical_14_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_14_Qty.Text)) > 0 Then
            rstSave("Chemical_14_Qty") = Me.f_Chemical_14_Qty.Text
        Else
            rstSave("Chemical_14_Qty") = 0
        End If
    Else
        rstSave("Chemical_14_Code") = 0
        rstSave("Chemical_14_Qty") = 0
    End If
    If Me.f_Chemical_15_Code.ListIndex > -1 Then
        rstSave("Chemical_15_Code") = Me.f_Chemical_15_Code.ItemData(Me.f_Chemical_15_Code.ListIndex)
        If Len(Trim(Me.f_Chemical_15_Qty.Text)) > 0 Then
            rstSave("Chemical_15_Qty") = Me.f_Chemical_15_Qty.Text
        Else
            rstSave("Chemical_15_Qty") = 0
        End If
    Else
        rstSave("Chemical_15_Code") = 0
        rstSave("Chemical_15_Qty") = 0
    End If
    If Me.f_R_Color_7.ListIndex > -1 And Me.f_R_Color_7 <> "-- Select --" Then
        rstSave("R_Color_7") = Me.f_R_Color_7.ItemData(Me.f_R_Color_7.ListIndex)
        rstSave("R_Color_7_Qty") = IIf(IsNull(Me.f_R_Color_7_Qty.Text), 0, Me.f_R_Color_7_Qty.Text)
    End If
    If Me.f_R_Color_8.ListIndex > -1 And Me.f_R_Color_8 <> "-- Select --" Then
        rstSave("R_Color_8") = Me.f_R_Color_8.ItemData(Me.f_R_Color_8.ListIndex)
        rstSave("R_Color_8_Qty") = IIf(IsNull(Me.f_R_Color_8_Qty.Text), 0, Me.f_R_Color_8_Qty.Text)
    End If
    If Me.f_R_Color_9.ListIndex > -1 And Me.f_R_Color_9 <> "-- Select --" Then
        rstSave("R_Color_9") = Me.f_R_Color_9.ItemData(Me.f_R_Color_9.ListIndex)
        rstSave("R_Color_9_Qty") = IIf(IsNull(Me.f_R_Color_9_Qty.Text), 0, Me.f_R_Color_9_Qty.Text)
    End If
    If Me.f_R_Color_10.ListIndex > -1 And Me.f_R_Color_10 <> "-- Select --" Then
        rstSave("R_Color_10") = Me.f_R_Color_10.ItemData(Me.f_R_Color_10.ListIndex)
        rstSave("R_Color_10_Qty") = IIf(IsNull(Me.f_R_Color_10_Qty.Text), 0, Me.f_R_Color_10_Qty.Text)
    End If
    If Me.f_Soap3.ListIndex > -1 Then
        rstSave("Soap3") = Me.f_Soap3.ItemData(Me.f_Soap3.ListIndex)
        rstSave("Soap3_Qty") = IIf(IsNull(Me.f_Soap3_Qty.Text), 0, Me.f_Soap3_Qty.Text)
        rstSave("Soap_Temp") = Me.f_SoapTime2.Text
    End If
    If Me.f_Castic2.ListIndex > -1 Then
        rstSave("Castic2") = Me.f_Castic2.ItemData(Me.f_Castic2.ListIndex)
        rstSave("Castic2_Qty") = IIf(IsNull(Me.f_Castic_Qty2.Text), 0, Me.f_Castic_Qty2.Text)
        rstSave("CasticTime2") = Me.f_CasticTime2.Text
    End If
    If Me.f_Castic3.ListIndex > -1 Then
        rstSave("Castic3") = Me.f_Castic3.ItemData(Me.f_Castic3.ListIndex)
        rstSave("Castic3_Qty") = IIf(IsNull(Me.f_Castic_Qty3.Text), 0, Me.f_Castic_Qty3.Text)
        rstSave("CasticTime3") = Me.f_CasticTime3.Text
    End If
    If Me.f_Hydro2.ListIndex > -1 Then
        rstSave("Hydro2") = Me.f_Hydro2.ItemData(Me.f_Hydro2.ListIndex)
        rstSave("Hydro2_Qty") = IIf(IsNull(Me.f_Hydro_Qty2.Text), 0, Me.f_Hydro_Qty2.Text)
    End If
    If Me.f_Hydro3.ListIndex > -1 Then
        rstSave("Hydro3") = Me.f_Hydro3.ItemData(Me.f_Hydro3.ListIndex)
        rstSave("Hydro3_Qty") = IIf(IsNull(Me.f_Hydro_Qty3.Text), 0, Me.f_Hydro_Qty3.Text)
    End If
    rstSave("Temp2") = Me.f_Temp2.Text
    rstSave("TempTime2") = Me.f_TempTime2.Text
    rstSave("Temp3") = Me.f_Temp3.Text
    rstSave("TempTime3") = Me.f_TempTime3.Text
    If Me.f_RecipeColor.ListIndex > -1 And Me.f_RecipeColor <> "-- Select --" Then
        rstSave("RecipeColor1") = Me.f_RecipeColor.ItemData(Me.f_RecipeColor.ListIndex)
        rstSave("RecipeColor1_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty.Text), 0, Me.f_RecipeColor_Qty.Text)
    End If
    If Me.f_RecipeColor_2.ListIndex > -1 And Me.f_RecipeColor_2 <> "-- Select --" Then
        rstSave("RecipeColor2") = Me.f_RecipeColor_2.ItemData(Me.f_RecipeColor_2.ListIndex)
        rstSave("RecipeColor2_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_2.Text), 0, Me.f_RecipeColor_Qty_2.Text)
    End If
    If Me.f_RecipeColor_3.ListIndex > -1 And Me.f_RecipeColor_3 <> "-- Select --" Then
        rstSave("RecipeColor3") = Me.f_RecipeColor_3.ItemData(Me.f_RecipeColor_3.ListIndex)
        rstSave("RecipeColor3_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_3.Text), 0, Me.f_RecipeColor_Qty_3.Text)
    End If
    If Me.f_RecipeColor_4.ListIndex > -1 And Me.f_RecipeColor_4 <> "-- Select --" Then
        rstSave("RecipeColor4") = Me.f_RecipeColor_4.ItemData(Me.f_RecipeColor_4.ListIndex)
        rstSave("RecipeColor4_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_4.Text), 0, Me.f_RecipeColor_Qty_4.Text)
    End If
    If Me.f_RecipeColor_5.ListIndex > -1 And Me.f_RecipeColor_5 <> "-- Select --" Then
        rstSave("RecipeColor5") = Me.f_RecipeColor_5.ItemData(Me.f_RecipeColor_5.ListIndex)
        rstSave("RecipeColor5_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_5.Text), 0, Me.f_RecipeColor_Qty_5.Text)
    End If
    If Me.f_RecipeColor_6.ListIndex > -1 And Me.f_RecipeColor_6 <> "-- Select --" Then
        rstSave("RecipeColor6") = Me.f_RecipeColor_6.ItemData(Me.f_RecipeColor_6.ListIndex)
        rstSave("RecipeColor6_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_6.Text), 0, Me.f_RecipeColor_Qty_6.Text)
    End If
    If Me.f_RecipeColor_7.ListIndex > -1 And Me.f_RecipeColor_7 <> "-- Select --" Then
        rstSave("RecipeColor7") = Me.f_RecipeColor_7.ItemData(Me.f_RecipeColor_7.ListIndex)
        rstSave("RecipeColor7_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_7.Text), 0, Me.f_RecipeColor_Qty_7.Text)
    End If
    If Me.f_RecipeColor_8.ListIndex > -1 And Me.f_RecipeColor_8 <> "-- Select --" Then
        rstSave("RecipeColor8") = Me.f_RecipeColor_8.ItemData(Me.f_RecipeColor_8.ListIndex)
        rstSave("RecipeColor8_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_8.Text), 0, Me.f_RecipeColor_Qty_8.Text)
    End If
    If Me.f_RecipeColor_9.ListIndex > -1 And Me.f_RecipeColor_9 <> "-- Select --" Then
        rstSave("RecipeColor9") = Me.f_RecipeColor_9.ItemData(Me.f_RecipeColor_9.ListIndex)
        rstSave("RecipeColor9_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_9.Text), 0, Me.f_RecipeColor_Qty_9.Text)
    End If
    If Me.f_RecipeColor_10.ListIndex > -1 And Me.f_RecipeColor_10 <> "-- Select --" Then
        rstSave("RecipeColor10") = Me.f_RecipeColor_10.ItemData(Me.f_RecipeColor_10.ListIndex)
        rstSave("RecipeColor10_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_10.Text), 0, Me.f_RecipeColor_Qty_10.Text)
    End If
    If Me.f_RecipeColor_11.ListIndex > -1 And Me.f_RecipeColor_11 <> "-- Select --" Then
        rstSave("RecipeColor11") = Me.f_RecipeColor_11.ItemData(Me.f_RecipeColor_11.ListIndex)
        rstSave("RecipeColor11_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_11.Text), 0, Me.f_RecipeColor_Qty_11.Text)
    End If
    If Me.f_RecipeColor_12.ListIndex > -1 And Me.f_RecipeColor_12 <> "-- Select --" Then
        rstSave("RecipeColor12") = Me.f_RecipeColor_12.ItemData(Me.f_RecipeColor_12.ListIndex)
        rstSave("RecipeColor12_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_12.Text), 0, Me.f_RecipeColor_Qty_12.Text)
    End If
    If Me.f_RecipeColor_13.ListIndex > -1 And Me.f_RecipeColor_13 <> "-- Select --" Then
        rstSave("RecipeColor13") = Me.f_RecipeColor_13.ItemData(Me.f_RecipeColor_13.ListIndex)
        rstSave("RecipeColor13_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_13.Text), 0, Me.f_RecipeColor_Qty_13.Text)
    End If
    If Me.f_RecipeColor_14.ListIndex > -1 And Me.f_RecipeColor_14 <> "-- Select --" Then
        rstSave("RecipeColor14") = Me.f_RecipeColor_14.ItemData(Me.f_RecipeColor_14.ListIndex)
        rstSave("RecipeColor14_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_14.Text), 0, Me.f_RecipeColor_Qty_14.Text)
    End If
    If Me.f_RecipeColor_15.ListIndex > -1 And Me.f_RecipeColor_15 <> "-- Select --" Then
        rstSave("RecipeColor15") = Me.f_RecipeColor_15.ItemData(Me.f_RecipeColor_15.ListIndex)
        rstSave("RecipeColor15_Qty") = IIf(IsNull(Me.f_RecipeColor_Qty_15.Text), 0, Me.f_RecipeColor_Qty_15.Text)
    End If
    rstSave("Is_Active") = 1
rstSave.Update
rstSave.Close
Set rstSave = Nothing
m_AddMode = False
Call fillList
End Sub
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    
    Set rstList = FillRecordSet("SELECT top 200 ProcessCode, convert(varchar, ProcessDate, 103) as ProcessDate, isNull(SerialNo, 0) as SerialNo, PartyName, MachineNo, ItemTypeName, (Select ItemName from Item where ItemCode = Cone) as Cone, isNull(RecipeCode, 0) as RecipeCode, isNull(Re_RecipeCode, 0) as Re_RecipeCode  " & _
                                "FROM Party INNER JOIN (ItemType INNER JOIN Process ON ItemType.ItemTypeCode = Process.ItemTypeCode) ON Party.PartyCode = Process.PartyCode where Is_Active = 1 and Is_Cotton_Dyeing = 0 order by ProcessCode desc")
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!ProcessCode, _
                   Key:=CStr("Id=" & rstList!ProcessCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!ProcessDate
                 .Add Text:=rstList!SerialNo
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!MachineNo
                 .Add Text:=rstList!ItemTypeName
                 .Add Text:=rstList!Cone
                 .Add Text:=rstList!RecipeCode
                 .Add Text:=rstList!Re_RecipeCode
                 .Add Text:=rstList!ProcessCode
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
        Me.f_ProcessDate.value = IIf(IsNull(rstGetVal("ProcessDate")), Now, rstGetVal("ProcessDate"))
        Me.f_ProcessTime.value = IIf(IsNull(rstGetVal("ProcessTime")), Now, rstGetVal("ProcessTime"))
        Me.f_MachineNo.Text = rstGetVal("MachineNo")
        Me.f_SerialNo.Text = IIf(IsNull(rstGetVal("SerialNo")), 0, rstGetVal("SerialNo"))
        Me.f_NewColor.Text = IIf(IsNull(rstGetVal("NewColor")), " ", rstGetVal("NewColor"))
        Call selectValueInCombo(Me.f_PartyCode, rstGetVal("PartyCode"))
        Call selectValueInCombo(Me.f_PartyCode_2, IIf(IsNull(rstGetVal("PartyCode2")), -1, rstGetVal("PartyCode2")))
        Call selectValueInCombo(Me.f_PartyCode_3, IIf(IsNull(rstGetVal("PartyCode3")), -1, rstGetVal("PartyCode3")))
        Call selectValueInCombo(Me.f_ItemTypeCode, rstGetVal("ItemTypeCode"))
        Call selectValueInCombo(Me.f_ItemTypeCode_2, IIf(IsNull(rstGetVal("ItemTypeCode2")), -1, rstGetVal("ItemTypeCode2")))
        Call selectValueInCombo(Me.f_ItemTypeCode_3, IIf(IsNull(rstGetVal("ItemTypeCode3")), -1, rstGetVal("ItemTypeCode3")))
        Call selectValueInCombo(Me.f_Cone, rstGetVal("Cone"))
        Call selectValueInCombo(Me.f_Cone_2, IIf(IsNull(rstGetVal("Cone2")), -1, rstGetVal("Cone2")))
        Call selectValueInCombo(Me.f_Cone_3, IIf(IsNull(rstGetVal("Cone3")), -1, rstGetVal("Cone3")))
        Me.f_ConeKG.Text = rstGetVal("ConeKG")
        Me.f_ConeKG_2.Text = IIf(IsNull(rstGetVal("ConeKG2")), 0, rstGetVal("ConeKG2"))
        Me.f_ConeKG_3.Text = IIf(IsNull(rstGetVal("ConeKG3")), 0, rstGetVal("ConeKG3"))
        Me.f_Den.Text = rstGetVal("Den")
        Me.f_Den_2.Text = IIf(IsNull(rstGetVal("Den2")), 0, rstGetVal("Den2"))
        Me.f_Den_3.Text = IIf(IsNull(rstGetVal("Den3")), 0, rstGetVal("Den3"))
        Me.f_Temp.Text = rstGetVal("Temp")
        Me.f_TempTime.Text = rstGetVal("TempTime")
        Me.f_Temp2.Text = IIf(IsNull(rstGetVal("Temp2")), "", rstGetVal("Temp2"))
        Me.f_TempTime2.Text = IIf(IsNull(rstGetVal("TempTime2")), "", rstGetVal("TempTime2"))
        Me.f_Temp3.Text = IIf(IsNull(rstGetVal("Temp3")), "", rstGetVal("Temp3"))
        Me.f_TempTime3.Text = IIf(IsNull(rstGetVal("TempTime3")), "", rstGetVal("TempTime3"))
        Call selectValueInCombo(Me.f_Chemical, IIf(IsNull(rstGetVal("Chemical")), -1, rstGetVal("Chemical")))
        Call selectValueInCombo(Me.f_Chemical2, IIf(IsNull(rstGetVal("Chemical2")), -1, rstGetVal("Chemical2")))
        Call selectValueInCombo(Me.f_Chemical_3_Code, IIf(IsNull(rstGetVal("Chemical_3_Code")), -1, rstGetVal("Chemical_3_Code")))
        Call selectValueInCombo(Me.f_Chemical_4_Code, IIf(IsNull(rstGetVal("Chemical_4_Code")), -1, rstGetVal("Chemical_4_Code")))
        Call selectValueInCombo(Me.f_Chemical_5_Code, IIf(IsNull(rstGetVal("Chemical_5_Code")), -1, rstGetVal("Chemical_5_Code")))
        Call selectValueInCombo(Me.f_Chemical_6_Code, IIf(IsNull(rstGetVal("Chemical_6_Code")), -1, rstGetVal("Chemical_6_Code")))
        Call selectValueInCombo(Me.f_Chemical_7_Code, IIf(IsNull(rstGetVal("Chemical_7_Code")), -1, rstGetVal("Chemical_7_Code")))
        Call selectValueInCombo(Me.f_Chemical_8_Code, IIf(IsNull(rstGetVal("Chemical_8_Code")), -1, rstGetVal("Chemical_8_Code")))
        Call selectValueInCombo(Me.f_Chemical_9_Code, IIf(IsNull(rstGetVal("Chemical_9_Code")), -1, rstGetVal("Chemical_9_Code")))
        Call selectValueInCombo(Me.f_Chemical_10_Code, IIf(IsNull(rstGetVal("Chemical_10_Code")), -1, rstGetVal("Chemical_10_Code")))
        Call selectValueInCombo(Me.f_Chemical_11_Code, IIf(IsNull(rstGetVal("Chemical_11_Code")), -1, rstGetVal("Chemical_11_Code")))
        Call selectValueInCombo(Me.f_Chemical_12_Code, IIf(IsNull(rstGetVal("Chemical_12_Code")), -1, rstGetVal("Chemical_12_Code")))
        Call selectValueInCombo(Me.f_Chemical_13_Code, IIf(IsNull(rstGetVal("Chemical_13_Code")), -1, rstGetVal("Chemical_13_Code")))
        Call selectValueInCombo(Me.f_Chemical_14_Code, IIf(IsNull(rstGetVal("Chemical_14_Code")), -1, rstGetVal("Chemical_14_Code")))
        Call selectValueInCombo(Me.f_Chemical_15_Code, IIf(IsNull(rstGetVal("Chemical_15_Code")), -1, rstGetVal("Chemical_15_Code")))
        Me.f_Chemical_Qty.Text = IIf(IsNull(rstGetVal("Chemical_Qty")), 0, rstGetVal("Chemical_Qty"))
        Me.f_Chemical2_Qty.Text = IIf(IsNull(rstGetVal("Chemical2_Qty")), 0, rstGetVal("Chemical2_Qty"))
        Me.f_Chemical_3_Qty.Text = IIf(IsNull(rstGetVal("Chemical_3_Qty")), 0, rstGetVal("Chemical_3_Qty"))
        Me.f_Chemical_4_Qty.Text = IIf(IsNull(rstGetVal("Chemical_4_Qty")), 0, rstGetVal("Chemical_4_Qty"))
        Me.f_Chemical_5_Qty.Text = IIf(IsNull(rstGetVal("Chemical_5_Qty")), 0, rstGetVal("Chemical_5_Qty"))
        Me.f_Chemical_6_Qty.Text = IIf(IsNull(rstGetVal("Chemical_6_Qty")), 0, rstGetVal("Chemical_6_Qty"))
        Me.f_Chemical_7_Qty.Text = IIf(IsNull(rstGetVal("Chemical_7_Qty")), 0, rstGetVal("Chemical_7_Qty"))
        Me.f_Chemical_8_Qty.Text = IIf(IsNull(rstGetVal("Chemical_8_Qty")), 0, rstGetVal("Chemical_8_Qty"))
        Me.f_Chemical_9_Qty.Text = IIf(IsNull(rstGetVal("Chemical_9_Qty")), 0, rstGetVal("Chemical_9_Qty"))
        Me.f_Chemical_10_Qty.Text = IIf(IsNull(rstGetVal("Chemical_10_Qty")), 0, rstGetVal("Chemical_10_Qty"))
        Me.f_Chemical_11_Qty.Text = IIf(IsNull(rstGetVal("Chemical_11_Qty")), 0, rstGetVal("Chemical_11_Qty"))
        Me.f_Chemical_12_Qty.Text = IIf(IsNull(rstGetVal("Chemical_12_Qty")), 0, rstGetVal("Chemical_12_Qty"))
        Me.f_Chemical_13_Qty.Text = IIf(IsNull(rstGetVal("Chemical_13_Qty")), 0, rstGetVal("Chemical_13_Qty"))
        Me.f_Chemical_14_Qty.Text = IIf(IsNull(rstGetVal("Chemical_14_Qty")), 0, rstGetVal("Chemical_14_Qty"))
        Me.f_Chemical_15_Qty.Text = IIf(IsNull(rstGetVal("Chemical_15_Qty")), 0, rstGetVal("Chemical_15_Qty"))
        Me.f_RecipeCode.Text = IIf(IsNull(rstGetVal("RecipeCode")), 0, rstGetVal("RecipeCode"))
        If rstGetVal("Re_RecipeCode") = 1 Then
            Me.f_Re_RecipeCode.value = Checked
        Else
            Me.f_Re_RecipeCode.value = Unchecked
        End If
'        Dim sql As String
'        If Me.f_Re_RecipeCode.value = Checked And Len(Trim(Me.f_RecipeCode.Text)) > 0 Then
'            sql = "Select ItemCode, ItemName from Item where ItemCode in (Select ItemCode from RecipeDetail where RecipeMasterCode = " & Me.f_RecipeCode.Text & ")"
'            FillColorCombo sql, f_R_Color_1, "ItemName", "ItemCode"
'            FillColorCombo sql, f_R_Color_2, "ItemName", "ItemCode"
'            FillColorCombo sql, f_R_Color_3, "ItemName", "ItemCode"
'            FillColorCombo sql, f_R_Color_4, "ItemName", "ItemCode"
'            FillColorCombo sql, f_R_Color_5, "ItemName", "ItemCode"
'            FillColorCombo sql, f_R_Color_6, "ItemName", "ItemCode"
            
'            FillColorCombo sql, f_RecipeColor_6, "ItemName", "ItemCode"
'            FillColorCombo sql, f_RecipeColor_7, "ItemName", "ItemCode"
'            FillColorCombo sql, f_RecipeColor_8, "ItemName", "ItemCode"
'            FillColorCombo sql, f_RecipeColor_9, "ItemName", "ItemCode"
'            FillColorCombo sql, f_RecipeColor_10, "ItemName", "ItemCode"
'            FillColorCombo sql, f_RecipeColor_11, "ItemName", "ItemCode"
'        ElseIf Me.f_Re_RecipeCode.value = Unchecked And Len(Trim(Me.f_RecipeCode.Text)) > 0 Then
'            sql = "Select ItemCode, ItemName from Item where ItemCode in (Select ItemCode from RecipeDetail where RecipeMasterCode = " & Me.f_RecipeCode.Text & ")"
'            FillColorCombo sql, f_Color_1, "ItemName", "ItemCode"
'            FillColorCombo sql, f_Color_2, "ItemName", "ItemCode"
'            FillColorCombo sql, f_Color_3, "ItemName", "ItemCode"
'            FillColorCombo sql, f_Color_4, "ItemName", "ItemCode"
'            FillColorCombo sql, f_Color_5, "ItemName", "ItemCode"
            
'            FillColorCombo sql, f_RecipeColor, "ItemName", "ItemCode"
'            FillColorCombo sql, f_RecipeColor_2, "ItemName", "ItemCode"
'            FillColorCombo sql, f_RecipeColor_3, "ItemName", "ItemCode"
'            FillColorCombo sql, f_RecipeColor_4, "ItemName", "ItemCode"
'            FillColorCombo sql, f_RecipeColor_5, "ItemName", "ItemCode"
'        End If
       
        Call selectValueInCombo(Me.f_Color_1, IIf(IsNull(rstGetVal("Color_1")), -1, rstGetVal("Color_1")))
        Call selectValueInCombo(Me.f_Color_2, IIf(IsNull(rstGetVal("Color_2")), -1, rstGetVal("Color_2")))
        Call selectValueInCombo(Me.f_Color_3, IIf(IsNull(rstGetVal("Color_3")), -1, rstGetVal("Color_3")))
        Call selectValueInCombo(Me.f_Color_4, IIf(IsNull(rstGetVal("Color_4")), -1, rstGetVal("Color_4")))
        Call selectValueInCombo(Me.f_Color_5, IIf(IsNull(rstGetVal("Color_5")), -1, rstGetVal("Color_5")))
        Call selectValueInCombo(Me.f_R_Color_1, IIf(IsNull(rstGetVal("R_Color_1")), -1, rstGetVal("R_Color_1")))
        Call selectValueInCombo(Me.f_R_Color_2, IIf(IsNull(rstGetVal("R_Color_2")), -1, rstGetVal("R_Color_2")))
        Call selectValueInCombo(Me.f_R_Color_3, IIf(IsNull(rstGetVal("R_Color_3")), -1, rstGetVal("R_Color_3")))
        Call selectValueInCombo(Me.f_R_Color_4, IIf(IsNull(rstGetVal("R_Color_4")), -1, rstGetVal("R_Color_4")))
        Call selectValueInCombo(Me.f_R_Color_5, IIf(IsNull(rstGetVal("R_Color_5")), -1, rstGetVal("R_Color_5")))
        Call selectValueInCombo(Me.f_R_Color_6, IIf(IsNull(rstGetVal("R_Color_6")), -1, rstGetVal("R_Color_6")))
        Call selectValueInCombo(Me.f_R_Color_7, IIf(IsNull(rstGetVal("R_Color_7")), -1, rstGetVal("R_Color_7")))
        Call selectValueInCombo(Me.f_R_Color_8, IIf(IsNull(rstGetVal("R_Color_8")), -1, rstGetVal("R_Color_8")))
        Call selectValueInCombo(Me.f_R_Color_9, IIf(IsNull(rstGetVal("R_Color_9")), -1, rstGetVal("R_Color_9")))
        Call selectValueInCombo(Me.f_R_Color_10, IIf(IsNull(rstGetVal("R_Color_10")), -1, rstGetVal("R_Color_10")))
        Me.f_Color_1_Qty.Text = IIf(IsNull(rstGetVal("Color_1_Qty")), 0, rstGetVal("Color_1_Qty"))
        Me.f_Color_2_Qty.Text = IIf(IsNull(rstGetVal("Color_2_Qty")), 0, rstGetVal("Color_2_Qty"))
        Me.f_Color_3_Qty.Text = IIf(IsNull(rstGetVal("Color_3_Qty")), 0, rstGetVal("Color_3_Qty"))
        Me.f_Color_4_Qty.Text = IIf(IsNull(rstGetVal("Color_4_Qty")), 0, rstGetVal("Color_4_Qty"))
        Me.f_Color_5_Qty.Text = IIf(IsNull(rstGetVal("Color_5_Qty")), 0, rstGetVal("Color_5_Qty"))
        Me.f_R_Color_1_Qty.Text = IIf(IsNull(rstGetVal("R_Color_1_Qty")), 0, rstGetVal("R_Color_1_Qty"))
        Me.f_R_Color_2_Qty.Text = IIf(IsNull(rstGetVal("R_Color_2_Qty")), 0, rstGetVal("R_Color_2_Qty"))
        Me.f_R_Color_3_Qty.Text = IIf(IsNull(rstGetVal("R_Color_3_Qty")), 0, rstGetVal("R_Color_3_Qty"))
        Me.f_R_Color_4_Qty.Text = IIf(IsNull(rstGetVal("R_Color_4_Qty")), 0, rstGetVal("R_Color_4_Qty"))
        Me.f_R_Color_5_Qty.Text = IIf(IsNull(rstGetVal("R_Color_5_Qty")), 0, rstGetVal("R_Color_5_Qty"))
        Me.f_R_Color_6_Qty.Text = IIf(IsNull(rstGetVal("R_Color_6_Qty")), 0, rstGetVal("R_Color_6_Qty"))
        Me.f_R_Color_7_Qty.Text = IIf(IsNull(rstGetVal("R_Color_7_Qty")), 0, rstGetVal("R_Color_7_Qty"))
        Me.f_R_Color_8_Qty.Text = IIf(IsNull(rstGetVal("R_Color_8_Qty")), 0, rstGetVal("R_Color_8_Qty"))
        Me.f_R_Color_9_Qty.Text = IIf(IsNull(rstGetVal("R_Color_9_Qty")), 0, rstGetVal("R_Color_9_Qty"))
        Me.f_R_Color_10_Qty.Text = IIf(IsNull(rstGetVal("R_Color_10_Qty")), 0, rstGetVal("R_Color_10_Qty"))
        Call selectValueInCombo(Me.f_Soap, IIf(IsNull(rstGetVal("Soap")), 0, rstGetVal("Soap")))
        Me.f_Soap_Qty.Text = IIf(IsNull(rstGetVal("Soap_Qty")), 0, rstGetVal("Soap_Qty"))
        Me.f_SoapTime.Text = IIf(IsNull(rstGetVal("SoapTime")), "", rstGetVal("SoapTime"))
        Call selectValueInCombo(Me.f_Soap2, IIf(IsNull(rstGetVal("Soap2")), -1, rstGetVal("Soap2")))
        Me.f_Soap2_Qty.Text = IIf(IsNull(rstGetVal("Soap2_Qty")), 0, rstGetVal("Soap2_Qty"))
        Me.f_SoapTime2.Text = IIf(IsNull(rstGetVal("SoapTime2")), "", rstGetVal("SoapTime2"))
        Call selectValueInCombo(Me.f_Soap3, IIf(IsNull(rstGetVal("Soap3")), -1, rstGetVal("Soap3")))
        Me.f_Soap3_Qty.Text = IIf(IsNull(rstGetVal("Soap3_Qty")), 0, rstGetVal("Soap3_Qty"))
        Call selectValueInCombo(Me.f_Acid, IIf(IsNull(rstGetVal("Acid")), -1, rstGetVal("Acid")))
        Me.f_Acid_Qty.Text = IIf(IsNull(rstGetVal("Acid_Qty")), 0, rstGetVal("Acid_Qty"))
        Call selectValueInCombo(Me.f_Acid2, IIf(IsNull(rstGetVal("Acid2")), -1, rstGetVal("Acid2")))
        Me.f_Acid2_Qty.Text = IIf(IsNull(rstGetVal("Acid2_Qty")), 0, rstGetVal("Acid2_Qty"))
        Call selectValueInCombo(Me.f_Hydro, IIf(IsNull(rstGetVal("Hydro")), 0, rstGetVal("Hydro")))
        Me.f_Hydro_Qty.Text = IIf(IsNull(rstGetVal("Hydro_Qty")), 0, rstGetVal("Hydro_Qty"))
        Call selectValueInCombo(Me.f_Hydro2, IIf(IsNull(rstGetVal("Hydro2")), -1, rstGetVal("Hydro2")))
        Me.f_Hydro_Qty2.Text = IIf(IsNull(rstGetVal("Hydro2_Qty")), 0, rstGetVal("Hydro2_Qty"))
        Call selectValueInCombo(Me.f_Hydro3, IIf(IsNull(rstGetVal("Hydro3")), -1, rstGetVal("Hydro3")))
        Me.f_Hydro_Qty.Text = IIf(IsNull(rstGetVal("Hydro_Qty")), 0, rstGetVal("Hydro_Qty"))
        Call selectValueInCombo(Me.f_Castic, IIf(IsNull(rstGetVal("Castic")), 0, rstGetVal("Castic")))
        Me.f_Castic_Qty.Text = IIf(IsNull(rstGetVal("Castic_Qty")), 0, rstGetVal("Castic_Qty"))
        Me.f_CasticTime.Text = IIf(IsNull(rstGetVal("CasticTime")), "", rstGetVal("CasticTime"))
        Call selectValueInCombo(Me.f_Castic2, IIf(IsNull(rstGetVal("Castic2")), -1, rstGetVal("Castic2")))
        Me.f_Castic_Qty2.Text = IIf(IsNull(rstGetVal("Castic2_Qty")), 0, rstGetVal("Castic2_Qty"))
        Me.f_CasticTime2.Text = IIf(IsNull(rstGetVal("CasticTime2")), "", rstGetVal("CasticTime2"))
        Call selectValueInCombo(Me.f_Castic3, IIf(IsNull(rstGetVal("Castic3")), -1, rstGetVal("Castic3")))
        Me.f_Castic_Qty3.Text = IIf(IsNull(rstGetVal("Castic3_Qty")), 0, rstGetVal("Castic3_Qty"))
        Me.f_CasticTime3.Text = IIf(IsNull(rstGetVal("CasticTime3")), "", rstGetVal("CasticTime3"))
        Me.f_Remarks.Text = IIf(IsNull(rstGetVal("Remarks")), " ", rstGetVal("Remarks"))
        Call selectValueInCombo(Me.f_RecipeColor, IIf(IsNull(rstGetVal("RecipeColor1")), -1, rstGetVal("RecipeColor1")))
        Call selectValueInCombo(Me.f_RecipeColor_2, IIf(IsNull(rstGetVal("RecipeColor2")), -1, rstGetVal("RecipeColor2")))
        Call selectValueInCombo(Me.f_RecipeColor_3, IIf(IsNull(rstGetVal("RecipeColor3")), -1, rstGetVal("RecipeColor3")))
        Call selectValueInCombo(Me.f_RecipeColor_4, IIf(IsNull(rstGetVal("RecipeColor4")), -1, rstGetVal("RecipeColor4")))
        Call selectValueInCombo(Me.f_RecipeColor_5, IIf(IsNull(rstGetVal("RecipeColor5")), -1, rstGetVal("RecipeColor5")))
        Call selectValueInCombo(Me.f_RecipeColor_6, IIf(IsNull(rstGetVal("RecipeColor6")), -1, rstGetVal("RecipeColor6")))
        Call selectValueInCombo(Me.f_RecipeColor_7, IIf(IsNull(rstGetVal("RecipeColor7")), -1, rstGetVal("RecipeColor7")))
        Call selectValueInCombo(Me.f_RecipeColor_8, IIf(IsNull(rstGetVal("RecipeColor8")), -1, rstGetVal("RecipeColor8")))
        Call selectValueInCombo(Me.f_RecipeColor_9, IIf(IsNull(rstGetVal("RecipeColor9")), -1, rstGetVal("RecipeColor9")))
        Call selectValueInCombo(Me.f_RecipeColor_10, IIf(IsNull(rstGetVal("RecipeColor10")), -1, rstGetVal("RecipeColor10")))
        Call selectValueInCombo(Me.f_RecipeColor_11, IIf(IsNull(rstGetVal("RecipeColor11")), -1, rstGetVal("RecipeColor11")))
        Call selectValueInCombo(Me.f_RecipeColor_12, IIf(IsNull(rstGetVal("RecipeColor12")), -1, rstGetVal("RecipeColor12")))
        Call selectValueInCombo(Me.f_RecipeColor_13, IIf(IsNull(rstGetVal("RecipeColor13")), -1, rstGetVal("RecipeColor13")))
        Call selectValueInCombo(Me.f_RecipeColor_14, IIf(IsNull(rstGetVal("RecipeColor14")), -1, rstGetVal("RecipeColor14")))
        Call selectValueInCombo(Me.f_RecipeColor_15, IIf(IsNull(rstGetVal("RecipeColor15")), -1, rstGetVal("RecipeColor15")))
        Me.f_RecipeColor_Qty.Text = IIf(IsNull(rstGetVal("RecipeColor1_Qty")), 0, rstGetVal("RecipeColor1_Qty"))
        Me.f_RecipeColor_Qty_2.Text = IIf(IsNull(rstGetVal("RecipeColor2_Qty")), 0, rstGetVal("RecipeColor2_Qty"))
        Me.f_RecipeColor_Qty_3.Text = IIf(IsNull(rstGetVal("RecipeColor3_Qty")), 0, rstGetVal("RecipeColor3_Qty"))
        Me.f_RecipeColor_Qty_4.Text = IIf(IsNull(rstGetVal("RecipeColor4_Qty")), 0, rstGetVal("RecipeColor4_Qty"))
        Me.f_RecipeColor_Qty_5.Text = IIf(IsNull(rstGetVal("RecipeColor5_Qty")), 0, rstGetVal("RecipeColor5_Qty"))
        Me.f_RecipeColor_Qty_6.Text = IIf(IsNull(rstGetVal("RecipeColor6_Qty")), 0, rstGetVal("RecipeColor6_Qty"))
        Me.f_RecipeColor_Qty_7.Text = IIf(IsNull(rstGetVal("RecipeColor7_Qty")), 0, rstGetVal("RecipeColor7_Qty"))
        Me.f_RecipeColor_Qty_8.Text = IIf(IsNull(rstGetVal("RecipeColor8_Qty")), 0, rstGetVal("RecipeColor8_Qty"))
        Me.f_RecipeColor_Qty_9.Text = IIf(IsNull(rstGetVal("RecipeColor9_Qty")), 0, rstGetVal("RecipeColor9_Qty"))
        Me.f_RecipeColor_Qty_10.Text = IIf(IsNull(rstGetVal("RecipeColor10_Qty")), 0, rstGetVal("RecipeColor10_Qty"))
        Me.f_RecipeColor_Qty_11.Text = IIf(IsNull(rstGetVal("RecipeColor11_Qty")), 0, rstGetVal("RecipeColor11_Qty"))
        Me.f_RecipeColor_Qty_12.Text = IIf(IsNull(rstGetVal("RecipeColor12_Qty")), 0, rstGetVal("RecipeColor12_Qty"))
        Me.f_RecipeColor_Qty_13.Text = IIf(IsNull(rstGetVal("RecipeColor13_Qty")), 0, rstGetVal("RecipeColor13_Qty"))
        Me.f_RecipeColor_Qty_14.Text = IIf(IsNull(rstGetVal("RecipeColor14_Qty")), 0, rstGetVal("RecipeColor14_Qty"))
        Me.f_RecipeColor_Qty_15.Text = IIf(IsNull(rstGetVal("RecipeColor15_Qty")), 0, rstGetVal("RecipeColor15_Qty"))
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
End Sub
Private Sub lvwphase_Click()
    cmdSave.Enabled = True
    m_AddMode = False
    If Me.lvwphase.ListItems.Count > 0 Then
        m_ListID = Me.lvwphase.SelectedItem.ListSubItems(9).Text
        ClickPane = 1
        Call getVal
    End If
End Sub
Private Sub lvwphase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSave.Enabled = True
        m_AddMode = False
        If Me.lvwphase.ListItems.Count > 0 Then
            m_ListID = Me.lvwphase.SelectedItem.ListSubItems(9).Text
            ClickPane = 1
            Call getVal
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
Private Sub SERChk_Click()
    If SERChk.value = Checked Then
        Me.srSER1.Enabled = True
        Me.srSER2.Enabled = True
    Else
        Me.srSER1.Enabled = False
        Me.srSER2.Enabled = False
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
    srPC2 = srPC1
    Call SrfillList
End Sub
Private Sub SrPC2_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrSER1_keyup(KeyCode As Integer, Shift As Integer)
    srSER2 = srSER1
    Call SrfillList
End Sub
Private Sub SrSER2_keyup(KeyCode As Integer, Shift As Integer)
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
    Dim cbo6 As String
    Dim cbo7 As String
    If dtChk.value = Checked Then
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
    
    If SERChk.value = Checked And Len(Trim(Me.srSER1)) > 0 And Len(Trim(Me.srSER2)) > 0 Then
        cbo7 = " And (SerialNo like '%" & Me.srSER1 & "%')"
    Else
        cbo7 = ""
    End If
    
    sql = " SELECT top 200 ProcessCode, convert(varchar, ProcessDate, 103) as ProcessDate, isNull(SerialNo, 0) as SerialNo, PartyName, MachineNo, ItemTypeName, Cone, isNull(RecipeCode, 0) as RecipeCode, isNull(Re_RecipeCode, 0) as Re_RecipeCode " & _
          " FROM Party INNER JOIN (ItemType INNER JOIN Process ON ItemType.ItemTypeCode = Process.ItemTypeCode) ON Party.PartyCode = Process.PartyCode " & _
          " Where Process.Is_Active = 1 and Is_Cotton_Dyeing = 0 " & _
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
                   Text:=rstList!ProcessCode, _
                   Key:=CStr("Id=" & rstList!ProcessCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!ProcessDate
                 .Add Text:=rstList!SerialNo
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!MachineNo
                 .Add Text:=rstList!ItemTypeName
                 .Add Text:=rstList!Cone
                 .Add Text:=rstList!RecipeCode
                 .Add Text:=rstList!Re_RecipeCode
                 .Add Text:=rstList!ProcessCode
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
Private Sub chkQty(vItemCode As Integer, vQty As Double)
    Dim AvbQty As Double
    Dim strAns As String
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
        If (IsNull(vQty)) Or vQty <= 0 Then
            MsgBox "Quantity must be greater then zero"
            MsgBit = 0
        ElseIf (IIf(m_AddMode = False, (CDbl(AvbQty) + CDbl(PreQty)), CDbl(AvbQty)) < IIf(m_AddMode = False, CDbl(vQty), CDbl(vQty))) Then
            strAns = MsgBox("Quantity not Available !" & Chr(13) & "Would your like to Continue ", vbYesNo + vbInformation)
            If strAns = vbNo Then
                MsgBit = 0
            Else
                MsgBit = 1
            End If
        Else
            MsgBit = 1
        End If
    End If
End Sub
