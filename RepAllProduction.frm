VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RepAllProduction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Report"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4815
   Begin Crystal.CrystalReport crptDaily 
      Left            =   3480
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   4710
      Begin VB.Frame Frame3 
         Caption         =   "Date"
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2895
         Begin MSComCtl2.DTPicker dt2 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   2650
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Format          =   19791873
            CurrentDate     =   38301
         End
         Begin MSComCtl2.DTPicker dt1 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2650
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Format          =   19791873
            CurrentDate     =   38301
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2760
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RepAllProduction.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RepAllProduction.frx":0278
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton cmdCancel 
         Height          =   405
         Left            =   3120
         TabIndex        =   2
         Top             =   795
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Cancel"
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
         MICON           =   "RepAllProduction.frx":0310
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
      Begin LVbuttons.LaVolpeButton cmdReport 
         Height          =   405
         Left            =   3120
         TabIndex        =   1
         Top             =   345
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Preview"
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
         MICON           =   "RepAllProduction.frx":032C
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
End
Attribute VB_Name = "RepAllProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdReport_Click()
Dim str As String
Dim rstGetVal As New ADODB.Recordset
        crptDaily.ReportFileName = App.Path & "\Reports\Rpt_Production.rpt"
        crptDaily.Connect = conStr
        
        selcformula = "{Delivery.DeliveryMasterDate} >= #" & Format(dt1, "dd-mmm-yy") & "#  and {Delivery.DeliveryMasterDate} <= #" & Format(dt2, "dd-mmm-yy") & "#  "
        
        crptDaily.Formulas(0) = "Date_1 ='" & Format(Me.dt1.value, "dd-mm-yyyy") & "'"
        crptDaily.Formulas(1) = "Date_2 ='" & Format(Me.dt2.value, "dd-mm-yyyy") & "'"
        crptDaily.SelectionFormula = selcformula
        crptDaily.WindowState = crptMaximized
        crptDaily.Action = 1
End Sub
Private Sub Form_Load()
    mdlGeneral.DBConn
    dt1 = Now
    dt2 = Now
End Sub
