VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RepDeliveryParty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery Report"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4725
   Begin Crystal.CrystalReport crptDaily 
      Left            =   3960
      Top             =   360
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
      Height          =   2445
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   4605
      Begin VB.Frame Frame3 
         Caption         =   "Date"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
         Begin MSComCtl2.DTPicker dt2 
            Height          =   315
            Left            =   120
            TabIndex        =   7
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
            TabIndex        =   6
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
         Left            =   3240
         Top             =   240
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
               Picture         =   "RepDeliveryParty.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RepDeliveryParty.frx":0278
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton cmdCancel 
         Height          =   405
         Left            =   3120
         TabIndex        =   2
         Top             =   1400
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
         MICON           =   "RepDeliveryParty.frx":0310
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
         Top             =   840
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
         MICON           =   "RepDeliveryParty.frx":032C
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
      Begin VB.Frame Frame2 
         Caption         =   "Party"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2895
         Begin VB.ComboBox Party 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Text            =   "Party"
            Top             =   240
            Width           =   2650
         End
      End
   End
End
Attribute VB_Name = "RepDeliveryParty"
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
   If Me.Party.ListIndex > -1 Then
        crptDaily.ReportFileName = App.Path & "\Reports\Rpt_PartyDelivery.rpt"
        crptDaily.Connect = conStr

        selcformula = "({Delivery.PartyCode}=" & Me.Party.ItemData(Me.Party.ListIndex) & " or {Delivery.DeliveryPartyCode}=" & Me.Party.ItemData(Me.Party.ListIndex) & ") and  ({Delivery.DeliveryMasterDate} >= #" & Format(dt1, "dd-mmm-yy") & "#  and {Delivery.DeliveryMasterDate} <= #" & Format(dt2, "dd-mmm-yy") & "# ) "
        vPartyName = getPartyName(Me.Party.ItemData(Me.Party.ListIndex))
        
        crptDaily.Formulas(0) = "PartyHeading ='" & vPartyName & "'"
        crptDaily.Formulas(1) = "Date_1 ='" & Format(Me.dt1.value, "dd-mm-yyyy") & "'"
        crptDaily.Formulas(2) = "Date_2 ='" & Format(Me.dt2.value, "dd-mm-yyyy") & "'"
        crptDaily.SelectionFormula = selcformula
        crptDaily.WindowState = crptMaximized
        crptDaily.Action = 1
    End If
End Sub
Private Sub Form_Load()
    mdlGeneral.DBConn
    dt1 = Now
    dt2 = Now
    FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", Party, "PartyName", "PartyCode"
End Sub
