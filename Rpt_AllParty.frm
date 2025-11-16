VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Rpt_AllParty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Party Report"
   ClientHeight    =   1155
   ClientLeft      =   2520
   ClientTop       =   3495
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4740
   Begin Crystal.CrystalReport crptDaily 
      Left            =   840
      Top             =   1320
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
      Height          =   1125
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   4605
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2160
         Top             =   480
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
               Picture         =   "Rpt_AllParty.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Rpt_AllParty.frx":0278
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton cmdCancel 
         Height          =   405
         Left            =   2880
         TabIndex        =   2
         Top             =   400
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
         MICON           =   "Rpt_AllParty.frx":06B3
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
         Left            =   480
         TabIndex        =   1
         Top             =   400
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
         MICON           =   "Rpt_AllParty.frx":06CF
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
Attribute VB_Name = "Rpt_AllParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdReport_Click()
Dim str As String
    

    crptDaily.ReportFileName = App.Path & "\Reports\Rpt_AllParty.rpt"
    
   ' MsgBox "report file name =" & crptDaily.ReportFileName
    
    crptDaily.Connect = conStr
   ' Str = "date(" & Year(dtDaily.VALUE) & "," & Month(dtDaily.VALUE) & "," & Day(dtDaily.VALUE) & ")"
   ' crptDaily.SelectionFormula = "{outgoing.DATE_RECEIVED}= " & Str & "  and {outgoing.Drawee_code} = 0"
    crptDaily.WindowState = crptMaximized
    crptDaily.Action = 1
    

End Sub
Private Sub Form_Load()
   ' dtDaily.VALUE = Date
End Sub

