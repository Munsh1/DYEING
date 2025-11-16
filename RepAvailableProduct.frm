VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RepAvailableProducts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Available Products"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1215
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
         Left            =   2280
         Top             =   360
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
               Picture         =   "RepAvailableProduct.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RepAvailableProduct.frx":0278
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton cmdCancel 
         Height          =   400
         Left            =   3000
         TabIndex        =   3
         Top             =   625
         Width           =   1400
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
         MICON           =   "RepAvailableProduct.frx":0310
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
         Left            =   3000
         TabIndex        =   2
         Top             =   200
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
         MICON           =   "RepAvailableProduct.frx":032C
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
      Begin VB.ComboBox ItemType 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "ItemType"
         Top             =   480
         Width           =   2115
      End
   End
End
Attribute VB_Name = "RepAvailableProducts"
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
   If Me.ItemType.ItemData(Me.ItemType.ListIndex) > 0 Then
   
        vItemType = Me.ItemType.ItemData(Me.ItemType.ListIndex)
        
        crptDaily.Connect = conStr
       ' crptDaily.ReportFileName = App.Path & "\Reports\AvailableQty.rpt"
        
       ' selcformula = "{vwAvailableQty.ItemTypeCode}=" & Me.ItemType.ItemData(Me.ItemType.ListIndex)
      
        If Me.ItemType.ItemData(Me.ItemType.ListIndex) = 1 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type1.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 2 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type2.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 3 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type3.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 4 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type4.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 5 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type5.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 6 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type6.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 7 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type7.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 8 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type8.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 9 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type9.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 10 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type10.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 11 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type11.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 12 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type12.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 13 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type13.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 14 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type14.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 15 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type15.rpt"
        ElseIf Me.ItemType.ItemData(Me.ItemType.ListIndex) = 16 Then
            crptDaily.ReportFileName = App.Path & "\Reports\vwAvailableQty_Type16.rpt"
        Else
            selcformula = "{vwAvailableQty.ItemTypeCode}=" & Me.ItemType.ItemData(Me.ItemType.ListIndex)
            crptDaily.ReportFileName = App.Path & "\Reports\AvailableQty.rpt"
        End If
        
         
'       ItemType = getFieldValue(ItemType, "ItemType", "ItemTypeName", "ItemTypeCode")
        vItemTypeName = getItemTypeName(Me.ItemType.ItemData(Me.ItemType.ListIndex))
       
        crptDaily.Formulas(0) = "ItemType ='" & vItemTypeName & "'"
        crptDaily.SelectionFormula = selcformula
        crptDaily.WindowState = crptMaximized
        crptDaily.Action = 1

    End If
End Sub

Private Sub Form_Load()
    mdlGeneral.DBConn
    FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1", ItemType, "ItemTypeName", "ItemTypeCode"
End Sub
