VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RepReceiveQuantity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Activity"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3075
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
      Height          =   3045
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   4605
      Begin VB.Frame Frame4 
         Caption         =   "Party"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   2895
         Begin VB.ComboBox Party 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Date"
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   2895
         Begin MSComCtl2.DTPicker dt2 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   2650
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Format          =   61538305
            CurrentDate     =   38301
         End
         Begin MSComCtl2.DTPicker dt1 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2650
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Format          =   61538305
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
               Picture         =   "RepReceiveQuantity.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RepReceiveQuantity.frx":0278
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton cmdCancel 
         Height          =   405
         Left            =   3120
         TabIndex        =   2
         Top             =   1755
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
         MICON           =   "RepReceiveQuantity.frx":0310
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
         Top             =   1260
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
         MICON           =   "RepReceiveQuantity.frx":032C
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
         Caption         =   "Item"
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2895
         Begin VB.ComboBox Item 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   2650
         End
         Begin VB.ComboBox ItemType 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2650
         End
      End
   End
End
Attribute VB_Name = "RepReceiveQuantity"
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
    If Me.ItemType.ListIndex > -1 And Me.Item.ListIndex > -1 Then
    
        sql = "if exists (select 1 from sysobjects where xtype = 'U' and name = 'vw_ReceiveQuantity_support') begin drop table dbo.vw_ReceiveQuantity_support end"
        cnDatabase.Execute sql
        
        sql = "if exists (select 1 from sysobjects where xtype = 'U' and name = 'vw_ReceiveQuantity') begin drop table dbo.vw_ReceiveQuantity end"
        cnDatabase.Execute sql
        
        sql = "select * into dbo.vw_ReceiveQuantity_support from dbo.vw_ReceivedQuantity where partycode = " & Me.Party.ItemData(Me.Party.ListIndex) & " and itemcode = " & Me.Item.ItemData(Me.Item.ListIndex) & " order by MasterInputDate "
        cnDatabase.Execute sql
        
        sql = "alter table dbo.vw_ReceiveQuantity_support add Id int identity(1, 1)"
        cnDatabase.Execute sql
        
        sql = "select Type, MasterInputDate, ProcessCode, PartyName,ItemTypeName, ItemName, NewColor, ChallanCode, Receive, Pending, Delivery, (select sum(isNull(Receive,0) - (isNull(Pending, 0) + isNull(Delivery, 0))) from dbo.vw_ReceiveQuantity_support where Id <= a.id) as Balance, Cone , PartyCode, ItemTypeCode, ItemCode into dbo.vw_ReceiveQuantity from dbo.vw_ReceiveQuantity_support a"
        cnDatabase.Execute sql
   
        ' ------ Changed ------
        
        'sql = "insert into Dyeing.dbo.vw_ReceiveQuantity (type, MasterInputDate, NewColor, Balance) select -1,  '" & Format(Me.dt1.value, "yyyy-mm-dd") & "', 'Opening Balance', sum(Balance) from Dyeing.dbo.vw_ReceiveQuantity where MasterInputDate < '" & Format(Me.dt1.value, "yyyy-mm-dd") & "'"
        'cnDatabase.Execute sql
   
       ' sql = "delete from Dyeing.dbo.vw_ReceiveQuantity where MasterInputDate < '" & Format(Me.dt1.value, "yyyy-mm-dd") & "'"
        'cnDatabase.Execute sql
        
        'sql = "alter table Dyeing.dbo.vw_ReceiveQuantity add Id int identity(1, 1)"
        'cnDatabase.Execute sql
        
        'sql = "update Dyeing.dbo.vw_ReceiveQuantity set Balance = (select isNull(Balance, 0) from vw_ReceiveQuantity a where id = vw_ReceiveQuantity.id -1) + isNull(Receive, 0) - (isNull(Pending, 0) + isNull(Delivery, 0)) where type > -1"
        'cnDatabase.Execute sql

        crptDaily.ReportFileName = App.Path & "\Reports\Rpt_ReceiveQuantity.rpt"
        crptDaily.Connect = conStr
        
        selcformula = "{vw_ReceiveQuantity.PartyCode}= " & Me.Party.ItemData(Me.Party.ListIndex) & " and {vw_ReceiveQuantity.ItemCode}= " & Me.Item.ItemData(Me.Item.ListIndex) & " and {vw_ReceiveQuantity.MasterInputDate} >= #" & Format(dt1, "dd-mmm-yy") & "#  and {vw_ReceiveQuantity.MasterInputDate} <= #" & Format(dt2, "dd-mm-yy") & "#  "
        
        vItemName = getItemName(Me.Item.ItemData(Me.Item.ListIndex))
        vPartyName = getPartyName(Me.Party.ItemData(Me.Party.ListIndex))
        
        crptDaily.Formulas(0) = "PartyHeading ='" & vPartyName & "'"
        crptDaily.Formulas(1) = "ItemHeading ='" & vItemName & "'"
        crptDaily.Formulas(2) = "Date_1 ='" & Format(Me.dt1.value, "dd-mm-yyyy") & "'"
        crptDaily.Formulas(3) = "Date_2 ='" & Format(Me.dt2.value, "dd-mm-yyyy") & "'"
        crptDaily.SelectionFormula = selcformula
        crptDaily.WindowState = crptMaximized
        crptDaily.Action = 1
    End If
End Sub
Private Sub Form_Load()
    mdlGeneral.DBConn
    dt1 = Now
    dt2 = Now
    FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1", ItemType, "ItemTypeName", "ItemTypeCode"
    FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", Party, "PartyName", "PartyCode"
End Sub

Private Sub ItemType_Click()
    If Me.ItemType.ListIndex > -1 Then
        i = Me.ItemType.ItemData(Me.ItemType.ListIndex)
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = " & i & " order by ItemName ", Item, "ItemName", "ItemCode"
    Else
        Me.Item.Clear
    End If
End Sub
