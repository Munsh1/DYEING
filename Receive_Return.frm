VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Receive_Return 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                           -----  Receive Return -----"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleMode       =   0  'User
   ScaleWidth      =   11385
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   120
      TabIndex        =   37
      Top             =   2400
      Width           =   7935
      Begin MSComctlLib.ListView lvwphase 
         Height          =   2520
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   4445
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
      Caption         =   "Detail Block"
      Height          =   1095
      Left            =   120
      TabIndex        =   32
      Top             =   1200
      Width           =   7935
      Begin VB.TextBox Rates 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6360
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Qty 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5160
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox Item 
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Text            =   "Item"
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox Item_Type 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Text            =   "Item_Type"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Unit-Price"
         Height          =   255
         Left            =   6600
         TabIndex        =   36
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   5400
         TabIndex        =   35
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Item"
         Height          =   255
         Left            =   3000
         TabIndex        =   34
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Item Type"
         Height          =   255
         Left            =   720
         TabIndex        =   33
         Top             =   280
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Master Block"
      Height          =   1095
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   7935
      Begin VB.TextBox Challan 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6360
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox Party 
         Height          =   315
         Left            =   2400
         TabIndex        =   2
         Text            =   "Party"
         Top             =   600
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker ReturnDate 
         Height          =   300
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   44695555
         CurrentDate     =   38216
      End
      Begin VB.Label Label3 
         Caption         =   "GatPass No."
         Height          =   255
         Left            =   6405
         TabIndex        =   31
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Party"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   8160
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   2895
         Begin VB.CheckBox dtChk 
            Caption         =   "Date"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   735
         End
         Begin MSComCtl2.DTPicker SrDate 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   280
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   44695553
            CurrentDate     =   38224
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   2895
         Begin VB.CheckBox PtChk 
            Caption         =   "Party"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox SrParty 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Text            =   "SrParty"
            Top             =   280
            Width           =   2655
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   2895
         Begin VB.CheckBox ImTChk 
            Caption         =   "Item Type"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   0
            Width           =   1095
         End
         Begin VB.ComboBox SrItemType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   20
            Text            =   "SrItemType"
            Top             =   280
            Width           =   2655
         End
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   2880
         Width           =   2895
         Begin VB.CheckBox ImChk 
            Caption         =   "Item"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox SrItem 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Text            =   "SrItem"
            Top             =   280
            Width           =   2655
         End
      End
      Begin LVbuttons.LaVolpeButton Cmdhide 
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   4440
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
         MICON           =   "Receive_Return.frx":0000
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Receive_Return.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Receive_Return.frx":0284
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Receive_Return.frx":06DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Receive_Return.frx":0AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Receive_Return.frx":0F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Receive_Return.frx":134C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Receive_Return.frx":1788
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton CmdAllSearch 
      Height          =   405
      Left            =   5655
      TabIndex        =   12
      Top             =   5400
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
      MICON           =   "Receive_Return.frx":181C
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
      Left            =   6915
      TabIndex        =   13
      Top             =   5400
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
      MICON           =   "Receive_Return.frx":1838
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
   Begin LVbuttons.LaVolpeButton CmdDel 
      Height          =   405
      Left            =   4395
      TabIndex        =   11
      Top             =   5400
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
      MICON           =   "Receive_Return.frx":1854
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
   Begin LVbuttons.LaVolpeButton CmdNew 
      Height          =   405
      Left            =   3360
      TabIndex        =   10
      Top             =   5400
      Width           =   1005
      _ExtentX        =   1773
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
      MICON           =   "Receive_Return.frx":1870
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
      Left            =   2295
      TabIndex        =   8
      Top             =   5400
      Width           =   1005
      _ExtentX        =   1773
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
      MICON           =   "Receive_Return.frx":188C
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
   Begin LVbuttons.LaVolpeButton CMDdetail 
      Height          =   405
      Left            =   1155
      TabIndex        =   9
      Top             =   5400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Detail"
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
      MICON           =   "Receive_Return.frx":18A8
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
   Begin VB.Label lblCaption 
      Caption         =   "Label12"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   5520
      Width           =   1215
   End
End
Attribute VB_Name = "Receive_Return"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ListID As String
Dim d_ListID As String
Dim PreQty As Double
Private Sub Challan_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Item_Type.SetFocus
    End If
End Sub
Private Sub CmdAllSearch_Click()
        Receive_Return.Left = 200
        Receive_Return.Top = 500
        Receive_Return.Width = 11500
                
        Call SrfillList
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CmdDel_Click()
    If (Len(Trim(m_ListID)) > 0 And Len(Trim(d_ListID)) > 0) Then
        Dim strAns As String
        Dim AvbQty As Integer
        Dim rstGetQty As New ADODB.Recordset
        
        strAns = MsgBox("Do you want to delete this record...?", vbYesNo + vbInformation)
        If strAns = vbYes Then
            cnDatabase.Execute "update ReturnDetail set IsActive = 0 where ReturnDetailCode =" & d_ListID
            Call fillList
            MsgBox ("Record deleted succesfully..."), vbInformation
            Me.CmdDel.Enabled = False
            Me.cmdSave.Enabled = False
            Call addNewMaster
        End If
        d_ListID = ""
        m_AddMode = True
        Me.CmdNew.SetFocus
        End If
End Sub
Private Sub CMDdetail_Click()
        isdetail = True
        Call addNewDetail
        Me.Item_Type.SetFocus
End Sub
Private Sub Cmdhide_Click()
        Receive_Return.Width = 8250
        Receive_Return.Left = 2000
        Receive_Return.Top = 500
        Me.SrItem.ListIndex = -1
        Me.SrItemType.ListIndex = -1
        Me.srParty.ListIndex = -1
        Call fillList
End Sub
Private Sub CmdNew_Click()
    Call addNewMaster
    Me.Party.SetFocus
End Sub
Private Sub cmdSave_Click()
If Len(Trim(Party)) > 0 And Len(Trim(Challan)) > 0 And Len(Trim(Item_Type)) > 0 And Len(Trim(Item)) > 0 And Len(Trim(Qty)) > 0 And Len(Trim(Rates)) > 0 Then
            Call setValMaster
            Call setValDetail
            Call fillList
            Call addNewDetail
            Me.Item_Type.SetFocus
            MsgBox ("Record saved successfully"), vbInformation
Else
    MsgBox "Provide data in all Fields"
End If
End Sub
Private Sub Form_Load()
    isdetail = True
    m_AddMode = True
    DBConn
    ReturnDate = Date
    SrDate = Date
    FillCombo "SELECT Distinct vwPartyReceives.PartyCode, Party.PartyName FROM Party INNER JOIN vwPartyReceives ON Party.PartyCode = vwPartyReceives.PartyCode where IsActive = 1 order by 2", Party, "PartyName", "PartyCode"
    
    FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", SrItemType, "ItemTypeName", "ItemTypeCode"
    FillCombo "Select PartyCode, PartyName from Party where IsActive = 1 order by 2", srParty, "PartyName", "PartyCode"

    lvwphase.ColumnHeaders.Add Text:="Detail Code", Width:=0
    lvwphase.ColumnHeaders.Add Text:="Master Code", Width:=0
    lvwphase.ColumnHeaders.Add Text:="Party Name", Width:=2100
    lvwphase.ColumnHeaders.Add Text:="Item Type", Width:=2000
    lvwphase.ColumnHeaders.Add Text:="Item Name", Width:=2000
    lvwphase.ColumnHeaders.Add Text:="Quantity", Width:=800
    lvwphase.ColumnHeaders.Add Text:="Rates", Width:=800
    
    Call fillList
    lblCaption.Caption = "Add Master"
End Sub
Private Sub Item_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Qty.SetFocus
    End If
End Sub
Private Sub Item_Type_Click()
    If (Me.Party.ListIndex > -1 And Me.Item_Type.ListIndex > -1) Then
        i = Me.Party.ItemData(Me.Party.ListIndex)
        j = Me.Item_Type.ItemData(Me.Item_Type.ListIndex)
        FillCombo "Select Distinct vwPartyReceives.ItemCode, ItemName from Item inner join vwPartyReceives on vwPartyReceives.ItemCode = Item.ItemCode where IsActive = 1 and PartyCode = " & i & " and vwPartyReceives.ItemTypeCode = " & j, Item, "ItemName", "ItemCode"
    Else
        Me.Item.Clear
    End If
End Sub
Public Sub addNewDetail()
    d_ListID = ""
    Me.Item_Type.ListIndex = -1
    Me.Item.ListIndex = -1
    Me.Qty.Text = ""
    Me.Rates.Text = ""
    lblCaption.Caption = "Add Detail"
End Sub
Private Sub getValMaster()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("Select ReturnMasterCode, ReturnMasterDate, PartyCode, ChallanCode From ReturnMaster Where ReturnMasterCode = " & m_ListID)
    If Not (rstGetVal.EOF) Then
        Me.ReturnDate.value = IIf(IsNull(rstGetVal("ReturnMasterDate")), Now, rstGetVal("ReturnMasterDate"))
        Call selectValueInCombo(Me.Party, rstGetVal("PartyCode"))
        Me.Challan.Text = rstGetVal("ChallanCode")
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
End Sub
Private Sub getValDetail()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("Select ItemTypeCode, ItemCode, Quantity, Rate From ReturnDetail Where ReturnDetailCode = " & d_ListID)
    If Not (rstGetVal.EOF) Then
    Debug.Print rstGetVal("ItemTypeCode")
        Call selectValueInCombo(Me.Item_Type, rstGetVal("ItemTypeCode"))
        Call selectValueInCombo(Me.Item, rstGetVal("ItemCode"))
        Me.Qty.Text = rstGetVal("Quantity")
        Me.Rates.Text = rstGetVal("Rate")
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
End Sub
Public Sub setValMaster()
    Dim rstSave As New ADODB.Recordset
    If (Len(Trim(m_ListID)) = 0) Then
        Set rstSave = FillRecordSet("select * from ReturnMaster Where 1 = 2")
        rstSave.AddNew
        m_ListID = ValAutoNumber("ReturnMaster", "ReturnMasterCode")
        rstSave("ReturnMasterCode") = m_ListID
    Else
       Set rstSave = FillRecordSet("select * from ReturnMaster where ReturnMasterCode =" & m_ListID)
    End If
    
    rstSave("ReturnMasterDate") = Me.ReturnDate.value
    rstSave("PartyCode") = Me.Party.ItemData(Party.ListIndex)
    rstSave("ChallanCode") = Me.Challan.Text
    rstSave("IsActive") = "1"
    
    rstSave.Update
    rstSave.Close
    Set rstSave = Nothing
End Sub
Public Sub setValDetail()
    Dim rstSave As New ADODB.Recordset
    If (Len(Trim(d_ListID)) = 0) Then
        Set rstSave = FillRecordSet("select * from ReturnDetail Where 1 = 2")
        rstSave.AddNew
        d_ListID = ValAutoNumber("ReturnDetail", "ReturnDetailCode")
        rstSave("ReturnDetailCode") = d_ListID
    Else
       Set rstSave = FillRecordSet("select * from ReturnDetail where ReturnDetailCode =" & d_ListID)
    End If
    
    rstSave("ReturnMasterCode") = m_ListID
    rstSave("ItemTypeCode") = Item_Type.ItemData(Item_Type.ListIndex)
    rstSave("ItemCode") = Me.Item.ItemData(Item.ListIndex)
    rstSave("Quantity") = Me.Qty.Text
    rstSave("Rate") = Me.Rates.Text
    rstSave("IsActive") = "1"
    
    rstSave.Update
    rstSave.Close
    Set rstSave = Nothing
    Call addNewDetail
End Sub
Public Sub addNewMaster()
    m_ListID = ""
    Me.ReturnDate.value = Now
    Me.Party.ListIndex = -1
    Me.Challan.Text = ""
    
    d_ListID = ""
    Me.Item_Type.ListIndex = -1
    Me.Item.ListIndex = -1
    Me.Qty.Text = ""
    Me.Rates.Text = ""
    
    lblCaption.Caption = "Add Master"
End Sub
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Set rstList = FillRecordSet("SELECT top 60 ReturnDetailCode, ReturnMaster.ReturnMasterCode, PartyName, ReturnMaster.PartyCode, ReturnDetail.ItemTypeCode, ItemTypeName, ReturnDetail.ItemCode, ItemName, Quantity, Rate " & _
                                "FROM (Party INNER JOIN (ReturnMaster INNER JOIN ReturnDetail ON ReturnMaster.ReturnMasterCode = ReturnDetail.ReturnMasterCode) ON Party.PartyCode = ReturnMaster.PartyCode) INNER JOIN (ItemType INNER JOIN Item ON ItemType.ItemTypeCode = Item.ItemTypeCode) ON (ReturnDetail.ItemCode = Item.ItemCode) AND (ReturnDetail.ItemTypeCode = ItemType.ItemTypeCode) " & _
                                "Where ReturnDetail.IsActive = 1 order by ReturnMaster.ReturnMasterCode desc, ReturnDetailCode desc")
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!ReturnDetailCode, _
                   Key:=CStr("Id=" & rstList!ReturnDetailCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!ReturnMasterCode
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!ItemTypeName
                 .Add Text:=rstList!ItemName
                 .Add Text:=rstList!Quantity
                 .Add Text:=rstList!Rate
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub Item_Type_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Item.SetFocus
    End If
End Sub
Private Sub lvwphase_Click()
    isdetail = False
    cmdSave.Enabled = True
    CmdDel.Enabled = True
    m_AddMode = False
    d_ListID = Me.lvwphase.SelectedItem.Text
    m_ListID = Me.lvwphase.ListItems.Item(Me.lvwphase.SelectedItem.Index).ListSubItems(1).Text
    
    Call getValMaster
    Call getValDetail
End Sub
Private Sub lvwphase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isdetail = False
        cmdSave.Enabled = True
        CmdDel.Enabled = True
        m_AddMode = False
        d_ListID = Me.lvwphase.SelectedItem.Text
        m_ListID = Me.lvwphase.ListItems.Item(Me.lvwphase.SelectedItem.Index).ListSubItems(1).Text
        
        Call getValMaster
        Call getValDetail
    End If
End Sub
Private Sub Party_Click()
    If (Me.Party.ListIndex > -1) Then
        i = Me.Party.ItemData(Me.Party.ListIndex)
        FillCombo "Select Distinct vwPartyReceives.ItemTypeCode, ItemTypeName from ItemType inner join vwPartyReceives on vwPartyReceives.ItemTypeCode = ItemType.ItemTypeCode where IsActive = 1 and PartyCode = " & i, Item_Type, "ItemTypeName", "ItemTypeCode"
        Me.Item.Clear
    Else
        Me.Item_Type.Clear
    End If
End Sub
Private Sub Party_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Challan.SetFocus
    End If
End Sub
Private Sub Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Rates.SetFocus
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub Qty_LostFocus()
    If (Item_Type.ListIndex <> -1 And Item.ListIndex <> -1) Then
        Dim AvbQty As Double
        Dim rstGetQty As New ADODB.Recordset
        Set rstGetQty = FillRecordSet("Select Qty from vwAvailableQty where ItemCode = " & Me.Item.ItemData(Item.ListIndex) & " and ItemTypeCode = " & Me.Item_Type.ItemData(Item_Type.ListIndex))
        AvbQty = 0
        If Not (rstGetQty.EOF) Then
            If (Not IsNull(rstGetQty("Qty"))) Then
                AvbQty = CDbl(rstGetQty("Qty"))
            End If
        End If
        rstGetQty.Close
        Set rstGetQty = Nothing
        If (Len(Trim(Me.Qty.Text)) > 0) Then
            If (IsNull(Me.Qty.Text)) Then
                MsgBox "Quantity must be greater then zero"
                Me.cmdSave.Enabled = False
            ElseIf (IIf(m_AddMode = False, (CLng(AvbQty) + CLng(PreQty - Me.Qty.Text)), CLng(AvbQty)) < IIf(m_AddMode = False, Abs(CLng(Me.Qty.Text) - PreQty), CLng(Me.Qty.Text))) Then
                MsgBox "Quantity not Available !"
                Me.cmdSave.Enabled = False
            ElseIf CLng(Me.Qty.Text) = 0 Then
                MsgBox "Quantity must be greater then zero"
                Me.cmdSave.Enabled = False
            Else
                Me.cmdSave.Enabled = True
                Me.Rates.SetFocus
            End If
        End If
    End If
End Sub
Private Sub Qty_GotFocus()
    If Len(Trim(Qty)) > 0 Then
        PreQty = Me.Qty.Text
    Else
        PreQty = 0
    End If
End Sub
Private Sub Rates_KeyPress(KeyAscii As Integer)
    Dim numVar As Integer
    Call EnableSave
    If KeyAscii = 13 Then
        If Len(Trim(Party)) > 0 And Len(Trim(Challan)) > 0 And Len(Trim(Item_Type)) > 0 And Len(Trim(Item)) > 0 And Len(Trim(Qty)) > 0 And Len(Trim(Rates)) > 0 Then
            Me.cmdSave.SetFocus
        End If
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub EnableSave()
    If Len(Trim(Party)) > 0 And Len(Trim(Challan)) > 0 And Len(Trim(Item_Type)) > 0 And Len(Trim(Item)) > 0 And Len(Trim(Qty)) > 0 And Len(Trim(Rates)) > 0 Then
        Me.cmdSave.Enabled = True
        Me.CmdDel.Enabled = True
    Else
        Me.cmdSave.Enabled = False
        Me.CmdDel.Enabled = False
    End If
End Sub
Private Sub SrItemType_Click()
    If Me.SrItemType.ListIndex > -1 Then
        i = Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = " & i, SrItem, "ItemName", "ItemCode"
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
    If dtChk.value = Checked Then
        'srdt = " And ReturnMasterDate = #" & Me.SrDate.value & " #"
        srdt = " And (ReturnMasterDate between Convert(datetime, '" & Me.SrDate.value - 1 & "')  and Convert(datetime, '" & Me.SrDate.value + 1 & "'))"
    Else
        srdt = ""
    End If
    
    If PtChk.value = Checked And Me.srParty.ListIndex > -1 Then
        cbo1 = " And ReturnMaster.partycode = " & Me.srParty.ItemData(Me.srParty.ListIndex)
    Else
        cbo1 = ""
    End If
    
    If ImTChk.value = Checked And Me.SrItemType.ListIndex > -1 Then
        cbo2 = " And ReturnDetail.ItemTypeCode = " & Me.SrItemType.ItemData(Me.SrItemType.ListIndex)
    Else
        cbo2 = ""
    End If
    
    If ImChk.value = Checked And Me.SrItem.ListIndex > -1 Then
        cbo3 = " And ReturnDetail.ItemCode = " & Me.SrItem.ItemData(Me.SrItem.ListIndex)
    Else
        cbo3 = ""
    End If
    
    sql = " SELECT top 60 ReturnDetailCode, ReturnMaster.ReturnMasterCode, PartyName, ReturnMaster.PartyCode, ReturnDetail.ItemTypeCode, ItemTypeName, ReturnDetail.ItemCode, ItemName, Quantity, Rate " & _
          " FROM (Party INNER JOIN (ReturnMaster INNER JOIN ReturnDetail ON ReturnMaster.ReturnMasterCode = ReturnDetail.ReturnMasterCode) ON Party.PartyCode = ReturnMaster.PartyCode) INNER JOIN (ItemType INNER JOIN Item ON ItemType.ItemTypeCode = Item.ItemTypeCode) ON (ReturnDetail.ItemCode = Item.ItemCode) AND (ReturnDetail.ItemTypeCode = ItemType.ItemTypeCode) " & _
          " Where ReturnDetail.IsActive = 1 " & _
          srdt & _
          cbo1 & _
          cbo2 & _
          cbo3 & _
          " Order by ReturnMaster.ReturnMasterCode desc, ReturnDetailCode desc"
                                
    Debug.Print sql
    Set rstList = FillRecordSet(sql)
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!ReturnDetailCode, _
                   Key:=CStr("Id=" & rstList!ReturnDetailCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!ReturnMasterCode
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!ItemTypeName
                 .Add Text:=rstList!ItemName
                 .Add Text:=rstList!Quantity
                 .Add Text:=rstList!Rate
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub SrParty_Click()
    Call SrfillList
End Sub
Private Sub SrItem_Click()
    Call SrfillList
End Sub
Private Sub SrDate_Change()
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
Private Sub dtChk_Click()
    If dtChk.value = Checked Then
        Me.SrDate.Enabled = True
    Else
        Me.SrDate.Enabled = False
    End If
    Call SrfillList
End Sub
