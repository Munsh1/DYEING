VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form Item_Search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                 ----- Item Information -----"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6900
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   960
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
            Picture         =   "Item_Search.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Item_Search.frx":0404
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdClose 
      Height          =   405
      Left            =   5520
      TabIndex        =   3
      Top             =   4245
      Width           =   1200
      _ExtentX        =   2117
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
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Item_Search.frx":0840
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
   Begin LVbuttons.LaVolpeButton CmdAllSearch 
      Height          =   405
      Left            =   3960
      TabIndex        =   2
      Top             =   4245
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
      MICON           =   "Item_Search.frx":085C
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
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   6645
      Begin VB.TextBox TXTSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   250
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Search"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3045
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   6615
      Begin MSComctlLib.ListView lvwphase 
         Height          =   2790
         Left            =   75
         TabIndex        =   4
         Top             =   180
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   4921
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
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "Item_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    If (Len(Trim(Me.TXTSearch.Text)) > 0) Then
         Set rstList = FillRecordSet("SELECT Item.ItemCode, Item.ItemName, Item.IsActive, Item.MinimumQuantity, ItemType.ItemTypeName FROM ItemType INNER JOIN Item ON ItemType.ItemTypeCode = Item.ItemTypeCode where Item.IsActive = 1 and ItemName + ItemType.ItemTypeName like '%" & Me.TXTSearch.Text & "%' order by Item.ItemCode")
    Else
         Set rstList = FillRecordSet("SELECT Item.ItemCode, Item.ItemName, Item.IsActive, Item.MinimumQuantity, ItemType.ItemTypeName FROM ItemType INNER JOIN Item ON ItemType.ItemTypeCode = Item.ItemTypeCode where Item.IsActive = 1 order by Item.ItemCode desc")
    End If
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!ItemCode, _
                   Key:=CStr("Id=" & rstList!ItemCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!ItemName
                 .Add Text:=rstList!MinimumQuantity
                 .Add Text:=rstList!ItemTypeName
                 .Add Text:=rstList!ItemCode
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub CmdAllSearch_Click()
    TXTSearch = ""
    Call fillList
End Sub
Private Sub cmdClose_Click()
    Unload Item_Search
End Sub
Private Sub Form_Load()
    DBConn
    lvwphase.ColumnHeaders.Add Text:="Item Code", Width:=1000
    lvwphase.ColumnHeaders.Add Text:="Item Name", Width:=2900
    lvwphase.ColumnHeaders.Add Text:="Min Quantity", Width:=1100
    lvwphase.ColumnHeaders.Add Text:="Type", Width:=1350
    Call fillList
End Sub


Private Sub txtsearch_KeyUp(KeyCode As Integer, Shift As Integer)
    Call fillList
End Sub
