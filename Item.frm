VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                              ----- Item Information -----"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6900
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Item.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Item.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Item.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Item.frx":0AFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   405
      Left            =   960
      TabIndex        =   3
      Top             =   4245
      Width           =   1200
      _ExtentX        =   2117
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
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   14737632
      EMBOSSS         =   14737632
      MPTR            =   0
      MICON           =   "Item.frx":0F38
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
   Begin LVbuttons.LaVolpeButton CmdNew 
      Height          =   405
      Left            =   2400
      TabIndex        =   4
      Top             =   4245
      Width           =   1200
      _ExtentX        =   2117
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
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   14737632
      EMBOSSS         =   14737632
      MPTR            =   0
      MICON           =   "Item.frx":0F54
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
      Left            =   3840
      TabIndex        =   5
      Top             =   4245
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Delete"
      ENAB            =   0   'False
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
      EMBOSSM         =   14737632
      EMBOSSS         =   14737632
      MPTR            =   0
      MICON           =   "Item.frx":0F70
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
   Begin LVbuttons.LaVolpeButton cmdClose 
      Height          =   405
      Left            =   5280
      TabIndex        =   6
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
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   14737632
      EMBOSSS         =   14737632
      MPTR            =   0
      MICON           =   "Item.frx":0F8C
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
   Begin VB.Frame Item_Frame 
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6585
      Begin VB.TextBox Minimum_Quantity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox Item_Name 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   250
         Width           =   4935
      End
      Begin VB.ComboBox Item_Type 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   575
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Min Qty."
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Lb_Item_Name 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Lb_Item_Type 
         Caption         =   "Item Type"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   625
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2805
      Left            =   120
      TabIndex        =   8
      Top             =   1380
      Width           =   6615
      Begin MSComctlLib.ListView lvwphase 
         Height          =   2500
         Left            =   75
         TabIndex        =   7
         Top             =   180
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   4419
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
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ListID As Long
Dim m_AddMode As Boolean
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Set rstList = FillRecordSet("SELECT Item.ItemCode, Item.ItemName, Item.MinimumQuantity, Item.IsActive, ItemType.ItemTypeName FROM ItemType INNER JOIN Item ON ItemType.ItemTypeCode = Item.ItemTypeCode where Item.IsActive = 1 order by Item.ItemCode desc")
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
Public Sub setVal()
    Dim rstSave As New ADODB.Recordset
        If m_AddMode = True Then
            Set rstSave = FillRecordSet("select * from Item where 1 = 2")
            rstSave.AddNew
            rstSave("ItemCode") = ValAutoNumber("Item", "ItemCode")
        Else
           Set rstSave = FillRecordSet("select * from Item where ItemCode =" & m_ListID)
        End If
    rstSave("ItemName") = Me.Item_Name.Text
    rstSave("ItemTypeCode") = Me.Item_Type.ItemData(Item_Type.ListIndex)
    rstSave("MinimumQuantity") = Me.Minimum_Quantity.Text
    rstSave("IsActive") = 1
    rstSave.Update
    rstSave.Close
    Set rstSave = Nothing
    m_AddMode = False
End Sub
Private Sub getVal()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("select * from Item where ItemCode =" & m_ListID)
    If Not (rstGetVal.EOF) Then
        Call selectValueInCombo(Me.Item_Type, rstGetVal("ItemTypeCode"))
        Me.Item_Name.Text = rstGetVal("ItemName")
        Me.Minimum_Quantity.Text = rstGetVal("MinimumQuantity")
   End If
    rstGetVal.Close
    Set rstGetVal = Nothing
End Sub
Private Sub cmdClose_Click()
    Unload Item
End Sub
Private Sub CmdDel_Click()
    Dim strAns As String
    strAns = MsgBox("Do you want to delete this record...?", vbYesNo + vbInformation)
    If strAns = vbYes Then
        cnDatabase.Execute "update Item set IsActive=0  where ItemCode=" & m_ListID
        Call fillList
        MsgBox ("Record deleted succesfully..."), vbInformation
        Me.CmdDel.Enabled = False
        Me.cmdSave.Enabled = False
        Item_Name.SetFocus
    End If
    m_ListID = 0
    Call ClearField
    m_AddMode = True
End Sub

Private Sub CmdNew_Click()
    Call ClearField
    Me.Item_Name.SetFocus
    m_AddMode = True
    cmdSave.Enabled = False
    CmdDel.Enabled = False
End Sub
Private Sub cmdSave_Click()
    If Len(Trim(Item_Type)) > 0 And Len(Trim(Item_Name)) > 0 And Len(Trim(Minimum_Quantity)) > 0 Then
        Dim rstSearch As New ADODB.Recordset
        Call setVal
        Call fillList
        cmdSave.Enabled = False
        CmdDel.Enabled = False
        Me.Item_Name.Text = ""
        Me.Minimum_Quantity.Text = ""
        Me.Item_Name.SetFocus
        MsgBox ("Record saved successfully"), vbInformation
        m_AddMode = True
    Else
        MsgBox "Provide data in all Fields"
    End If
End Sub
Private Sub Form_Load()
    m_AddMode = True
    CmdDel.Enabled = False
    cmdSave.Enabled = False
    DBConn
    FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 order by 2", Item_Type, "ItemTypeName", "ItemTypeCode"
    lvwphase.ColumnHeaders.Add Text:="Item Code", Width:=1000
    lvwphase.ColumnHeaders.Add Text:="Item Name", Width:=2900
    lvwphase.ColumnHeaders.Add Text:="Min Quantity", Width:=1100
    lvwphase.ColumnHeaders.Add Text:="Type", Width:=1350
    Call fillList
End Sub
Private Sub Item_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Item_Type.SetFocus
    End If
    Call validateItem
End Sub
Private Sub Item_Name_KeyUp(KeyCode As Integer, Shift As Integer)
    Call validateItem
End Sub
Private Sub Item_Type_Change()
 Call validateItem
End Sub

Private Sub Item_Type_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Minimum_Quantity.SetFocus
    End If
    Call validateItem
End Sub

Private Sub Item_Type_KeyUp(KeyCode As Integer, Shift As Integer)
    Call validateItem
End Sub

Private Sub lvwphase_Click()
  cmdSave.Enabled = True
'  CmdDel.Enabled = True
  m_AddMode = False
  m_ListID = Me.lvwphase.SelectedItem.ListSubItems(4).Text
  Call getVal
End Sub
Private Sub lvwphase_ItemClick(ByVal Item As MSComctlLib.ListItem)
    m_ListID = Mid(Item.Key, 4, Len(Item.Key))
End Sub
Private Sub lvwphase_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.Enabled = True
 '   CmdDel.Enabled = True
    m_AddMode = False
    m_ListID = Me.lvwphase.SelectedItem.ListSubItems(4).Text
    Call getVal
End If
End Sub
Private Sub ClearField()
    Me.Item_Name.Text = ""
    Me.Minimum_Quantity.Text = ""
    Me.Item_Type.ListIndex = 0
End Sub
Private Sub validateItem()
    If (Me.Item_Name.Text <> "" And Me.Item_Type.ListIndex > -1) Then
        Me.cmdSave.Enabled = True
'        Me.CmdDel.Enabled = True
    Else
        Me.cmdSave.Enabled = False
        Me.CmdDel.Enabled = False
    End If
End Sub

Private Sub Minimum_Quantity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdSave.SetFocus
    End If
    Call validateItem
End Sub
