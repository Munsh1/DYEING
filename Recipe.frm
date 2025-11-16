VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Recipe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                  ----- Recipe -----"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6149.02
   ScaleMode       =   0  'User
   ScaleWidth      =   11460
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Recipe.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Recipe.frx":0268
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Recipe.frx":06C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Recipe.frx":0ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Recipe.frx":0F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Recipe.frx":1330
            Key             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   26
      Top             =   0
      Width           =   3135
      Begin VB.Frame Frame19 
         Height          =   735
         Left            =   120
         TabIndex        =   40
         Top             =   3960
         Width           =   2895
         Begin VB.CheckBox PCChk 
            Caption         =   "Recipe Code"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox srPC2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   42
            Top             =   360
            Width           =   1000
         End
         Begin VB.TextBox srPC1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   1000
         End
      End
      Begin VB.Frame Frame8 
         Height          =   2535
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   2895
         Begin VB.TextBox SrQty6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   39
            Top             =   2050
            Width           =   615
         End
         Begin VB.TextBox SrQty5 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   38
            Top             =   1700
            Width           =   615
         End
         Begin VB.TextBox SrQty4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   37
            Top             =   1350
            Width           =   615
         End
         Begin VB.ComboBox SrItem6 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Text            =   "SrItem6"
            Top             =   2050
            Width           =   2055
         End
         Begin VB.ComboBox SrItem5 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Text            =   "SrItem5"
            Top             =   1700
            Width           =   2055
         End
         Begin VB.ComboBox SrItem4 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   34
            Text            =   "SrItem4"
            Top             =   1350
            Width           =   2055
         End
         Begin VB.TextBox SrQty3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   21
            Top             =   1000
            Width           =   615
         End
         Begin VB.TextBox SrQty2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   19
            Top             =   650
            Width           =   615
         End
         Begin VB.TextBox SrQty1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   17
            Top             =   280
            Width           =   615
         End
         Begin VB.ComboBox SrItem3 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   20
            Text            =   "SrItem3"
            Top             =   1000
            Width           =   2055
         End
         Begin VB.ComboBox SrItem2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Text            =   "SrItem2"
            Top             =   650
            Width           =   2055
         End
         Begin VB.CheckBox ImChk 
            Caption         =   "Item"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox SrItem1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Text            =   "SrItem1"
            Top             =   280
            Width           =   2055
         End
      End
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
            TabIndex        =   13
            Top             =   0
            Width           =   735
         End
         Begin MSComCtl2.DTPicker SrDate 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   280
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   44433409
            CurrentDate     =   38224
         End
      End
      Begin LVbuttons.LaVolpeButton Cmdhide 
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   4920
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
         MICON           =   "Recipe.frx":176C
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
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   7935
      Begin MSComctlLib.ListView lvwphase 
         Height          =   2505
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
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
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detail Block"
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   7935
      Begin VB.TextBox Qty 
         Height          =   315
         Left            =   6360
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Item 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Text            =   "Item"
         Top             =   600
         Width           =   2895
      End
      Begin VB.ComboBox Item_Type 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Text            =   "Item_Type"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   6720
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Item"
         Height          =   255
         Left            =   3960
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Item Type"
         Height          =   255
         Left            =   840
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Master Block"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.TextBox Rem 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   5895
      End
      Begin MSComCtl2.DTPicker RecipeDate 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   44433409
         CurrentDate     =   38224
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   255
         Left            =   3480
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   360
         Width           =   495
      End
   End
   Begin LVbuttons.LaVolpeButton CmdAllSearch 
      Height          =   405
      Left            =   5655
      TabIndex        =   10
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
      MICON           =   "Recipe.frx":1788
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
      TabIndex        =   11
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
      MICON           =   "Recipe.frx":17A4
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
      TabIndex        =   9
      Top             =   5400
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
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Recipe.frx":17C0
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
      TabIndex        =   8
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
      MICON           =   "Recipe.frx":17DC
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
      TabIndex        =   6
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
      MICON           =   "Recipe.frx":17F8
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
      TabIndex        =   7
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
      MICON           =   "Recipe.frx":1814
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
      TabIndex        =   25
      Top             =   5520
      Width           =   1215
   End
End
Attribute VB_Name = "Recipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ListID As String
Dim d_ListID As String
Private Sub CmdAllSearch_Click()
        Recipe.Left = 200
        Recipe.Top = 500
        Recipe.Width = 11500
        
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
            cnDatabase.Execute "update RecipeDetail set IsActive = 0 where RecipeDetailCode =" & d_ListID
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
        Recipe.Width = 8250
        Recipe.Left = 2000
        Recipe.Top = 500
        Me.SrItem1.ListIndex = -1
        Me.SrItem2.ListIndex = -1
        Me.SrItem3.ListIndex = -1
        Call fillList
End Sub
Private Sub CmdNew_Click()
    Call addNewMaster
    Me.Rem.SetFocus
End Sub
Private Sub cmdSave_Click()
If Len(Trim(Item_Type)) > 0 And Len(Trim(Item)) > 0 And Len(Trim(Qty)) > 0 Then
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
    RecipeDate = Date
    SrDate = Date
    FillCombo "Select ItemTypeCode, ItemTypeName from ItemType where IsActive = 1 and ItemTypeName like " & "'%Color%'" & "  order by 2", Item_Type, "ItemTypeName", "ItemTypeCode"
    
    lvwphase.ColumnHeaders.Add Text:="Detail Code", Width:=0
    lvwphase.ColumnHeaders.Add Text:="Recipe Code", Width:=1500
    lvwphase.ColumnHeaders.Add Text:="Item Name", Width:=5000
    lvwphase.ColumnHeaders.Add Text:="Quantity", Width:=1150
    
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
    If (Me.Item_Type.ListIndex > -1) Then
        i = Me.Item_Type.ItemData(Me.Item_Type.ListIndex)
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = " & i & " order by 2", Item, "ItemName", "ItemCode"
    Else
        Me.Item.Clear
    End If
End Sub
Public Sub addNewDetail()
    d_ListID = ""
    Me.Item_Type.ListIndex = -1
    Me.Item.ListIndex = -1
    Me.Qty.Text = ""
    lblCaption.Caption = "Add Detail"
End Sub
Private Sub getValMaster()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("Select RecipeMasterCode, RecipeMasterDate, Remarks From RecipeMaster Where RecipeMasterCode = " & m_ListID)
    If Not (rstGetVal.EOF) Then
        Me.RecipeDate.value = IIf(IsNull(rstGetVal("RecipeMasterDate")), Now, rstGetVal("RecipeMasterDate"))
        Me.Rem.Text = rstGetVal("Remarks")
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
End Sub
Private Sub getValDetail()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("Select ItemTypeCode, ItemCode, Quantity From RecipeDetail Where RecipeDetailCode = " & d_ListID)
    If Not (rstGetVal.EOF) Then
    Debug.Print rstGetVal("ItemTypeCode")
        Call selectValueInCombo(Me.Item_Type, rstGetVal("ItemTypeCode"))
        Call selectValueInCombo(Me.Item, rstGetVal("ItemCode"))
        Me.Qty.Text = rstGetVal("Quantity")
   End If
   rstGetVal.Close
   Set rstGetVal = Nothing
End Sub
Public Sub setValMaster()
    Dim rstSave As New ADODB.Recordset
    If (Len(Trim(m_ListID)) = 0) Then
        Set rstSave = FillRecordSet("select * from RecipeMaster Where 1 = 2")
        rstSave.AddNew
        m_ListID = ValAutoNumber("RecipeMaster", "RecipeMasterCode")
        rstSave("RecipeMasterCode") = m_ListID
    Else
       Set rstSave = FillRecordSet("select * from RecipeMaster where RecipeMasterCode =" & m_ListID)
    End If
    
    rstSave("RecipeMasterDate") = Me.RecipeDate.value
    rstSave("Remarks") = IIf(IsNull(Me.Rem.Text), " ", Me.Rem.Text)
    rstSave("IsActive") = "1"
    
    rstSave.Update
    rstSave.Close
    Set rstSave = Nothing
End Sub
Public Sub setValDetail()
    Dim rstSave As New ADODB.Recordset
    If (Len(Trim(d_ListID)) = 0) Then
        Set rstSave = FillRecordSet("select * from RecipeDetail Where 1 = 2")
        rstSave.AddNew
        d_ListID = ValAutoNumber("RecipeDetail", "RecipeDetailCode")
        rstSave("RecipeDetailCode") = d_ListID
    Else
       Set rstSave = FillRecordSet("select * from RecipeDetail where RecipeDetailCode =" & d_ListID)
    End If
    
    rstSave("RecipeMasterCode") = m_ListID
    rstSave("ItemTypeCode") = Item_Type.ItemData(Item_Type.ListIndex)
    rstSave("ItemCode") = Me.Item.ItemData(Item.ListIndex)
    rstSave("Quantity") = Me.Qty.Text
    rstSave("IsActive") = "1"
    
    rstSave.Update
    rstSave.Close
    Set rstSave = Nothing
    Call addNewDetail
End Sub
Public Sub addNewMaster()
    m_ListID = ""
    Me.RecipeDate.value = Now
    Me.Rem.Text = ""
    
    d_ListID = ""
    Me.Item_Type.ListIndex = -1
    Me.Item.ListIndex = -1
    Me.Qty.Text = ""
    
    lblCaption.Caption = "Add Master"
End Sub
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Set rstList = FillRecordSet("SELECT top 60 RecipeDetail.RecipeDetailCode, RecipeDetail.ItemTypeCode, RecipeDetail.ItemCode, RecipeDetail.Quantity, RecipeMaster.RecipeMasterCode, RecipeMaster.Remarks, Item.ItemName, ItemType.ItemTypeName " & _
                                "FROM ItemType INNER JOIN (Item INNER JOIN (RecipeMaster INNER JOIN RecipeDetail ON RecipeMaster.RecipeMasterCode = RecipeDetail.RecipeMasterCode) ON Item.ItemCode = RecipeDetail.ItemCode) ON (ItemType.ItemTypeCode = RecipeDetail.ItemTypeCode) AND (ItemType.ItemTypeCode = Item.ItemTypeCode) " & _
                                "Where RecipeDetail.IsActive = 1 order by RecipeMaster.RecipeMasterCode desc, RecipeDetailCode desc")
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!RecipeDetailCode, _
                   Key:=CStr("Id=" & rstList!RecipeDetailCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!RecipeMasterCode
                 .Add Text:=rstList!ItemName
                 .Add Text:=rstList!Quantity
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
        m_AddMode = False
        d_ListID = Me.lvwphase.SelectedItem.Text
        m_ListID = Me.lvwphase.ListItems.Item(Me.lvwphase.SelectedItem.Index).ListSubItems(1).Text
        
        Call getValMaster
        Call getValDetail
    End If
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

Private Sub SrPC1_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub SrPC2_keyup(KeyCode As Integer, Shift As Integer)
    Call SrfillList
End Sub
Private Sub Qty_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 And Me.cmdSave.Enabled = True Then
         cmdSave.SetFocus
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub EnableSave()
    If Len(Trim(Item_Type)) > 0 And Len(Trim(Item)) > 0 And Len(Trim(Qty)) > 0 Then
        Me.cmdSave.Enabled = True
    Else
        Me.cmdSave.Enabled = False
        Me.CmdDel.Enabled = False
    End If
End Sub
Private Sub Rem_KeyPress(KeyAscii As Integer)
    Call EnableSave
    If KeyAscii = 13 Then
         Item_Type.SetFocus
    End If
End Sub
Private Sub SrfillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Dim sql As String
    Dim cbo3 As String
    Dim cbo4 As String
    If dtChk.value = Checked Then
        'srdt = " And RecipeMasterDate = #" & Format(Me.SrDate.value, "mm/dd/yyyy") & " #"
        srdt = " And (RecipeMasterDate between Convert(datetime, '" & Me.SrDate.value - 1 & "')  and Convert(datetime, '" & Me.SrDate.value + 1 & "'))"
    Else
        srdt = ""
    End If
    
    If ImChk.value = Checked And Me.SrItem1.ListIndex > -1 And Me.SrItem2.ListIndex > -1 And Me.SrItem3.ListIndex > -1 And Me.SrItem4.ListIndex > -1 And Me.SrItem5.ListIndex > -1 And Me.SrItem6.ListIndex > -1 And Len(Trim(Me.SrQty1.Text)) > 0 And Len(Trim(Me.SrQty2.Text)) > 0 And Len(Trim(Me.SrQty3.Text)) > 0 And Len(Trim(Me.SrQty4.Text)) > 0 And Len(Trim(Me.SrQty5.Text)) > 0 And Len(Trim(Me.SrQty6.Text)) > 0 Then
        cbo3 = " And RecipeMaster.RecipeMasterCode in (select RecipeMasterCode from RecipeDetail where ((RecipeDetail.ItemCode = " & Me.SrItem1.ItemData(Me.SrItem1.ListIndex) & " and quantity = " & Me.SrQty1.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem2.ItemData(Me.SrItem2.ListIndex) & " and quantity = " & Me.SrQty2.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem3.ItemData(Me.SrItem3.ListIndex) & " and quantity = " & Me.SrQty3.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem4.ItemData(Me.SrItem4.ListIndex) & " and quantity = " & Me.SrQty4.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem5.ItemData(Me.SrItem5.ListIndex) & " and quantity = " & Me.SrQty5.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem6.ItemData(Me.SrItem6.ListIndex) & " and quantity = " & Me.SrQty6.Text & ")) group by RecipeMasterCode having count(RecipeMasterCode) > 5 )"
    ElseIf ImChk.value = Checked And Me.SrItem1.ListIndex > -1 And Me.SrItem2.ListIndex > -1 And Me.SrItem3.ListIndex > -1 And Me.SrItem4.ListIndex > -1 And Me.SrItem5.ListIndex > -1 And Len(Trim(Me.SrQty1.Text)) > 0 And Len(Trim(Me.SrQty2.Text)) > 0 And Len(Trim(Me.SrQty3.Text)) > 0 And Len(Trim(Me.SrQty4.Text)) > 0 And Len(Trim(Me.SrQty5.Text)) > 0 Then
        cbo3 = " And RecipeMaster.RecipeMasterCode in (select RecipeMasterCode from RecipeDetail where ((RecipeDetail.ItemCode = " & Me.SrItem1.ItemData(Me.SrItem1.ListIndex) & " and quantity = " & Me.SrQty1.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem2.ItemData(Me.SrItem2.ListIndex) & " and quantity = " & Me.SrQty2.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem3.ItemData(Me.SrItem3.ListIndex) & " and quantity = " & Me.SrQty3.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem4.ItemData(Me.SrItem4.ListIndex) & " and quantity = " & Me.SrQty4.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem5.ItemData(Me.SrItem5.ListIndex) & " and quantity = " & Me.SrQty5.Text & ")) group by RecipeMasterCode having count(RecipeMasterCode) > 4 )"
    ElseIf ImChk.value = Checked And Me.SrItem1.ListIndex > -1 And Me.SrItem2.ListIndex > -1 And Me.SrItem3.ListIndex > -1 And Me.SrItem4.ListIndex > -1 And Len(Trim(Me.SrQty1.Text)) > 0 And Len(Trim(Me.SrQty2.Text)) > 0 And Len(Trim(Me.SrQty3.Text)) > 0 And Len(Trim(Me.SrQty4.Text)) > 0 Then
        cbo3 = " And RecipeMaster.RecipeMasterCode in (select RecipeMasterCode from RecipeDetail where ((RecipeDetail.ItemCode = " & Me.SrItem1.ItemData(Me.SrItem1.ListIndex) & " and quantity = " & Me.SrQty1.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem2.ItemData(Me.SrItem2.ListIndex) & " and quantity = " & Me.SrQty2.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem3.ItemData(Me.SrItem3.ListIndex) & " and quantity = " & Me.SrQty3.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem4.ItemData(Me.SrItem4.ListIndex) & " and quantity = " & Me.SrQty4.Text & ")) group by RecipeMasterCode having count(RecipeMasterCode) > 3 )"
    ElseIf ImChk.value = Checked And Me.SrItem1.ListIndex > -1 And Me.SrItem2.ListIndex > -1 And Me.SrItem3.ListIndex > -1 And Len(Trim(Me.SrQty1.Text)) > 0 And Len(Trim(Me.SrQty2.Text)) > 0 And Len(Trim(Me.SrQty3.Text)) > 0 Then
        cbo3 = " And RecipeMaster.RecipeMasterCode in (select RecipeMasterCode from RecipeDetail where ((RecipeDetail.ItemCode = " & Me.SrItem1.ItemData(Me.SrItem1.ListIndex) & " and quantity = " & Me.SrQty1.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem2.ItemData(Me.SrItem2.ListIndex) & " and quantity = " & Me.SrQty2.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem3.ItemData(Me.SrItem3.ListIndex) & " and quantity = " & Me.SrQty3.Text & ")) group by RecipeMasterCode having count(RecipeMasterCode) > 2 )"
    ElseIf ImChk.value = Checked And Me.SrItem1.ListIndex > -1 And Me.SrItem2.ListIndex > -1 And Len(Trim(Me.SrQty1.Text)) > 0 And Len(Trim(Me.SrQty2.Text)) > 0 Then
        cbo3 = " And RecipeMaster.RecipeMasterCode in (select RecipeMasterCode from RecipeDetail where ((RecipeDetail.ItemCode = " & Me.SrItem1.ItemData(Me.SrItem1.ListIndex) & " and quantity = " & Me.SrQty1.Text & ")"
        cbo3 = cbo3 & " OR (RecipeDetail.ItemCode = " & Me.SrItem2.ItemData(Me.SrItem2.ListIndex) & " and quantity = " & Me.SrQty2.Text & ")) group by RecipeMasterCode having count(RecipeMasterCode) > 1 )"
    ElseIf ImChk.value = Checked And Me.SrItem1.ListIndex > -1 And Len(Trim(Me.SrQty1.Text)) > 0 Then
        cbo3 = " And RecipeMaster.RecipeMasterCode in (select RecipeMasterCode from RecipeDetail where (RecipeDetail.ItemCode = " & Me.SrItem1.ItemData(Me.SrItem1.ListIndex) & " and quantity = " & Me.SrQty1.Text & "))"
    Else
        cbo3 = ""
    End If
    
    If PCChk.value = Checked And Len(Trim(Me.srPC1)) > 0 And Len(Trim(Me.srPC2)) > 0 Then
        cbo4 = " And (RecipeMaster.RecipeMasterCode between " & Me.srPC1 & " and " & Me.srPC2 & " )"
    Else
        cbo4 = ""
    End If
   
    sql = " SELECT top 100 RecipeDetail.RecipeDetailCode, RecipeDetail.ItemTypeCode, RecipeDetail.ItemCode, RecipeDetail.Quantity, RecipeMaster.RecipeMasterCode, RecipeMaster.Remarks, Item.ItemName, ItemType.ItemTypeName " & _
          " FROM ItemType INNER JOIN (Item INNER JOIN (RecipeMaster INNER JOIN RecipeDetail ON RecipeMaster.RecipeMasterCode = RecipeDetail.RecipeMasterCode) ON Item.ItemCode = RecipeDetail.ItemCode) ON (ItemType.ItemTypeCode = RecipeDetail.ItemTypeCode) AND (ItemType.ItemTypeCode = Item.ItemTypeCode) " & _
          " Where RecipeDetail.IsActive = 1 " & _
          srdt & _
          cbo2 & _
          cbo3 & _
          cbo4 & _
          " Order by RecipeMaster.RecipeMasterCode desc, RecipeDetailCode desc"
                                
    Debug.Print sql
    Set rstList = FillRecordSet(sql)
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!RecipeDetailCode, _
                   Key:=CStr("Id=" & rstList!RecipeDetailCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!RecipeMasterCode
                 .Add Text:=rstList!ItemName
                 .Add Text:=rstList!Quantity
            End With
        rstList.MoveNext
      Loop
    End If
    rstList.Close
    Set rstList = Nothing
End Sub
Private Sub SrItem_Click()
    Call SrfillList
End Sub
Private Sub SrDate_Change()
    Call SrfillList
End Sub
Private Sub ImChk_Click()
  
    If ImChk.value = Checked Then
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = (Select ItemTypeCode from ItemType where IsActive = 1 and ItemTypeName like " & "'%Color%') order by 2", SrItem1, "ItemName", "ItemCode"
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = (Select ItemTypeCode from ItemType where IsActive = 1 and ItemTypeName like " & "'%Color%') order by 2", SrItem2, "ItemName", "ItemCode"
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = (Select ItemTypeCode from ItemType where IsActive = 1 and ItemTypeName like " & "'%Color%') order by 2", SrItem3, "ItemName", "ItemCode"
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = (Select ItemTypeCode from ItemType where IsActive = 1 and ItemTypeName like " & "'%Color%') order by 2", SrItem4, "ItemName", "ItemCode"
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = (Select ItemTypeCode from ItemType where IsActive = 1 and ItemTypeName like " & "'%Color%') order by 2", SrItem5, "ItemName", "ItemCode"
        FillCombo "Select ItemCode, ItemName from Item where IsActive = 1 and ItemTypeCode = (Select ItemTypeCode from ItemType where IsActive = 1 and ItemTypeName like " & "'%Color%') order by 2", SrItem6, "ItemName", "ItemCode"
                
        Me.SrItem1.Enabled = True
        Me.SrItem2.Enabled = True
        Me.SrItem3.Enabled = True
        Me.SrItem4.Enabled = True
        Me.SrItem5.Enabled = True
        Me.SrItem6.Enabled = True

        Me.SrQty1.Enabled = True
        Me.SrQty2.Enabled = True
        Me.SrQty3.Enabled = True
        Me.SrQty4.Enabled = True
        Me.SrQty5.Enabled = True
        Me.SrQty6.Enabled = True
        
    Else
        Me.SrItem1.Clear
        Me.SrItem2.Clear
        Me.SrItem3.Clear
        Me.SrItem4.Clear
        Me.SrItem5.Clear
        Me.SrItem6.Clear
        
        Me.SrQty1.Text = ""
        Me.SrQty2.Text = ""
        Me.SrQty3.Text = ""
        Me.SrQty4.Text = ""
        Me.SrQty5.Text = ""
        Me.SrQty6.Text = ""
        
        Me.SrItem1.Enabled = False
        Me.SrItem2.Enabled = False
        Me.SrItem3.Enabled = False
        Me.SrItem4.Enabled = False
        Me.SrItem5.Enabled = False
        Me.SrItem6.Enabled = False
        
        Me.SrQty1.Enabled = False
        Me.SrQty2.Enabled = False
        Me.SrQty3.Enabled = False
        Me.SrQty4.Enabled = False
        Me.SrQty5.Enabled = False
        Me.SrQty6.Enabled = False
    
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

Private Sub SrItem1_Change()
    Call SrfillList
End Sub

Private Sub SrItem2_Change()
    Call SrfillList
End Sub

Private Sub SrItem3_Change()
    Call SrfillList
End Sub
Private Sub SrItem4_Change()
    Call SrfillList
End Sub
Private Sub SrItem5_Change()
    Call SrfillList
End Sub
Private Sub SrItem6_Change()
    Call SrfillList
End Sub

Private Sub SrQty1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         SrItem2.SetFocus
         Call SrfillList
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub SrQty2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         SrItem3.SetFocus
         Call SrfillList
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub SrQty3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         SrItem4.SetFocus
         Call SrfillList
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub SrQty4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         SrItem5.SetFocus
         Call SrfillList
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub SrQty5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         SrItem6.SetFocus
         Call SrfillList
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub SrQty6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Call SrfillList
    End If
    If KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub SrQty1_LostFocus()
    Call SrfillList
End Sub
Private Sub SrQty2_LostFocus()
    Call SrfillList
End Sub
Private Sub SrQty3_LostFocus()
    Call SrfillList
End Sub
Private Sub SrQty4_LostFocus()
    Call SrfillList
End Sub
Private Sub SrQty5_LostFocus()
    Call SrfillList
End Sub
Private Sub SrQty6_LostFocus()
    Call SrfillList
End Sub
