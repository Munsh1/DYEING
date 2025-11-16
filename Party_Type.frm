VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form PartyType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                   ------ Party Type ------"
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
      Left            =   6240
      Top             =   1080
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
            Picture         =   "Party_Type.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Party_Type.frx":043C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Party_Type.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Party_Type.frx":0AD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   405
      Left            =   960
      TabIndex        =   1
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
      MICON           =   "Party_Type.frx":0F30
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
      Left            =   2400
      TabIndex        =   2
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
      MICON           =   "Party_Type.frx":0F4C
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
      TabIndex        =   3
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
      MICON           =   "Party_Type.frx":0F68
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
      TabIndex        =   4
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
      MICON           =   "Party_Type.frx":0F84
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
   Begin VB.Frame Item_Frame 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6585
      Begin VB.TextBox Party_Type_Name 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   350
         Width           =   4935
      End
      Begin VB.Label Lb_Item_Name 
         Caption         =   "Party Type"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   350
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   6615
      Begin MSComctlLib.ListView lvwphase 
         Height          =   2985
         Left            =   75
         TabIndex        =   6
         Top             =   180
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   5265
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
Attribute VB_Name = "PartyType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ListID As Long
Dim m_AddMode As Boolean
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Set rstList = FillRecordSet("SELECT PartyTypeCode, PartyTypeName, IsActive FROM PartyType where IsActive = 1 order by PartyTypeCode desc")
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!PartyTypeCode, _
                   Key:=CStr("Id=" & rstList!PartyTypeCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!PartyTypeName
                 .Add Text:=rstList!PartyTypeCode
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
            Set rstSave = FillRecordSet("select * from PartyType where 1 = 2")
            rstSave.AddNew
            rstSave("PartyTypeCode") = ValAutoNumber("PartyType", "PartyTypeCode")
        Else
           Set rstSave = FillRecordSet("select * from PartyType where PartyTypeCode =" & m_ListID)
        End If
    rstSave("PartyTypeName") = Me.Party_Type_Name.Text
    rstSave.Update
    rstSave.Close
    Set rstSave = Nothing
    m_AddMode = False
End Sub
Private Sub getVal()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("select * from PartyType where PartyTypeCode =" & m_ListID)
    If Not (rstGetVal.EOF) Then
        Me.Party_Type_Name.Text = rstGetVal("PartyTypeName")
   End If
    rstGetVal.Close
    Set rstGetVal = Nothing
End Sub
Private Sub cmdClose_Click()
    Unload PartyType
End Sub
Private Sub CmdDel_Click()
    Dim strAns As String
    strAns = MsgBox("Do you want to delete this record...?", vbYesNo + vbInformation)
    If strAns = vbYes Then
        cnDatabase.Execute "update PartyType set IsActive = 0  where PartyTypeCode = " & m_ListID
        Call fillList
        MsgBox ("Record deleted succesfully..."), vbInformation
        Me.CmdDel.Enabled = False
        Me.cmdSave.Enabled = False
        Me.Party_Type_Name.SetFocus
    End If
    m_ListID = 0
    Call ClearField
    m_AddMode = True
End Sub

Private Sub CmdNew_Click()
    Call ClearField
    Me.Party_Type_Name.SetFocus
    m_AddMode = True
    cmdSave.Enabled = False
    CmdDel.Enabled = False
End Sub
Private Sub cmdSave_Click()
    Dim rstSearch As New ADODB.Recordset
    Call setVal
    Call fillList
    cmdSave.Enabled = False
    CmdDel.Enabled = False
    Me.Party_Type_Name.Text = ""
    Me.Party_Type_Name.SetFocus
    MsgBox ("Record saved successfully"), vbInformation
    m_AddMode = True
End Sub
Private Sub Form_Load()
    m_AddMode = True
    CmdDel.Enabled = False
    cmdSave.Enabled = False
    DBConn
    lvwphase.ColumnHeaders.Add Text:="Party Type Code", Width:=1500
    lvwphase.ColumnHeaders.Add Text:="Party Type Name", Width:=4830
    Call fillList
End Sub
Private Sub Party_Type_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdSave.SetFocus
    End If
    Call validateItem
End Sub
Private Sub Party_Type_Name_KeyUp(KeyCode As Integer, Shift As Integer)
    Call validateItem
End Sub

Private Sub lvwphase_Click()
  cmdSave.Enabled = True
'  CmdDel.Enabled = True
  m_AddMode = False
  m_ListID = Me.lvwphase.SelectedItem.ListSubItems(2).Text
  Call getVal
End Sub
Private Sub lvwphase_ItemClick(ByVal Item As MSComctlLib.ListItem)
    m_ListID = Mid(Item.Key, 4, Len(Item.Key))
End Sub
Private Sub lvwphase_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.Enabled = True
'    CmdDel.Enabled = True
    m_AddMode = False
    m_ListID = Me.lvwphase.SelectedItem.ListSubItems(2).Text
    Call getVal
End If
End Sub
Private Sub ClearField()
    Me.Party_Type_Name.Text = ""
End Sub
Private Sub validateItem()
    If (Me.Party_Type_Name.Text <> "") Then
        Me.cmdSave.Enabled = True
'        Me.CmdDel.Enabled = True
    Else
        Me.cmdSave.Enabled = False
        Me.CmdDel.Enabled = False
    End If
End Sub
