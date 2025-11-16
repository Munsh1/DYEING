VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form Party 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                      ----- Party Information -----"
   ClientHeight    =   4830
   ClientLeft      =   1275
   ClientTop       =   1665
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7260
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   960
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
            Picture         =   "Party.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Party.frx":0458
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Party.frx":06C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Party.frx":0AF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdClose 
      Height          =   405
      Left            =   5880
      TabIndex        =   8
      Top             =   4350
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
      MICON           =   "Party.frx":0F30
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
   Begin LVbuttons.LaVolpeButton CmdDel 
      Height          =   405
      Left            =   4440
      TabIndex        =   7
      Top             =   4350
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
      MICON           =   "Party.frx":0F4C
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
   Begin LVbuttons.LaVolpeButton CmdNew 
      Height          =   405
      Left            =   3000
      TabIndex        =   6
      Top             =   4350
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
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Party.frx":0F68
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
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   405
      Left            =   1560
      TabIndex        =   5
      Top             =   4350
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
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Party.frx":0F84
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
   Begin VB.Frame Frame1 
      Height          =   2325
      Left            =   120
      TabIndex        =   16
      Top             =   1965
      Width           =   6975
      Begin MSComctlLib.ListView lvwphase 
         Height          =   2010
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   3545
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
   Begin VB.Frame Detail 
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6975
      Begin VB.TextBox Party_Cell 
         Height          =   285
         Left            =   1515
         TabIndex        =   4
         Top             =   1520
         Width           =   4500
      End
      Begin VB.CheckBox Is_Active 
         Caption         =   "Either Active or not?"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Party_Phone 
         Height          =   285
         Left            =   1515
         TabIndex        =   3
         Top             =   1200
         Width           =   4500
      End
      Begin VB.TextBox Party_Address 
         Height          =   285
         Left            =   1515
         TabIndex        =   2
         Top             =   900
         Width           =   4500
      End
      Begin VB.ComboBox Party_Type 
         Height          =   315
         Left            =   1515
         TabIndex        =   1
         Text            =   "Party_Type"
         Top             =   555
         Width           =   4500
      End
      Begin VB.TextBox Party_Name 
         Height          =   285
         Left            =   1515
         TabIndex        =   0
         Top             =   240
         Width           =   4500
      End
      Begin VB.Label Label1 
         Caption         =   "Party Cell"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Lb_Party_Phone 
         Caption         =   "Party Phone"
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   1260
         Width           =   1200
      End
      Begin VB.Label Lb_Party_Address 
         Caption         =   "Party Address"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   915
         Width           =   1200
      End
      Begin VB.Label Lb_Party_Type 
         Caption         =   "Party Type"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label Lb_Party_Name 
         Caption         =   "Party Name"
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   315
         Width           =   1200
      End
   End
End
Attribute VB_Name = "Party"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ListID As Long
Dim m_AddMode As Boolean
Dim CMDSearch As Boolean
Private Sub fillList()
    Dim lstItem As ListItem
    Dim rstList  As New ADODB.Recordset
    Set rstList = FillRecordSet("SELECT Party.PartyCode, Party.PartyName, Party.PartyAddress, Party.PartyPhone, Party.PartyCell, PartyType.PartyTypeName " & _
            " FROM PartyType INNER JOIN Party ON PartyType.PartyTypeCode = Party.PartyTypeCode where Party.IsActive = 1 order by Party.PartyCode desc")
    lvwphase.ListItems.Clear
    If Not rstList.EOF Then
      Do While Not rstList.EOF
            Set lstItem = lvwphase.ListItems.Add( _
                   Text:=rstList!PartyCode, _
                   Key:=CStr("Id=" & rstList!PartyCode))
            With lstItem.ListSubItems
                 .Add Text:=rstList!PartyName
                 .Add Text:=rstList!PartyTypeName
                 .Add Text:=rstList!PartyAddress
                 .Add Text:=rstList!PartyPhone
                 .Add Text:=rstList!PartyCell
                 .Add Text:=rstList!PartyCode
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
            Set rstSave = FillRecordSet("select * from Party where 1 = 2")
            rstSave.AddNew
            rstSave("PartyCode") = ValAutoNumber("Party", "PartyCode")
        Else
           Set rstSave = FillRecordSet("select * from Party where PartyCode =" & m_ListID)
        End If
        rstSave("PartyName") = Me.Party_Name.Text
        rstSave("PartyAddress") = Me.Party_Address.Text
        rstSave("PartyTypeCode") = Me.Party_Type.ItemData(Party_Type.ListIndex)
        rstSave("PartyPhone") = Me.Party_Phone.Text
        rstSave("PartyCell") = Me.Party_Cell.Text
        rstSave("IsActive") = 1
    rstSave.Update
    rstSave.Close
    Set rstSave = Nothing
    m_AddMode = False
    Call fillList
End Sub
Private Sub getVal()
    Dim rstGetVal As New ADODB.Recordset
    Set rstGetVal = FillRecordSet("select * from Party where PartyCode =" & m_ListID)
    If Not (rstGetVal.EOF) Then
        Call selectValueInCombo(Me.Party_Type, rstGetVal("PartyTypeCode"))
        Me.Party_Name.Text = rstGetVal("PartyName")
        Me.Party_Address.Text = rstGetVal("PartyAddress")
        Me.Party_Phone.Text = rstGetVal("PartyPhone")
        Me.Party_Cell.Text = rstGetVal("PartyCell")
    End If
    rstGetVal.Close
    Set rstGetVal = Nothing
End Sub
Private Sub cmdClose_Click()
    Unload Party
End Sub
Private Sub CmdDel_Click()
Dim strAns As String
    strAns = MsgBox("Do you want to delete this record...?", vbYesNo + vbInformation)
    If strAns = vbYes Then
        cnDatabase.Execute "update Party set IsActive = 0  where PartyCode=" & m_ListID
        Call fillList
        MsgBox ("Record deleted succesfully..."), vbInformation
        Me.CmdDel.Enabled = False
        Me.cmdSave.Enabled = False
        Party_Name.SetFocus
    End If
    m_ListID = 0
    Call ClearField
    m_AddMode = True
End Sub
Private Sub CmdNew_Click()
    Call ClearField
    Me.Party_Name.SetFocus
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
    Call ClearField
    MsgBox ("Record saved successfully"), vbInformation
    m_AddMode = True
End Sub
Private Sub Form_Load()
    m_AddMode = True
    CmdDel.Enabled = False
    cmdSave.Enabled = False
    DBConn
    FillCombo "Select PartyTypeCode, PartyTypeName from PartyType where IsActive = 1 order by 2", Party_Type, "PartyTypeName", "PartyTypeCode"
    lvwphase.ColumnHeaders.Add Text:="Party Code", Width:=1000
    lvwphase.ColumnHeaders.Add Text:="Party Name", Width:=1300
    lvwphase.ColumnHeaders.Add Text:="Type", Width:=1000
    lvwphase.ColumnHeaders.Add Text:="Address", Width:=1500
    lvwphase.ColumnHeaders.Add Text:="Phone", Width:=900
    lvwphase.ColumnHeaders.Add Text:="Cell", Width:=900
    Call fillList
End Sub
Private Sub lvwphase_Click()
    FillCombo "Select PartyTypeCode, PartyTypeName from PartyType", Party_Type, "PartyTypeName", "PartyTypeCode"
    cmdSave.Enabled = True
   ' CmdDel.Enabled = True
    m_AddMode = False
    m_ListID = Me.lvwphase.SelectedItem.ListSubItems(6).Text
    Call getVal
End Sub
Private Sub lvwphase_ItemClick(ByVal Item As MSComctlLib.ListItem)
    m_ListID = Mid(Item.Key, 4, Len(Item.Key))
End Sub
Private Sub lvwphase_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FillCombo "Select PartyTypeCode, PartyTypeName from PartyType", Party_Type, "PartyTypeName", "PartyTypeCode"
    cmdSave.Enabled = True
'    CmdDel.Enabled = True
    m_AddMode = False
    m_ListID = Me.lvwphase.SelectedItem.ListSubItems(6).Text
    Call getVal
End If
End Sub
Private Sub ClearField()
    Me.Party_Name.Text = ""
    Me.Party_Address.Text = ""
    Me.Party_Phone.Text = ""
    Me.Party_Cell.Text = ""
    Me.Party_Type.ListIndex = 0
End Sub
Private Sub Party_Address_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Party_Phone.SetFocus
    End If
    Call EnableSave
End Sub

Private Sub Party_Name_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Me.Party_Type.SetFocus
        End If
    Call EnableSave
End Sub
Private Sub Party_Phone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Party_Cell.SetFocus
    End If
End Sub
Private Sub Party_Cell_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdSave.SetFocus
    End If
End Sub

Private Sub Party_Type_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Party_Address.SetFocus
    End If
    Call EnableSave
End Sub
Private Sub EnableSave()
    If Len(Trim(Me.Party_Name.Text)) > 0 And Me.Party_Type.ListIndex >= 0 Then
        Me.cmdSave.Enabled = True
'        Me.CmdDel.Enabled = True
    Else
        Me.cmdSave.Enabled = False
        Me.CmdDel.Enabled = False
    End If
End Sub
